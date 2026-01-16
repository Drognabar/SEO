"""
Скрипт проверки URL из sitemap сайта.

Что делает:
- Берёт корневой sitemap (sitemap.xml) по переданному URL.
- Рекурсивно обходит все вложенные sitemap-файлы (sitemapindex -> sitemap -> urlset).
- Собирает все уникальные URL из карт сайта.
- Проверяет HTTP-статусы этих URL (HEAD-запросом с учётом редиректов и ретраями).
- Формирует два CSV:
  1) со всеми URL и их статусами;
  2) только с проблемными URL:
     - редиректы 3xx,
     - ошибки 4xx/5xx/сетевые (status=0),
     - 18+ URL (по паттернам cid_3896, cid_3920, cid_3936, cid_3940, cid_3944).

Как запустить:
1. Установить зависимости (нужен Python 3 и пакет requests):
   py -m pip install requests
   или
   python -m pip install requests

2. Сохранить этот файл под именем, например: check_sitemap_urls.py

3. Запустить из терминала, например:
   py check_sitemap_urls.py ^
     --root-sitemap https://kotofoto.ru/sitemap.xml ^
     --output kotofoto_sitemap_issues.csv ^
     --output-all kotofoto_sitemap_all.csv ^
     --timeout 25 ^
     --workers 10

Основные параметры:
- --root-sitemap  (обязательно) — URL корневого sitemap.xml.
- --output        (опционально) — CSV только с ПРОБЛЕМНЫМИ URL (по умолчанию sitemap_issues.csv).
- --output-all    (опционально) — CSV со ВСЕМИ URL и статусами (по умолчанию sitemap_all_urls.csv).
- --timeout       (опционально) — таймаут HTTP-запросов в секундах (по умолчанию 10, для тяжёлых сайтов лучше 20–45).
- --workers       (опционально) — количество параллельных потоков для проверки URL (по умолчанию 10).

Что будет на выходе:
1) CSV-файл с проблемными URL:
   - sitemap_source — из какого sitemap-файла взят URL.
   - url           — сам URL.
   - final_status  — финальный HTTP-статус (после всех редиректов).
   - redirect_chain — цепочка редиректов (если были).
   - is_18plus     — true/false, относится ли URL к 18+ категориям.
   - issue_type    — тип проблемы: redirect_in_sitemap, error_in_sitemap, adult_ok, adult_with_error.

2) CSV-файл со всеми URL:
   - те же поля, но вместо issue_type используется raw_issue_type:
     ok / redirect_in_sitemap / error_in_sitemap / adult_ok / adult_with_error.

Скрипт рассчитан на аудит больших sitemap (как у интернет-магазинов) и помогает быстро найти
битые URL, лишние редиректы и отдельный слой 18+ страниц, которые попадают в карты сайта.
"""

import argparse
import csv
import sys
import xml.etree.ElementTree as ET
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib.parse import urlparse

import requests

NAMESPACE = {"sm": "http://www.sitemaps.org/schemas/sitemap/0.9"}

# Паттерны для 18+ категорий
ADULT_PATTERNS = [
    "cid_3896",  # товары для взрослых
    "cid_3920",  # секс-игрушки
    "cid_3936",  # клиторальные стимуляторы
    "cid_3940",  # мастурбаторы
    "cid_3944",  # анальные пробки
]


def fetch_xml(url, timeout):
    """Загрузка XML sitemap по URL."""
    headers = {
        "User-Agent": "SitemapAuditBot/1.0 (+https://example.com/contact)",
    }
    resp = requests.get(url, headers=headers, timeout=timeout)
    resp.raise_for_status()
    return resp.text


def parse_sitemap(url, timeout, seen_sitemaps, all_urls):
    """Рекурсивный разбор sitemapindex/urlset и сбор URL."""
    if url in seen_sitemaps:
        return
    seen_sitemaps.add(url)

    try:
        xml_text = fetch_xml(url, timeout)
    except Exception as e:
        print(f"[ERROR] Не удалось загрузить sitemap {url}: {e}", file=sys.stderr)
        return

    try:
        root = ET.fromstring(xml_text)
    except Exception as e:
        print(f"[ERROR] Не удалось распарсить sitemap {url}: {e}", file=sys.stderr)
        return

    tag = root.tag
    if tag.endswith("sitemapindex"):
        # индекс карт
        for sm_el in root.findall("sm:sitemap", NAMESPACE):
            loc_el = sm_el.find("sm:loc", NAMESPACE)
            if loc_el is not None and loc_el.text:
                child_url = loc_el.text.strip()
                parse_sitemap(child_url, timeout, seen_sitemaps, all_urls)
    elif tag.endswith("urlset"):
        # обычная карта
        for url_el in root.findall("sm:url", NAMESPACE):
            loc_el = url_el.find("sm:loc", NAMESPACE)
            if loc_el is not None and loc_el.text:
                page_url = loc_el.text.strip()
                all_urls.append((url, page_url))  # (sitemap_source, url)
    else:
        print(f"[WARN] Неизвестный корневой тег в {url}: {tag}", file=sys.stderr)


def is_adult_url(url):
    """Проверка, относится ли URL к 18+ категориям по паттернам в пути."""
    path = urlparse(url).path
    return any(pat in path for pat in ADULT_PATTERNS)


def check_url(session, sitemap_source, url, timeout):
    """
    Проверка одного URL:
    - HEAD-запрос с allow_redirects=True
    - несколько попыток при таймаутах/ошибках сети
    - возвращает данные по URL + признак проблемы.
    """
    headers = {
        "User-Agent": "SitemapAuditBot/1.0 (+https://example.com/contact)",
    }
    redirect_chain = []
    final_status = None

    attempts = 3
    for attempt in range(1, attempts + 1):
        try:
            resp = session.head(
                url, headers=headers, allow_redirects=True, timeout=timeout
            )
            final_status = resp.status_code
            history = resp.history
            redirect_chain = [f"{h.status_code}:{h.url}" for h in history]
            if history:
                redirect_chain.append(f"{resp.status_code}:{resp.url}")
            break  # успешный ответ – выходим из цикла
        except Exception as e:
            if attempt == attempts:
                final_status = 0
                redirect_chain = [f"error:{e.__class__.__name__}"]
            # иначе пробуем ещё раз

    is_adult = is_adult_url(url)

    issue_type = None
    if final_status in range(300, 400):
        issue_type = "redirect_in_sitemap"
    elif final_status >= 400 or final_status == 0:
        issue_type = "error_in_sitemap"
    elif is_adult and final_status == 200:
        issue_type = "adult_ok"
    elif is_adult and (final_status >= 300 or final_status == 0):
        issue_type = "adult_with_error"

    # raw_issue_type — всегда заполнен (для полного файла)
    if issue_type is None:
        raw_issue_type = "ok"
    else:
        raw_issue_type = issue_type

    return {
        "sitemap_source": sitemap_source,
        "url": url,
        "final_status": final_status,
        "redirect_chain": " -> ".join(redirect_chain),
        "is_18plus": str(is_adult).lower(),
        "issue_type": issue_type,         # может быть None
        "raw_issue_type": raw_issue_type, # всегда != None
    }


def main():
    parser = argparse.ArgumentParser(
        description="Проверка URL из sitemap на статусы и 18+ категории"
    )
    parser.add_argument(
        "--root-sitemap", required=True, help="URL корневого sitemap.xml"
    )
    parser.add_argument(
        "--output",
        default="sitemap_issues.csv",
        help="Путь к CSV-файлу с ПРОБЛЕМНЫМИ URL",
    )
    parser.add_argument(
        "--output-all",
        default="sitemap_all_urls.csv",
        help="Путь к CSV-файлу со ВСЕМИ URL и статусами",
    )
    parser.add_argument(
        "--timeout", type=int, default=10, help="Таймаут HTTP-запросов (сек)"
    )
    parser.add_argument(
        "--workers", type=int, default=10, help="Количество параллельных потоков"
    )

    args = parser.parse_args()

    seen_sitemaps = set()
    all_urls = []  # list of (sitemap_source, url)

    print(f"[INFO] Загружаем и обходим sitemap: {args.root_sitemap}")
    parse_sitemap(args.root_sitemap, args.timeout, seen_sitemaps, all_urls)

    if not all_urls:
        print("[WARN] Не найдено ни одного URL в sitemap.")
        return

    # Убираем дубли
    unique = {}
    for src, u in all_urls:
        if u not in unique:
            unique[u] = src

    urls_to_check = [(src, u) for u, src in unique.items()]
    total = len(urls_to_check)
    print(f"[INFO] Всего уникальных URL для проверки: {total}")

    all_results = []      # все URL
    problem_results = []  # только проблемные

    def print_progress(done, total_):
        if total_ == 0:
            return
        percent = done / total_ * 100
        bar_len = 30
        filled = int(bar_len * done / total_)
        bar = "#" * filled + "." * (bar_len - filled)
        print(
            f"\r[PROGRESS] [{bar}] {percent:5.1f}% ({done}/{total_})",
            end="",
            flush=True,
        )

    with requests.Session() as session:
        with ThreadPoolExecutor(max_workers=args.workers) as executor:
            futures = [
                executor.submit(check_url, session, sitemap_source, url, args.timeout)
                for sitemap_source, url in urls_to_check
            ]

            done = 0
            print_progress(done, total)
            for fut in as_completed(futures):
                res = fut.result()
                if res:
                    all_results.append(res)
                    if res["issue_type"] is not None:
                        problem_results.append(res)
                done += 1
                print_progress(done, total)

    print()  # перенос строки после прогресс-бара
    print(f"[INFO] Всего URL (all): {len(all_results)}")
    print(f"[INFO] Проблемных URL: {len(problem_results)}")

    # ----- пишем ПРОБЛЕМНЫЕ URL -----
    problem_fieldnames = [
        "sitemap_source",
        "url",
        "final_status",
        "redirect_chain",
        "is_18plus",
        "issue_type",
    ]

    with open(args.output, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=problem_fieldnames)
        writer.writeheader()
        for row in problem_results:
            writer.writerow({k: row[k] for k in problem_fieldnames})

    print(f"[INFO] Проблемные URL сохранены в {args.output}")

    # ----- пишем ВСЕ URL -----
    all_fieldnames = [
        "sitemap_source",
        "url",
        "final_status",
        "redirect_chain",
        "is_18plus",
        "raw_issue_type",
    ]

    with open(args.output_all, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=all_fieldnames)
        writer.writeheader()
        for row in all_results:
            writer.writerow({k: row[k] for k in all_fieldnames})

    print(f"[INFO] Все URL сохранены в {args.output_all}")


if __name__ == "__main__":
    main()
