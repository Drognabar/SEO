"""
Доработанный скрипт аудита sitemap.xml для SEO (с обходом cookiesync-защиты и цветным прогрессбаром)

Цель:
- Скачать основную XML-карту сайта (sitemap.xml), рекурсивно обойти все вложенные карты,
- собрать полный список URL + их lastmod,
- найти дубли внутри sitemap,
- проверить HTTP-статус каждого URL (301/302, 404, 5xx и т.п.),
- сравнить lastmod из sitemap с заголовком Last-Modified на сервере (если есть),
- скорректировать влияние защиты вида api.action-media.ru/fake-pages/cookiesync,
- сохранить результат в CSV для дальнейшего анализа в Excel / Google Sheets.

Столбцы в CSV:
- url                — адрес из sitemap
- sitemap_lastmod    — значение lastmod из sitemap (как строка)
- final_url          — итоговый адрес после всех редиректов
- status_code        — HTTP-статус финального ответа
- is_redirect        — 1, если был хотя бы один редирект
- redirect_chain     — цепочка адресов через " -> "
- last_modified_hdr  — заголовок Last-Modified (если есть)
- is_duplicate       — 1, если URL в sitemap встречался более одного раза
- status_source      — источник статуса:
    * "head"                  — нормальный HEAD
    * "get"                   — нормальный GET, если HEAD был подозрительным
    * "cookiesync_protection" — и HEAD, и GET ушли на cookiesync-заглушку
    * "error"                 — ошибка сети/запроса

Инструкция по запуску:
1) Установить Python 3.8+ и requests (pip install requests).
2) Сохранить файл, например sitemap_audit.py.
3) При необходимости поменять SITEMAP_URL на свою карту.
4) В терминале/VS Code перейти в папку со скриптом и выполнить:
   python sitemap_audit.py
5) Открыть 26-2_sitemap_audit.csv в Excel / Google Sheets и анализировать.

Рекомендованные фильтры:
- status_code = 404 и status_source != 'cookiesync_protection' — реальные битые URL.
- status_code = 301/302 и final_url != url — редиректы.
- is_duplicate = 1 — дубли в sitemap.
- status_source = 'cookiesync_protection' — ответы, искажённые защитой (их не учитывать).
"""

import requests
import xml.etree.ElementTree as ET
from urllib.parse import urljoin
from collections import Counter
import csv
import time

SITEMAP_URL = "https://www.26-2.ru/sitemap/https_rubric.xml"
TIMEOUT = 20
SLEEP = 0.2
HEADERS = {
    "User-Agent": "26-2-SEO-audit/1.0 (+https://www.26-2.ru/; sitemap technical check)"
}

NS = {"sm": "http://www.sitemaps.org/schemas/sitemap/0.9"}

COOKIESYNC_HOST = "api.action-media.ru"
COOKIESYNC_PATH_PART = "/fake-pages/cookiesync"


def fetch(url):
    resp = requests.get(url, headers=HEADERS, timeout=TIMEOUT)
    resp.raise_for_status()
    return resp


def parse_sitemap(url):
    """
    Возвращает список словарей: {"url": ..., "lastmod": ...}
    из sitemap или sitemapindex (с рекурсией).
    """
    items = []
    resp = fetch(url)
    root = ET.fromstring(resp.content)

    # sitemapindex
    if root.tag.endswith("sitemapindex"):
        for sm_el in root.findall("sm:sitemap", NS):
            loc_el = sm_el.find("sm:loc", NS)
            if loc_el is not None and loc_el.text:
                child_sitemap = loc_el.text.strip()
                child_sitemap = urljoin(url, child_sitemap)
                items.extend(parse_sitemap(child_sitemap))
        return items

    # urlset
    if root.tag.endswith("urlset"):
        for url_el in root.findall("sm:url", NS):
            loc_el = url_el.find("sm:loc", NS)
            if loc_el is None or not loc_el.text:
                continue
            loc = loc_el.text.strip()

            lastmod_el = url_el.find("sm:lastmod", NS)
            lastmod = lastmod_el.text.strip() if lastmod_el is not None and lastmod_el.text else ""

            items.append({"url": loc, "lastmod": lastmod})
        return items

    return items


def is_cookiesync_url(u: str) -> bool:
    return COOKIESYNC_HOST in u and COOKIESYNC_PATH_PART in u


def _request_with_method(url, method="HEAD"):
    """
    Один запрос (HEAD/GET) с редиректами.
    Возвращает (resp, error_str).
    """
    try:
        if method.upper() == "HEAD":
            resp = requests.head(
                url,
                headers=HEADERS,
                timeout=TIMEOUT,
                allow_redirects=True,
            )
        else:
            resp = requests.get(
                url,
                headers=HEADERS,
                timeout=TIMEOUT,
                allow_redirects=True,
            )
        return resp, ""
    except requests.RequestException as e:
        return None, str(e)


def check_url(url):
    """
    1) HEAD-запрос.
    2) Если код нормальный и не уводит на cookiesync, используем его (status_source='head').
    3) Иначе пробуем GET.
    4) Если и GET уносит на cookiesync — status_source='cookiesync_protection'.
       Если GET нормальный — status_source='get'.
    При ошибках — status_source='error'.
    """
    # HEAD
    resp_head, err_head = _request_with_method(url, method="HEAD")
    if resp_head is None:
        return url, -1, False, f"error: {err_head}", "", "error"

    final_url_head = resp_head.url
    status_code_head = resp_head.status_code
    chain_head = [r.url for r in resp_head.history] + [resp_head.url]
    is_redirect_head = len(resp_head.history) > 0
    last_modified_hdr_head = resp_head.headers.get("Last-Modified", "").strip()

    suspicious_code = (
        status_code_head in (403, 405, 429) or status_code_head >= 500
    )
    hit_cookiesync = is_cookiesync_url(final_url_head)

    if not suspicious_code and not hit_cookiesync:
        return (
            final_url_head,
            status_code_head,
            is_redirect_head,
            " -> ".join(chain_head),
            last_modified_hdr_head,
            "head",
        )

    # GET
    resp_get, err_get = _request_with_method(url, method="GET")
    if resp_get is None:
        return (
            final_url_head,
            status_code_head,
            is_redirect_head,
            " -> ".join(chain_head),
            last_modified_hdr_head,
            "error",
        )

    final_url_get = resp_get.url
    status_code_get = resp_get.status_code
    chain_get = [r.url for r in resp_get.history] + [resp_get.url]
    is_redirect_get = len(resp_get.history) > 0
    last_modified_hdr_get = resp_get.headers.get("Last-Modified", "").strip()
    hit_cookiesync_get = is_cookiesync_url(final_url_get)

    if hit_cookiesync_get:
        return (
            final_url_get,
            status_code_get,
            is_redirect_get,
            " -> ".join(chain_get),
            last_modified_hdr_get,
            "cookiesync_protection",
        )

    return (
        final_url_get,
        status_code_get,
        is_redirect_get,
        " -> ".join(chain_get),
        last_modified_hdr_get,
        "get",
    )


def main():
    print(f"Скачиваю sitemap: {SITEMAP_URL}")
    all_items = parse_sitemap(SITEMAP_URL)
    print(f"Всего записей в sitemap (сырых): {len(all_items)}")

    urls = [it["url"] for it in all_items]
    counter = Counter(urls)
    duplicates = {u for u, c in counter.items() if c > 1}
    print(f"Найдено дублей (точное совпадение URL): {len(duplicates)}")

    unique_map = {}
    for it in all_items:
        url = it["url"]
        if url not in unique_map:
            unique_map[url] = it["lastmod"]

    total = len(unique_map)
    print(f"Уникальных URL: {total}")

    rows = []
    protection_cnt = 0
    error_cnt = 0

    LIGHT_GREEN = "\033[92m"   # ярко‑зелёный
    DARK_GREEN = "\033[32m"    # тёмно‑зелёный
    RESET = "\033[0m"

    for i, (url, sitemap_lastmod) in enumerate(unique_map.items(), start=1):
        final_url, status_code, is_redirect, redirect_chain, last_modified_hdr, status_source = check_url(url)

        if status_source == "cookiesync_protection":
            protection_cnt += 1
        elif status_source == "error":
            error_cnt += 1

        rows.append({
            "url": url,
            "sitemap_lastmod": sitemap_lastmod,
            "final_url": final_url,
            "status_code": status_code,
            "is_redirect": int(is_redirect),
            "redirect_chain": redirect_chain,
            "last_modified_hdr": last_modified_hdr,
            "is_duplicate": int(url in duplicates),
            "status_source": status_source,
        })

        progress = (i / total) * 100
        bar_len = 50
        filled = int(bar_len * i / total)
        bar_raw = "#" * filled + "-" * (bar_len - filled)

        # 0–50% — светло‑зелёный, 50–100% — тёмно‑зелёный
        color = LIGHT_GREEN if progress < 50 else DARK_GREEN

        ok_cnt = i - protection_cnt - error_cnt

        print(
            f"{color}[{bar_raw}]{RESET} {progress:5.1f}%  "
            f"({i}/{total})  ok={ok_cnt}  prot={protection_cnt}  err={error_cnt}",
            end="\r"
        )

        time.sleep(SLEEP)

    print()

    out_file = "26-2_sitemap_audit.csv"
    fieldnames = [
        "url",
        "sitemap_lastmod",
        "final_url",
        "status_code",
        "is_redirect",
        "redirect_chain",
        "last_modified_hdr",
        "is_duplicate",
        "status_source",
    ]
    with open(out_file, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, delimiter=";")
        writer.writeheader()
        writer.writerows(rows)

    print(f"Аудит завершен, результат: {out_file}")
    print("Итоги по запросам:")
    print(f" - Всего URL: {total}")
    print(f" - Под защитой (cookiesync_protection): {protection_cnt}")
    print(f" - Ошибки (status_source = 'error'): {error_cnt}")
    print("Фильтры для анализа CSV:")
    print(" - status_code = 404 — реальные битые URL (status_source != 'cookiesync_protection')")
    print(" - status_code = 301/302 и final_url != url — редиректы")
    print(" - is_duplicate = 1 — дубли в sitemap")
    print(" - status_source = 'cookiesync_protection' — ответы, искажённые защитой")


if __name__ == "__main__":
    main()
