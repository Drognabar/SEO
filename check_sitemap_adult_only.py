"""
Упрощённый скрипт аудита sitemap на наличие 18+ категорий.

Что делает:
- Берёт корневой sitemap (sitemap.xml) по переданному URL.
- Рекурсивно обходит все вложенные sitemap-файлы (sitemapindex -> sitemap -> urlset).
- Собирает все уникальные URL из карт сайта (без HTTP-проверок).
- По пути URL ищет паттерны 18+ категорий:
  cid_3896, cid_3920, cid_3921, cid_3936, cid_3940, cid_3941, cid_3944.
- На выходе:
  - печатает в консоль, есть ли 18+ в картах сайта и сколько таких URL;
  - по желанию пишет CSV со списком 18+ URL (sitemap_source, url).

Пример запуска:
  py check_sitemap_adult_only.py ^
    --root-sitemap https://kotofoto.ru/sitemap.xml ^
    --output-adult kotofoto_sitemap_adult_urls.csv ^
    --timeout 25
"""

import argparse
import csv
import sys
import xml.etree.ElementTree as ET
from urllib.parse import urlparse

import requests

NAMESPACE = {"sm": "http://www.sitemaps.org/schemas/sitemap/0.9"}

# Паттерны для 18+ категорий (категории и товары внутри них)
ADULT_PATTERNS = [
    "cid_3896",  # tovary_dlya_vzroslyh
    "cid_3920",  # seks_igrushki
    "cid_3921",  # vibratory
    "cid_3936",  # klitoralnye_stimulyatory
    "cid_3940",  # masturbatory
    "cid_3941",  # nasadki_na_stimulyatory
    "cid_3944",  # plagi_probki_analnye
]


def fetch_xml(url, timeout):
    """Загрузка XML sitemap по URL."""
    headers = {
        "User-Agent": "SitemapAdultAuditBot/1.0 (+https://example.com/contact)",
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


def main():
    parser = argparse.ArgumentParser(
        description="Проверка sitemap на наличие 18+ категорий (без HTTP-запросов)"
    )
    parser.add_argument(
        "--root-sitemap", required=True, help="URL корневого sitemap.xml"
    )
    parser.add_argument(
        "--output-adult",
        default="sitemap_adult_urls.csv",
        help="Путь к CSV-файлу с 18+ URL (опционально)",
    )
    parser.add_argument(
        "--timeout", type=int, default=10, help="Таймаут HTTP-запросов при загрузке XML"
    )

    args = parser.parse_args()

    seen_sitemaps = set()
    all_urls = []  # list of (sitemap_source, url)

    print(f"[INFO] Загружаем и обходим sitemap: {args.root_sitemap}")
    parse_sitemap(args.root_sitemap, args.timeout, seen_sitemaps, all_urls)

    if not all_urls:
        print("[WARN] Не найдено ни одного URL в sitemap.")
        return

    # Убираем дубли по URL
    unique = {}
    for src, u in all_urls:
        if u not in unique:
            unique[u] = src

    urls = [(src, u) for u, src in unique.items()]
    total_urls = len(urls)
    print(f"[INFO] Всего уникальных URL в картах сайта: {total_urls}")

    adult_urls = []
    for sitemap_source, url in urls:
        if is_adult_url(url):
            adult_urls.append({"sitemap_source": sitemap_source, "url": url})

    adult_count = len(adult_urls)
    if adult_count == 0:
        print("[RESULT] В картах сайта НЕ найдено URL с 18+ категориями.")
    else:
        print(
            f"[RESULT] В картах сайта найдены 18+ URL: всего {adult_count} шт."
        )

        # Пишем CSV со списком 18+ URL
        try:
            with open(args.output_adult, "w", newline="", encoding="utf-8") as f:
                writer = csv.DictWriter(f, fieldnames=["sitemap_source", "url"])
                writer.writeheader()
                for row in adult_urls:
                    writer.writerow(row)
            print(f"[INFO] 18+ URL сохранены в {args.output_adult}")
        except Exception as e:
            print(f"[ERROR] Не удалось записать CSV {args.output_adult}: {e}", file=sys.stderr)


if __name__ == "__main__":
    main()
