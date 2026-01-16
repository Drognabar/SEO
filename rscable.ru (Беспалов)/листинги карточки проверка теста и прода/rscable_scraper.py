"""
Скрипт сравнивает SEO‑тексты под листингом товаров на боевом и тестовом
(закрытом по логину/паролю) вариантах сайта rscable и сохраняет результаты в Excel.

Что делает:
- Берёт из Excel‑файла INPUT_FILE список пар URL со столбцами old_url (боевая) и new_url (тестовая).
- Для каждой пары:
  - скачивает HTML боевой страницы без авторизации;
  - скачивает HTML тестовой страницы с HTTP Basic Auth (LOGIN / PASSWORD);
  - вытаскивает SEO‑текст под листингом по нужному CSS‑селектору:
    * для боевой версии: div.group_description_block.qweqweqwe.bottom.muted777
    * для тестовой версии: div.group_description_block.bottom.font_15
- Очищает текст от лишних пробелов.
- Считает:
  - процент похожести текстов (SequenceMatcher, от 0 до 100);
  - разницу длины текста (новый минус старый);
  - статус:
    * EMPTY_BOTH — текстов нет на обеих версиях;
    * EMPTY_OLD — нет текста на боевой;
    * EMPTY_NEW — нет текста на тестовой;
    * MATCH — тексты совпадают на 100%;
    * ALMOST — похожесть ≥ 95%;
    * SIMILAR — похожесть ≥ 80%;
    * DIFF — сильно различаются.
- Складывает результаты в Excel‑файл OUTPUT_FILE со столбцами:
  A_URL_OLD, B_TEXT_OLD, C_URL_NEW, D_TEXT_NEW, E_SIMILARITY_%,
  F_STATUS, G_LENGTH_DIFF.
- Между запросами делает паузу SLEEP_BETWEEN_REQUESTS секунд, чтобы не долбить сайт.

Как подготовить:
1. Установить зависимости:
   pip install requests beautifulsoup4 pandas lxml
2. В начале файла задать:
   - LOGIN и PASSWORD — доступ к тестовому сайту (HTTP Basic Auth).
   - INPUT_FILE — путь к Excel со столбцами old_url и new_url.
   - OUTPUT_FILE — имя/путь для итогового отчёта.
   - SLEEP_BETWEEN_REQUESTS и REQUEST_TIMEOUT при необходимости.
3. В INPUT_FILE каждая строка должна содержать пару URL:
   - old_url — боевой адрес;
   - new_url — соответствующий тестовый адрес.

Как пользоваться:
1. Поместить INPUT_FILE и скрипт в одну папку (или прописать полный путь к файлу в настройках).
2. Запустить скрипт из консоли из папки со скриптом:
   py rscable_scraper.py
3. Дождаться завершения (в консоли будет лог по каждой паре URL и финальное сообщение
   «Готово, файл сохранён: ...»).
4. Открыть OUTPUT_FILE в Excel и анализировать:
   - где тексты пустые,
   - где сильно отличаются,
   - где длина и похожесть в норме.
"""

import requests
from requests.auth import HTTPBasicAuth
from bs4 import BeautifulSoup
import pandas as pd
import time
from difflib import SequenceMatcher


# ---- НАСТРОЙКИ ----

LOGIN = "super-rscable"
PASSWORD = "jF9&<qL1$<xE"

INPUT_FILE = "urls.xlsx"          # файл с колонками old_url / new_url
OUTPUT_FILE = "rscable_seo_text.xlsx"

REQUEST_TIMEOUT = 30
SLEEP_BETWEEN_REQUESTS = 0  # можно поставить 0 или 0.2–0.5 при необходимости


# ---- ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ----

def get_soup(html: str) -> BeautifulSoup:
    try:
        return BeautifulSoup(html, "lxml")
    except Exception:
        return BeautifulSoup(html, "html.parser")


def extract_seo_text(html: str, version: str) -> str:
    """
    SEO‑текст под листингом:
    - боевая:  div.group_description_block.qweqweqwe.bottom.muted777
    - тестовая: div.group_description_block.bottom.font_15
    """
    soup = get_soup(html)

    if version == "old":
        sel = "div.group_description_block.qweqweqwe.bottom.muted777"
    else:  # new
        sel = "div.group_description_block.bottom.font_15"

    block = soup.select_one(sel)
    if not block:
        return ""

    text = block.get_text(separator=" ", strip=True)
    return " ".join(text.split())


def fetch_html(url: str, use_auth: bool = False) -> str:
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
    }
    try:
        if use_auth:
            resp = requests.get(
                url,
                timeout=REQUEST_TIMEOUT,
                auth=HTTPBasicAuth(LOGIN, PASSWORD),
                headers=headers,
                verify=False,
            )
        else:
            resp = requests.get(
                url,
                timeout=REQUEST_TIMEOUT,
                headers=headers,
                verify=False,
            )
        resp.raise_for_status()
        return resp.text
    except Exception as e:
        return f"ERROR_HTML: {e}"


def calculate_similarity(text_old: str, text_new: str) -> float:
    """Процент похожести текстов от 0 до 100."""
    if not text_old and not text_new:
        return 100.0

    if not text_old or not text_new:
        return 0.0

    matcher = SequenceMatcher(None, text_old, text_new)
    ratio = matcher.ratio()
    return round(ratio * 100, 1)


def get_status(similarity: float, text_old: str, text_new: str) -> str:
    """Статус для столбца F."""
    if not text_old and not text_new:
        return "EMPTY_BOTH"

    if not text_old:
        return "EMPTY_OLD"

    if not text_new:
        return "EMPTY_NEW"

    if similarity == 100:
        return "MATCH"
    elif similarity >= 95:
        return "ALMOST"
    elif similarity >= 80:
        return "SIMILAR"
    else:
        return "DIFF"


# ---- ОСНОВНОЙ ПРОЦЕСС ----

def main():
    # читаем список пар URL
    df_urls = pd.read_excel(INPUT_FILE)  # ожидаются колонки old_url / new_url
    results = []

    total = len(df_urls)
    print(f"\nНайдено пар URL: {total}\n")

    for idx, row in df_urls.iterrows():
        old_url = str(row["old_url"]).strip()
        new_url = str(row["new_url"]).strip()

        print(f"[{idx + 1}/{total}]")
        print(f"  Боевая:  {old_url}")
        html_old = fetch_html(old_url, use_auth=False)
        if html_old.startswith("ERROR_HTML:"):
            print(f"    ✗ Ошибка боевой: {html_old}")
            seo_old = ""
        else:
            seo_old = extract_seo_text(html_old, version="old")
            print(f"    ✓ SEO‑текст (боевой): {len(seo_old)} символов")

        time.sleep(SLEEP_BETWEEN_REQUESTS)

        print(f"  Тестовая: {new_url}")
        html_new = fetch_html(new_url, use_auth=True)
        if html_new.startswith("ERROR_HTML:"):
            print(f"    ✗ Ошибка тестовой: {html_new}")
            seo_new = ""
        else:
            seo_new = extract_seo_text(html_new, version="new")
            print(f"    ✓ SEO‑текст (тест): {len(seo_new)} символов")

        # метрики сравнения
        similarity = calculate_similarity(seo_old, seo_new)
        status = get_status(similarity, seo_old, seo_new)
        length_diff = len(seo_new) - len(seo_old)

        print(f"    → Статус: {status}, похожесть: {similarity}%, Δдлины: {length_diff}\n")

        results.append({
            "A_URL_OLD": old_url,
            "B_TEXT_OLD": seo_old,
            "C_URL_NEW": new_url,
            "D_TEXT_NEW": seo_new,
            "E_SIMILARITY_%": similarity,
            "F_STATUS": status,
            "G_LENGTH_DIFF": length_diff
        })

    # собираем итоговый DataFrame и сохраняем
    df_out = pd.DataFrame(results, columns=[
        "A_URL_OLD",
        "B_TEXT_OLD",
        "C_URL_NEW",
        "D_TEXT_NEW",
        "E_SIMILARITY_%",
        "F_STATUS",
        "G_LENGTH_DIFF",
    ])

    df_out.to_excel(OUTPUT_FILE, index=False)
    print(f"\n✅ Готово, файл сохранён: {OUTPUT_FILE}\n")


if __name__ == "__main__":
    main()
