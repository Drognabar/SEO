python
"""
Скрипт сравнивает два текста в карточках товаров (боевой и тестовый
вариант сайта rscable) и фиксирует результаты по каждой паре URL в Excel.

Что делает:
- Берёт из Excel‑файла INPUT_FILE список пар URL со столбцами old_url (боевая карточка)
  и new_url (тестовая карточка).
- Для каждой пары:
  - скачивает HTML боевой карточки без авторизации;
  - скачивает HTML тестовой карточки с HTTP Basic Auth (LOGIN / PASSWORD);
  - вытаскивает по 2 текстовых блока из каждой карточки по CSS‑селекторам.

  Боевая карточка:
  - Текст 1: div.font_xs.text-block.preview-text-replace
  - Текст 2: div.content.detail-text-wrap[itemprop="description"]

  Тестовая карточка:
  - Текст 1: div.grid-list__item.catalog-detail__previewtext[itemprop="description"]
  - Текст 2: div.content.content--max-width.js-detail-description[itemprop="description"]

- Очищает тексты от лишних пробелов.
- Для каждой пары текстов (Текст1 старый/новый и Текст2 старый/новый):
  - считает процент похожести (SequenceMatcher, от 0 до 100);
  - присваивает статус:
    * EMPTY_BOTH — оба текста пустые;
    * EMPTY_OLD — пустой только старый текст;
    * EMPTY_NEW — пустой только новый текст;
    * MATCH — 100% совпадение;
    * ALMOST — похожесть ≥ 95%;
    * SIMILAR — похожесть ≥ 80%;
    * DIFF — тексты заметно отличаются.
- Формирует Excel‑отчёт OUTPUT_FILE с колонками:
  A_URL_OLD      — URL боевой карточки,
  B_TEXT1_OLD    — Текст 1 (боевой),
  C_TEXT2_OLD    — Текст 2 (боевой),
  D_URL_NEW      — URL тестовой карточки,
  E_TEXT1_NEW    — Текст 1 (тест),
  F_TEXT2_NEW    — Текст 2 (тест),
  G_SIM_TEXT1_%  — похожесть текстов 1, %,
  H_STATUS_TEXT1 — статус по текстам 1,
  I_SIM_TEXT2_%  — похожесть текстов 2, %,
  J_STATUS_TEXT2 — статус по текстам 2.
- Между запросами делает паузу SLEEP_BETWEEN_REQUESTS секунд.

Как подготовить:
1. Установить зависимости:
   pip install requests beautifulsoup4 pandas lxml
2. В начале файла задать:
   - LOGIN и PASSWORD — логин/пароль для доступа к тестовому сайту (HTTP Basic Auth).
   - INPUT_FILE — путь к Excel со столбцами old_url и new_url.
   - OUTPUT_FILE — имя/путь итогового отчёта.
   - REQUEST_TIMEOUT и SLEEP_BETWEEN_REQUESTS при необходимости.
3. В INPUT_FILE каждая строка должна содержать пару URL:
   - old_url — адрес боевой карточки;
   - new_url — соответствующий адрес тестовой карточки.

Как пользоваться:
1. Поместить INPUT_FILE и скрипт в одну папку (или прописать полный путь до файла в настройках).
2. Запустить скрипт из консоли из папки со скриптом:
   py rscable_product_compare.py
3. Дождаться завершения (в консоли будут логи по каждой карточке и финальное
   сообщение «Готово, файл сохранён: ...»).
4. Открыть OUTPUT_FILE в Excel и анализировать:
   - где тексты отсутствуют,
   - где тексты полностью совпадают / почти совпадают,
   - где сильно отличаются по содержанию.
"""

import requests
from requests.auth import HTTPBasicAuth
from bs4 import BeautifulSoup
import pandas as pd
from difflib import SequenceMatcher
import time


# ---- НАСТРОЙКИ ----

LOGIN = "super-rscable"
PASSWORD = "jF9&<qL1$<xE"

INPUT_FILE = "urls_products.xlsx"        # файл с колонками old_url / new_url
OUTPUT_FILE = "rscable_product_texts.xlsx"

REQUEST_TIMEOUT = 30
SLEEP_BETWEEN_REQUESTS = 0   # при необходимости можно поставить 0.2–0.5


# ---- ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ----

def get_soup(html: str) -> BeautifulSoup:
    try:
        return BeautifulSoup(html, "lxml")
    except Exception:
        return BeautifulSoup(html, "html.parser")


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


def clean_text(block) -> str:
    if not block:
        return ""
    text = block.get_text(separator=" ", strip=True)
    return " ".join(text.split())


def extract_texts_old(html: str) -> tuple[str, str]:
    """
    Боевая карточка товара:
    - Текст 1: <div class="font_xs text-block preview-text-replace">
    - Текст 2: <div class="content detail-text-wrap 45" itemprop="description">
    """
    soup = get_soup(html)

    t1_block = soup.select_one("div.font_xs.text-block.preview-text-replace")
    t2_block = soup.select_one(
        'div.content.detail-text-wrap[itemprop="description"]'
    )

    return clean_text(t1_block), clean_text(t2_block)


def extract_texts_new(html: str) -> tuple[str, str]:
    """
    Тестовая карточка товара:
    - Текст 1: <div class="grid-list__item catalog-detail__previewtext" itemprop="description">
    - Текст 2: <div class="content content--max-width js-detail-description" itemprop="description">
    """
    soup = get_soup(html)

    t1_block = soup.select_one(
        'div.grid-list__item.catalog-detail__previewtext[itemprop="description"]'
    )
    t2_block = soup.select_one(
        'div.content.content--max-width.js-detail-description[itemprop="description"]'
    )

    return clean_text(t1_block), clean_text(t2_block)


def calculate_similarity(text_a: str, text_b: str) -> float:
    """Процент похожести текстов от 0 до 100."""
    if not text_a and not text_b:
        return 100.0
    if not text_a or not text_b:
        return 0.0
    matcher = SequenceMatcher(None, text_a, text_b)
    return round(matcher.ratio() * 100, 1)


def get_status(similarity: float, text_a: str, text_b: str) -> str:
    """Статус сравнения."""
    if not text_a and not text_b:
        return "EMPTY_BOTH"
    if not text_a:
        return "EMPTY_OLD"
    if not text_b:
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
    df_urls = pd.read_excel(INPUT_FILE)  # ожидаются колонки old_url / new_url
    results = []

    total = len(df_urls)
    print(f"\nНайдено пар URL (карточки): {total}\n")

    for idx, row in df_urls.iterrows():
        old_url = str(row["old_url"]).strip()
        new_url = str(row["new_url"]).strip()

        print(f"[{idx + 1}/{total}]")
        print(f"  Боевая:  {old_url}")
        html_old = fetch_html(old_url, use_auth=False)
        if html_old.startswith("ERROR_HTML:"):
            print(f"    ✗ Ошибка боевой: {html_old}")
            old_t1, old_t2 = "", ""
        else:
            old_t1, old_t2 = extract_texts_old(html_old)
            print(
                f"    ✓ Текст1 боевой: {len(old_t1)} симв., "
                f"Текст2 боевой: {len(old_t2)} симв."
            )

        time.sleep(SLEEP_BETWEEN_REQUESTS)

        print(f"  Тестовая: {new_url}")
        html_new = fetch_html(new_url, use_auth=True)
        if html_new.startswith("ERROR_HTML:"):
            print(f"    ✗ Ошибка тестовой: {html_new}")
            new_t1, new_t2 = "", ""
        else:
            new_t1, new_t2 = extract_texts_new(html_new)
            print(
                f"    ✓ Текст1 тест: {len(new_t1)} симв., "
                f"Текст2 тест: {len(new_t2)} симв."
            )

        # Сравнение текстов 1 и 2 по отдельности
        sim1 = calculate_similarity(old_t1, new_t1)
        status1 = get_status(sim1, old_t1, new_t1)

        sim2 = calculate_similarity(old_t2, new_t2)
        status2 = get_status(sim2, old_t2, new_t2)

        print(
            f"    → Текст1: {status1}, {sim1}% | "
            f"Текст2: {status2}, {sim2}%\n"
        )

        results.append({
            "A_URL_OLD": old_url,
            "B_TEXT1_OLD": old_t1,
            "C_TEXT2_OLD": old_t2,
            "D_URL_NEW": new_url,
            "E_TEXT1_NEW": new_t1,
            "F_TEXT2_NEW": new_t2,
            "G_SIM_TEXT1_%": sim1,
            "H_STATUS_TEXT1": status1,
            "I_SIM_TEXT2_%": sim2,
            "J_STATUS_TEXT2": status2,
        })

    df_out = pd.DataFrame(results, columns=[
        "A_URL_OLD",
        "B_TEXT1_OLD",
        "C_TEXT2_OLD",
        "D_URL_NEW",
        "E_TEXT1_NEW",
        "F_TEXT2_NEW",
        "G_SIM_TEXT1_%",
        "H_STATUS_TEXT1",
        "I_SIM_TEXT2_%",
        "J_STATUS_TEXT2",
    ])

    df_out.to_excel(OUTPUT_FILE, index=False)
    print(f"\n✅ Готово, файл сохранён: {OUTPUT_FILE}\n")


if __name__ == "__main__":
    main()
