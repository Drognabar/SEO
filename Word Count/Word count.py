"""
Скрипт считает объём текста на одной веб‑странице по её URL.

Что делает:
- Загружает HTML страницы по адресу URL с помощью requests.
- Удаляет теги script, style, noscript и вытаскивает только видимый текст (BeautifulSoup).
- Очищает лишние пробелы и пустые строки.
- Считает количество абзацев, символов с пробелами и без, слов и предложений.
- Выводит эти показатели в консоль с русскими названиями полей.

Как пользоваться:
1. Установить зависимости:
   pip install requests beautifulsoup4
2. Вверху файла изменить значение переменной URL на адрес нужной страницы.
3. Запустить скрипт из консоли из папки со скриптом:
   py Word-count.py
4. Посмотреть в консоли статистику:
   "адрес", "абзацы", "символы_с_пробелами",
   "символы_без_пробелов", "слова", "предложения".
"""

import requests
from bs4 import BeautifulSoup
import re

# сюда можно подставлять любой адрес
URL = "https://ncarus.ru/"


def get_text_stats(url: str) -> dict:
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
    }

    resp = requests.get(url, headers=headers, timeout=10)
    resp.raise_for_status()

    soup = BeautifulSoup(resp.text, "html.parser")

    # убираем скрипты/стили
    for tag in soup(["script", "style", "noscript"]):
        tag.decompose()

    # берём только видимый текст
    text = soup.get_text(separator="\n")

    # чистим лишние пробелы
    lines = [line.strip() for line in text.splitlines()]
    lines = [line for line in lines if line]  # убираем пустые строки
    clean_text = "\n".join(lines)

    # абзацы считаем по непустым строкам
    paragraphs = [line for line in clean_text.split("\n") if line.strip()]
    paragraphs_count = len(paragraphs)

    # считаем символы
    chars_with_spaces = len(clean_text)
    chars_no_spaces = len(clean_text.replace(" ", "").replace("\n", ""))

    # слова: всё, что похоже на последовательность букв/цифр
    words = re.findall(r"\w+", clean_text, flags=re.UNICODE)
    words_count = len(words)

    # предложения: грубо по . ! ? (для русского обычно достаточно)
    sentences = re.split(r"[.!?]+", clean_text)
    sentences = [s.strip() for s in sentences if s.strip()]
    sentences_count = len(sentences)

    return {
        "адрес": url,
        "абзацы": paragraphs_count,
        "символы_с_пробелами": chars_with_spaces,
        "символы_без_пробелов": chars_no_spaces,
        "слова": words_count,
        "предложения": sentences_count,
    }


if __name__ == "__main__":
    stats = get_text_stats(URL)
    for key, value in stats.items():
        print(f"{key}: {value}")
