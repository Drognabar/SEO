"""
Скрипт проверяет мета‑теги и H1 для списка страниц сайта.

Что делает:
- Берёт список URL из текстового файла (по умолчанию pages.txt).
- Отбрасывает мусорные адреса: wp‑служебные пути, параметрические дубли,
  страницы поиска, сортировки, пагинации, ссылки на файлы/картинки.
- По каждому «чистому» URL получает статус ответа, <title>, meta description и H1.
- Формирует CSV‑отчёт clean_meta_report.csv с колонками
  url, status, title, description, h1, отсортированный по статусу и заголовку.

Как пользоваться:
1. Установить зависимости:
   pip install requests beautifulsoup4 pandas
2. Подготовить файл ncarus_pages.txt в той же папке — по одному URL в строке.
3. При необходимости заменить значения INPUT_FILE и OUTPUT_FILE в настройках.
4. Запустить скрипт из консоли:
   py meta_checker.py
5. Открыть clean_meta_report.csv в Excel и анализировать статусы, пустые title,
   description и H1.
"""
import requests
from bs4 import BeautifulSoup
import pandas as pd
import time
import re

# --- НАСТРОЙКИ ---
INPUT_FILE = "pages.txt"
OUTPUT_FILE = "clean_meta_report.csv"
TIMEOUT = 10
USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"

# Логика "мусора" (как в нашем промте)
def is_junk(url):
    # 1. Технические файлы WordPress
    if '/wp-content/' in url or '/wp-includes/' in url:
        return True
    # 2. Параметрические дубли (?p=123, ?post_type=)
    if '?p=' in url or '?post_type=' in url or '?id=' in url:
        return True
    # 3. Поиск, сортировка, фильтры
    if '?q=' in url or '/search/' in url or '?sort=' in url or 'price_min' in url:
        return True
    # 4. Пагинация (page/2, page-2)
    if re.search(r'/page/\d+', url) or re.search(r'page-\d+', url):
        return True
    # 5. Ссылки на картинки/файлы (если вдруг попали как html)
    if url.lower().endswith(('.jpg', '.png', '.pdf', '.jpeg', '.gif', '.xml')):
        return True
    
    return False

def get_page_data(url):
    headers = {'User-Agent': USER_AGENT}
    try:
        response = requests.get(url, headers=headers, timeout=TIMEOUT)
        
        # Если редирект, отслеживаем конечный статус
        status_code = response.status_code
        if status_code != 200:
            return {'url': url, 'status': status_code, 'title': '', 'description': '', 'h1': ''}
            
        response.encoding = response.apparent_encoding
        soup = BeautifulSoup(response.text, 'html.parser')
        
        title = soup.find('title').get_text(strip=True) if soup.find('title') else ''
        
        desc_tag = soup.find('meta', attrs={'name': 'description'})
        description = desc_tag['content'].strip() if desc_tag and desc_tag.get('content') else ''
        
        h1_tag = soup.find('h1')
        h1 = h1_tag.get_text(strip=True) if h1_tag else ''
        
        return {
            'url': url,
            'status': status_code,
            'title': title,
            'description': description,
            'h1': h1
        }
    except Exception as e:
        return {'url': url, 'status': 'Error', 'title': str(e), 'description': '', 'h1': ''}

def main():
    try:
        with open(INPUT_FILE, 'r', encoding='utf-8') as f:
            raw_urls = [line.strip() for line in f if line.strip()]
    except FileNotFoundError:
        print("Файл не найден!")
        return

    # 1. ФИЛЬТРАЦИЯ
    clean_urls = [url for url in raw_urls if not is_junk(url)]
    junk_count = len(raw_urls) - len(clean_urls)

    print(f"Всего URL в файле: {len(raw_urls)}")
    print(f"Отброшено мусора: {junk_count}")
    print(f"Осталось полезных: {len(clean_urls)}")
    print("-" * 30)

    # 2. ПАРСИНГ
    results = []
    for i, url in enumerate(clean_urls, 1):
        print(f"[{i}/{len(clean_urls)}] Парсим: {url}")
        data = get_page_data(url)
        results.append(data)
        time.sleep(0.1) 

    # 3. СОХРАНЕНИЕ
    df = pd.DataFrame(results)
    # Сортировка: сначала ошибки (404), потом пустые теги, потом нормальные
    df.sort_values(by=['status', 'title'], inplace=True)
    
    df.to_csv(OUTPUT_FILE, index=False, sep=';', encoding='utf-8-sig')
    print(f"\nГотово! Отчет сохранен в {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
