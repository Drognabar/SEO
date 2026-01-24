"""
generation_text_for_product.py - Универсальная генерация SEO-текстов для kotofoto.ru

ЛОГИКА РАБОТЫ ПО ШАГАМ:

1. Загружает 'input_urls.xlsx' (A: URL, B: Model 'small'/'large').

2. Парсит название с kotofoto.ru (h1/title).

3. Определяет категорию по ключевым словам (видеокарта/монитор/NAS/SSD/general).

4. Выбирает блоки под категорию.

5. Формирует универсальный промпт (URL + абзац + блоки).

6. Генерирует через Perplexity API.

7. Сохраняет в 'output_texts.xlsx' (C: SEO_Текст).

8. Пауза 3 сек между запросами.

ИНСТРУКЦИЯ ПО ЗАПУСКУ:

1. pip install pandas openai python-dotenv requests beautifulsoup4 openpyxl

2. .env: PERPLEXITY_API_KEY=pplx-...

3. input_urls.xlsx:

| URL | Model |
| https://kotofoto.ru/... | small |

4. python generation_text_for_product.py

5. Результат: output_texts.xlsx

Мониторь: Settings → API Usage
"""

import pandas as pd
from openai import OpenAI
from dotenv import load_dotenv
import os
import requests
from bs4 import BeautifulSoup
import time

load_dotenv()

client = OpenAI(api_key=os.getenv("PERPLEXITY_API_KEY"), base_url="https://api.perplexity.ai")

def scrape_title(url):
    try:
        resp = requests.get(url, timeout=10)
        soup = BeautifulSoup(resp.text, 'html.parser')
        title = (soup.find('h1') or soup.find('title') or soup.find('meta', {'property': 'og:title'}))
        return title.get_text(strip=True) if title else "Товар"
    except:
        return "Парсинг не удался"

def detect_category(title):
    title_lower = title.lower()
    if any(word in title_lower for word in ['видеокарта', 'gpu', 'rtx', 'rx', 'arc', 'vga']):
        return "videocard"
    elif any(word in title_lower for word in ['монитор', 'monitor', 'display']):
        return "monitor"
    elif any(word in title_lower for word in ['nas', 'сетевое хранилище']):
        return "nas"
    elif any(word in title_lower for word in ['ssd', 'nvme', 'sata']):
        return "storage"
    else:
        return "general"

def get_structure(category):
    structures = {
        "videocard": ['Для кого и под какие задачи', 'Архитектура и производительность', 'Подключения и охлаждение', 'Важные нюансы перед покупкой'],
        "monitor": ['Для кого и под какие задачи', 'Картинка и параметры дисплея', 'Гейминг и технологии', 'Подключения и эргономика', 'Важные нюансы перед покупкой'],
        "nas": ['Для кого и под какие задачи', 'Аппаратные возможности и хранение', 'Сети и расширения', 'Важные нюансы перед покупкой'],
        "storage": ['Для кого и под какие задачи', 'Характеристики накопителя', 'Интерфейсы и совместимость', 'Важные нюансы перед покупкой'],
        "general": ['Для кого и под какие задачи', 'Основные характеристики', 'Преимущества и применение', 'Важные нюансы перед покупкой']
    }
    return structures.get(category, structures["general"])

# Основной код
df = pd.read_excel('input_urls.xlsx')
df['Title'] = df['URL'].apply(scrape_title)
df['Category'] = df['Title'].apply(detect_category)
df['Structure'] = df['Category'].apply(get_structure)

for idx, row in df.iterrows():
    prompt = f"""
    Напиши SEO-текст для товара: {row['Title']}
    URL: {row['URL']}
    
    Структура:
    {chr(10).join(row['Structure'])}
    
    Текст на русском, 800-1200 слов, оптимизирован для kotofoto.ru.
    """
    
    response = client.chat.completions.create(
        model=row.get('Model', 'llama-3.1-sonar-small-128k-online'),
        messages=[{"role": "user", "content": prompt}]
    )
    
    df.at[idx, 'SEO_Текст'] = response.choices[0].message.content
    print(f"Обработан: {row['Title']}")
    
    time.sleep(3)

df.to_excel('output_texts.xlsx', index=False)
print("Готово! output_texts.xlsx создан.")
