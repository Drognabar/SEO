"""
Скрипт — простой однопоточный краулер сайта https://dataru.ru/,
который обходит внутренние ссылки, собирает до 2000 уникальных
страниц домена и сохраняет их список в текстовый файл dataru.ru.txt.
"""
import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin, urlparse
import time

start_url = "https://dataru.ru/"
domain = urlparse(start_url).netloc

visited = set()
to_visit = {start_url}
max_pages = 2000  # можно увеличить, если нужно

while to_visit and len(visited) < max_pages:
    url = to_visit.pop()
    if url in visited:
        continue
    visited.add(url)

    try:
        headers = {"User-Agent": "Mozilla/5.0"}
        response = requests.get(url, headers=headers, timeout=10)
        if response.status_code != 200:
            continue

        soup = BeautifulSoup(response.text, "html.parser")
        print(f"Обработано: {len(visited)} | {url}")

        for a in soup.find_all("a", href=True):
            href = urljoin(url, a["href"])
            parsed = urlparse(href)
            if parsed.netloc == domain and href.startswith("https://"):
                if "#" in href:
                    href = href.split("#")[0]
                if href not in visited:
                    to_visit.add(href)

        time.sleep(0.3)
    except Exception as e:
        print(f"Ошибка {url}: {e}")

print(f"\nНайдено уникальных страниц: {len(visited)}")
with open("dataru.ru.txt", "w", encoding="utf-8") as f:
    f.write("\n".join(sorted(visited)))
print("Список сохранён в dataru.ru.txt")
