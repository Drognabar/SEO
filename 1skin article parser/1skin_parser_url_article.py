"""
╔════════════════════════════════════════════════════════════════════════════╗
║                   1SKIN PARSER - URL ARTICLE SCRAPER                       ║
║                                                                            ║
║  Парсер для сбора данных статей сайта 1skin.ru                           ║
║  Собирает: URL, дату публикации, дату обновления, просмотры, время чтения║
║                                                                            ║
║  ПОДДЕРЖИВАЕМЫЕ РАЗДЕЛЫ:                                                  ║
║  • /article/ - основные статьи                                            ║
║  • /want/    - статьи о желаемом (с версии 2.1)                           ║
║                                                                            ║
║  ИСПОЛЬЗОВАНИЕ:                                                            ║
║  1. Установи зависимости: pip install requests beautifulsoup4 pandas      ║
║     openpyxl lxml                                                         ║
║  2. Создай файл urls.txt рядом со скриптом                               ║
║  3. Добавь туда URL (каждый на новой строке)                             ║
║  4. Запусти: python 1skin_parser_url_article.py                          ║
║                                                                            ║
║  РЕЗУЛЬТАТ:                                                                ║
║  ✓ Файл 1skin_articles.xlsx с полной таблицей                            ║
║  ✓ Статистика в консоли                                                   ║
║  ✓ Примеры данных                                                          ║
╚════════════════════════════════════════════════════════════════════════════╝
"""

import requests
from bs4 import BeautifulSoup
import pandas as pd
import time
import re
import os
import json
from pathlib import Path


class SkinArticleScraper:
    """Парсер статей сайта 1skin.ru с поддержкой /article/ и /want/"""
    
    def __init__(self, urls_file='urls.txt'):
        """
        Инициализация парсера
        
        Args:
            urls_file (str): Имя файла с URL статей
        """
        self.base_url = "https://1skin.ru"
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        })
        self.articles = []
        script_dir = Path(__file__).parent.absolute()
        self.urls_file = script_dir / urls_file
    
    def load_urls_from_file(self):
        """Загрузить список URL статей из текстового файла"""
        try:
            if not self.urls_file.exists():
                print(f"✗ Файл не найден: {self.urls_file}")
                print(f"\n📁 Текущая рабочая директория: {os.getcwd()}")
                print(f"📁 Директория скрипта: {self.urls_file.parent}")
                print(f"\n💡 Пожалуйста:")
                print(f"  1. Создай файл 'urls.txt' в директории: {self.urls_file.parent}")
                print(f"  2. Добавь туда URL статей (каждый на новой строке)")
                print(f"  3. Запусти скрипт снова\n")
                
                print(f"📋 Файлы в директории {self.urls_file.parent}:")
                for file in self.urls_file.parent.iterdir():
                    if file.is_file():
                        print(f"   - {file.name} ({file.stat().st_size} байт)")
                
                return []
            
            with open(self.urls_file, 'r', encoding='utf-8') as f:
                urls = [line.strip() for line in f if line.strip()]
            
            print(f"✓ Загружено {len(urls)} URL из файла {self.urls_file.name}")
            return urls
        
        except Exception as e:
            print(f"✗ Ошибка при чтении файла: {e}")
            return []
    
    def get_sitemap_urls(self):
        """Получить список всех URL статей из sitemap.xml сайта"""
        try:
            sitemap_url = f"{self.base_url}/sitemap.xml"
            response = self.session.get(sitemap_url, timeout=10)
            soup = BeautifulSoup(response.content, 'xml')
            
            # Ищем URL из обоих разделов: /article/ и /want/
            urls = [loc.text for loc in soup.find_all('loc') 
                   if '/article/' in loc.text or '/want/' in loc.text]
            
            print(f"✓ Найдено {len(urls)} статей в sitemap")
            return urls
        
        except Exception as e:
            print(f"✗ Ошибка при парсинге sitemap: {e}")
            return []
    
    def extract_article_data(self, url):
        """Загрузить и спарсить данные одной статьи"""
        try:
            response = self.session.get(url, timeout=10)
            response.raise_for_status()
            soup = BeautifulSoup(response.content, 'html.parser')
            
            # Извлекаем четыре значения
            publish_date = self._get_publish_date(soup)
            update_date = self._get_update_date(soup)
            views = self._get_views_count(soup)
            read_time = self._get_read_time(soup)
            
            return {
                'URL страницы': url,
                'Дата публикации': publish_date,
                'Дата обновления': update_date,
                'Количество просмотров': views,
                'Время чтения (мин)': read_time
            }
        
        except Exception as e:
            return None
    
    def _get_publish_date(self, soup):
        """
        Извлечь дату публикации статьи
        Структура: "Дата публикации: 14.01.2019"
        """
        try:
            # МЕТОД 1: Ищем текст "Дата публикации:" в странице
            text = soup.get_text()
            match = re.search(r'Дата\s+публикации[:\s]+(\d{1,2}\.\d{1,2}\.\d{4})', text, re.IGNORECASE)
            if match:
                return match.group(1)
            
            # МЕТОД 2: Ищем через JSON-LD структурированные данные
            scripts = soup.find_all('script', {'type': 'application/ld+json'})
            for script in scripts:
                try:
                    data = json.loads(script.string)
                    if isinstance(data, dict):
                        if 'datePublished' in data:
                            date_str = data['datePublished'][:10]
                            date_obj = __import__('datetime').datetime.strptime(date_str, '%Y-%m-%d')
                            return date_obj.strftime('%d.%m.%Y')
                except:
                    pass
            
            # МЕТОД 3: Ищем в section с классом section-purple
            section = soup.find('section', class_='section-purple')
            if section:
                divs = section.find_all('div', class_='pt-3')
                if divs:
                    pt3_div = divs[0]
                    date_div = pt3_div.find('div')
                    if date_div:
                        date_text = date_div.get_text(strip=True)
                        if re.match(r'\d{1,2}\.\d{1,2}\.\d{4}', date_text):
                            return date_text
            
            # МЕТОД 4: Ищем любую дату в формате DD.MM.YYYY в начале страницы
            dates = re.findall(r'\d{1,2}\.\d{1,2}\.\d{4}', text[:2000])
            if dates:
                return dates[0]
            
            return 'Не найдена'
        
        except Exception as e:
            return 'Не найдена'
    
    def _get_update_date(self, soup):
        """
        Извлечь дату обновления статьи
        Структура: "Дата обновления: 03.06.2025"
        Это поле может быть отсутствовать на старых статьях
        """
        try:
            text = soup.get_text()
            
            # МЕТОД 1: Ищем текст "Дата обновления:" в странице
            match = re.search(r'Дата\s+обновления[:\s]+(\d{1,2}\.\d{1,2}\.\d{4})', text, re.IGNORECASE)
            if match:
                return match.group(1)
            
            # МЕТОД 2: Ищем через JSON-LD структурированные данные (dateModified)
            scripts = soup.find_all('script', {'type': 'application/ld+json'})
            for script in scripts:
                try:
                    data = json.loads(script.string)
                    if isinstance(data, dict):
                        if 'dateModified' in data:
                            date_str = data['dateModified'][:10]
                            date_obj = __import__('datetime').datetime.strptime(date_str, '%Y-%m-%d')
                            return date_obj.strftime('%d.%m.%Y')
                except:
                    pass
            
            # МЕТОД 3: Ищем в section с классом section-purple (вторая дата в pt-3)
            # Структура может быть: [дата публикации, дата обновления, время чтения]
            section = soup.find('section', class_='section-purple')
            if section:
                divs = section.find_all('div', class_='pt-3')
                if divs:
                    pt3_div = divs[0]
                    child_divs = pt3_div.find_all('div', recursive=False)
                    
                    # Если есть больше чем одна дата, вторая - это дата обновления
                    if len(child_divs) >= 2:
                        # Проверяем, это ли вторая дата
                        second_text = child_divs[1].get_text(strip=True)
                        if re.match(r'\d{1,2}\.\d{1,2}\.\d{4}', second_text):
                            return second_text
            
            return 'Не указана'
        
        except Exception as e:
            return 'Не указана'
    
    def _get_views_count(self, soup):
        """
        Извлечь количество просмотров статьи
        Структура: "Просмотрено: 135"
        """
        try:
            text = soup.get_text()
            
            # МЕТОД 1: Ищем текст "Просмотрено:" или "Просмотро"
            match = re.search(r'Просмотрено[:\s]+(\d+)', text, re.IGNORECASE)
            if match:
                return int(match.group(1))
            
            # МЕТОД 2: Ищем через meta теги
            meta_views = soup.find('meta', {'property': 'article:view_count'})
            if meta_views:
                views = meta_views.get('content', '')
                if views.isdigit():
                    return int(views)
            
            # МЕТОД 3: Ищем в разных data атрибутах
            for elem in soup.find_all(['div', 'span']):
                if elem.get('data-views'):
                    views = re.findall(r'\d+', str(elem.get('data-views')))
                    if views:
                        return int(views[0])
                
                classes = elem.get('class', [])
                if any('view' in c.lower() for c in classes if isinstance(c, str)):
                    numbers = re.findall(r'\d+', elem.get_text())
                    if numbers:
                        return int(numbers[0])
            
            # МЕТОД 4: Исходный метод - через pt-3 селектор (третья позиция)
            section = soup.find('section', class_='section-purple')
            if section:
                divs = section.find_all('div', class_='pt-3')
                if divs:
                    pt3_div = divs[0]
                    child_divs = pt3_div.find_all('div', recursive=False)
                    if len(child_divs) >= 3:  # 3-я позиция для просмотров
                        views_text = child_divs[2].get_text(strip=True)
                        numbers = re.findall(r'\d+', views_text)
                        if numbers:
                            return int(numbers[0])
            
            return 'Не указано'
        
        except Exception as e:
            return 'Не указано'
    
    def _get_read_time(self, soup):
        """
        Извлечь время чтения статьи в минутах
        Структура: "Время чтения: 18 мин"
        """
        try:
            text = soup.get_text()
            
            # МЕТОД 1: Ищем текст "Время чтения:"
            match = re.search(r'Время\s+чтения[:\s]+(\d+)\s*мин', text, re.IGNORECASE)
            if match:
                return int(match.group(1))
            
            # МЕТОД 2: Ищем через meta теги
            meta_time = soup.find('meta', {'property': 'article:reading_time'})
            if meta_time:
                time_str = meta_time.get('content', '')
                numbers = re.findall(r'\d+', time_str)
                if numbers:
                    return int(numbers[0])
            
            # МЕТОД 3: Ищем в разных data атрибутах
            for elem in soup.find_all(['div', 'span']):
                if elem.get('data-read-time'):
                    times = re.findall(r'\d+', str(elem.get('data-read-time')))
                    if times:
                        return int(times[0])
                
                classes = elem.get('class', [])
                if any('time' in c.lower() or 'read' in c.lower() for c in classes if isinstance(c, str)):
                    if 'мин' in elem.get_text().lower():
                        numbers = re.findall(r'\d+', elem.get_text())
                        if numbers:
                            return int(numbers[0])
            
            # МЕТОД 4: Исходный метод - через pt-3 селектор (четвертая позиция)
            section = soup.find('section', class_='section-purple')
            if section:
                divs = section.find_all('div', class_='pt-3')
                if divs:
                    pt3_div = divs[0]
                    child_divs = pt3_div.find_all('div', recursive=False)
                    if len(child_divs) >= 4:  # 4-я позиция для времени
                        time_text = child_divs[3].get_text(strip=True)
                        numbers = re.findall(r'\d+', time_text)
                        if numbers:
                            return int(numbers[0])
            
            return 'Не указано'
        
        except Exception as e:
            return 'Не указано'
    
    def scrape(self, use_file=True, max_articles=None):
        """Запустить процесс парсинга статей"""
        print("🔄 Запуск парсинга 1skin.ru...\n")
        
        if use_file:
            urls = self.load_urls_from_file()
        else:
            urls = self.get_sitemap_urls()
        
        if not urls:
            print("✗ Не удалось получить URL")
            return []
        
        if max_articles:
            urls = urls[:max_articles]
        
        print(f"📄 Парсинг {len(urls)} статей...\n")
        
        success_count = 0
        
        for i, url in enumerate(urls, 1):
            article_name = url.split('/')[-2]
            print(f"[{i}/{len(urls)}] {article_name}...", end=' ', flush=True)
            
            data = self.extract_article_data(url)
            
            if data:
                self.articles.append(data)
                print(f"✓ {data['Дата публикации']} | UPD: {data['Дата обновления']} | {data['Количество просмотров']} v | {data['Время чтения (мин)']} мин")
                success_count += 1
            else:
                print("✗")
            
            time.sleep(1)
        
        print(f"\n✓ Успешно собрано {success_count}/{len(urls)} статей")
        return self.articles
    
    def save_to_excel(self, filename='1skin_articles.xlsx'):
        """Сохранить спарсенные данные в файл Excel"""
        
        if not self.articles:
            print("✗ Нет данных для сохранения")
            return False
        
        try:
            df = pd.DataFrame(self.articles)
            
            # Используем текущую директорию для сохранения
            output_path = Path.cwd() / filename
            
            print(f"\n💾 Сохранение в: {output_path}")
            
            # Пытаемся сохранить с openpyxl
            try:
                with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='Статьи')
                    
                    worksheet = writer.sheets['Статьи']
                    worksheet.column_dimensions['A'].width = 50  # URL
                    worksheet.column_dimensions['B'].width = 15  # Дата публикации
                    worksheet.column_dimensions['C'].width = 15  # Дата обновления
                    worksheet.column_dimensions['D'].width = 20  # Просмотры
                    worksheet.column_dimensions['E'].width = 18  # Время чтения
            
            except ImportError:
                # Если openpyxl не установлен, используем xlsxwriter
                print("  (openpyxl не найден, использую xlsxwriter)")
                with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False, sheet_name='Статьи')
            
            except Exception as e:
                # Если оба не работают, сохраняем как CSV
                print(f"  Ошибка с Excel: {e}")
                print("  Сохраняю как CSV вместо этого...")
                csv_path = Path.cwd() / filename.replace('.xlsx', '.csv')
                df.to_csv(csv_path, index=False, encoding='utf-8-sig')
                output_path = csv_path
            
            print(f"✅ Файл сохранен: {output_path}")
            print(f"   Размер: {output_path.stat().st_size / 1024:.1f} KB")
            
            print(f"\n📊 Статистика:")
            print(f"  • Всего строк: {len(df)}")
            
            # Статистика по просмотрам
            numeric_views = pd.to_numeric(df['Количество просмотров'], errors='coerce')
            if numeric_views.notna().any():
                avg_views = numeric_views.mean()
                max_views = numeric_views.max()
                min_views = numeric_views.min()
                print(f"  • Среднее количество просмотров: {avg_views:.0f}")
                print(f"  • Макс просмотров: {max_views:.0f}")
                print(f"  • Мин просмотров: {min_views:.0f}")
            
            # Статистика по времени чтения
            numeric_time = pd.to_numeric(df['Время чтения (мин)'], errors='coerce')
            if numeric_time.notna().any():
                avg_time = numeric_time.mean()
                print(f"  • Среднее время чтения: {avg_time:.1f} мин")
            
            # Статистика по датам обновления
            updated_articles = df[df['Дата обновления'] != 'Не указана'].shape[0]
            print(f"  • Статей с датой обновления: {updated_articles}/{len(df)}")
            
            print(f"\n📋 Примеры данных (первые 3 строки):")
            print(df.head(3).to_string(index=False))
            
            return True
        
        except Exception as e:
            print(f"✗ Ошибка при сохранении файла: {e}")
            print(f"  Полная ошибка: {type(e).__name__}: {str(e)}")
            return False


# ═══════════════════════════════════════════════════════════════════════════
# ГЛАВНАЯ ПРОГРАММА
# ═══════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    print("═" * 80)
    print("1SKIN PARSER - URL ARTICLE SCRAPER v2.1")
    print("═" * 80 + "\n")
    
    scraper = SkinArticleScraper(urls_file='urls.txt')
    articles = scraper.scrape(use_file=True, max_articles=None)
    
    if articles:
        scraper.save_to_excel('1skin_articles.xlsx')
    else:
        print("✗ Не было собрано никаких данных")
    
    print("\n✅ Парсинг завершен!")
