import pandas as pd
import os
import re
import numpy as np
from datetime import datetime

# --- Вспомогательные функции ---

def create_chpu(url):
    """Преобразует полный URL в ЧПУ-адрес (относительный путь)."""
    if isinstance(url, str):
        # Удаляем протокол (http/https) и домен
        chpu = re.sub(r'https?://[^/]+', '', url)
        return chpu if chpu else '/'
    return ''

def find_files(current_dir):
    """Находит три типа файлов в директории."""
    files = os.listdir(current_dir)
    file_map = {}
    
    csv_suffixes = [
        'ключевые слова.csv', 'теги.csv', 'urls.csv', 
        '_keywords_.csv', '_tags_.csv', '_urls_.csv'
    ]
    
    for f in files:
        f_lower = f.lower()
        if f.endswith(('.xlsx', '.csv')) and not f.startswith('~$'):
            
            is_keyword = any(k in f_lower for k in ['keyword', 'ключевые слова'])
            is_tag = any(k in f_lower for k in ['tag', 'теги'])
            is_url = any(k in f_lower for k in ['url', '_urls'])

            if is_keyword and 'keywords' not in file_map:
                file_map['keywords'] = f
            elif is_tag and 'tags' not in file_map:
                file_map['tags'] = f
            elif is_url and 'urls' not in file_map:
                file_map['urls'] = f
                
    if len(file_map) < 3:
        print("⛔ Не все три типа файлов (_keywords_, _tags_, _urls_) найдены в папке. Анализ невозможен.")
        return None
        
    return file_map

def read_data_file(filepath):
    """Читает CSV или XLSX файл с умным определением разделителя."""
    if filepath.endswith('.xlsx'):
        return pd.read_excel(filepath, header=0)
    elif filepath.endswith('.csv'):
        try:
            return pd.read_csv(filepath, header=0, sep=',', encoding='utf-8')
        except:
            try:
                return pd.read_csv(filepath, header=0, sep=';', encoding='cp1251')
            except:
                return pd.read_csv(filepath, header=0, sep=',', encoding='cp1251')
    return pd.DataFrame()

# Функция create_recommendation_df() удалена

def normalize_columns(df, file_type):
    """Динамически нормализует названия столбцов."""
    
    current_cols_norm = {str(col): str(col).strip().lower().replace('.', '').replace(',', '') for col in df.columns}
    mapping = {}
    
    if file_type == 'keywords':
        mapping = {
            'ключевое слово': 'Ключевое слово', 'теги': 'Теги', 'трафик': 'Трафик_Ключ',
            'частотность': 'Частотность', 'динамика': 'Динамика_Ключ', 'url': 'URL'
        }
        # Динамическое обнаружение столбцов позиций/дат
        position_cols_map = {}
        for original_col, norm_col_name in current_cols_norm.items():
            if 'позиция' in norm_col_name or any(keyword in norm_col_name for keyword in ['дата', '2025', '2024']) and not 'средняя' in norm_col_name:
                position_cols_map[norm_col_name] = original_col 
        
        sorted_pos_cols_norm = sorted(position_cols_map.keys())
        if len(sorted_pos_cols_norm) >= 2:
            original_col_date1 = position_cols_map[sorted_pos_cols_norm[0]]
            original_col_date2 = position_cols_map[sorted_pos_cols_norm[1]]
            
            mapping[original_col_date1] = 'Позиция_Дата_1'
            mapping[original_col_date2] = 'Позиция_Дата_2'
            
    elif file_type == 'tags':
        mapping = {
            'тег': 'Тег', 'трафик': 'Трафик_Тег', 'частотность': 'Частотность_Тег',
            'видимость': 'Видимость_Тег', 'динамика': 'Динамика_Видимости_Тег', 
            'средняя позиция': 'Средняя_Позиция_Тег', 'динамика_позиция': 'Динамика_Позиции_Тег'
        }
    elif file_type == 'urls':
        mapping = {
            'url': 'URL', 'количество ключей': 'Количество_Ключей_URL', 'теги': 'Теги_URL',
            'трафик': 'Трафик_URL', 'частотность': 'Частотность_URL', 'видимость': 'Видимость_URL',
            'динамика': 'Динамика_Видимости_URL', 'средняя позиция': 'Средняя_Позиция_URL'
        }

    # Применяем сопоставление
    new_cols = {}
    for col in df.columns:
        norm_col = str(col).strip().lower().replace('.', '').replace(',', '')
        
        if col in mapping:
            new_cols[col] = mapping[col]
        else:
            for old_col_norm, new_col_name in mapping.items():
                if norm_col == old_col_norm:
                    new_cols[col] = new_col_name
                    break
    
    df = df.rename(columns=new_cols)
    
    # Добавляем недостающие столбцы (для устойчивости)
    required_cols = list(set(mapping.values()))
    for col in required_cols:
        if col not in df.columns:
            df[col] = np.nan 
            
    return df


# --- Основная функция анализа ---

def complex_seo_analysis(file_map):
    """Загружает, объединяет и анализирует все три таблицы."""
    
    print("Загрузка и динамическая нормализация данных...")
    
    df_keywords = read_data_file(file_map['keywords'])
    df_tags_agg = read_data_file(file_map['tags'])
    df_urls_agg = read_data_file(file_map['urls'])

    df_keywords = normalize_columns(df_keywords, 'keywords')
    df_tags_agg = normalize_columns(df_tags_agg, 'tags')
    df_urls_agg = normalize_columns(df_urls_agg, 'urls')

    if 'Позиция_Дата_1' not in df_keywords.columns or 'Позиция_Дата_2' not in df_keywords.columns:
        raise ValueError(f"Критическая ошибка: В файле keywords ({file_map['keywords']}) не удалось найти два столбца с позициями/датами. Проверьте заголовки.")
    
    # --- Очистка и нормализация типов ---
    df_keywords['Динамика_Ключ'] = pd.to_numeric(df_keywords['Динамика_Ключ'].astype(str), errors='coerce').fillna(0)
    df_keywords['Частотность'] = pd.to_numeric(df_keywords['Частотность'].astype(str), errors='coerce').fillna(0)
    df_keywords['Позиция_Дата_1'] = pd.to_numeric(df_keywords['Позиция_Дата_1'].astype(str), errors='coerce').fillna(1000)
    df_keywords['Позиция_Дата_2'] = pd.to_numeric(df_keywords['Позиция_Дата_2'].astype(str), errors='coerce').fillna(1000)
    
    df_urls_agg['Количество_Ключей_URL'] = pd.to_numeric(df_urls_agg['Количество_Ключей_URL'].astype(str), errors='coerce').fillna(0)
    df_urls_agg['Динамика_Видимости_URL'] = pd.to_numeric(df_urls_agg['Динамика_Видимости_URL'].astype(str), errors='coerce').fillna(0)
    df_urls_agg['Частотность_URL'] = pd.to_numeric(df_urls_agg['Частотность_URL'].astype(str), errors='coerce').fillna(0)
    df_tags_agg['Динамика_Видимости_Тег'] = pd.to_numeric(df_tags_agg['Динамика_Видимости_Тег'].astype(str), errors='coerce').fillna(0)


    # Извлекаем домен для имени выходного файла
    first_url = df_urls_agg['URL'].dropna().iloc[0] if not df_urls_agg.empty and df_urls_agg['URL'].dropna().iloc[0] else "http://example.com"
    match = re.search(r'https?://([^/]+)', first_url)
    domain = match.group(1).replace('.', '_') if match else 'unknown_site'
    
    results = {}
    
    # ----------------------------------------------------------------------
    # ОБЩАЯ АГРЕГАЦИЯ
    # ----------------------------------------------------------------------
    keywords_sum = df_keywords.groupby('URL')['Динамика_Ключ'].sum().reset_index()
    keywords_sum.rename(columns={'Динамика_Ключ': 'Суммарная_Динамика_Ключей'}, inplace=True)
    
    df_combined_url = df_urls_agg.merge(keywords_sum, on='URL', how='left')
    df_combined_url['Суммарная_Динамика_Ключей'] = df_combined_url['Суммарная_Динамика_Ключей'].fillna(0)
    
    # Расширенный датафрейм для работы с тегами
    df_url_tag_expanded = df_urls_agg.copy()
    df_url_tag_expanded['Теги_URL'] = df_url_tag_expanded['Теги_URL'].astype(str).replace('nan', '')
    df_url_tag_expanded['Тег'] = df_url_tag_expanded['Теги_URL'].str.split(r',\s*')
    df_url_tag_expanded = df_url_tag_expanded.explode('Тег')
    df_url_tag_expanded['Тег'] = df_url_tag_expanded['Тег'].str.strip()


    # ----------------------------------------------------------------------
    # АНАЛИЗ 1: Ключевой Рост 101 -> ТОП-20 (Прорыв)
    # ----------------------------------------------------------------------
    print("Проведение анализа '1. Ключевой Рост 101 -> ТОП-20'...")

    df_rise = df_keywords[
        (df_keywords['Позиция_Дата_1'] > 20) & 
        (df_keywords['Позиция_Дата_2'] <= 20)
    ].copy()

    rise_counts = df_rise.groupby('URL').agg(
        Прирост_101_ТОП20=('Ключевое слово', 'count'),
        Сумма_Частотности_Прироста=('Частотность', 'sum')
    ).reset_index()

    df_final_rise = rise_counts.merge(
        df_combined_url[['URL', 'Частотность_URL', 'Динамика_Видимости_URL']],
        on='URL',
        how='left'
    )
    
    df_final_rise['ЧПУ-адрес'] = df_final_rise['URL'].apply(create_chpu)

    results['1.Ключ_Рост_101_ТОП20'] = df_final_rise.sort_values(
        by=['Прирост_101_ТОП20', 'Сумма_Частотности_Прироста'], 
        ascending=[False, False]
    )[['URL', 'ЧПУ-адрес', 'Прирост_101_ТОП20', 'Сумма_Частотности_Прироста', 'Частотность_URL', 'Динамика_Видимости_URL']].head(20).copy()


    # ----------------------------------------------------------------------
    # АНАЛИЗ 2: Критический Спад URL (Угрозы)
    # ----------------------------------------------------------------------
    print("Проведение анализа '2. Критический Спад URL'...")
    
    freq_median = df_combined_url['Частотность_URL'].median()
    
    df_critical_fall = df_combined_url[
        (df_combined_url['Динамика_Видимости_URL'].fillna(0) < 0) & 
        (df_combined_url['Частотность_URL'] > freq_median)
    ].copy()

    df_critical_fall['Цена_Просадки'] = df_critical_fall['Частотность_URL'] * df_critical_fall['Динамика_Видимости_URL'].abs()
    
    results['2.Критич_Спад_URL'] = df_critical_fall.sort_values(
        by='Цена_Просадки', ascending=False
    )[['URL', 'Частотность_URL', 'Динамика_Видимости_URL', 'Цена_Просадки', 'Теги_URL']].head(15).copy()

    # ----------------------------------------------------------------------
    # АНАЛИЗ 3: Выявление "Тегов-сирот" (Упал URL, Вырос Тег)
    # ----------------------------------------------------------------------
    print("Проведение анализа '3. Теги-Сироты'...")
    
    df_merged_full = df_url_tag_expanded.merge(
        df_tags_agg[['Тег', 'Динамика_Видимости_Тег']], 
        on='Тег', 
        how='left'
    )
    
    df_orphans = df_merged_full[
        (df_merged_full['Динамика_Видимости_Тег'].fillna(0) > 0) & 
        (df_merged_full['Динамика_Видимости_URL'].fillna(0) < 0)
    ].drop_duplicates(subset=['URL', 'Тег']).copy()
    
    results['3.Теги_Сироты_Диагност'] = df_orphans[[
        'URL', 'Тег', 'Динамика_Видимости_URL', 'Динамика_Видимости_Тег', 'Средняя_Позиция_URL'
    ]].sort_values(by='Динамика_Видимости_Тег', ascending=False).copy()
    
    # ----------------------------------------------------------------------
    # АНАЛИЗ 4: Скрытый Потенциал (Hidden Gems)
    # ----------------------------------------------------------------------
    print("Проведение анализа '4. Скрытый Потенциал'...")
    
    freq_median_key = df_keywords['Частотность'].median()
    df_gems = df_keywords[
        (df_keywords['Частотность'] > freq_median_key) & 
        (df_keywords['Позиция_Дата_2'] <= 10) &
        (df_keywords['Динамика_Ключ'] <= 0) 
    ].copy()
    
    df_gems['Приоритет_Gems'] = df_gems['Частотность'] / (df_gems['Позиция_Дата_2'] + 1)
    
    results['4.Скрытый_Потенциал_Gems'] = df_gems.sort_values(
        by='Приоритет_Gems', ascending=False
    )[['Ключевое слово', 'Частотность', 'Позиция_Дата_2', 'Динамика_Ключ', 'Теги', 'URL']].head(20).copy()

    # ----------------------------------------------------------------------
    # АНАЛИЗ 5: Пограничные ТОП-10 (Just Missed)
    # ----------------------------------------------------------------------
    print("Проведение анализа '5. Пограничные ТОП-10'...")
    
    df_just_missed = df_keywords[
        (df_keywords['Позиция_Дата_2'] >= 11) & 
        (df_keywords['Позиция_Дата_2'] <= 20)
    ].copy()
    
    df_just_missed['Приоритет_Just_Missed'] = df_just_missed['Частотность'] / (df_just_missed['Позиция_Дата_2'] - 10)
    
    results['5.Пограничные_ТОП-10'] = df_just_missed.sort_values(
        by='Приоритет_Just_Missed', ascending=False
    )[['Ключевое слово', 'Частотность', 'Позиция_Дата_2', 'Динамика_Ключ', 'Теги', 'URL']].head(20).copy()

    # ----------------------------------------------------------------------
    # АНАЛИЗ 6: Анализ Отсутствия Соответствия (Content Gap)
    # ----------------------------------------------------------------------
    print("Проведение анализа '6. Content Gap'...")

    freq_median_tag = df_tags_agg['Частотность_Тег'].median() if not df_tags_agg.empty else 0
    df_high_freq_tags = df_tags_agg[df_tags_agg['Частотность_Тег'] >= freq_median_tag].copy()
    
    df_merged_gap = df_high_freq_tags.merge(
        df_url_tag_expanded[['Тег', 'URL']].drop_duplicates(), 
        on='Тег', 
        how='left'
    )
    
    df_gap_raw = df_merged_gap[df_merged_gap['URL'].isna()].drop_duplicates(subset=['Тег']).copy()
    
    df_gap_low_vis = df_high_freq_tags[df_high_freq_tags['Видимость_Тег'].fillna(0) < 10].copy()
    
    df_gap = pd.concat([df_gap_raw.drop(columns='URL', errors='ignore'), df_gap_low_vis]).drop_duplicates(subset=['Тег']).copy()
    
    # УДАЛЕНЫ пустые столбцы из финального вывода
    results['6.Анализ_Content_Gap'] = df_gap.sort_values(
        by=['Частотность_Тег', 'Динамика_Видимости_Тег'], 
        ascending=[False, True]
    )[['Тег', 'Частотность_Тег', 'Динамика_Видимости_Тег']].copy() 

    # ----------------------------------------------------------------------
    # АНАЛИЗ 7: Оценка эффективности многотеговых URL
    # ----------------------------------------------------------------------
    print("Проведение анализа '7. Многотеговые URL'...")
    
    df_many_tags = df_combined_url[df_combined_url['Теги_URL'].astype(str).apply(lambda x: len(x.split(',')) > 3)].copy()
    
    df_many_tags_fall = df_many_tags[df_many_tags['Динамика_Видимости_URL'].fillna(0) < 0]
    
    results['7.Многотеговые_URL_Спад'] = df_many_tags_fall.sort_values(
        by='Динамика_Видимости_URL', ascending=True
    )[['URL', 'Теги_URL', 'Количество_Ключей_URL', 'Частотность_URL', 'Динамика_Видимости_URL']].copy()
    
    # ----------------------------------------------------------------------
    # АНАЛИЗ 8: Общая Динамика URL (Приоритет на Рост)
    # ----------------------------------------------------------------------
    df_combined_url['Результат_Видимость'] = df_combined_url['Динамика_Видимости_URL'].fillna(0).apply(
        lambda x: 'Выросла' if x > 0 else ('Упала' if x < 0 else 'Без изменений')
    )
    
    results['8.Общая_Динамика_URL'] = df_combined_url.sort_values(
        by='Суммарная_Динамика_Ключей', ascending=False
    )[['URL', 'Результат_Видимость', 'Динамика_Видимости_URL', 
      'Частотность_URL', 'Трафик_URL', 'Суммарная_Динамика_Ключей', 'Теги_URL']].copy()
      
    # ----------------------------------------------------------------------
    # АНАЛИЗ 9: Таблица Рекомендаций (УДАЛЕН ПОЛНОСТЬЮ)
    # ----------------------------------------------------------------------
    # УДАЛЕН: results['9.Угрозы_и_Рекомендации'] = create_recommendation_df().copy() 
    
    # ----------------------------------------------------------------------
    # АНАЛИЗ 10: Формат для Пользователя (УБРАН ОФИЦИАЛЬНЫЙ_САЙТ)
    # ----------------------------------------------------------------------
    df_user_format = df_combined_url.copy()
    df_user_format['Ссылка_Каталог'] = df_user_format['URL'].apply(create_chpu)
    
    # УДАЛЕН столбец Официальный_Сайт
    results['10.Формат_для_Пользоват'] = df_user_format.sort_values(
        by='Динамика_Видимости_URL', ascending=False
    )[['Ссылка_Каталог', 'Динамика_Видимости_URL', 'Частотность_URL']].copy()

    # ----------------------------------------------------------------------
    # АНАЛИЗ 11: Точечные Потери (Ключ vs. Тег) 
    # ----------------------------------------------------------------------
    print("Проведение анализа '11. Точечные Потери (Ключ vs. Тег)'...")
    
    df_key_tag = df_keywords.copy()
    df_key_tag['Основной_Тег'] = df_key_tag['Теги'].astype(str).str.split(',').str[0].str.strip()
    
    df_merged_conflict = df_key_tag.merge(
        df_tags_agg[['Тег', 'Динамика_Видимости_Тег']],
        left_on='Основной_Тег',
        right_on='Тег',
        how='left'
    )
    
    df_conflict = df_merged_conflict[
        (df_merged_conflict['Динамика_Ключ'] < -30) & 
        (df_merged_conflict['Динамика_Видимости_Тег'].fillna(0) >= 0)
    ].copy()
    
    df_conflict['Приоритет_Конфликта'] = df_conflict['Частотность'] * df_conflict['Динамика_Ключ'].abs()
    
    results['11.Точечные_Потери_К_Т'] = df_conflict.sort_values(
        by='Приоритет_Конфликта', ascending=False
    )[['Ключевое слово', 'Частотность', 'Динамика_Ключ', 'Позиция_Дата_2', 
      'URL', 'Основной_Тег', 'Динамика_Видимости_Тег']].head(20).copy()

    # ----------------------------------------------------------------------
    # АНАЛИЗ 12: Быстрые Победы (Low-Hanging Fruit) 
    # ----------------------------------------------------------------------
    print("Проведение анализа '12. Быстрые Победы (Low-Hanging Fruit)'...")

    df_low_hanging = df_keywords[
        (df_keywords['Позиция_Дата_2'] > 10) & 
        (df_keywords['Позиция_Дата_2'] <= 20) &
        (df_keywords['Динамика_Ключ'] <= 5) 
    ].copy()

    df_low_hanging['Приоритет_Победы'] = df_low_hanging['Частотность'] * (21 - df_low_hanging['Позиция_Дата_2'])

    results['12.Быстрые_Победы_LHF'] = df_low_hanging.sort_values(
        by='Приоритет_Победы', ascending=False
    )[['Ключевое слово', 'Частотность', 'Позиция_Дата_2', 'Динамика_Ключ', 'URL']].head(20).copy()

    # ----------------------------------------------------------------------
    # АНАЛИЗ 13: Несбывшиеся Надежды (Gap ТОП-20) 
    # ----------------------------------------------------------------------
    print("Проведение анализа '13. Несбывшиеся Надежды (Gap ТОП-20)'...")

    high_freq_url_median = df_combined_url['Частотность_URL'].median()
    
    df_gap_top20 = df_combined_url[
        (df_combined_url['Частотность_URL'] > high_freq_url_median) & 
        (df_combined_url['Количество_Ключей_URL'] == 0) & 
        (df_combined_url['Динамика_Видимости_URL'] <= 0)
    ].copy()
    
    results['13.Gap_ТОП20_Надежды'] = df_gap_top20.sort_values(
        by='Частотность_URL', ascending=False
    )[['URL', 'Частотность_URL', 'Динамика_Видимости_URL', 'Теги_URL']].copy()

    # ----------------------------------------------------------------------
    # АНАЛИЗ 14: Технические Угрозы (Падение НЧ-Страниц) 
    # ----------------------------------------------------------------------
    print("Проведение анализа '14. Технические Угрозы (Падение НЧ-Страниц)'...")

    low_freq_url_median = df_combined_url['Частотность_URL'].median()
    
    df_technical_threats = df_combined_url[
        (df_combined_url['Частотность_URL'] <= low_freq_url_median) & 
        (df_combined_url['Динамика_Видимости_URL'] < 
         df_combined_url['Динамика_Видимости_URL'].quantile(0.25)) 
    ].copy()

    results['14.Технические_Угрозы_НЧ'] = df_technical_threats.sort_values(
        by='Динамика_Видимости_URL', ascending=True
    )[['URL', 'Частотность_URL', 'Динамика_Видимости_URL', 'Теги_URL']].head(20).copy()


    # --- Сохранение результатов ---
    output_filename = f"{domain}-complex-ra-analis_v11_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    print(f"📝 Сохранение результатов в файл: {output_filename}")
    
    try:
        with pd.ExcelWriter(output_filename, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            for sheet_name, df_result in results.items():
                # Записываем DataFrame на лист
                df_result.to_excel(writer, sheet_name=sheet_name, index=False)
                
                # Получаем объект листа (Worksheet)
                worksheet = writer.sheets[sheet_name]
                
                # Автоматическая настройка ширины столбцов
                for i, col in enumerate(df_result.columns):
                    # Находим максимальную длину контента в столбце (включая заголовок)
                    max_len = max(
                        df_result[col].astype(str).map(len).max(),
                        len(str(col))
                    ) + 2 # Добавляем небольшой запас
                    
                    # Ограничиваем ширину столбца, чтобы избежать слишком широких ячеек (например, для длинных URL)
                    final_len = min(max_len, 70) 
                    
                    # Устанавливаем ширину столбца
                    worksheet.set_column(i, i, final_len)

        print(f"🎉 Результаты для {domain} сохранены и доступны в файле: {output_filename}")
        return output_filename
    except Exception as e:
        print(f"❌ Ошибка при записи файла {output_filename}: {e}. Проверьте, что файл не открыт.")
        return None

# --- Запуск анализа ---
def run_complex_analysis():
    current_dir = os.getcwd()
    file_map = find_files(current_dir)

    if file_map is None:
        return

    print("--- ⚙️ Запуск комплексного SEO-анализа (v11) ---")
    
    try:
        output_file = complex_seo_analysis(file_map)
        if output_file:
            print(f"Файл {output_file} успешно создан.")
    except Exception as e:
        print(f"❌ Критическая ошибка при выполнении анализа: {e}")

if __name__ == "__main__":
    run_complex_analysis()