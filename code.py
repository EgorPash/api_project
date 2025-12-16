import requests
import json
import pandas as pd
import openpyxl

# Конфигурация
API_KEY = 'uiabkzvbelqx'
API_URL = 'https://seo-utils.ru/task/search_engine_parser_3/'
HEADERS = {'Authorization': f'Bearer {API_KEY}'}

# Загрузка регионов из файла regions.json
with open('regions.json', 'r', encoding='utf-8') as f:
    regions = json.load(f)

# Пример выбора региона
region_name = "Москва"  # Укажите нужный регион
selected_region = None

for region in regions:
    if region_name in region['region']:
        selected_region = region['region']
        break

if selected_region:
    print(f"Выбранный регион: {selected_region}")
else:
    print("Регион не найден.")

# Функция для проверки отзыва
def check_review(review_text, region='Москва'):
    params = {
        'query': review_text,
        'search_engine': 'yandex',
        'region': region,
        'count': 1  # Получаем только первую позицию
    }
    response = requests.get(API_URL, headers=HEADERS, params=params)
    return response.json()


# Чтение отзывов из файла Excel
reviews_df = pd.read_excel('Отзывы.xlsx')  # Читаем файл Excel
reviews_df = reviews_df[reviews_df['Отзыв'].str.len() > 30]  # Фильтрация отзывов по длине

# Список для хранения результатов
results = []

# Проверка отзывов
for index, row in reviews_df.iterrows():
    review_text = row['Отзыв']  # Предполагается, что колонка с отзывами называется 'Отзыв'
    result = check_review(review_text)

    # Проверка на наличие результатов
    if result and 'results' in result and len(result['results']) > 0:
        first_result = result['results'][0]
        site = first_result['site']
        link = first_result['link']

        if 'napopravku' in site.lower():
            results.append({'Отзыв': review_text, 'Ссылка': link, 'Статус': 'Уникальный отзыв'})
        else:
            results.append({'Отзыв': review_text, 'Ссылка': link, 'Статус': 'Первоисточник конкурент'})
    else:
        results.append({'Отзыв': review_text, 'Ссылка': None, 'Статус': 'Нет результатов'})

# Создание DataFrame с результатами
results_df = pd.DataFrame(results)

# Сохранение результатов в файл Excel
results_df.to_excel('Результаты.xlsx', index=False)

# Подсчет статистики
total_reviews = len(results_df)
unique_reviews = len(results_df[results_df['Статус'] == 'Уникальный отзыв'])
competitor_sources = len(results_df[results_df['Статус'] == 'Первоисточник конкурент'])

# Вывод статистики
print(f'Всего проверено: {total_reviews}')
print(f'Уникальных отзывов: {unique_reviews}')
print(f'Первоисточник конкурент: {competitor_sources}')
