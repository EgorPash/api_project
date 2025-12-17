import requests
import pandas as pd
import json
import time
import openpyxl

# Константы
API_KEY = 'uiabkzvbelqx'
API_URL_SUBMIT = f'http://seo-utils.ru/api/submit_task/{API_KEY}/'
API_URL_RESULT = f'http://seo-utils.ru/api/get_task_result/{API_KEY}/'


# Функция для отправки задачи
def submit_task(queries, region):
    task_data = {
        'name': 'SearchEngineParser3',
        'args': {
            'search-engine': 'yandex',
            'host': 'www.yandex.ru',
            'region': region,  # Передаем регион как строку
            'queries': queries  # Список всех текстов отзывов
        },
        'opts': {
            'domains': ['napopravku.ru']
        }
    }
    response = requests.post(API_URL_SUBMIT, json=task_data)
    return response.json()


# Функция для получения результата задачи
def get_task_result(task_id, max_retries=5):
    for attempt in range(max_retries):
        time.sleep(10)
        response = requests.get(f'{API_URL_RESULT}/{task_id}/')
        result = response.json()
        if result.get('success') and result['result'].get('is_finished'):
            return result
        elif attempt == max_retries - 1:
            raise Exception(f"Задача {task_id} не завершилась после {max_retries} попыток.")
    return None


# Функция для анализа отзывов
def analyze_reviews(reviews_df, region):
    valid_reviews = reviews_df[reviews_df['text'].str.len() > 30].dropna(subset=['text'])
    queries = valid_reviews['text'].tolist()

    if not queries:
        return [], 0, 0, 0  # Нет валидных отзывов

    # Отправляем одну задачу со всеми запросами
    response = submit_task(queries, region)
    if not response.get('success'):
        raise Exception(f"Ошибка отправки задачи: {response.get('reason', 'Неизвестная ошибка')}")

    task_id = response['result']['task_id']
    result_response = get_task_result(task_id)

    if not result_response or not result_response.get('success'):
        raise Exception("Ошибка получения результата задачи.")

    data = result_response['result']['data']

    results = []
    unique_count = 0
    competitor_count = 0
    total_checked = len(data)

    for i, item in enumerate(data):
        review_text = queries[i]  # Текст отзыва
        first_result = item[0] if item else None
        if first_result:
            link = first_result['link']
            domain = link.split('/')[2] if '//' in link else link.split('/')[0]
            if 'napopravku.ru' in domain:
                unique_count += 1
                results.append((review_text, link, 'Уникальный отзыв'))
            else:
                competitor_count += 1
                results.append((review_text, link, 'Первоисточник конкурент'))
        else:
            results.append((review_text, 'Нет результатов', 'Нет данных'))

    return results, total_checked, unique_count, competitor_count


# Основной код
def main():
    # Загружаем отзывы
    reviews_df = pd.read_excel('Отзывы.xlsx')

    # Загружаем регионы и выбираем (например, по ID "1" — Москва)
    with open('regions.json', 'r', encoding='utf-8') as f:
        regions = json.load(f)

    # Выбираем регион
    selected_region = next((r for r in regions if r['props']['id'] == '1'), regions[0])['region']
    print(f"Используемый регион: {selected_region}")

    # Анализируем
    results, total_checked, unique_count, competitor_count = analyze_reviews(reviews_df, selected_region)

    # Сохраняем результаты
    results_df = pd.DataFrame(results, columns=['Отзыв', 'Ссылка', 'Тип'])
    results_df.to_excel('Итоги.xlsx', index=False)

    # Выводим статистику
    print(f'Всего проверено: {total_checked}')
    print(f'Уникальных отзывов: {unique_count}')
    print(f'Первоисточник конкурент: {competitor_count}')


if __name__ == '__main__':
    main()
