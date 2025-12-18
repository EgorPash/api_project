import requests
import pandas as pd
import json
import time
import openpyxl

# Константы
API_KEY = 'uiabkzvbelqx'
API_URL_SUBMIT = f'http://seo-utils.ru/api/submit_task/{API_KEY}/'
API_URL_RESULT = f'http://seo-utils.ru/api/get_task_result/{API_KEY}/'

# Жестко заданные регион и city ID для Москвы (номер региона: 263, city ID: 263)
REGION_ID = '213'  # Номер региона для Yandex (Москва)
CITY_ID = '263'    # Соответствующий city ID

# Основной код
def main():
    try:
        # Загружаем отзывы из CSV
        reviews_df = pd.read_csv('Отзывы.csv', sep=';', header=0, encoding='utf-8')
        print("CSV загружен успешно.")
    except pd.errors.ParserError as e:
        print(f"Ошибка парсинга CSV: {e}. Проверьте разделитель и кодировку.")
        return
    except FileNotFoundError:
        print("Файл 'Отзывы.csv' не найден.")
        return
    except Exception as e:
        print(f"Ошибка загрузки CSV: {e}")
        return

    # Фильтруем отзывы (длина >30, без NaN)
    valid_reviews = reviews_df[reviews_df['text'].str.len() > 30].dropna(subset=['text'])
    queries = valid_reviews['text'].tolist()

    if not queries:
        print("Нет валидных отзывов.")
        return

    # Отправляем задачу
    task_data = {
        'name': 'SearchEngineParser3',
        'args': {
            'search-engine': 'yandex',
            'host': 'www.yandex.ru',
            'region': REGION_ID,  # Используем ID региона
            'city': CITY_ID,      # Добавляем city ID для точности
            'queries': queries
        },
        'opts': {}
    }
    response = requests.post(API_URL_SUBMIT, json=task_data)
    try:
        response_data = response.json()
    except json.JSONDecodeError:
        print(f"Ошибка декодирования JSON: {response.text}")
        return

    if not response_data.get('success'):
        print(f"Ошибка отправки: {response_data.get('reason')}")
        return

    task_id = response_data['result']['task_id']

    # Ждем результат
    for attempt in range(5):
        time.sleep(30)
        result_response = requests.get(f'{API_URL_RESULT}/{task_id}/')
        try:
            result = result_response.json()
        except json.JSONDecodeError:
            print(f"Попытка {attempt + 1}: Ошибка JSON, повтор.")
            continue
        if result.get('success') and result['result'].get('is_finished'):
            break
    else:
        print("Задача не завершилась.")
        return

    data = result['result']['data']

    # Анализируем результаты
    results = []
    unique_count = 0
    competitor_count = 0
    total_checked = len(data)

    for i, item in enumerate(data):
        review_text = queries[i]
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

    # Сохраняем
    results_df = pd.DataFrame(results, columns=['Отзыв', 'Ссылка', 'Тип'])
    results_df.to_csv('Итоги.csv', index=False, encoding='utf-8')

    # Статистика
    print(f'Всего проверено: {total_checked}')
    print(f'Уникальных отзывов: {unique_count}')
    print(f'Первоисточник конкурент: {competitor_count}')

if __name__ == '__main__':
    main()
