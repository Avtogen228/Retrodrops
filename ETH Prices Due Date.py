import requests
from datetime import datetime, timedelta
import pandas as pd

def get_eth_price_at_date(date_str):
    """
    Получает цену ETH (пара ETHUSDT) на заданную дату в формате ДД.ММ.ГГГГ.
    
    :param date_str: Дата в формате ДД.М.ММ.ГГГГ
    :return: Цена закрытия ETHUSDT на заданную дату или None, если данные недоступны
    """
    try:
        # Преобразуем строку даты в datetime-объект
        date_obj = datetime.strptime(date_str, "%d.%m.%Y")
        
        # Рассчитаем временные метки начала и конца дня в миллисекундах
        start_timestamp = int(date_obj.timestamp() * 1000)
        end_timestamp = int((date_obj + timedelta(days=1)).timestamp() * 1000)
        
        # Параметры API
        base_url = "https://api.bybit.com/v5/market/kline"
        params = {
            "category": "spot",
            "symbol": "ETHUSDT",
            "interval": "D",
            "start": start_timestamp,
            "end": end_timestamp
        }

        # Запрос к API
        response = requests.get(base_url, params=params)

        if response.status_code != 200:
            print(f"Ошибка: Невозможно получить цену ETH с Bybit. Код статуса: {response.status_code}")
            return None

        # Парсим данные
        data = response.json()
        candles = data.get("result", {}).get("list", [])

        if candles:
            return float(candles[0][4])  # Цена закрытия свечи
        else:
            return None

    except ValueError:
        print("Ошибка: Неверный формат даты. Используйте ДД.ММ.ГГГГ.")
        return None

def process_csv_to_excel(input_csv, output_excel):
    """
    Читает даты из .csv файла, получает цены ETH для каждой даты
    и записывает результаты в .xlsx файл в таком же формате.
    
    :param input_csv: Путь к входному .csv файлу с датами
    :param output_excel: Путь к выходному .xlsx файлу с ценами
    """
    try:
        # Чтение CSV файла
        dates_df = pd.read_csv(input_csv, header=None)

        # Функция для получения цены по каждой дате
        def get_price(date):
            if pd.isna(date):
                return None
            return get_eth_price_at_date(date)

        # Применяем функцию ко всем элементам DataFrame
        prices_df = dates_df.apply(lambda row: row.map(get_price), axis=1)

        # Сохраняем результат в Excel
        prices_df.to_excel(output_excel, index=False, header=False, engine='openpyxl')
        print(f"Результаты успешно сохранены в файл: {output_excel}")

    except Exception as e:
        print(f"Ошибка обработки файла: {e}")

# Пример использования
if __name__ == "__main__":
    input_csv_path = input("Введите путь к входному .csv файлу: ")
    output_excel_path = input("Введите путь к выходному .xlsx файлу: ")
    process_csv_to_excel(input_csv_path, output_excel_path)
