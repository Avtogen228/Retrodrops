import requests
import csv
import xlrd
from datetime import datetime, timezone, timedelta

# Bybit API endpoint
base_url = "https://api.bybit.com/v5/market/kline"

def get_listing_date_bybit(symbol):
    """
    Быстрый поиск даты листинга монеты с использованием бинарного поиска по времени.
    """
    today = datetime.now(timezone.utc)
    end_time = int(today.timestamp() * 1000)  # Конец (текущее время)
    start_time = 0  # Начало (эпоха Unix)

    interval = "D"  # Используем дневные свечи для точности
    print(f"Начинаем поиск даты листинга для {symbol}...")

    while start_time <= end_time:
        mid_time = (start_time + end_time) // 2  # Середина диапазона
        params = {
            "category": "spot",
            "symbol": symbol,
            "interval": interval,
            "start": mid_time,
            "end": end_time
        }

        response = requests.get(base_url, params=params)
        if response.status_code != 200:
            print(f"Ошибка: Невозможно получить данные с Bybit. Код статуса: {response.status_code}")
            return None, None

        data = response.json()
        candles = data.get("result", {}).get("list", [])

        if candles:
            # Если свечи найдены, сужаем диапазон к более раннему времени
            end_time = int(candles[-1][0]) - 1  # Последняя свеча
            listing_timestamp = int(candles[-1][0])
        else:
            # Если свечей нет, сужаем диапазон к более позднему времени
            start_time = mid_time + 1

    # Возвращаем найденную дату листинга
    if 'listing_timestamp' in locals():
        listing_date = datetime.fromtimestamp(listing_timestamp / 1000, tz=timezone.utc)
        return listing_date, listing_timestamp
    print("Дата листинга не найдена.")
    return None, None

def get_listing_price(symbol, listing_timestamp):
    """
    Получает цену монеты на листинге, используя закрытие дневной свечи в день листинга.
    """
    interval = "D"  # Используем дневные свечи
    start_time = listing_timestamp
    end_time = listing_timestamp + 86400000  # Конец того же дня (1 день в миллисекундах)

    params = {
        "category": "spot",
        "symbol": symbol,
        "interval": interval,
        "start": start_time,
        "end": end_time
    }

    response = requests.get(base_url, params=params)
    if response.status_code != 200:
        print(f"Ошибка: Невозможно получить данные о свечах с Bybit. Код статуса: {response.status_code}")
        return None

    data = response.json()
    candles = data.get("result", {}).get("list", [])

    if not candles:
        print("Данные дневной свечи отсутствуют для определения цены на листинге.")
        return None

    # Цена закрытия первой дневной свечи
    listing_price = float(candles[0][4])  # Индекс 0 — первая дневная свеча
    return listing_price

def get_price_after_days(symbol, listing_timestamp, days):
    """
    Получает цену монеты спустя определённое количество дней после даты листинга.
    """
    interval = "D"  # Дневные свечи
    start_time = listing_timestamp + days * 86400000  # Начало через `days` дней
    end_time = start_time + 86400000  # Конец дня

    params = {
        "category": "spot",
        "symbol": symbol,
        "interval": interval,
        "start": start_time,
        "end": end_time
    }

    response = requests.get(base_url, params=params)
    if response.status_code != 200:
        print(f"Ошибка: Невозможно получить данные о свечах с Bybit. Код статуса: {response.status_code}")
        return "-"  # Если запрос не удался, вернуть прочерк

    data = response.json()
    candles = data.get("result", {}).get("list", [])

    if not candles:
        print(f"Данные свечей отсутствуют для {symbol} спустя {days} дней.")
        return "-"  # Если свечи отсутствуют, вернуть прочерк

    # Цена закрытия дневной свечи
    return float(candles[0][4])


def get_current_price(symbol):
    """
    Получает текущую цену монеты с помощью эндпоинта /tickers.
    """
    params = {
        "category": "spot",
        "symbol": symbol
    }

    response = requests.get(f"https://api.bybit.com/v5/market/tickers", params=params)
    if response.status_code != 200:
        print(f"Ошибка: Невозможно получить текущую цену с Bybit. Код статуса: {response.status_code}")
        return None

    data = response.json()
    result = data.get("result", {}).get("list", [])
    if result:
        price = result[0].get("lastPrice")
        return float(price) if price else None

    print(f"Ошибка: Текущая цена для пары {symbol} отсутствует.")
    return None

def calculate_change(current, base):
    """
    Вычисляет изменение в процентах относительно базовой цены.
    """
    if current == "-" or base == "-" or current is None or base is None:
        return "-"  # Если данные отсутствуют, вернуть прочерк
    change = (current - base) / base * 100
    return change



def format_price_with_change(price, change):
    """
    Форматирует цену с отображением изменения.
    """
    if price == "-" or change == "-":
        return "-"  # Если данные отсутствуют, вернуть прочерк
    change = round(change, 0)
    if price is None:
        return "-"
    if price < 0.01:
        return f"{price:.5f} ({change:+.0f}%)"
    if price < 0.1:
        return f"{price:.4f} ({change:+.0f}%)"
    if price < 1:
        return f"{price:.3f} ({change:+.0f}%)"
    if price < 10:
        return f"{price:.2f} ({change:+.0f}%)"
    if price < 100:
        return f"{price:.1f} ({change:+.0f}%)"
    if price > 100:
        return f"{price:.0f} ({change:+.0f}%)"

def get_eth_price_at_time(timestamp):
    """
    Получает цену ETH (пара ETHUSDT) на заданный момент времени.
    """
    params = {
        "category": "spot",
        "symbol": "ETHUSDT",
        "interval": "1d",
        "start": timestamp,
        "end": timestamp + 86400000  # Один день в миллисекундах
    }

    response = requests.get(base_url, params=params)
    if response.status_code != 200:
        print(f"Ошибка: Невозможно получить цену ETH с Bybit. Код статуса: {response.status_code}")
        return None

    data = response.json()
    candles = data.get("result", {}).get("list", [])
    if candles:
        return int(candles[0][4])  # Цена закрытия свечи
    return None


def get_peak_and_lowest_price(symbol, listing_timestamp, days=180):
    """
    Определяет пиковую и наименьшую цены монеты в диапазоне с даты листинга до `days` дней.
    Также возвращает даты, когда эти цены были зафиксированы.
    """
    interval = "D"  # Дневные свечи
    start_time = listing_timestamp
    end_time = listing_timestamp + days * 86400000  # `days` дней в миллисекундах

    params = {
        "category": "spot",
        "symbol": symbol,
        "interval": interval,
        "start": start_time,
        "end": end_time
    }

    response = requests.get(base_url, params=params)
    if response.status_code != 200:
        print(f"Ошибка: Невозможно получить данные свечей с Bybit. Код статуса: {response.status_code}")
        return None, None, None, None

    data = response.json()
    candles = data.get("result", {}).get("list", [])

    if not candles:
        print("Данные дневных свечей отсутствуют для анализа диапазона цен.")
        return None, None, None, None

    # Ищем пиковую и наименьшую цены
    peak_price = max(candles, key=lambda x: float(x[4]))  # Цена закрытия
    lowest_price = min(candles, key=lambda x: float(x[4]))

    # Извлекаем значения цен и дат
    peak_price_value = float(peak_price[4])
    peak_price_date = datetime.fromtimestamp(int(peak_price[0]) / 1000, tz=timezone.utc)

    lowest_price_value = float(lowest_price[4])
    lowest_price_date = datetime.fromtimestamp(int(lowest_price[0]) / 1000, tz=timezone.utc)

    return peak_price_value, peak_price_date, lowest_price_value, lowest_price_date


def get_eth_peak_and_low_on_date(eth_symbol, target_date_timestamp):
    """
    Получает пиковую и минимальную цены ETH в указанный день.

    :param eth_symbol: Символ пары ETH (например, "ETHUSDT").
    :param target_date_timestamp: Временная метка начала дня в миллисекундах.
    :return: Пиковая и минимальная цены ETH.
    """
    interval = "D"  # Используем 15-минутные свечи
    end_time = target_date_timestamp + 86400000  # Конец дня (1 день в миллисекундах)

    params = {
        "category": "spot",
        "symbol": eth_symbol,
        "interval": interval,
        "start": target_date_timestamp,
        "end": end_time
    }

    response = requests.get(base_url, params=params)
    if response.status_code != 200:
        print(f"Ошибка: Невозможно получить данные свечей ETH с Bybit. Код статуса: {response.status_code}")
        return None, None

    data = response.json()
    candles = data.get("result", {}).get("list", [])

    if not candles:
        print("Данные 15-минутных свечей для ETH отсутствуют.")
        return None, None

    # Ищем пиковую и минимальную цены ETH
    high_prices = [float(candle[3]) for candle in candles]  # Максимальные цены
    low_prices = [float(candle[4]) for candle in candles]  # Минимальные цены

    peak_eth_price = max(high_prices) if high_prices else None
    lowest_eth_price = min(low_prices) if low_prices else None

    return int(peak_eth_price), int(lowest_eth_price)

def calculate_ratio(eth_change, token_change):
    """
    Вычисляет отношение первого числа ко второму по указанному принципу.
    
    1. Делит числа на 100.
    2. Если число положительное, прибавляет 1.
    3. Если число отрицательное, вычитает из 1.
    4. Делит первое число на второе.
    
    :param eth_change: Первое число (проценты).
    :param token_change: Второе число (проценты).
    :return: Отношение двух чисел или "-" при ошибке.
    """
    try:
        # Деление на 100
        factor1 = eth_change / 100
        factor2 = token_change / 100

        # Преобразование по правилам
        if factor1 >= 0:
            factor1 = 1 + factor1
        else:
            factor1 = 1 - abs(factor1)

        if factor2 >= 0:
            factor2 = 1 + factor2
        else:
            factor2 = 1 - abs(factor2)

        # Деление и возврат результата
        return round(factor1 / factor2, 2)
    except (TypeError, ZeroDivisionError):
        return "-"

def read_symbols_from_xls(file_path):
    """
    Читает список монет из Excel файла.
    
    :param file_path: Путь к .xls файлу
    :return: Список символов монет
    """
    symbols = []
    workbook = xlrd.open_workbook(file_path)
    sheet = workbook.sheet_by_index(0)
    for row_idx in range(sheet.nrows):
        symbol = sheet.cell_value(row_idx, 0).strip().upper()
        if symbol.endswith("USDT"):
            symbols.append(symbol)
    return symbols

def save_results_to_csv(data, output_file):
    """
    Сохраняет результаты в CSV файл.
    
    :param data: Данные для записи (список списков)
    :param output_file: Путь к выходному CSV файлу
    """
    headers = [
        "Монета", "Дата Листинга", "Цена Листинга",
        "Цена спустя 90 дней", "Цена спустя 180 дней", "Текущая цена",
        "Цена ETH на листинге", "Цена ETH спустя 90 дней", "Цена ETH спустя 180 дней", "Текущая цена ETH",
        "Пиковая цена", "Минимальная цена", "ETH на пике монеты", "ETH на минимуме монеты",
        "Отношение на пике", "Отношение на минимуме", "Отношение текущей цены", 
        "Отношение спустя 90 дней", "Отношение спустя 180 дней"
    ]
    
    with open(output_file, mode="w", encoding="utf-8", newline="") as file:
        writer = csv.writer(file)
        writer.writerow(headers)
        writer.writerows(data)

def main():
    input_file = "inputs Bybit.xls"  # Входной Excel файл
    output_file = "output Bybit.csv"  # Выходной CSV файл

    # Чтение символов из Excel файла
    symbols = read_symbols_from_xls(input_file)
    if not symbols:
        print("Файл ввода пуст или не содержит символов.")
        return

    all_results = []
    for symbol in symbols:
        print(f"Обработка {symbol}...")
        result = process_symbol(symbol)
        all_results.append(result)

    # Сохранение результатов в CSV файл
    save_results_to_csv(all_results, output_file)
    print(f"Результаты успешно сохранены в {output_file}")

def process_symbol(symbol):
    """
    Обрабатывает один символ и возвращает данные для вывода в таблицу.
    """
    result = get_listing_date_bybit(symbol)
    if not result:
        return [symbol, "-", "-", "-", "-", "-", "-", "-", "-", "-", "-", "-", "-", "-", "-", "-", "-", "-"]
    
    listing_date, listing_timestamp = result
    price_listing = get_listing_price(symbol, listing_timestamp)
    price_90_days = get_price_after_days(symbol, listing_timestamp, 90)
    price_180_days = get_price_after_days(symbol, listing_timestamp, 180)
    current_price = get_current_price(symbol)

    eth_price_listing = get_price_after_days("ETHUSDT", listing_timestamp, 0)
    eth_price_90_days = get_price_after_days("ETHUSDT", listing_timestamp, 90)
    eth_price_180_days = get_price_after_days("ETHUSDT", listing_timestamp, 180)
    eth_current_price = get_current_price("ETHUSDT")

    peak_price, peak_date, lowest_price, lowest_date = get_peak_and_lowest_price(symbol, listing_timestamp)
    eth_peak_price, _ = get_eth_peak_and_low_on_date("ETHUSDT", int(peak_date.timestamp() * 1000)) if peak_date else (None, None)
    eth_low_price, _ = get_eth_peak_and_low_on_date("ETHUSDT", int(lowest_date.timestamp() * 1000)) if lowest_date else (None, None)

    ratio_at_peak = calculate_ratio(calculate_change(eth_peak_price, eth_price_listing), calculate_change(peak_price, price_listing))
    ratio_at_low = calculate_ratio(calculate_change(eth_low_price, eth_price_listing), calculate_change(lowest_price, price_listing))
    ratio_current = calculate_ratio(calculate_change(eth_current_price, eth_price_listing), calculate_change(current_price, price_listing))
    ratio_90_days = calculate_ratio(calculate_change(eth_price_90_days, eth_price_listing), calculate_change(price_90_days, price_listing))
    ratio_180_days = calculate_ratio(calculate_change(eth_price_180_days, eth_price_listing), calculate_change(price_180_days, price_listing))

    formatted_listing_date = listing_date.strftime('%d.%m.%Y') if listing_date else "-"
    formatted_price_listing = price_listing
    formatted_price_90_days = format_price_with_change(price_90_days, calculate_change(price_90_days, price_listing))
    formatted_price_180_days = format_price_with_change(price_180_days, calculate_change(price_180_days, price_listing))
    formatted_current_price = format_price_with_change(current_price, calculate_change(current_price, price_listing))
    formatted_eth_price_listing = int(eth_price_listing)
    formatted_eth_price_90_days = format_price_with_change(eth_price_90_days, calculate_change(eth_price_90_days, eth_price_listing))
    formatted_eth_price_180_days = format_price_with_change(eth_price_180_days, calculate_change(eth_price_180_days, eth_price_listing))
    formatted_eth_current_price = format_price_with_change(eth_current_price, calculate_change(eth_current_price, eth_price_listing))
    formatted_peak_price = format_price_with_change(peak_price, calculate_change(peak_price, price_listing))
    formatted_lowest_price = format_price_with_change(lowest_price, calculate_change(lowest_price, price_listing))

    return [
        symbol, formatted_listing_date, formatted_price_listing, formatted_price_90_days,
        formatted_price_180_days, formatted_current_price, formatted_eth_price_listing,
        formatted_eth_price_90_days, formatted_eth_price_180_days, formatted_eth_current_price,
        formatted_peak_price, formatted_lowest_price, eth_peak_price, eth_low_price,
        ratio_at_peak, ratio_at_low, ratio_current, ratio_90_days, ratio_180_days
    ]

if __name__ == "__main__":
    main()







