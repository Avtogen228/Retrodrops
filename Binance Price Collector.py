import xlwt
import xlrd
import csv
import time
from datetime import datetime, timedelta, timezone
import requests

# Binance API endpoint for klines (candlesticks data)
url = "https://api.binance.com/api/v3/klines"

# Function to fetch the first available data (listing date and price) from the Kline data
def fetch_listing_date(symbol):
    params = {
        "symbol": symbol,
        "interval": "1d",  # Daily candles
        "startTime": 0  # Fetch from the very beginning
    }
    response = requests.get(url, params=params)
    
    if response.status_code != 200:
        print(f"Error: Unable to fetch Kline data for {symbol}. Status code: {response.status_code}")
        return None, None
    
    data = response.json()
    
    if not data:
        print(f"Error: No historical data found for {symbol}. It may not be listed or available.")
        return None, None

    try:
        listing_date = datetime.fromtimestamp(data[0][0] / 1000, tz=timezone.utc)
        listing_price = round(float(data[0][4]), 4)
        return listing_date, listing_price
    except (IndexError, ValueError) as e:
        print(f"Error processing the response data: {e}")
        return None, None

# Function to fetch price at a specific time
def fetch_price(symbol, start_time, end_time):
    params = {
        "symbol": symbol,
        "interval": "1d",
        "startTime": start_time,
        "endTime": end_time
    }
    response = requests.get(url, params=params)
    data = response.json()
    if data:
        return round(float(data[0][4]), 4)
    return None

# Function to fetch the peak price within a specific time range
def fetch_peak_price(symbol, start_time, end_time):
    params = {
        "symbol": symbol,
        "interval": "1d",
        "startTime": start_time,
        "endTime": end_time
    }
    response = requests.get(url, params=params)
    data = response.json()
    if data:
        high_prices = [float(kline[2]) for kline in data]
        return round(max(high_prices), 4) if high_prices else None
    return None

# Function to fetch the lowest price within a specific time range
def fetch_lowest_price(symbol, start_time, end_time):
    params = {
        "symbol": symbol,
        "interval": "1d",
        "startTime": start_time,
        "endTime": end_time
    }
    response = requests.get(url, params=params)
    data = response.json()
    if data:
        low_prices = [float(kline[3]) for kline in data]
        return round(min(low_prices), 4) if low_prices else None
    return None

# Function to fetch the current price of the symbol
def fetch_current_price(symbol):
    ticker_url = "https://api.binance.com/api/v3/ticker/price"
    response = requests.get(ticker_url, params={"symbol": symbol})
    if response.status_code == 200:
        price = float(response.json()["price"])
        return round(price, 4)
    else:
        print(f"Error: Unable to fetch current price for {symbol}. Status code: {response.status_code}")
        return None

# Function to calculate relative change between ETH and coin
def calculate_relative_change(coin_change_percent, eth_change_percent):
    try:
        coin_factor = 1 + coin_change_percent / 100
        eth_factor = 1 + eth_change_percent / 100
        if coin_factor == 0:
            raise ValueError("Coin factor cannot be zero.")
        return round(eth_factor / coin_factor, 2)
    except Exception as e:
        print(f"Error calculating relative change: {e}")
        return None

# Function to fetch the timestamp of peak or lowest price
def fetch_timestamp_of_extreme(symbol, start_time, end_time, extreme_type="peak"):
    params = {
        "symbol": symbol,
        "interval": "1d",
        "startTime": start_time,
        "endTime": end_time
    }
    response = requests.get(url, params=params)
    data = response.json()
    if data:
        if extreme_type == "peak":
            extreme_value = max(data, key=lambda kline: float(kline[2]))
        else:
            extreme_value = min(data, key=lambda kline: float(kline[3]))
        return extreme_value[0]  # Return timestamp
    return None

def save_to_excel(data, filename="output.xls"):
    """
    Сохраняет данные в Excel файл.

    :param data: Список строк (каждая строка - список значений для таблицы)
    :param filename: Имя файла для сохранения
    """
    wb = xlwt.Workbook()
    sheet = wb.add_sheet("Data")

    # Заголовки
    headers = [
        "Symbol", "Listing Date", "Listing Price",
        "Price After 90 Days", "Price After 180 Days", "Current Price",
        "ETH Listing Price", "ETH Price After 90 Days", "ETH Price After 180 Days", "ETH Current Price",
        "Peak Price", "Lowest Price", "Peak-to-ETH Ratio", "Lowest-to-ETH Ratio",
        "Rel Change Current", "Rel Change 90 Days", "Rel Change 180 Days"
    ]

    # Заполняем заголовки
    for col_num, header in enumerate(headers):
        sheet.write(0, col_num, header)

    # Заполняем данные
    for row_num, row in enumerate(data, start=1):
        for col_num, value in enumerate(row):
            sheet.write(row_num, col_num, value)

    # Сохраняем файл
    wb.save(filename)
    print(f"Данные успешно сохранены в {filename}")


def read_symbols_from_csv(filename):
    """
    Читает символы из CSV-файла.

    :param filename: Имя CSV-файла
    :return: Список символов
    """
    symbols = []
    try:
        with open(filename, newline='', encoding='utf-8') as csvfile:
            reader = csv.reader(csvfile)
            for row in reader:
                if row:  # Пропускаем пустые строки
                    symbols.append(row[0].strip())
    except FileNotFoundError:
        print(f"Файл {filename} не найден.")
    except Exception as e:
        print(f"Ошибка при чтении файла: {e}")
    return symbols



def get_ticker_data(symbol):
    # Удаляем "USDT" из названия монеты для отображения
    base_symbol = symbol.replace("USDT", "")

    try:
        listing_date, listing_price = fetch_listing_date(symbol)
        if not listing_date:
            raise ValueError(f"No listing data found for {symbol}.")

        ninety_days_later = listing_date + timedelta(days=90)
        one_eighty_days_later = listing_date + timedelta(days=180)
        start_time_for_highs_dips = listing_date + timedelta(minutes=15)

        price_90_days = fetch_price(symbol, int(ninety_days_later.timestamp() * 1000), int(ninety_days_later.timestamp() * 1000 + 86400000))
        price_180_days = fetch_price(symbol, int(one_eighty_days_later.timestamp() * 1000), int(one_eighty_days_later.timestamp() * 1000 + 86400000))
        current_price = fetch_current_price(symbol)

        peak_price_180 = fetch_peak_price(symbol, int(start_time_for_highs_dips.timestamp() * 1000), int(one_eighty_days_later.timestamp() * 1000))
        lowest_price_180 = fetch_lowest_price(symbol, int(start_time_for_highs_dips.timestamp() * 1000), int(one_eighty_days_later.timestamp() * 1000))

        peak_timestamp = fetch_timestamp_of_extreme(symbol, int(start_time_for_highs_dips.timestamp() * 1000), int(one_eighty_days_later.timestamp() * 1000), "peak")
        lowest_timestamp = fetch_timestamp_of_extreme(symbol, int(start_time_for_highs_dips.timestamp() * 1000), int(one_eighty_days_later.timestamp() * 1000), "low")

        eth_symbol = "ETHUSDT"
        eth_listing_price = fetch_price(eth_symbol, int(listing_date.timestamp() * 1000), int(listing_date.timestamp() * 1000 + 86400000))
        eth_price_90_days = fetch_price(eth_symbol, int(ninety_days_later.timestamp() * 1000), int(ninety_days_later.timestamp() * 1000 + 86400000))
        eth_price_180_days = fetch_price(eth_symbol, int(one_eighty_days_later.timestamp() * 1000), int(one_eighty_days_later.timestamp() * 1000 + 86400000))
        eth_current_price = fetch_current_price(eth_symbol)

        eth_price_at_peak = fetch_price(eth_symbol, peak_timestamp, peak_timestamp + 86400000) if peak_timestamp else None
        eth_price_at_lowest = fetch_price(eth_symbol, lowest_timestamp, lowest_timestamp + 86400000) if lowest_timestamp else None

        def calculate_change(current, base):
            if current is None or base is None:
                return "-"
            return round((current - base) / base * 100, 2)

        def format_price_with_change(price, change):
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
            return f"{price:.2f} ({change:+.0f}%)" if change != "-" else f"{price:.2f}"

        rel_change_current = calculate_relative_change(calculate_change(current_price, listing_price), calculate_change(eth_current_price, eth_listing_price))
        rel_change_90 = calculate_relative_change(calculate_change(price_90_days, listing_price), calculate_change(eth_price_90_days, eth_listing_price))
        rel_change_180 = calculate_relative_change(calculate_change(price_180_days, listing_price), calculate_change(eth_price_180_days, eth_listing_price))

        peak_to_eth_ratio = calculate_relative_change(calculate_change(peak_price_180, listing_price), calculate_change(eth_price_at_peak, eth_listing_price)) if peak_price_180 and eth_price_at_peak else "-"
        lowest_to_eth_ratio = calculate_relative_change(calculate_change(lowest_price_180, listing_price), calculate_change(eth_price_at_lowest, eth_listing_price)) if lowest_price_180 and eth_price_at_lowest else "-"

        # Формируем строку для таблицы
        row = [
            base_symbol,  # Symbol
            listing_date.strftime('%d.%m.%y'),  # Listing Date
            format_price_with_change(listing_price, 0),  # Listing Price (no change)
            format_price_with_change(price_90_days, calculate_change(price_90_days, listing_price)),  # Price After 90 Days
            format_price_with_change(price_180_days, calculate_change(price_180_days, listing_price)),  # Price After 180 Days
            format_price_with_change(current_price, calculate_change(current_price, listing_price)),  # Current Price
            format_price_with_change(eth_listing_price, 0),  # ETH Listing Price
            format_price_with_change(eth_price_90_days, calculate_change(eth_price_90_days, eth_listing_price)),  # ETH Price After 90 Days
            format_price_with_change(eth_price_180_days, calculate_change(eth_price_180_days, eth_listing_price)),  # ETH Price After 180 Days
            format_price_with_change(eth_current_price, calculate_change(eth_current_price, eth_listing_price)),  # ETH Current Price
            format_price_with_change(peak_price_180, calculate_change(peak_price_180, listing_price)),  # Peak Price
            format_price_with_change(lowest_price_180, calculate_change(lowest_price_180, listing_price)),  # Lowest Price
            peak_to_eth_ratio, lowest_to_eth_ratio,  # Ratios
            rel_change_current, rel_change_90, rel_change_180  # Relative Changes
        ]

    except Exception as e:
        print(f"Ошибка при обработке {symbol}: {e}")
        # Если данные не удалось получить, создаём пустую строку
        row = [base_symbol] + [""] * 15

    return row

def main():
    input_file = "input.csv"  # Имя входного CSV-файла
    output_file = "ticker_data.xls"  # Имя выходного файла

    symbols = read_symbols_from_csv(input_file)
    if not symbols:
        print("Список символов пуст. Проверьте файл.")
        return

    all_data = []
    for symbol in symbols:
        row = get_ticker_data(symbol)  # Получаем строку данных
        all_data.append(row)  # Добавляем её в общий список
        time.sleep(1)

    save_to_excel(all_data, filename=output_file)  # Сохраняем весь список сразу


if __name__ == "__main__":
    main()
