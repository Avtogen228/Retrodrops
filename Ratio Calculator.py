import csv
import xlwt
import os

def calculate_percentage_difference_from_csv():
    try:
        # Устанавливаем путь к файлу input.csv
        file_path = os.path.join(os.path.dirname(__file__), "input.csv")

        # Создаем Excel-файл для основных результатов
        workbook = xlwt.Workbook()
        sheet = workbook.add_sheet("Результаты")

        # Создаем Excel-файл для данных только об изменениях
        ratio_workbook = xlwt.Workbook()
        ratio_sheet = ratio_workbook.add_sheet("Изменения")

        # Записываем заголовки для основного файла
        sheet.write(0, 0, "Опорное значение")

        with open(file_path, newline='', encoding='utf-8') as csvfile:
            reader = csv.reader(csvfile, delimiter=';')
            row_index = 1
            ratio_row_index = 0

            for row in reader:
                if len(row) < 2:
                    print("Ошибка: строка должна содержать хотя бы два числа.")
                    continue

                try:
                    # Преобразуем первую ячейку в число (опорное значение)
                    base_number = float(row[0].replace(',', '.'))
                    if base_number == 0:
                        print("Ошибка: первое число в строке равно 0, вычисление невозможно.")
                        continue

                    # Записываем опорное значение в основной Excel
                    sheet.write(row_index, 0, str(base_number).replace('.', ','))

                    # Обрабатываем остальные числа в строке
                    ratio_row = []
                    for col, value in enumerate(row[1:], start=1):
                        target_number = float(value.replace(',', '.'))
                        percentage_difference = ((target_number - base_number) / base_number) * 100

                        # Форматируем вывод для основного Excel
                        sign = "+" if percentage_difference > 0 else ""
                        result = f"{str(target_number).replace('.', ',')} ({sign}{round(percentage_difference)}%)"
                        sheet.write(row_index, col, result)

                        # Вычисляем значение изменения для файла изменений
                        if percentage_difference < 100:
                            ratio_value = 1 - abs(percentage_difference) / 100
                        else:
                            ratio_value = 1 + abs(percentage_difference) / 100

                        ratio_row.append(round(ratio_value, 2))

                    # Записываем строку изменений в ratio.xls
                    for ratio_col, ratio_value in enumerate(ratio_row):
                        ratio_sheet.write(ratio_row_index, ratio_col, ratio_value)

                    row_index += 1
                    ratio_row_index += 1
                except ValueError:
                    print("Ошибка: строка содержит некорректные данные.")

        # Сохраняем Excel-файлы
        output_file = "output.xls"
        ratio_output_file = "ratio.xls"
        workbook.save(output_file)
        ratio_workbook.save(ratio_output_file)

        print(f"Результаты сохранены в файлы: {output_file} и {ratio_output_file}")
    except FileNotFoundError:
        print("Ошибка: файл input.csv не найден.")
    except Exception as e:
        print(f"Произошла ошибка: {e}")

# Запуск программы
if __name__ == "__main__":
    calculate_percentage_difference_from_csv()
