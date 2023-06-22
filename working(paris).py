import json
from openpyxl import load_workbook

# Загрузка таблицы
workbook = load_workbook('Вставить путь до таблицы')
sheet = workbook.active

previous_value = None  # Переменная для хранения предыдущего значения
results = []  # Список для хранения результатов

# Поиск значений 0 и вывод информации о ячейках
for column in sheet.iter_cols():
    for cell in column:
        if cell.value == 0:
            row_above = cell.row - 2
            day_of_week = sheet.cell(row=cell.row, column=cell.column - 1).value
            day_mapping = {
                'понедельник': 'Monday',
                'вторник': 'Tuesday',
                'среда': 'Wednesday',
                'четверг': 'Thursday',
                'пятница': 'Friday',
                'суббота': 'Saturday',
                'воскресенье': 'Sunday'
            }
            day_of_week_eng = day_mapping.get(day_of_week)
            value_above = sheet.cell(row=row_above, column=cell.column).value
            previous_value = sheet.cell(row=row_above, column=cell.column).value if value_above != 5 else previous_value

            if value_above == 5:
                output = {"groupName": previous_value, "weekDay": day_of_week_eng, "paris": []}
            else:
                output = {"groupName": value_above, "weekDay": day_of_week_eng, "paris": []}

            # Добавление значений из соседних ячеек
            for i in range(0, 6):  # 6 строк вниз
                pair_column = cell.column + 1
                room_column = cell.column + 2
                pair_number_column = cell.column

                pair_value = sheet.cell(row=cell.row + i, column=pair_column).value
                room_value = sheet.cell(row=cell.row + i, column=room_column).value
                pair_number_value = sheet.cell(row=cell.row + i, column=pair_number_column).value

                if pair_value is not None:
                    pair_value = ' '.join(pair_value.split())
                    if room_value and room_value != 0:
                        room_value = ' '.join(str(room_value).split())
                        row_values = {"pairNum": pair_number_value, "pairName": pair_value, "pairCab": room_value}
                    else:
                        row_values = {"pairNum": pair_number_value, "pairName": pair_value}
                    output["paris"].append(row_values)

            if output["paris"]:
                results.append(output)

# Закрытие таблицы
workbook.close()

# Изменение пути сохранения файла JSON
output_path = "путь"

# Сохранение в файл JSON с отступами
with open(output_path, "w", encoding="utf-8") as file:
    json.dump(results, file, indent=4, ensure_ascii=False)
