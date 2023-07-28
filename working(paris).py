import json
import requests
from openpyxl import load_workbook

# Загрузка таблицы
workbook = load_workbook('Укажите путь до файла')

results = []  # Список для хранения результатов

# Проход по всем листам в книге
for sheet_name in workbook.sheetnames:
    sheet = workbook[sheet_name]

   

    previous_value = None  # Переменная для хранения предыдущего значения
    for column in sheet.iter_cols():
        for cell in column:
            if cell.value == 0:
                row_above = cell.row - 2
                day_of_week = sheet.cell(row=cell.row, column=cell.column - 1).value
                day_mapping = {
                    'понедельник': 'monday',
                    'вторник': 'tuesday',
                    'среда': 'wednesday',
                    'четверг': 'thursday',
                    'пятница': 'friday',
                    'суббота': 'saturday',
                    'воскресенье': 'sunday'
                }
                day_of_week_eng = day_mapping.get(day_of_week)
                value_above = sheet.cell(row=row_above, column=cell.column).value
                previous_value = sheet.cell(row=row_above, column=cell.column).value if value_above != 5 else previous_value

                if value_above == 5:
                    output = {"groupName": previous_value, "weekDay": day_of_week_eng, "pairs": []}
                else:
                    output = {"groupName": value_above, "weekDay": day_of_week_eng, "pairs": []}

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
                            row_values = {"pairNum": pair_number_value, "pairName": pair_value, "pairCab": ""}
                        output["pairs"].append(row_values)

                if output["pairs"]:
                    results.append(output)


payload = {
    "password": "укажите пароль с send запроса",
    "groups": results
}


json_payload = json.dumps(payload, separators=(",", ":"), ensure_ascii=False)

    # Отправка файла по ссылке
server_url = "ссылка на send запрос"  # Замените на вашу ссылку
response = requests.post(server_url, json=payload, headers={'Content-Type': 'application/json'})
print(f"Отправка результатов на сайт для листа '{sheet_name}': {response.status_code} - {response.text}")
 
# Закрытие таблицы
workbook.close()