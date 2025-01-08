import openpyxl
import datetime
import os
import shutil
import logging

# Настройка логирования
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Загружаем файл с данными пациентов
patients_file = 'patients.xlsx'  # Укажите путь к вашему файлу
blank_file = 'blank.xlsx'  # Шаблон

# Проверка существования файлов
if not os.path.exists(patients_file):
    raise FileNotFoundError(f"Файл {patients_file} не найден.")
if not os.path.exists(blank_file):
    raise FileNotFoundError(f"Файл {blank_file} не найден.")

try:
    patients_wb = openpyxl.load_workbook(patients_file)
except Exception as e:
    logging.error(f"Ошибка при загрузке файла {patients_file}: {e}")
    exit()

if 'Лист1' not in patients_wb.sheetnames:
    raise ValueError("Лист 'Лист1' не найден в файле patients.xlsx.")

patients_ws = patients_wb['Лист1']  # Получаем конкретный лист с данными (Лист1)

# Функция для определения адреса ячейки по индексу буквы
def get_cell_address(base_col, row, index):
    col_index = openpyxl.utils.column_index_from_string(base_col) + index
    col_letter = openpyxl.utils.get_column_letter(col_index)
    return f"{col_letter}{row}"

# Функция для записи значения в ячейку, учитывая объединённые ячейки
def write_to_cell_safe(ws, cell_address, value):
    cell = ws[cell_address]
    if isinstance(cell, openpyxl.cell.cell.MergedCell):
        for merge_range in ws.merged_cells.ranges:
            if cell_address in merge_range:
                top_left_cell = merge_range.min_row, merge_range.min_col
                ws.cell(row=top_left_cell[0], column=top_left_cell[1]).value = value
                return
    else:
        cell.value = value

#Функция для записи номера паспорта
def write_passport(ws, passport):
    if not passport:  # Пропускаем, если паспорт отсутствует
        return

    passport_str = str(passport)  # Преобразуем номер паспорта в строку
    start_col = 'AO'  # Начинаем с ячейки AO33
    row = 33

    # Записываем первые 4 символа
    for i in range(4):
        if i < len(passport_str):
            cell_address = get_cell_address(start_col, row, i * 2)  # Смещение через одну ячейку
            write_to_cell_safe(ws, cell_address, passport_str[i])

    # Пропускаем две ячейки (смещение на 2 ячейки после первых 4 символов)
    # Записываем оставшиеся символы
    for i in range(4, len(passport_str)):
        cell_address = get_cell_address(start_col, row, i * 2 + 2)  # Смещение + пропуск двух ячеек
        write_to_cell_safe(ws, cell_address, passport_str[i])

# Функция для записи суммы в строку 40 с учетом смещения
def write_amount(ws, amount):
    if amount is None:
        return  # Пропускаем, если сумма отсутствует

    amount_str = str(amount)  # Преобразуем сумму в строку
    start_col = 'BQ'  # Начинаем с ячейки BQ40
    row = 40

    # Вычисляем начальную ячейку с учетом смещения
    start_col_index = openpyxl.utils.column_index_from_string(start_col)
    start_col_index -= len(amount_str) * 2  # Смещаем влево на количество символов суммы * 2
    start_col_letter = openpyxl.utils.get_column_letter(start_col_index)

    # Записываем сумму посимвольно
    for index, char in enumerate(amount_str):
        cell_address = get_cell_address(start_col_letter, row, (index * 2)+2)
        write_to_cell_safe(ws, cell_address, char)

# Подготавливаем список для хранения имён новых файлов
new_files = []

# Проходим по каждой строке в файле пациентов, начиная с 2-й строки (0 — это заголовок)
for row in patients_ws.iter_rows(min_row=2, values_only=True):
    surname, name, patronymic, birthdate, period, amount, inn, passport, issue_date, uploaded = row

    # Пропускаем строки, которые уже были обработаны
    if uploaded is not None:
        continue

    # Создаем имя нового файла
    date_created = datetime.datetime.now().strftime('%Y-%m-%d')
    new_file_name = f'{surname}{name[0]}{patronymic[0]}_{date_created}.xlsx'
    new_file_path = os.path.join(os.getcwd(), new_file_name)

    # Копируем шаблон
    shutil.copy(blank_file, new_file_path)

    try:
        new_wb = openpyxl.load_workbook(new_file_path)
        new_ws = new_wb.active

        # Функция для записи текста посимвольно в ячейки
        def write_to_cells(base_col, row, text):
            for index, char in enumerate(text):
                cell_address = get_cell_address(base_col, row, index * 2)
                write_to_cell_safe(new_ws, cell_address, char)

        # Заполняем фамилию, имя и отчество
        write_to_cells('I', 24, surname)
        write_to_cells('I', 26, name)
        write_to_cells('I', 28, patronymic)

        # Заполняем период
        def write_period(base_col, row, period_value):
            period_str = str(period_value)
            for index, char in enumerate(period_str):
                cell_address = get_cell_address(base_col, row, index * 2)
                write_to_cell_safe(new_ws, cell_address, char)

        write_period('BU', 11, period)

        # Заполняем сумму
        write_amount(new_ws, amount)

        # Заполняем ИНН (если есть)
        if inn:
            write_to_cells('A', 10, str(inn))  # Пример: запись ИНН в ячейку A10

        # Заполняем паспорт (если есть)
        write_passport(new_ws, passport)

        # Заполняем паспорт (если есть)
        if passport:
            write_to_cells('B', 10, str(passport))  # Пример: запись паспорта в ячейку B10

        # Заполняем дату выдачи (если есть)
        if issue_date:
            if isinstance(issue_date, datetime.datetime):
                issue_date_str = issue_date.strftime('%d.%m.%Y')
            else:
                issue_date_str = str(issue_date)
            write_to_cells('C', 10, issue_date_str)  # Пример: запись даты выдачи в ячейку C10

        # Сохраняем изменения
        new_wb.save(new_file_path)
        new_files.append(new_file_name)
        logging.info(f"Файл {new_file_name} успешно создан.")

    except Exception as e:
        logging.error(f"Ошибка при обработке файла {new_file_name}: {e}")
    finally:
        if 'new_wb' in locals():
            new_wb.close()

# Выводим список созданных файлов
logging.info("Созданные файлы:")
for file in new_files:
    logging.info(file)

patients_wb.close()