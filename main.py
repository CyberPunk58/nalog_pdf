import openpyxl
import datetime
import os
import shutil

# Загружаем файл с данными пациентов
patients_file = 'patients.xlsx'  # Укажите путь к вашему файлу
patients_wb = openpyxl.load_workbook(patients_file)
patients_ws = patients_wb['Лист1']  # Получаем конкретный лист с данными (Лист1)

# Загружаем файл-шаблон
blank_file = 'blank.xlsx'  # Укажите путь к вашему шаблону
blank_file_path = os.path.join(os.getcwd(), blank_file)

# Функция для определения адреса ячейки по индексу буквы
def get_cell_address(base_col, row, index):
    col_index = openpyxl.utils.column_index_from_string(base_col) + index
    col_letter = openpyxl.utils.get_column_letter(col_index)
    return f"{col_letter}{row}"

# Функция для записи значения в ячейку, учитывая объединённые ячейки
def write_to_cell_safe(ws, cell_address, value):
    cell = ws[cell_address]
    if isinstance(cell, openpyxl.cell.cell.MergedCell):
        # Получаем диапазон объединения, к которому принадлежит ячейка
        for merge_range in ws.merged_cells.ranges:
            if cell_address in merge_range:
                # Записываем значение в верхнюю левую ячейку диапазона
                top_left_cell = merge_range.min_row, merge_range.min_col
                ws.cell(row=top_left_cell[0], column=top_left_cell[1]).value = value
                return
    else:
        cell.value = value

# Подготавливаем список для хранения имён новых файлов
new_files = []

# Проходим по каждой строке в файле пациентов, начиная с 2-й строки (0 — это заголовок)
for row in patients_ws.iter_rows(min_row=2, values_only=True):
    surname, name, patronymic, birthdate, period, amount, uploaded = row

    # Проверяем, нужно ли обрабатывать запись (если 'Загружено' равно None)
    if uploaded is None:
        # Конструируем имя нового файла
        date_created = datetime.datetime.now().strftime('%Y-%m-%d')
        new_file_name = f'{surname}{name[0]}{patronymic[0]}_{date_created}.xlsx'
        new_file_path = os.path.join(os.getcwd(), new_file_name)

        # Копируем файл-шаблон в новый файл
        shutil.copy(blank_file_path, new_file_path)

        # Открываем новый файл для редактирования
        new_wb = openpyxl.load_workbook(new_file_path)
        new_ws = new_wb.active  # Работаем с активным листом

        # Функция для записи текста посимвольно в ячейки
        def write_to_cells(base_col, row, text):
            for index, char in enumerate(text):
                cell_address = get_cell_address(base_col, row, index * 2)  # Смещение через одну ячейку
                write_to_cell_safe(new_ws, cell_address, char)

        # Заполняем фамилию, имя и отчество
        write_to_cells('I', 24, surname)
        write_to_cells('I', 26, name)
        write_to_cells('I', 28, patronymic)

        # Заполняем период
        def write_period(base_col, row, period_value):
            period_str = str(period_value)  # Преобразуем численное значение в строку
            for index, char in enumerate(period_str):
                cell_address = get_cell_address(base_col, row, index * 2)
                write_to_cell_safe(new_ws, cell_address, char)

        write_period('BU', 11, period)

        # Сохраняем изменения в файле
        new_wb.save(new_file_path)
        new_files.append(new_file_name)  # Добавляем имя нового файла в список

# Выводим список созданных файлов
print("Созданные файлы:")
for file in new_files:
    print(file)
