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
    col_index = ord(base_col) - ord('A') + 1 + index  # Преобразование в числовое значение столбца
    col_letter = ''
    while col_index > 0:
        col_index, remainder = divmod(col_index - 1, 26)
        col_letter = chr(65 + remainder) + col_letter
    return f"{col_letter}{row}"

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
                new_ws[cell_address].value = char

        # Заполняем фамилию, имя и отчество
        write_to_cells('I', 24, surname)
        write_to_cells('I', 26, name)
        write_to_cells('I', 28, patronymic)

        # Сохраняем изменения в файле
        new_wb.save(new_file_path)
        new_files.append(new_file_name)  # Добавляем имя нового файла в список

# Выводим список созданных файлов
print("Созданные файлы:")
for file in new_files:
    print(file)
