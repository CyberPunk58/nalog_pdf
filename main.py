import openpyxl
import datetime
import os
import shutil  # Импортируем библиотеку для копирования файлов

# Загружаем файл с данными пациентов
patients_file = 'patients.xlsx'  # Укажите путь к вашему файлу
patients_wb = openpyxl.load_workbook(patients_file)  # Загружаем книгу
patients_ws = patients_wb['Лист1']  # Получаем конкретный лист с данными (Лист1)

# Загружаем файл-шаблон
blank_file = 'blank.xlsx'  # Укажите путь к вашему шаблону
# Прежде чем работать, убедитесь, что файл 'blank.xlsx' находится в той же директории, что и ваш скрипт
blank_file_path = os.path.join(os.getcwd(), blank_file)

# Подготавливаем список для хранения имён новых файлов
new_files = []

# Проходим по каждой строке в файле пациентов, начиная с 2-й строки (0 — это заголовок)
for row in patients_ws.iter_rows(min_row=2, values_only=True):
    surname, name, patronymic, birthdate, period, amount, uploaded = row

    # Проверяем, нужно ли обрабатывать запись (если 'Загружено' равно None)
    if uploaded is None:
        # Конструируем имя нового файла
        date_created = datetime.datetime.now().strftime('%Y-%m-%d')  # Получаем текущую дату в формате ГГГГ-ММ-ДД
        new_file_name = f'{surname}{name[0]}{patronymic[0]}_{date_created}.xlsx'  # Формируем имя файла
        new_file_path = os.path.join(os.getcwd(), new_file_name)  # Полный путь к новому файлу

        # Копируем файл-шаблон в новый файл
        shutil.copy(blank_file_path, new_file_path)  # Используем shutil для копирования файла
        new_files.append(new_file_name)  # Добавляем имя нового файла в список

# Выводим список созданных файлов
print("Созданные файлы:")
for file in new_files:
    print(file)


