# ------------------------------------------------------------------------
# -------------------------- Обработка файла .xlsb -----------------------
# ------------------------------------------------------------------------

import glob

# Проверка, есть ли в корневом каталоге файл с расширением .xlsb
file_list = glob.glob('../*.xlsb')
if (len(file_list) == 0):
    print('File *.xlsx not found')
    exit()

# Ипрорт pandas
import pandas as pd

# Путь к исходному файлу
file_src = file_list[0]

# Получение dataframe исходного файла
df = pd.read_excel(file_src, engine='pyxlsb')

# Удаление последних двух ненужных столбцов, если они существуют
if 'Unnamed: 28' in df.columns:
    df.drop('Unnamed: 28', inplace=True, axis=1)
if 'Unnamed: 27' in df.columns:
    df.drop('Unnamed: 27', inplace=True, axis=1)

# Очистка всех значений столбца 'ФИО'
for i in range(df['ФИО'].size):
    df['ФИО'][i] = ''

# Создание нового файла, запись в него данных и добавление форматирования
with pd.ExcelWriter('../Таблица_безфио.xlsx', engine='xlsxwriter') as wb:
    df.to_excel(wb, sheet_name='Данные', index=False)
    sheet = wb.sheets['Данные']
    sheet.autofilter('A1:AA'+str(df.shape[0]))

# -------------------------------- Форматирование нового файла -----------------------------------

# Форматирование по умолчанию для всех строк в таблице
# !!!!!!!!!!!!!!!!!!!!!!!!! Сделать итерацию по всем строкам и установить базовое форматирование !!!!!!!!!!!!!!!!!!
    sheet.set_default_row(15)

# Форматирование строки заголовков
    cell_format = wb.book.add_format()
    cell_format.set_bold()
    cell_format.set_font_color('red')
    sheet.set_row(0, 26.5, cell_format)

# Установка ширины всех столбцов
    sheet.set_column(0, 0, 9.5)
    sheet.set_column(1, 1, 38)
    sheet.set_column(2, 2, 19)
    sheet.set_column(3, 3, 6)
    sheet.set_column(4, 4, 4.5)
    sheet.set_column(5, 5, 21)

# Удаление старого файла
import os
os.remove(file_src)

print('done')