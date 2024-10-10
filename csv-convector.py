import pandas as pd
import os

def convert_excel_to_csv(input_file):
    # Чтение всех листов из Excel файла
    xlsx = pd.ExcelFile(input_file)
    base_name = os.path.splitext(input_file)[0]

    for sheet in xlsx.sheet_names:
        df = pd.read_excel(xlsx, sheet_name=sheet, header=0)
        output_file = f"{input_file.replace('.xlsx', '')}.csv"
        df.to_csv(output_file, index=False, encoding='utf-8-sig')
        print(f"Файл {output_file} успешно сохранен.")

# Получение списка всех файлов в текущей директории
files = [f for f in os.listdir() if f.endswith('.xlsx')]

# Конвертация каждого файла
for file in files:
    convert_excel_to_csv(file)
