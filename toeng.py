# How does the code works? 
# First of all it gets all lists from the excel file 
# Load it to the dataframe using pandas
# After this code is splitting data frames by chuncks 
# Each size of chuck is 20,000 lines
# Transilliterate the chunk
# And the final step is creating new xlsx file where we write all data from chucnk

# How to start the script?
# Place script in folder with xlsx files you need to translliterate 
# In concole start the script by writing "py toeng.py"
# Thats all!
# You'll see the messages how the work is going

#RU
# Как работает код?
# Первым делом скрипт получает все страницы из эксель файлов
# После этого загружает информацию со всех листов в dataframe 
# Далее скрипт разделяет дата фрейм на чанки по 20,0000 строк
# Транслиттерирует чанк (просто заменяет русские буквы на английские по сути)
# После того как он полностью "перевел" чанк. Создает эксель документ куда записывает всю информацию из чанка
# Таким образом большие файлы пилятся на маленькие файлы


# Как запустить код?
# Положите скрипт в папку где находятся необходимые xlsx файлы 
# В консоле запустите скрипт написва "py toeng.py"
# Все код работает! 
# В консоле будут выводится сообщения о том как проходит работа 

import pandas as pd
from transliterate import translit
import os



def transliterate_excel_chunked(input_path, chunk_size=20000):
    # Чтение всех листов в Excel-файле
    # Reading all info from all lists in excel file
    xls = pd.ExcelFile(input_path)
    
    for sheet_name in xls.sheet_names:
        # Чтение текущего листа целиком
        # Getting data from current list 
        df = pd.read_excel(xls, sheet_name=sheet_name)
        
        # Разделение DataFrame на чанки
        # Split the DataFrame by chuncks
        for chunk_number, start_row in enumerate(range(0, df.shape[0], chunk_size)):
            end_row = min(start_row + chunk_size, df.shape[0])
            df_chunk = df.iloc[start_row:end_row]
            
            # Применение транслитерации ко всем строковым ячейкам
            # Translitterate the chunck
            df_transliterated = df_chunk.map(
                lambda x: translit(x, 'ru', reversed=True) if isinstance(x, str) else x
            )
            
            # Создание имени файла для текущего чанка
            # Creating file name for the current chucnk
            output_file = os.path.join(f"{os.path.splitext(os.path.basename(input_path))[0]}_{sheet_name}_part{chunk_number + 1}.xlsx")
            
            # Запись чанка в новый Excel-файл
            # Writing all data from chunck to new excel file
            output_file = output_file.replace(' ', '_')
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                df_transliterated.to_excel(writer, sheet_name=sheet_name, index=False)

#Start point
#Точка запуска кода

files = [f for f in os.listdir() if f.endswith('.xlsx')]
for file in files:
    print(f'Starting {file}')
    transliterate_excel_chunked(file)
    print(f'{file} complete!')