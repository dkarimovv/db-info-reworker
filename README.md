# EN

# Universal Excel File Processor with Transliteration

This project is designed for processing Excel files and converting them into a more readable format by cleaning phone numbers, emails, and transliterating text where needed. It also includes the capability to convert Excel files into CSV format and process large files in chunks.

## Key Features:
- **Excel File Processing**: Processes Excel files, cleans and formats contact details (phones and emails).
- **Transliteration**: Replaces Russian characters with their English equivalents using chunk processing for large files.
- **CSV Conversion**: Converts Excel files into CSV format for easier manipulation.
- **Flexible Column Management**: Automatically detects columns and renames them for consistency.
- **Custom Descriptions**: Generates detailed descriptions based on contact details.

## Technology Stack:
- **Python**: The main programming language for data processing.
- **pandas**: For reading and manipulating Excel data.
- **openpyxl**: For writing Excel files.
- **transliterate**: For converting Cyrillic text to Latin equivalents.
- **os**: For file system operations.

## Project Structure:
-> config.py # Configuration file for column sets and matching. 
-> csv-convector.py # Script for converting Excel files to CSV format.
-> main.py # Main script for processing Excel files and generating results. 
-> toeng.py # Script for chunked transliteration of large Excel files.

## How to Use:

1. **Main Processing Script (`main.py`)**:
   - Run `main.py` to process the Excel files and generate the final output.
   - The script will clean phone numbers, format emails, and create detailed descriptions based on your dataset.

2. **CSV Conversion (`csv-convector.py`)**:
   - Run `csv-convector.py` to convert all Excel files in the folder to CSV format.
   
3. **Transliteration (`toeng.py`)**:
   - Use `toeng.py` to transliterate Cyrillic text to Latin by processing large Excel files in chunks of 20,000 rows.

4. Place your Excel files in the same directory as the scripts and run the relevant script based on your needs.


# RU

# Универсальный обработчик Excel-файлов с транслитерацией

Этот проект предназначен для обработки Excel-файлов, включая их очистку, форматирование контактных данных (телефоны, электронные письма) и транслитерацию текста при необходимости. Также проект включает возможность конвертировать Excel-файлы в формат CSV и обрабатывать большие файлы частями.

## Основные функции:
- **Обработка Excel-файлов**: Обрабатывает Excel-файлы, очищает и форматирует контактные данные (телефоны и электронные письма).
- **Транслитерация**: Заменяет русские символы на английские с использованием частичной обработки для больших файлов.
- **Конвертация в CSV**: Конвертирует Excel-файлы в формат CSV для дальнейшей обработки.
- **Гибкое управление столбцами**: Автоматически определяет столбцы и переименовывает их для консистентности.
- **Создание описаний**: Генерирует детальные описания на основе контактных данных.

## Технологический стек:
- **Python**: Основной язык программирования для обработки данных.
- **pandas**: Для работы с Excel-файлами и манипуляции данными.
- **openpyxl**: Для записи и обработки Excel-файлов.
- **transliterate**: Для транслитерации текста с кириллицы на латиницу.
- **os**: Для работы с файловой системой.

## Структура проекта:
-> config.py # Конфигурационный файл для наборов столбцов. 
-> csv-convector.py # Скрипт для конвертации Excel-файлов в CSV формат. 
-> main.py # Основной скрипт для обработки Excel-файлов и генерации результатов. 
-> toeng.py # Скрипт для транслитерации больших Excel-файлов по частям.


## Как использовать:

1. **Основной скрипт обработки (`main.py`)**:
   - Запустите `main.py` для обработки Excel-файлов и генерации финального результата.
   - Скрипт очищает телефонные номера, форматирует электронные письма и создает детальные описания на основе вашего набора данных.

2. **Конвертация в CSV (`csv-convector.py`)**:
   - Запустите `csv-convector.py` для конвертации всех Excel-файлов в директории в формат CSV.
   
3. **Транслитерация (`toeng.py`)**:
   - Используйте `toeng.py` для транслитерации кириллического текста на латиницу, обрабатывая большие Excel-файлы по частям (по 20 000 строк).

4. Поместите свои Excel-файлы в ту же директорию, где находятся скрипты, и запустите соответствующий скрипт в зависимости от ваших нужд.

