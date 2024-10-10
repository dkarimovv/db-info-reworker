import pandas as pd
import os 

from transliterate import translit
from config import column_sets

def byebrackets(details_str):
    if '8 (' in details_str:
        details_str = details_str.replace('8 (', '7').replace(') ', '')
    return details_str

def process_contact_details(details_str):
    """Функция для обработки строки с контактами (номера телефонов или электронные почты)"""
    if pd.isna(details_str):
        return None, None
    details_str = str(details_str)
    if '8 (' in details_str:
        details_str = details_str.replace('8 (', '7').replace(') ', '')
    elif '+7 (' in details_str:
        details_str = details_str.replace('+7 (' , '7').replace(') ', '')
    elif '+7' in details_str:
        details_str = details_str.replace('+7' , '7')

    if ';' in details_str:
        details_str = details_str.replace(';' , ',')
    if '-' in details_str:
        details_str = details_str.replace('-' , '')
    
    details = details_str.split(',')
    main_detail = details[0].strip() if details else None
    additional_details = ', '.join(details[1:]).strip() if len(details) > 1 else None
    return main_detail, additional_details

def combine_columns(df, columns):
    combined = df[columns].apply(lambda row: ', '.join(row.dropna().astype(str)), axis=1)
    return combined

def generate_company_name(row):
    return f"IP {row['ФИО']}" if 'ФИО' in row else None

def excel_format_universal(input_file):
    max_matched_columns = 0
    assignedto = 'nameplace'
    descriptions = []
    dataframes = []
    selected_set = None
    output_file_path = input_file.replace('.xlsx', '') + '_mod' + '.xlsx'

    # Чтение всех листов из Excel файла
    xlsx = pd.ExcelFile(input_file)

    for sheet in xlsx.sheet_names:
        df = pd.read_excel(xlsx, sheet_name=sheet, header=0)
        dataframes.append(df)

    # Объединение всех листов в один DataFrame
    df = pd.concat(dataframes, ignore_index=True)

    # Определение соответствующего набора столбцов из config.py
    for _, set_info in column_sets.items():
        matched_columns = sum(col in df.columns for col in set_info['columns'])
        if matched_columns > max_matched_columns:
            selected_set = set_info
            max_matched_columns = matched_columns
            
    #     print(f"Set: {set_name}, Matched Columns: {matched_columns} ||| Total len {len(col in df.columns)}")

    for _, set_info in column_sets.items():
        if all(col in df.columns for col in set_info['columns']):
            selected_set = set_info
            break

    if selected_set is None:
        for _, set_info in column_sets.items():
            matched_columns = sum(col in df.columns for col in set_info['columns'])
            if matched_columns > max_matched_columns:
                selected_set = set_info
                max_matched_columns = matched_columns
        if selected_set is None:
            print(f"{input_file} Отсутствуют необходимые столбцы")
            return


    # Переименование столбцов для соответствия требованиям
    try:
        df_selected = df[selected_set['columns']].rename(columns=selected_set['rename'])
    except KeyError as e:
        print(f"{input_file} ||| Ошибка: отсутствует один из необходимых столбцов: {e}")
        return

    # Обработка телефонов и электронной почты
    phone_columns = ['Телефон директора', 'Сотовый директора', 'Телефоны из открытых источников', 'Сотовый', 'Телефон']
    email_columns = ['Емайл директора', 'Электронный адрес', 'емайлы из открытых источников']

    # Объединение всех вариантов столбцов с телефонами и почтами
    df_selected['Все телефоны'] = combine_columns(df_selected, [col for col in phone_columns if col in df_selected.columns])
    df_selected['Все почты'] = combine_columns(df_selected, [col for col in email_columns if col in df_selected.columns])

    # Проверка и создание столбца "Наименование компании"
    if 'Наименование компании' not in df_selected.columns:
        df_selected['Наименование компании'] = df_selected.apply(generate_company_name, axis=1)

    # Создание нового столбца "Описание"
    
    for index, row in df_selected.iterrows():
        main_phone, additional_phones = process_contact_details(row.get('Сотовый директора', ''))
        main_email, additional_emails = process_contact_details(row.get('Электронный адрес', ''))
        open_source_phones, additional_open_source_phones = process_contact_details(row.get('Телефоны из открытых источников', ''))
        open_source_emails, additional_open_source_emails = process_contact_details(row.get('емайлы из открытых источников', ''))

        if open_source_phones:
            open_source_phones = byebrackets(str(open_source_phones))
            additional_open_source_phones = byebrackets(str(additional_open_source_phones))

        description = (
            (f"WhatsApp: {row.get('Наличие вацап', '')}\n" if 'Наличие вацап' in row else "") +
            (f"Sotovii direktora: {main_phone}\n" if main_phone else "") +
            (f"Dop nomera: {additional_phones}\n" if additional_phones else "No dop numbers\n") +
            (f"Dop nomera (iz otkritix istochnikov): {additional_open_source_phones}\n" if additional_open_source_phones else "") +
            (f"Opensource telephoni: {str(row.get('Телефоны из открытых источников', '')).replace('8 (', '7').replace(') ', '')}\n" if 'Телефоны из открытых источников' in row else "") +
            (f"Email: {main_email}\n" if main_email else "") +
            (f"Dop email: {additional_emails}\n" if additional_emails in row else "") +
            (f"Dop email (iz otkritix istochnikov): {open_source_emails}" if open_source_emails else "") + 
            (f", {additional_open_source_emails}\n" if additional_open_source_emails else "") +
            (f"Vid dejatel'nosti/otrasl': {row.get('Вид деятельности/отрасль', '')}\n" if 'Вид деятельности/отрасль' in row else "") +
            (f"Kod osnovnogo vida dejatel'nosti: {row.get('Код основного вида деятельности', '')}" if 'Код основного вида деятельности' in row else "")
        )
        descriptions.append(description)

        df_selected.at[index, 'Сотовый директора'] = main_phone
        df_selected.at[index, 'Телефоны из открытых источников'] = open_source_phones
        df_selected.at[index, 'Электронный адрес'] = main_email

    df_selected['Описание'] = descriptions

    df_selected['Lead Status'] = 'NewLead'
    df_selected['Lead Source'] , _ = str(translit(input_file.replace('.xlsx' , ''), 'ru', reversed=True)).split('_part')
    df_selected['Assigned to'] = assignedto

    # Перевод названий столбцов на английский
    column_translation = {
        'Наименование компании': 'Company Name',
        'ОГРН': 'OGRN',
        'Руководитель': 'Director',
        'Сотовый директора': 'Director Mobile',
        'Телефоны из открытых источников': 'Public Phones',
        'Электронный адрес': 'Email',
        'ИНН': 'INN',
        'Описание': 'Description'
    }


    df_selected = df_selected.rename(columns=column_translation)

    final_columns = ['Company Name', 'OGRN', 'Director', 'Director Mobile',
                     'Public Phones', 'Email', 'INN', 'Description',
                     'Lead Status', 'Lead Source', 'Assigned to']
    
    # Фильтрация столбцов для исключения отсутствующих
    final_columns = [col for col in final_columns if col in df_selected.columns]

    df_result = df_selected[final_columns]

    # Сохранение результата в новый файл
    df_result.to_excel(output_file_path, index=False)
    print("Файл успешно сохранен:", output_file_path)


fs = [f for f in os.listdir() if f.endswith('.xlsx')]
for f in fs:
    excel_format_universal(f)
