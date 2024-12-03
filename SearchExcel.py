import os
import pandas as pd

def find_excel_files(directory):
    """Ищет все файлы Excel в указанной директории и её подкаталогах."""
    excel_files = []
    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.endswith(('.xls', '.xlsx')):
                excel_files.append(os.path.join(root, file))
    return excel_files

def search_word_in_excel(file_path, word):
    """Ищет заданное слово в файле Excel."""
    if file_path.endswith('.xls'):
        engine = 'xlrd'
    elif file_path.endswith('.xlsx'):
        engine = 'openpyxl'
    try:
        df = pd.read_excel(file_path, engine = engine)
        for column in df.columns:
            if df[column].astype(str).str.contains(word, na=False).any():
                print(f'Слово "{word}" найдено в файле: {file_path}')
                finding_files_lst.append(file_path)
                return True
    except Exception as e:
        print(f'Ошибка при обработке файла {file_path}: {e}')
    return False

def main(search_directory, search_word):
    """Основная функция для поиска файлов и слов."""
    excel_files = find_excel_files(search_directory)
    print(f'Найдено {len(excel_files)} файлов Excel.')

    for file in excel_files:
        search_word_in_excel(file, search_word)
    for file_name in finding_files_lst:
        print(file_name)

# Задайте директорию для поиска и слово для поиска
search_directory = 'C:\\Users\\ita\\Downloads\\'  # Замените на нужный путь
search_word = 'плата'  # Замените на нужное слово
finding_files_lst = []

main(search_directory, search_word)