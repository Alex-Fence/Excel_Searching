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
    errors = []
    if file_path.endswith('.xls'):
        engine = 'xlrd'
    elif file_path.endswith('.xlsx'):
        engine = 'openpyxl'
    try:
        df = pd.read_excel(file_path, engine=engine)
        for column in df.columns:
            if df[column].astype(str).str.contains(word, na=False).any():
                found_file_str = f'Слово "{word}" найдено в файле: {file_path}'
                print(found_file_str)
                write_to_file(result_file, found_file_str + '\n')

                finding_files_lst.append(file_path)
                return True
    except Exception as e:
        errors.append(f'Ошибка при обработке файла {file_path}: {e}')
    return False

def write_to_file(file_path, write_str):
    with open(file_path, 'a', encoding='utf-8') as f:
        f.write(write_str)

def main_func(search_directory, search_word):
    """Основная функция для поиска файлов и слов."""
    excel_files = find_excel_files(search_directory)
    print(f'Найдено {len(excel_files)} файлов Excel.')

    for file in excel_files:
        search_word_in_excel(file, search_word)
    for file_name in finding_files_lst:
        print(file_name)


def input_data():
    """Функция для ввода данных пользователем."""
    sourse_directory = input('Введите путь к директории для поиска: ')
    search_word = input('Введите слово для поиска: ')
    return sourse_directory, search_word


if __name__ == '__main__':
    # Задайте директорию для поиска и слово для поиска
    search_directory, search_word = input_data()
    finding_files_lst = []
    result_file = 'result.txt'

    main_func(search_directory, search_word)
