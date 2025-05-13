# БИБЛИОТЕКИ
import os.path
import shutil
import zipfile
import rarfile
import py7zr
import aspose.words as aw
import re
import pandas as pd
import jpype

jpype.startJVM()

# тессеракт
from text_convert import parse_pdf, image_to_text
from tables import *

# поиск совпадений
from match import match_finder, match_finder_extend


# УКАЗАНИЕ ПУТЕЙ К СВОДНОЙ АНКЕТЕ И АРХИВУ

# archive_path = input('введите путь до папки с документами от участника ')
# summary_form = input('введите путь до сводной анкеты ')
archive_path = r"D:\учеба\опд\example_of_archive"
summary_form = r"D:\учеба\опд\04. Сводная анкета КТЧ_ХВ-Мессояха-1_для претендента_ с ключевыми словами.xlsx"


# ФУНКЦИИ
def dir_creator(dirname):
    '''
    создание папки
    :param dirname: имя папки
    :return: имя и путь до папки
    '''
    path = os.path.join(dirname)
    if not os.path.exists(dirname):
        os.makedirs(dirname)
    else:
        files = os.listdir(path)
        if files:
            for file in files:
                os.remove(os.path.join(path, file))
    return dirname, path

def pdf_converter(filename, true_path, pdf_files_path):
    '''
    функция конвертирующая файл в пдф (если это doc/docx/html и сохраняющая их в папку с копиями
    :param filename: имя файла
    :param true_path: путь до оригинала
    :param pdf_files_path: путь до пдф копии
    :return: файл как объект класса Files
    '''

    basename, extension = os.path.splitext(filename)
    element_of_Files = Files(filename, true_path)

    if extension in (".pdf", ".PDF", ".png", ".jpg", ".jpeg"):
        set_into_class(element_of_Files, basename, extension, is_pdf=1, pdf_files_path=pdf_files_path, filename=filename)
        shutil.copy(true_path, pdf_files_path)

    elif extension in (".doc", ".docx", ".html"):
        file = aw.Document(true_path)
        like_a_path = os.path.join(pdf_files_path, f"{basename}.pdf")
        file.save(like_a_path)
        set_into_class(element_of_Files, basename, extension, like_a_path)

    return element_of_Files

def set_into_class(element_of_Files, basename, extension, like_a_path='', is_pdf=0, filename='', pdf_files_path=''):
    '''
    добавление атрибутов объектам класса Files
    :param element_of_Files: объект класса Files
    :param basename: название файла
    :param extension: расширение файла
    :param like_a_path: путь до пдф копии файла?
    :param is_pdf: пдф или нет
    :param filename: название файла + расширение
    :param pdf_files_path: путь до папки содержащей пдф копии файлов
    '''
    if is_pdf:
        element_of_Files.pdf_copy_path = os.path.join(pdf_files_path, filename)
    else:
        element_of_Files.pdf_copy_path = like_a_path
    element_of_Files.intended_name = basename
    element_of_Files.extension = extension

def dir_content_sorting(dir_path, unpacked_dir):
    '''
    сортировка всех файлов из данной директории по папкам в зависимости от их расширения
    :param dir_path: путь до данной директории
    '''

    # не полностью тестировано, но должно работать
    if dir_path.endswith('.zip'):
        with zipfile.ZipFile(dir_path, 'r') as zip_ref:
            zip_ref.extractall(unpacked_dir)
    elif dir_path.endswith('.rar'):
        with rarfile.RarFile(dir_path, 'r') as rar_ref:
            rar_ref.extractall(unpacked_dir)
    elif dir_path.endswith('.7z'):
        with py7zr.SevenZipFile(dir_path, mode='r') as sevenzip_ref:
            sevenzip_ref.extractall(unpacked_dir)

    for dirpath, dirnames, filenames in os.walk(dir_path):
        for filename in filenames:
            true_path = os.path.join(dirpath, filename)
            file_ext = os.path.splitext(filename)[1]
            if file_ext in ('.zip', '.rar', '.7z'):
                if file_ext == '.zip':
                    with zipfile.ZipFile(true_path, 'r') as zip_ref:
                        zip_ref.extractall(unpacked_dir)
                elif file_ext == '.rar':
                    with rarfile.RarFile(true_path, 'r') as rar_ref:
                        rar_ref.extractall(unpacked_dir)
                else:
                    with py7zr.SevenZipFile(true_path, mode='r') as sevenzip_ref:
                        sevenzip_ref.extractall(unpacked_dir)
            # помимо того что есть pdf бывает запись PDF, возможно такие штуки ещё встречаются, тогда нужно их сюда добавить
            elif file_ext not in (".pdf", ".PDF", ".doc", ".docx", ".html", ".png", ".jpg", ".jpeg", ".xls", ".xlsx"):
                shutil.copy(true_path, unknown_files_path)
            elif file_ext in (".xls", ".xlsx"):
                shutil.copy(true_path, table_files_path)
            else:
                element_of_Files = pdf_converter(filename, true_path, pdf_files_path)
                list_of_Files.append(element_of_Files)


# КЛАССЫ
class Files:
    def __init__(self, filename, true_path):
        self.filename = filename
        ''' название файла '''
        self.true_path = true_path
        ''' путь до файла '''
        self.pdf_copy_path = None
        ''' путь до пдф копии файла '''
        self.inside_text = None
        ''' содержимое файла '''
        self.intended_name = None
        ''' предполагаемое название документа '''
        self.extension = None
        ''' расширение файла '''

class Criteria:
    def __init__(self, documents, keywords, index):
        self.documents = documents
        ''' подтверждающие док-ы для данного критерия '''
        self.keywords = keywords
        ''' ключевые слова для данного критерия '''
        self.num_of_criteria = None
        ''' кол-во критериев '''
        self.index = index
        ''' номер данного набора ключ слов '''
        self.sub_index = None
        ''' номер данного ключ слова в наборе ключ слов '''


# ПОДГОТОВКА К РАБОТЕ

df = pd.read_excel(summary_form, header=None)

# Определяем критерий для начала таблицы (например, наличие ключевых слов)
criteria = "№ п/п"

# Ищем первую строку, содержащую ключевое слово
start_row = None
for index, row in df.iterrows():
    if criteria in row.values:
        start_row = index
        break

# Указываем номер строки таблицы, с которой начать считывание, у нас это start_row
df = pd.read_excel(summary_form, header=start_row)
for column in df.columns:
    if 'Unnamed' in column:
        df.drop(column, axis=1, inplace=True)

# Добавляем новые столбцы и сохраняем копию таблицы
df['по данным из анкеты участника'] = None
df['по названию из сводной анкеты'] = None
df['по ключевому слову'] = None
# df['по документам'] = None
df['информация'] = None
'''
1). по столбцу заполн претендентом из заполненной анкеты уч если есть
2). по столбцу подтвр док из исходной сводной анкеты
3). по ключ слову
4). извлеченная инфа (на будущее)
'''

# Создание папочек

# создаём папку для разархивированных файлов (для полной разархивации)
unpacked_dir, unpacked_files_path = dir_creator("unpacked_files")
unpacked_dir_2, unpacked_files_path_2 = dir_creator("unpacked_files_2")
# создаём папку для копий всех файлов (некоторые конвертируем в пдф)
files_copy_dir, pdf_files_path = dir_creator('files_copy')
# создаём папку для табличек (на будущее)
tables_dir, table_files_path = dir_creator("table_files")
# создаём папку для неопознанных файлов (неопознанный формат)
unknown_dir, unknown_files_path = dir_creator("unknown_files")
# создаём папку для файлов с неопознанным текстом
unknown_text_dir, unknown_text_files_path = dir_creator("unknown_text_files")


# РАБОТАЕМ

# раскидываем критерии из анкеты по элементам класса Criteria
criteria_list = []

for i, row in df.iterrows():
    documents = row["Подтверждающие документы"]
    keywords = row["Ключевые слова"]
    index = i
    # штука для пронумерованных столбцов, тут важно чтобы первый символ был 1 иначе не сработает

    keywords = str(keywords)
    if keywords[0].replace(' ', '') == '1':
        lines1 = keywords.split('\n')
        lines2 = documents.split('\n')
        for line1, line2 in zip(lines1, lines2):
            # разделяем строку по закрывающей скобке и берем вторую часть (после закрывающей скобки)
            sub_keyword = line1.split(')', 1)[1]
            # sub_document = line2.split(')', 1)[1]
            sub_index = line1.split(')', 1)[0]
            sub_document = line2.replace(str(sub_index), '', 1)

            obj = Criteria(sub_document, sub_keyword, index)
            obj.sub_index = sub_index
            criteria_list.append(obj)
    else:
        obj = Criteria(documents, keywords, index)
        criteria_list.append(obj)

# раскидываем всё по папочкам
list_of_Files = []

dir_content_sorting(archive_path, unpacked_files_path)

while (os.path.isdir('unpacked_files') and len(os.listdir('unpacked_files'))) or (os.path.isdir('unpacked_files_2') and len(os.listdir('unpacked_files_2'))):
    if (os.path.isdir('unpacked_files') and len(os.listdir('unpacked_files'))):
        dir_content_sorting(unpacked_files_path, unpacked_files_path_2)
        shutil.rmtree(unpacked_files_path)
        unpacked_dir, unpacked_files_path = dir_creator("unpacked_files")
    if (os.path.isdir('unpacked_files_2') and len(os.listdir('unpacked_files_2'))):
        dir_content_sorting(unpacked_files_path_2, unpacked_files_path)
        shutil.rmtree(unpacked_files_path_2)
        unpacked_dir_2, unpacked_files_path_2 = dir_creator("unpacked_files_2")
shutil.rmtree(unpacked_files_path)
shutil.rmtree(unpacked_files_path_2)

# Ищем заполненную претендентом сводную анкету
df_filled_form = None
for dirpath, dirnames, filenames in os.walk(table_files_path):
    for filename in filenames:
        file_name = os.path.splitext(filename)[0]
        file_ext = os.path.splitext(filename)[1]

        if match_finder('сводная анкета', file_name) or match_finder_extend('сводная анкета', file_name):
            if (match_finder('сводная анкета', file_name) or len(match_finder_extend('сводная анкета', file_name)) >= len(
            'сводная анкета') - 1):
                # извлекаем датафрейм из заполненной претендентом анкеты
                df_local = pd.read_excel(f'{archive_path}\{filename}')
                start_row = -1
                for index, row in df_local.iterrows():
                    if criteria in row.values:
                        start_row = index
                        break
                df_filled_form = pd.read_excel(f'{archive_path}\{filename}', header=start_row+1) # ПОЧЕМУУУУ

# если нашлась заполненная претендентом анкета, извлекаем столбец 'Заполняется Претендентом'
if isinstance(df_filled_form, pd.DataFrame):
    filled_by_candidate = dict(df_filled_form['Заполняется Претендентом'])
    for i, row in df.iterrows():
        # переписываем столбец в итоговую копию анкеты
        df.at[i, 'Заполняется Претендентом'] = filled_by_candidate[i]
        if str(filled_by_candidate[i]).lower().replace(' ', '') != str(df['Потребность Заказчика'][i]).lower().replace(' ', ''):
            df.at[i, 'по данным из анкеты участника'] = 'требуется проверка'



# тут если понадобится можно 'не запускать' этот модуль (например если нужно проверить только как всё раскидается по папочкам)
r = 1
if r == 1:
    # идём по элементам класса Files (пдфки + картинки)
    for element in list_of_Files:

        filename_ = os.path.splitext(element.filename)[0]
        print(filename_)

        #  СРАВНЕНИЕ НАЗВАНИЯ ФАЙЛА СО СТОЛБЦОМ "ПОДТВЕРЖДАЮЩИЕ ДОКУМЕНТЫ"
        for el in criteria_list:
            # print('документы', el.documents)
            if match_finder(filename_, el.documents) or match_finder_extend(filename_, el.documents):
                if df['по названию из сводной анкеты'][el.index]:
                    current_value = str(df.at[el.index, 'по ключевому слову'])
                    df.at[el.index, 'по названию из сводной анкеты'] = current_value + '\n' + str(element.pdf_copy_path)
                else:
                    df.at[el.index, 'по названию из сводной анкеты'] = str(element.pdf_copy_path)

            '''
            checker = 2
            print(el.documents, '!!!')
            documents = el.documents
            # print(el.index)
            documents_splitted_1 = re.split(" ИЛИ ", documents)
            for word_1 in documents_splitted_1:

                if checker == 1:
                    break
                # print(word_1)
                documents_splitted_2 = re.split(";", word_1)

                for word_2 in documents_splitted_2:

                    if checker == 0:
                        break
                    # print(word_2)
                    documents_splitted_3 = re.split(" или ", word_2)

                    for word_3 in documents_splitted_3:
                        # print(word_3)
                        if match_finder(filename_, word_3) or match_finder_extend(filename_, word_3):
                            checker = 1
                            break
                        checker = 0

            if checker == 1:
                print('есть совпадение')
                df.at[
                    el.index, 'по названию из сводной анкеты'] = element.pdf_copy_path  # надо заметить название столбика, ошибку пофиксить
                # checker = 2
            '''


        # ИЗВЛЕЧЕНИЕ ТЕКСТА ИЗ ДОКУМЕНТА
        # print(element.pdf_copy_path)
        if element.extension in ('.pdf', '.PDF'):
            try:
                text = parse_pdf(element.pdf_copy_path)
                # print(text)
            except:
                text = None
                shutil.copy(element.pdf_copy_path, unknown_text_files_path)

        else:  # png, jpg, jpeg
            try:
                text = image_to_text(element.pdf_copy_path)
            except:
                text = None
                shutil.copy(element.pdf_copy_path, unknown_text_files_path)


        # СРАВНЕНИЕ СО СТОЛБЦОМ ЗАПОЛНЕННЫМ УЧАСТНИКОМ
        for i, row in df.iterrows():
            # if df['по данным из анкеты участника'][i] == 'требуется проверка':
            if df['по данным из анкеты участника'][i]:
                match_check = match_finder_extend(str(df['Заполняется Претендентом'][i]), str(text))
                if (not isinstance(match_check, bool)) and len(match_check) >= len(str(df['Заполняется Претендентом'][i])) - 1:
                    current_value = str(df['по данным из анкеты участника'][i])
                    if current_value == 'требуется проверка':
                        df.at[i, 'по данным из анкеты участника'] = element.pdf_copy_path
                    else:
                        df.at[i, 'по данным из анкеты участника'] = current_value + '\n' + str(element.pdf_copy_path)


        # СРАВНЕНИЕ С КЛЮЧЕВЫМИ СЛОВАМИ
        for el in criteria_list:
            if el.keywords.lower().replace(' ','') == 'безключевыхслов' or el.keywords.lower().replace(' ','') == 'nan':  # тут мб лематизировать левую часть чтоб наверняка? мб убрать пробелы
                print('НЕТ КЛЮЧЕВЫХ СЛОВ')
            else:
                checker = 2
                print(el.keywords, '!!!')
                keywords = el.keywords
                # print(el.index)
                keywords_splitted_1 = re.split(" ИЛИ ", keywords)
                for word_1 in keywords_splitted_1:

                    if checker == 1:
                        break
                    # print(word_1)
                    keywords_splitted_2 = re.split(";", word_1)

                    for word_2 in keywords_splitted_2:

                        if checker == 0:
                            break
                        # print(word_2)
                        keywords_splitted_3 = re.split(" или ", word_2)

                        for word_3 in keywords_splitted_3:
                            # print(word_3)
                            match_check = match_finder(str(word_3), str(text))
                            if match_check:
                                checker = 1
                                break
                            checker = 0

                if checker == 1:
                    print('есть совпадение')
                    if (df.at[el.index, 'по ключевому слову']):
                        current_value = str(df.at[el.index, 'по ключевому слову'])
                        new_value = current_value +'\n'+ str(element.pdf_copy_path)
                        df.at[el.index, 'по ключевому слову'] = new_value
                    else:
                        df.at[el.index, 'по ключевому слову'] = str(element.pdf_copy_path)


# Сохраняем изменения в КОПИЮ таблицы
df.to_excel('анкета_копия.xlsx', index=False)