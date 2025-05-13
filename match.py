import stanza
nlp = stanza.Pipeline('ru', processors='tokenize,pos,lemma')
import re


# функции для анализа текста
def lemmatize_text(text):
    '''
    Функция для лемматизации текста
    :param text: исходный текст
    :return: words: список всех слов в текст,
    lemmatized_words: лемматизированные слова из теста записанные в список,
    lemmatized_text: лемматизированные слова из теста записанные в строку
    '''
    text = str(text)
    doc = nlp(text)
    # разбить предложения на слова и создать список в исходном порядке
    words = []
    for sentence in doc.sentences:
        for word in sentence.words:
            if word.upos != 'PUNCT':
                words.append(word.text.lower())
    # лемматизировать каждое слово и создать новый список
    lemmatized_words = [word.lemma.lower() for sentence in doc.sentences for word in sentence.words if word.upos != 'PUNCT']
    # создать из лемматизированных слов строку, разделенную пробелами
    lemmatized_text = ' '.join(lemmatized_words)
    return words, lemmatized_words, lemmatized_text

def full_match(key_str, main_str, lower=0):
    '''
    Функция для поиска подстроки в строке
    :param key_str: искомая строка
    :param main_str: строка в которой ищем
    :param lower: учитывать ли регистр
    :return: позиции в главной строке первого и последнего символов искомой подстроки,
    или None, если совпадение не найдено
    '''
    key_str = str(key_str)
    main_str = str(main_str)
    # необходимо ли учитывать заглавные буквы
    if lower:
        key_str = key_str.lower()
        main_str = main_str.lower()
    len_key = len(key_str)
    # поиск вхождения
    if key_str in main_str:
        for i in range(len(main_str)-len_key+1):
            if main_str[i:i+len_key] == key_str:
                return i, i+len_key
    else:
        return None

def match_finder(key_str, main_str):
    '''
    Функция для поиска подстроки в строке с учетом различных форм слов
    :param key_str: искомая строка
    :param main_str: строка в которой ищем
    :return: позиции в главной строке первого и последнего слова искомой подстроки и
    строка в той форме, в которой она содержится в исходной строке,
    или None, если совпадение не найдено
    '''
    key_str = str(key_str)
    main_str = str(main_str)
    if full_match(key_str, main_str, lower=1):
        return 'фулл метч', full_match(key_str, main_str, lower=1)
    else:
        words_main, lemm_words_main, lemm_text_main = lemmatize_text(main_str)
        words_key, lemm_words_key, lemm_text_key = lemmatize_text(key_str)
        if full_match(lemm_text_key, lemm_text_main):
            from_, to_ = full_match(lemm_text_key, lemm_text_main)
            founded_list = lemm_text_main[from_: to_].split(' ')
            for i in range(len(lemm_words_main)-len(founded_list)):
                if lemm_words_main[i:i+len(founded_list)] == founded_list:
                    return words_main[i:i+len(founded_list)], f'искомые слова идут от {i+1} до {i+len(founded_list)} слова в строке'
        else:
            return None

def match_finder_extend(text_1, text_2):
    '''
    Функция для поиска наибольшей совпадающей строки в двух больших строках
    :param text_1: одна большая строка
    :param text_2: другая большая строка
    :return: наибольшая совпадающая строка,
    или False, если совпадение не найдено
    '''
    text_1 = str(text_1)
    text_2 = str(text_2)
    # определяем меньшую строку
    if len(text_1) < len(text_2):
        text1, text2 = text_1, text_2
    else:
        text1, text2 = text_2, text_1
    # text2_len = len(text2)
    # удаляем знаки из строки
    text2 = re.sub(r'[",.,:,;,,,!,-]', '', text2)
    # ищем самое длинное совпадение
    for i in range(len(text1), 1, -1):
        for j in range(len(text1)-i+1):
            piece = re.sub(r'[",.,:,;,,,!,-]', '', text1[j: j+i])
            if full_match(piece, text2, lower=1):
                # print(text1[j: j+i], full_match(piece, text2, lower=1))
                # return text1[j: j+i], (i)/len(text1), (i)/text2_len
                return text1[j: j+i]
                # return True
            else:
                # ??
                return False
