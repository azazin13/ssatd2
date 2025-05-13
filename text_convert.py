import PyPDF2
import pytesseract
from pdfminer.high_level import extract_pages
from pdfminer.layout import LTTextContainer, LTRect, LTFigure
import pdfplumber
from pdf2image import convert_from_path
from pytesseract import Output
from tables import *
import pandas as pd

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
poppler_path = r'C:\Program Files\poppler-24.02.0\Library\bin'

# конвертация в строку текст из текстового элемента LTTextContainer
def text_extraction(element):
    return element.get_text()

# вырезание изображения в pdf-файле и сохранение pdf-файла только с обрезанным изображением
def crop_image(element, pageObj):
    [image_left, image_top, image_right, image_bottom] = [element.x0, element.y0, element.x1, element.y1]
    pageObj.mediabox.lower_left = (image_left, image_bottom)
    pageObj.mediabox.upper_right = (image_right, image_top)
    cropped_pdf_writer = PyPDF2.PdfWriter()
    cropped_pdf_writer.add_page(pageObj)
    with open('cropped_image.pdf', 'wb') as cropped_pdf_file:
        cropped_pdf_writer.write(cropped_pdf_file)

# конвертация pdf в png
def convert_to_images(input_file):
    images = convert_from_path(input_file, poppler_path=poppler_path)
    image = images[0]
    output_file = "PDF_image.png"
    image.save(output_file, "PNG")

# поворот изображения в зависимости от ориентации
def rotate_images(img):
    result = pytesseract.image_to_osd(img, output_type=Output.DICT,
                                      config='--psm 0 -c min_characters_to_try=5')
    if result['orientation'] == 90 or result['orientation'] == 180 or result['orientation'] == 270:
        rotated_image = cv2.rotate(img, cv2.ROTATE_90_CLOCKWISE if result['orientation'] == 270 else cv2.ROTATE_180 if
        result['orientation'] == 180 else cv2.ROTATE_90_COUNTERCLOCKWISE)
        return rotated_image
    else:
        return img

# конвертация изображение с текстом в строку
def image_to_text(image_path):
    image = cv2.imread(image_path, cv2.IMREAD_GRAYSCALE)
    adjusted_image = cv2.convertScaleAbs(image, alpha=2, beta=1.2)
    text = pytesseract.image_to_string(adjusted_image, lang='rus+eng')
    if text != '':
        rotated_image = rotate_images(adjusted_image)
        text = pytesseract.image_to_string(rotated_image, lang='rus+eng')
    return text

# извлечение таблицы с соответствующим номером из страницы
def extract_table(pdf_path, page_num, table_num):
    pdf = pdfplumber.open(pdf_path)
    table_page = pdf.pages[page_num]
    table = table_page.extract_tables()[table_num]
    pdf.close()
    return table

# конвертация текстовой таблицы в строку с разделителями и переносами
def table_converter(table):
    table_string = ''

    for row_num in range(len(table)):
        row = table[row_num]
        cleaned_row = [
            item.replace('\n', ' ') if item is not None and '\n' in item else 'None' if item is None else item for item
            in row]
        table_string += ('|' + '|'.join(cleaned_row) + '|' + '\n')
    return table_string[:-1]

# анализ pdf-файла
def parse_pdf(pdf_path):
    global lower_side, upper_side
    pdfFileObj = open(pdf_path, 'rb')
    pdfReader = PyPDF2.PdfReader(pdfFileObj)
    text_per_page = {}

    # итерирование по номерам страниц
    for pagenum, page in enumerate(extract_pages(pdf_path)):

        pageObj = pdfReader.pages[pagenum]
        page_content = []
        table_num = 0
        first_element = True
        table_extraction_flag = False
        table_index = 1

        pdf = pdfplumber.open(pdf_path)
        page_tables = pdf.pages[pagenum]
        tables = page_tables.find_tables()

        page_elements = [(element.y1, element) for element in page._objs]
        page_elements.sort(key=lambda a: a[0], reverse=True)

        # итерирование по элементам на странице
        for i, component in enumerate(page_elements):
            element = component[1]

            # if для текстового элемента
            if isinstance(element, LTTextContainer):
                if not table_extraction_flag:
                    line_text = text_extraction(element)
                    page_content.append(line_text + '\n')
                else:
                    pass

            # if для элемента, содержащего изображение
            if isinstance(element, LTFigure):
                crop_image(element, pageObj)
                convert_to_images('cropped_image.pdf')
                image_text = image_to_text('PDF_image.png')
                if image_text != '':
                    image = cv2.imread("PDF_image.png", cv2.IMREAD_GRAYSCALE)
                    cv2.imwrite("PDF_image.png", rotate_images(image))
                    cells = find_table_in_boxes()
                    # build_lines()
                    if len(cells):
                        extract_table_from_image('PDF_image.png', r'table %s.xlsx' % table_index)
                        if not len(pd.read_excel(r'table %s.xlsx' % table_index)):
                            os.remove(r'table %s.xlsx' % table_index)
                            page_content.extend(image_text)
                        table_index += 1
                    else:
                        page_content.extend(image_text)
                else:
                    pass

            # if для элемента, содержащего таблицу
            if isinstance(element, LTRect):
                if len(tables):
                    if first_element and (table_num + 1) <= len(tables):
                        lower_side = page.bbox[3] - tables[table_num].bbox[3]
                        upper_side = element.y1
                        table = extract_table(pdf_path, pagenum, table_num)
                        table_string = table_converter(table)
                        page_content.append(table_string)
                        table_extraction_flag = True
                        first_element = False

                    if element.y0 >= lower_side and element.y1 <= upper_side:
                        pass
                    elif i == len(page_elements) - 1:
                        pass
                    elif not isinstance(page_elements[i + 1][1], LTRect):
                        table_extraction_flag = False
                        first_element = True
                        table_num += 1

        key = 'Page_' + str(pagenum)
        text_per_page[key] = page_content # словарь со строками, содержащими текст из каждой страницы, ключ - номер страницы

    pdfFileObj.close()

    if os.path.isfile('cropped_image.pdf'):
        os.remove('cropped_image.pdf')
    if os.path.isfile('PDF_image.png'):
        os.remove('PDF_image.png')

    result = '\n'.join([''.join(text_per_page[page]) for page in text_per_page])
    print(result)

# тут можно отдельно проверить работу конвертатора на конкретном файле
if __name__ == '__main__':
    parse_pdf('pdf_name.pdf')
