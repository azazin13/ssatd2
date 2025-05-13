import os
import cv2
from img2table.ocr import TesseractOCR
from img2table.document import Image

# pre_file = os.path.join("data", "pre.png")
# out_file = os.path.join("data", "out.png")

# конвертация изображения, содержащего таблицу, в .xlsx
def extract_table_from_image(src, path):
    ocr = TesseractOCR(n_threads=1, lang="rus", tessdata_dir=r'C:\Program Files\Tesseract-OCR\tessdata')
    doc = Image(src)
    doc.to_xlsx(dest=path, ocr=ocr, implicit_rows=False, borderless_tables=False,  min_confidence=90)

# предобработка изображения, изменение контрастности для нахождения текстовых элементов и в дальнейшем контуров таблицы на изображении
def pre_process_image(img, save_in_file=None, morph_size=(8, 8)):
    pre = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    pre = cv2.threshold(pre, 250, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)[1]
    cpy = pre.copy()
    struct = cv2.getStructuringElement(cv2.MORPH_RECT, morph_size)
    cpy = cv2.dilate(~cpy, struct, anchor=(-1, -1), iterations=1)
    pre = ~cpy
    if save_in_file is not None:
        cv2.imwrite(save_in_file, pre)
    return pre

# получение контуров, ограничивающих текст, на основе предположений о размере текста
def find_text_boxes(min_text_height_limit=6, max_text_height_limit=40):
    in_file = r'PDF_image.png'
    img_table = cv2.imread(os.path.join(in_file))
    pre = pre_process_image(img_table)

    contours, hierarchy = cv2.findContours(pre, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
    boxes = []
    for contour in contours:
        box = cv2.boundingRect(contour)
        h = box[3]
        if min_text_height_limit < h < max_text_height_limit:
            boxes.append(box)
    return boxes

# сортировка ячеек в таблице по их расположению
def find_table_in_boxes(cell_threshold=10, min_columns=2):
    boxes = find_text_boxes()
    rows = {}
    cols = {}

    for box in boxes:
        (x, y, w, h) = box
        col_key = x // cell_threshold
        row_key = y // cell_threshold
        cols[row_key] = [box] if col_key not in cols else cols[col_key] + [box]
        rows[row_key] = [box] if row_key not in rows else rows[row_key] + [box]

    table_cells = list(filter(lambda r: len(r) >= min_columns, rows.values()))
    table_cells = [list(sorted(tb)) for tb in table_cells]
    table_cells = list(sorted(table_cells, key=lambda r: r[0][1]))

    return table_cells

# визуализация контуров таблицы на изображении (для анализа pdf не обязательно)
# просто рисует контуры и сохраняет пнг

# def build_lines():
#     img_table = cv2.imread(os.path.join(in_file))
#     table_cells = find_table_in_boxes()
#     if table_cells is None or len(table_cells) <= 0:
#         return [], []
#
#     max_last_col_width_row = max(table_cells, key=lambda b: b[-1][2])
#     max_x = max_last_col_width_row[-1][0] + max_last_col_width_row[-1][2]
#
#     max_last_row_height_box = max(table_cells[-1], key=lambda b: b[3])
#     max_y = max_last_row_height_box[1] + max_last_row_height_box[3]
#
#     hor_lines = []
#     ver_lines = []
#
#     for box in table_cells:
#         x = box[0][0]
#         y = box[0][1]
#         hor_lines.append((x, y, max_x, y))
#
#     for box in table_cells[0]:
#         x = box[0]
#         y = box[1]
#         ver_lines.append((x, y, x, max_y))
#
#     (x, y, w, h) = table_cells[0][-1]
#     ver_lines.append((max_x, y, max_x, max_y))
#     (x, y, w, h) = table_cells[0][0]
#     hor_lines.append((x, max_y, max_x, max_y))
#     for line in hor_lines:
#         [x1, y1, x2, y2] = line
#         cv2.line(img_table, (x1, y1), (x2, y2), (0, 0, 255), 1)
#
#     for line in ver_lines:
#         [x1, y1, x2, y2] = line
#         cv2.line(img_table, (x1, y1), (x2, y2), (0, 0, 255), 1)
#
#     cv2.imwrite(out_file, img_table)
