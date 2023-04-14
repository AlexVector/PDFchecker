import fitz
import PyPDF2
from past.builtins import raw_input
from pdf2docx import parse
from typing import Tuple
import zipfile
from xml.etree import ElementTree
import os
import camelot

files = os.listdir()

pdf_file = None
for file in files:
    if file.endswith(".pdf"):
        pdf_file = file
        break

if pdf_file:
    parse(file, "КОНВЕРТИРОВАНО.docx")
    print('!!!ВНИМАНИЕ!!!')
    print('Программа может не распознать некоторые графические элементы и таблицы.')
    print('Пожалуйста, сверьте количество изображений и таблиц в документе с результатом программы.')
    print('Погрешность размеров графических элементов и таблиц: 0.1 - 0.3 миллиметра.')
    pdf_reader = PyPDF2.PdfReader(open(pdf_file, 'rb'))  # открываем pdf файл для чтения
    total_chars = 0  # общее количество символов
    total_lines = 0  # общее количество строк
    full_total_img_area = 0 # общая площадь изображений по документу
    full_total_tab_area = 0 # общая площадь изображений по документу
    doc = fitz.open(pdf_file)  # открываем pdf файл для обработки изображений

    for page in range(len(pdf_reader.pages)):  # перебираем все страницы в pdf файле
        page_text = pdf_reader.pages[page].extract_text()
        lines = page_text.split('\n')
        for line in lines:
            total_lines += 1
            total_chars += len(line)

        tables = camelot.read_pdf(pdf_file, pages=str(page + 1)) # открываем файл через специальную библиотеку для извлечения таблиц
        # Цикл для поиска таблиц на каждом листе
        for table in tables:
            x1, y1, x2, y2 = table._bbox # Получение координат углов внешней границы таблицы в пиклесях
            width = x2 - x1 # Вычисление ширины таблицы в пикселях
            height = y2 - y1 # Вычисление высоты таблицы в пикселях
            width_cm = width * 0.035277778 # Перевод ширины из пикселей в сантиметры
            height_cm = height * 0.035277778 # Перевод высоты из пикселей в сантиметры
            table_area = width_cm * height_cm # Вычисление площади таблицы в сантиметрах квадратных
            full_total_tab_area += table_area
            print(f'Таблица на странице {page + 1}: {width_cm:.2f} см. x {height_cm:.2f} см. Площадь: {table_area:.2f} кв. сантиметров.') # Вывод результатов и площади таблицы


    with zipfile.ZipFile('КОНВЕРТИРОВАНО.docx') as docx:
        # Extract the document.xml file from the zip file
        with docx.open('word/document.xml') as f:
            # Parse the XML content
            tree = ElementTree.parse(f)
            root = tree.getroot()
            # Define the XML namespaces
            ns = {
                'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
                'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
                'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
                'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture'
            }
            # Find all inline images in the document
            img_counter = 1
            for pic in root.findall('.//wp:inline/a:graphic/a:graphicData/pic:pic', ns):
                # Get the image size in EMUs (English Metric Units)
                cx = int(pic.find('pic:spPr/a:xfrm/a:ext', ns).get('cx'))
                cy = int(pic.find('pic:spPr/a:xfrm/a:ext', ns).get('cy'))
                # Convert the size from EMUs to centimeters (1 cm = 360000 EMUs)
                width = cx / 360000
                height = cy / 360000
                this_img_area = width * height
                full_total_img_area += this_img_area
                print(f'Изображение {img_counter}: {width:.2f} x {height:.2f} см. Площадь: {this_img_area:.2f} кв. сантиметров.')
                img_counter += 1
    #os.remove("example.docx")

    print(f'Количество символов: {total_chars}')
    print(f'Количество строк: {total_lines}')
    print(f'Площадь графических материалов по всему документу: {round(full_total_img_area, 2)} кв. сантиметров.')
    print(f'Площадь таблиц по всему документу: {round(full_total_tab_area, 2)} кв. сантиметров.')
    print(f'ВНИМАНИЕ! Если таблицы в документе сохранены как картинки, то они посчитаются и в площади граф. '
          f'материалов, и в площади таблиц.')
    print('----------------------------------------------')
else:
    print('Отсутствуют pdf файлы в папке программы...')

print('Нажмите Enter, чтобы закрыть консоль...')
raw_input()