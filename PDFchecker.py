import fitz
import PyPDF2
from past.builtins import raw_input
from pdf2docx import parse
from typing import Tuple
import zipfile
from xml.etree import ElementTree
import os

files = os.listdir()

pdf_file = None
for file in files:
    if file.endswith(".pdf"):
        pdf_file = file
        break

if pdf_file:
    parse(file, "example.docx")
    print('!!!ВНИМАНИЕ!!!')
    print('Программа может не распознать некоторые графические элементы.')
    print('Пожалуйста, сверьте количество изображений в документе с результатом программы.')
    print('Погрешность размеров графических элементов: 0.1 - 0.3 миллиметра.')
    pdf_reader = PyPDF2.PdfReader(open(pdf_file, 'rb'))  # открываем pdf файл для чтения
    num_pages = len(pdf_reader.pages)  # получаем количество страниц в pdf файле
    total_chars = 0  # общее количество символов
    total_lines = 0  # общее количество строк
    full_total_area = 0 # общая площадь изображений по документу
    doc = fitz.open(pdf_file)  # открываем pdf файл для обработки изображений

    for i in range(num_pages):  # перебираем все страницы в pdf файле
        page = pdf_reader.pages[i]  # получаем текущую страницу
        text = page.extract_text()  # извлекаем текст со страницы
        num_chars = len(text)  # подсчитываем количество символов в тексте
        num_lines = len(text.splitlines())  # подсчитываем количество строк в тексте
        total_chars += num_chars  # добавляем к общему количеству символов
        total_lines += num_lines  # добавляем к общему количеству строк

    with zipfile.ZipFile('example.docx') as docx:
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
                full_total_area += this_img_area
                print(f'Изображение {img_counter}: {width:.2f} x {height:.2f} см. Площадь: {this_img_area:.2f} кв. сантиметров.')
                img_counter += 1
    os.remove("example.docx")

    print(f'Количество символов: {total_chars}')
    print(f'Количество строк: {total_lines}')
    print(f'Площадь графических материалов по всему документу: {round(full_total_area, 2)} кв. сантиметров.')
    print('----------------------------------------------')
else:
    print('Отсутствуют pdf файлы в папке программы...')

print('Нажмите Enter, чтобы закрыть консоль...')
raw_input()