import fitz
import PyPDF2
from PIL import Image
from past.builtins import raw_input
import os

files = os.listdir()

pdf_file = None
for file in files:
    if file.endswith(".pdf"):
        pdf_file = file
        break

if pdf_file:
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
        page_img = doc.load_page(i) # загружаем текущую страницу
        image_list = page_img.get_images()  # получаем список изображений
        total_area = 0  # обнуляем суммарную площадь
        if image_list:
            for image in image_list:  # проходим по всем изображениям
                xref = image[0]  # получаем ссылку на изображение
                width = image[2]  # получаем ширину изображения в пунктах
                height = image[3]  # получаем высоту изображения в пунктах
                print("Изображение", xref, "размеры:", width, "x", height, "пикселей,")  # выводим результат
                width_cm = width * 0.02635872298  # переводим ширину в сантиметры
                height_cm = height * 0.02635872298  # переводим высоту в сантиметры
                area_cm = width_cm * height_cm  # рассчитываем площадь в квадратных сантиметрах
                total_area += area_cm
                print("Изображение", xref, "размеры:", round(width_cm, 2), "x", round(height_cm, 2), "см,", "площадь:",
                      round(area_cm, 2), "кв.см")  # выводим результат
                print(f"Площадь графических материалов на странице {i + 1} составляет {round(total_area, 2)} кв.см.")  # выводим результат на экран
                full_total_area += total_area;

    print(f'Количество символов: {total_chars}')
    print(f'Количество строк: {total_lines}')
    print(f'Площадь графических материалов по всему документу: {full_total_area}')
    print('----------------------------------------------')
else:
    print('Отсутствуют pdf файлы в папке программы...')
print('Нажмите Enter, чтобы закрыть консоль...')
raw_input()
