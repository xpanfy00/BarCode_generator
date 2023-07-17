from reportlab.lib.pagesizes import portrait, landscape
from reportlab.lib.units import mm
import textwrap
import openpyxl
from reportlab.graphics import barcode
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas


def read_excel_file(file_path):
    wb = openpyxl.load_workbook(file_path)
    sheet = wb.active

    data = []
    headers = [cell.value for cell in sheet[1]]

    for row in sheet.iter_rows(min_row=2, values_only=True):
        item = {}
        for header, value in zip(headers, row):
            item[header] = value
        data.append(item)

    return data


def generate_barcode(barcode_value, text, conv, opis, brend, name, sost, color):
    pdfmetrics.registerFont(TTFont('ArialUnicodeMS', 'ArialUnicodeMS.ttf'))


    # Рисуем текст над штрих-кодом
    text_width = conv.stringWidth(text, "ArialUnicodeMS", 10)
    conv.setFont("ArialUnicodeMS", 10)
    barcode_width = 60 * mm  # Ширина штрих-кода
    barcode_height = 30 * mm  # Высота штрих-кода
    text_x = 50 * mm + (barcode_width - text_width) / 2  # Вычисляем координату x текста
    text_y = 50 * mm + barcode_height + 5 * mm  # Координата y текста
    conv.drawString(text_x, text_y, text)

    # Рисуем штрих-код
    barcode_width = 60 * mm  # Ширина штрих-кода
    barcode_height = 30 * mm  # Высота штрих-кода
    barcode_x = 50 * mm
    barcode_y = 50 * mm
    barcode_code = barcode.createBarcodeDrawing('EAN13', value=barcode_value, format='png', width=barcode_width,
                                                height=barcode_height)
    barcode_code.drawOn(conv, barcode_x, barcode_y)

    # Рисуем текст слева снизу
    left_text = "Арт: " + text
    left_text_x = barcode_x
    left_text_y = barcode_y - 10 * mm
    conv.setFont("ArialUnicodeMS", 10)
    conv.drawString(left_text_x, left_text_y, left_text)

    # Рисуем текст слева снизу
    left_text = "Состав:" + sost
    left_text_x = barcode_x
    left_text_y = barcode_y - 15 * mm
    conv.setFont("ArialUnicodeMS", 10)
    conv.drawString(left_text_x, left_text_y, left_text)

    # Рисуем текст слева снизу
    left_text = "Прод: " + name
    left_text_x = barcode_x
    left_text_y = barcode_y - 20 * mm
    conv.setFont("ArialUnicodeMS", 10)
    conv.drawString(left_text_x, left_text_y, left_text)

    # Рисуем текст справа снизу
    right_text = "Цвет: " + color
    right_text_width = conv.stringWidth(right_text, "ArialUnicodeMS", 10)
    right_text_x = barcode_x + barcode_width - right_text_width
    right_text_y = barcode_y - 10 * mm
    conv.setFont("ArialUnicodeMS", 10)
    conv.drawString(right_text_x, right_text_y, right_text)

    # Рисуем текст справа снизу
    right_text = " " + brend
    right_text_width = conv.stringWidth(right_text, "ArialUnicodeMS", 10)
    right_text_x = barcode_x + barcode_width - right_text_width
    right_text_y = barcode_y - 15 * mm
    conv.setFont("ArialUnicodeMS", 10)
    conv.drawString(right_text_x, right_text_y, right_text)

    # Рисуем текст под штрих-кодом
    text = opis
    text_width = conv.stringWidth(text, "ArialUnicodeMS", 10)
    text_x = barcode_x + (barcode_width - text_width) /2 + 57 * mm
    text_y = barcode_y - barcode_height +1 * mm
    conv.setFont("ArialUnicodeMS", 10)
    max_text_width = barcode_width +220
    lines = textwrap.wrap(text, width=int(max_text_width / 10))
    line_height = 12  # Высота строки
    for line in lines:
        conv.drawString(text_x, text_y, line)
        text_y -= line_height





data = read_excel_file("Книга1.xlsx")
pdf_file = "barcode.pdf"
barcode_width = 100 * mm
barcode_height = 160 * mm
conv = canvas.Canvas(pdf_file, pagesize=landscape((barcode_width, barcode_height)))

for item in data:
    barcode_value = item["Штрихкод"]
    text = item["Артикул"]
    opis = item["Описание"]
    brend = item["Бренд"]
    name = item["Название"]
    sost = item["Состав"]
    color = item["Цвет"]
    col = item["Кол-во"]
    for _ in range(col):
        generate_barcode(barcode_value, text, conv, opis, brend, name, sost, color)
        conv.showPage()

conv.save()
