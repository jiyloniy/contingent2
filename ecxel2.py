import os
import re
import os
import re

import django

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'config.settings')

django.setup()
from user.models import Faculty, Budjet, Shartnoma, Organization, Yonalish, Guruh

from datetime import datetime

from django.db.models import Sum
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill

from openpyxl.utils import get_column_letter


def set_cell_properties(cell, value, alignment, font, border):
    cell.value = value
    cell.alignment = alignment
    cell.font = font
    cell.border = border


import django

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'config.settings')

django.setup()
from user.models import Faculty, Budjet, Shartnoma, Organization, Yonalish, Guruh

org = Organization.objects.filter(name='kiuf').first()

FONT_NAME = 'Times New Roman'
FONT_SIZE = 8
FONT_COLOR = 'FF000000'
BORDER_STYLE = 'thin'
BORDER_COLOR = 'FF000000'
red_color = 'FF0000FF'
# color blue
blue_color = 'FFFF0000'
wb = Workbook()
ws = wb.active
# add auto size width of column

# Hozirgi vaqtni olib, formatini belgilash

now = datetime.now()
formatted_time = now.strftime("%Y-%m-%d")


def exporttoexcel(org):
    kurss = Guruh.objects.filter(org=org).values('kurs').distinct()
    organization_name = org.full_name
    ws.merge_cells('A1:P2')
    set_cell_properties(ws.cell(row=1, column=1),
                        f"{organization_name} talabalari kontingentining {formatted_time} holati haqida umumiy ma'lumot (o'zbek /rus)",
                        Alignment(horizontal='center', vertical='center'),
                        Font(name=FONT_NAME, size=18, bold=True, italic=False, color=FONT_COLOR),
                        Border(left=Side(border_style=BORDER_STYLE, color=BORDER_COLOR),
                               right=Side(border_style=BORDER_STYLE, color=BORDER_COLOR),
                               top=Side(border_style=BORDER_STYLE, color=BORDER_COLOR),
                               bottom=Side(border_style=BORDER_STYLE, color=BORDER_COLOR)))

    cell_properties = [
        {'row': 3, 'column': 1, 'value': '№', 'width': True, 'merge': 'A3:A4'},
        {'row': 3, 'column': 2, 'value': 'Ta\'lim yo\'nalishi kodi va nomi', 'width': True, 'merge': 'B3:B4'},
        {'row': 3, 'column': 3, 'value': 'Ta\'lim turi', 'width': True, 'merge': 'C3:C4'},
        {'row': 3, 'column': 4, 'value': 'Jami', 'width': True, 'merge': 'D3:D4'},
        {'row': 3, 'column': 5, 'value': 'Jami', 'width': True, 'merge': 'E3:F3'},
        {'row': 4, 'column': 5, 'value': 'O\'zbek', 'width': True},
        {'row': 4, 'column': 6, 'value': 'Rus', 'width': True},
        {'row': 3, 'column': 7, 'value': '1-kurs', 'width': True, 'merge': 'G3:I3'},
        {'row': 4, 'column': 7, 'value': 'Jami', 'width': True},
        {'row': 4, 'column': 8, 'value': 'O\'zbek', 'width': True},
        {'row': 4, 'column': 9, 'value': 'Rus', 'width': True},
        {'row': 3, 'column': 10, 'value': '2-kurs', 'width': True, 'merge': 'J3:L3'},
        {'row': 4, 'column': 10, 'value': 'Jami', 'width': True},
        {'row': 4, 'column': 11, 'value': 'O\'zbek', 'width': True},
        {'row': 4, 'column': 12, 'value': 'Rus', 'width': True},
        {'row': 3, 'column': 13, 'value': '3-kurs', 'width': True, 'merge': 'M3:O3'},
        {'row': 4, 'column': 13, 'value': 'Jami', 'width': True},
        {'row': 4, 'column': 14, 'value': 'O\'zbek', 'width': True},
        {'row': 4, 'column': 15, 'value': 'Rus', 'width': True},
        {'row': 3, 'column': 16, 'value': '4-kurs', 'width': True, 'merge': 'P3:R3'},
        {'row': 4, 'column': 16, 'value': 'Jami', 'width': True},
        {'row': 4, 'column': 17, 'value': 'O\'zbek', 'width': True},
        {'row': 4, 'column': 18, 'value': 'Rus', 'width': True},
        {'row': 3, 'column': 19, 'value': '4-kurs', 'width': True, 'merge': 'S3:U3'},
        {'row': 4, 'column': 19, 'value': 'Jami', 'width': True},
        {'row': 4, 'column': 20, 'value': 'O\'zbek', 'width': True},
        {'row': 4, 'column': 21, 'value': 'Rus', 'width': True},
        {'row': 3, 'column': 22, 'value': '5-kurs', 'width': True, 'merge': 'V3:X3'},
        {'row': 4, 'column': 22, 'value': 'Jami', 'width': True},
        {'row': 4, 'column': 23, 'value': 'O\'zbek', 'width': True},
        {'row': 4, 'column': 24, 'value': 'Rus', 'width': True},
        {'row': 3, 'column': 25, 'value': '5-kurs', 'width': True, 'merge': 'Y3:AA3'},
        {'row': 4, 'column': 25, 'value': 'Jami', 'width': True},
        {'row': 4, 'column': 26, 'value': 'O\'zbek', 'width': True},
        {'row': 4, 'column': 27, 'value': 'Rus', 'width': True},
        {'row': 3, 'column': 28, 'value': '6-kurs', 'width': True, 'merge': 'AB3:AD3'},
        {'row': 4, 'column': 28, 'value': 'Jami', 'width': True},
        {'row': 4, 'column': 29, 'value': 'O\'zbek', 'width': True},
        {'row': 4, 'column': 30, 'value': 'Rus', 'width': True},
    ]


    for properties in cell_properties:
        cell = ws.cell(row=properties['row'], column=properties['column'])
        set_cell_properties(cell,
                            properties['value'],
                            Alignment(horizontal='center', vertical='center'),
                            Font(name=FONT_NAME, size=FONT_SIZE, bold=True, italic=False, color=FONT_COLOR),
                            Border(left=Side(border_style=BORDER_STYLE, color=BORDER_COLOR),
                                   right=Side(border_style=BORDER_STYLE, color=BORDER_COLOR),
                                   top=Side(border_style=BORDER_STYLE, color=BORDER_COLOR),
                                   bottom=Side(border_style=BORDER_STYLE, color=BORDER_COLOR)))
        # if 'width' in properties and properties['width']:
        #     ws.column_dimensions[get_column_letter(properties['column'])].auto_size = True
        if 'merge' in properties:
            ws.merge_cells(properties['merge'])

    # turi_choices = (
    #     ('Kunduzgi', 'Kunduzgi'),
    #     ('Sirtqi', 'Sirtqi'),
    #     ('Masofaviy', 'Masofaviy'),
    # )
    yonalishlar = Yonalish.objects.filter(org=org)
    row = 5
    yonalish_kunduzgi = yonalishlar.filter(turi='Kunduzgi')
    yonalish_sirtqi = yonalishlar.filter(turi='Sirtqi')
    yonalish_masofaviy = yonalishlar.filter(turi='Masofaviy')
    yonalsih_bakalavr = yonalishlar.filter(yonalishguruh__bosqich='Bakalavr')
    yonalsih_magistr = yonalishlar.filter(yonalishguruh__bosqich='Magistratura')
    yonalsih_doktorantura = yonalishlar.filter(yonalishguruh__bosqich='Doktorantura')

#     1-talim turi misol uchun kunduzgi, sirqi,masofaviy
#     2-jami shu yo'nalishga bo'g'langan guruhlar bo'yicha
#     3 - jami o'zbeklar va ruslar ning soni
#     4 - 1-kurs jami rus +o'zbek
#     5- 1-kurs o'zbek jami
#     6- 1-kurs rus jami
# kunduzgi
    for yonalish in yonalish_kunduzgi:
        guruhlar = Guruh.objects.filter(org=org, yonalish=yonalish)
        jami = guruhlar.count()
        jami_o_uzbek = guruhlar.filter(yonalish__language='O\'zbek').count()
        jami_rus = guruhlar.filter(yonalish__language='Rus').count()
        kurslar = guruhlar.values('kurs').distinct()
        for kurs in kurss:
            jami_kurs = guruhlar.filter(kurs=kurs['kurs']).count()
            jami_o_uzbek_kurs = guruhlar.filter(kurs=kurs['kurs'], yonalish__language='O\'zbek').count()
            jami_rus_kurs = guruhlar.filter(kurs=kurs['kurs'], yonalish__language='Rus').count()
            ws.cell(row=row, column=1, value=row - 4)
            ws.cell(row=row, column=2, value=yonalish.name)
            ws.cell(row=row, column=3, value='Kunduzgi')
            ws.cell(row=row, column=4, value=jami)
            ws.cell(row=row, column=5, value=jami_o_uzbek)
            ws.cell(row=row, column=6, value=jami_rus)
            ws.cell(row=row, column=7, value=jami_kurs)
            ws.cell(row=row, column=8, value=jami_o_uzbek_kurs)
            ws.cell(row=row, column=9, value=jami_rus_kurs)
            row += 1

    wb.save('talabalar2.xlsx')
# sirtqi


exporttoexcel(org)
