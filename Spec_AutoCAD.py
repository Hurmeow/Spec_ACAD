import win32com.client
import re
from math import pi
from tkinter import filedialog as fd
from tkinter import *
import tkinter.messagebox as mb

ac = win32com.client.Dispatch('AutoCAD.Application')
color = win32com.client.Dispatch('AutoCAD.Application').GetInterfaceObject("AutoCAD.AcCmColor.23")

'''
GetInterfaceObject("AutoCAD.AcCmColor.XX")
XX:
CAD2015：AutoCAD.AcCmColor.20
CAD2016/17：AutoCAD.AcCmColor.21
CAD2018：AutoCAD.AcCmColor.22
CAD2019/2020：AutoCAD.AcCmColor.23
'''

'''
вероятно, подавляющее большинство приложений на основе tk помещают все компоненты в корневое окно по умолчанию. 
Это наиболее удобный способ сделать это, так как он уже существует. 
Выбор скрыть окно по умолчанию и создать свой собственный-это прекрасно, хотя это требует совсем немного дополнительной 
работы.

чтобы ответить на ваш конкретный вопрос о том, как это скрыть, используйте вывести метод корневого окна:

import Tkinter as tk
root = tk.Tk()
root.withdraw()

Если вы хотите сделайте окно видимым снова, вызовите правой кнопкой мышки (или wm_deiconify) метод.

root.deiconify()

как только вы закончите с диалогом, вы можете уничтожить корневое окно вместе со всеми другими виджетами tkinter 
с помощью уничтожить способ:
root.destroy()


'''
MS = ac.ActiveDocument

result = MS.Utility.GetEntity(None, None, "Укажите таблицу спецификаций: ")

color.SetRGB(255, 0, 0)

total_spec = dict()  # словарь для записи ведомости расхода стали


def total_mass(massa, string, total_dict=dict):
    d_str = re.search(r'%%C\d+', string).group()
    marka = re.search(r'[А-Яа-яA-Za-z]{1,2}\d{3,4}[А-Яа-яA-Za-z]{0,1}',
                      string).group()  # извлекаем d и длину элемента |[А-Яа-яA-Za-z]\d{3}

    title_name = str(marka + d_str)

    if not total_dict.get(marka, False):
        total_dict[marka] = {title_name: massa}
        print('IF')
    else:
        if not total_dict[marka].get(title_name, False):
            total_dict[marka][title_name] = massa
        else:
            total_dict[marka][title_name] += massa
        print('ELSE')


# result[0].SetCellTextStyle(5, 4, (result[0].GetCellTextStyle(4, 5))) установить стиль текста
# result[0].SetCellTextHeight(5, 4, 3.0)  установить высоту текста


for i in range(42, result[0].Rows):
    try:
        string1 = result[0].GetCellValue(i, 2)  # получаем инфу из столбца 'Наименование"
        if not string1:
            continue
    except AttributeError:
        print(f'ERROR : Cell {i + 1}')
        continue
    d, length = int(*re.findall(r'%%C(\d+)', string1)), int(
        *re.findall(r'L\W+(\d+)', string1))  # извлекаем d и длину элемента
    print(f'd= {d}, L= {length}')
    if d != 0 and length != 0:
        m = (pi * (d * 0.001) ** 2) * 1962.5  # 1962.5 = 0.25 * 7850  m - масса 1 кг на метр
        M = m * (length * 0.001)  # масса одного элемента
        total_m = int(result[0].GetCellValue(i, 3)) * M  # общая масса всех элементов
        result[0].SetCellValue(i, 4, M)  # запись в ячейку общей массы
        result[0].SetCellFormat(i, 4, (None, None, '%pr2%lu2'))  # изменяем тип ячейки

        result[0].SetCellValue(i, 5, total_m)  # запись в ячейку общей массы
        total_mass(total_m, string1, total_spec)
        result[0].SetCellFormat(i, 5, (None, None, '%pr2%lu2'))  # изменяем тип ячейки

    elif string1.find('м.п.') != -1:
        total_m = int(result[0].GetCellValue(i, 3)) * float(result[0].GetCellValue(i, 4))
        result[0].SetCellValue(i, 5, total_m)  # запись в ячейку общей массы
        result[0].SetCellFormat(i, 5, (None, None, '%pr2%lu2'))  # изменяем тип ячейки

        total_mass(total_m, string1, total_spec)
        for j in range(result[0].Columns):
            result[0].SetCellBackgroundColor(i, j, color)
        print(f'ПРОВЕРИТЬ!!! : Cell {i + 1}')

    elif string1.find('Каркас') != -1:  # Ячейки с каркасами должны содержать слово "Каркас"
        a = True
        while a:
            try:
                root = Tk()
                root.attributes("-topmost", True)
                root.withdraw()
                mb.showinfo(f'Выберите {string1}!!!!!!', f'Выберите {string1}!!!!!!')
                root.destroy()
                karkas_tab = MS.Utility.GetEntity(None, None,
                                                  string1.upper() + ':')  # Запрашиваем выбор таблицы спецификации Каркаса №
                a = False
            except:
                print('Ошибка выбора!!!!')
        karkas_m = 0
        for k in range(1, (karkas_tab[0].Rows)):
            try:
                string_k = karkas_tab[0].GetCellValue(k, 2)  # получаем инфу из столбца 'Наименование"
                if not string_k:
                    continue
            except AttributeError:
                print(f'ERROR Karkas : Cell {k + 1}')
                continue
            total_m = int(result[0].GetCellValue(i, 3)) * float(karkas_tab[0].GetCellValue(k, 5))
            total_mass(total_m, string_k, total_spec)

            karkas_m += float(karkas_tab[0].GetCellValue(k, 5))
            karkas_tab[0].SetRowHeight(k, 8)
        result[0].SetCellValue(i, 4, karkas_m)
        result[0].SetCellValue(i, 5, float(result[0].GetCellValue(i, 3)) * result[0].GetCellValue(i, 4))

        result[0].SetCellFormat(i, 4, (None, None, '%pr2%lu2'))  # изменяем тип ячейки
        result[0].SetCellFormat(i, 5, (None, None, '%pr2%lu2'))  # изменяем тип ячейки


    else:
        for j in range(result[0].Columns):
            result[0].SetCellBackgroundColor(i, j, color)
        print(f'НЕТ ДАННЫХ!!! : Cell {i + 1}')
    result[0].SetRowHeight(i, 8)

root = Tk()
root.attributes("-topmost", True)
root.withdraw()
file_name = fd.askopenfilename(defaultextension='.txt', filetypes=[('txt', '.txt')])
with open(file_name, 'w') as file:
    total_steel = 0
    for key, value in total_spec.items():
        print(key, file=file)
        for key2, value2 in total_spec[key].items():
            d_str = int(*re.findall(r'\w+%%C(\d{1,2})', key2))
            print(f'd{d_str}:     {round(value2, 2)}   кг', file=file)
            total_steel += round(value2, 2)
        print()
    print(f'Всего изделия арматурные:   {round(total_steel, 2)}   кг', file=file)
root.destroy()
