import openpyxl
import numpy as np


wb = openpyxl.load_workbook('wb.xlsx')
sheet = wb['Arkusz1']
tab = []


wave_speed = sheet['I2'].value


def wczytanie_danych(sheet, tab):
    for row_i in range(2, 10):
        for column_i in range(2, 5):
            print(row_i, sheet.cell(row=row_i, column=column_i).value)
            x1 = sheet.cell(row_i, column_i).value
            tab.append(x1)
    mat = np.array(tab).reshape(8, 3)
    return mat


print(wczytanie_danych(sheet, tab))










