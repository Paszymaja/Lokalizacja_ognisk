import openpyxl
import numpy as np


wb = openpyxl.load_workbook('wb.xlsx')
arkusz = wb['Arkusz1']
wave_speed = arkusz['A14'].value


def wczytanie_danych(sheet):
    tab = []
    for row_i in range(2, 10):
        for column_i in range(2, 6):
            x1 = float(sheet.cell(row_i, column_i).value)
            tab.append(x1)
    mat = np.array(tab).reshape(8, 4)
    return mat


macierzA = wczytanie_danych(arkusz)
print(macierzA)
print("   ")


def linearyzjaca(macierz, v):
    for i in range(1, 8):
        for j in range(0, 3):
            if j != 3:
                macierz[i][j] = 2*(-macierz[i][j]+macierz[0][j])
            else:
                macierz[i][j] = pow(v, 2)*2*(macierz[i][j]-macierz[0][j])
    return macierz


print(linearyzjaca(macierzA, wave_speed))














