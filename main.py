import openpyxl
import numpy as np


wb = openpyxl.load_workbook('miler.xlsx')
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


def linearyzjaca(macierz, v):
    tab = []
    for i in range(0, 8):
        for j in range(0, 4):
            if j != 3:
                tab.append((-2*(macierz[i][j])+2*macierz[0][j]))
            else:
                tab.append(pow(v, 2)*(2*macierz[i][j]-2*macierz[0][j]))
    mat = np.array(tab).reshape(8, 4)
    mat = np.delete(mat, 0, 0)
    return mat


def macierz_b(macierz, v):
    tab = []
    for i in range(0, 8):
        a = 0
        b = 0
        for j in range(0, 4):
            if j != 3:
                a = a+(-(pow(macierz[i][j], 2)) + pow(macierz[0][j], 2))
            else:
                b = b+(pow(v, 2)*(pow(macierz[i][j], 2) - pow(macierz[0][j], 2)))
        tab.append(a+b)
    mat = np.array(tab).reshape(8, 1)
    mat = np.delete(mat, 0, 0)
    return mat


macierzA = wczytanie_danych(arkusz)
print(macierzA)
print(linearyzjaca(macierzA, wave_speed))
macierzB = macierz_b(macierzA, wave_speed)
print(macierzB)















