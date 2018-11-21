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


def linearyzjaca(macierz, v): #TODO zrobic obosna macierz bo to nie dziala
    for i in range(0, 7):
        for j in range(0, 3):
            if j != 3:
                macierz[i][j] = (-2*(macierz[i][j])+2*macierz[7][j])
            else:
                macierz[i][j] = pow(v, 2)*(2*macierz[i][j]-2*macierz[7][j])
    print(macierz)
    macierz = np.delete(macierz, 7, 0)
    return macierz


def macierz_b(macierz, v):
    tab = []
    a = 0
    b = 0
    for i in range(0, 8):
        for j in range(0, 3):
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















