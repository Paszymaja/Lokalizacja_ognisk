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


def linearyzjaca(macierz, v):
    for i in range(0, 8):
        for j in range(0, 4):
            if j != 3:
                macierz[i][j] = 2*(-(macierz[i][j])+macierz[0][j])
            else:
                macierz[i][j] = pow(v, 2)*2*(macierz[i][j]-(macierz[0][j]))
    macierz = np.delete(macierz, 0, 0)
    return macierz


def macierz_b(macierz, v):
    tab = []
    for i in range(0, 8):
        for j in range(0, 4):
            if j != 3:
                a = -(pow(macierz[i][j], 2)) + pow(macierz[0][j], 2)
            else:
                b = pow(v, 2)*(pow(macierz[i][j], 2) - pow(macierz[0][j], 2))
        tab.append(a+b)
    mat = np.array(tab).reshape(8, 1)
    mat = np.delete(mat, 0, 0)
    return mat


macierzB = macierz_b(macierzA, wave_speed)
macierzA = linearyzjaca(macierzA, wave_speed)
print(macierzA, "\n")
print(macierzB, "\n")

macierzT = macierzA.transpose()
print(macierzT, "\n")
macierzA = np.dot(macierzT, macierzA)
macierzB = np.dot(macierzT, macierzB)
print(macierzA, "\n")
print(macierzB, "\n")

print(np.linalg.solve(macierzA, macierzB))
















