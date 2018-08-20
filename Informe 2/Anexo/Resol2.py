# -*- coding: utf-8 -*-
"""
Created on Fri Apr 27 15:36:13 2018

@author: Alicia
"""

import numpy as np
import openpyxl as pyxl
import matplotlib.pyplot as plt
import scipy.optimize as scop

wb = pyxl.load_workbook('PerfilPy.xlsx', data_only=True)
sheet = wb.get_active_sheet()
'''Recordar que las celdas estan en coordenadas que comienzan en el (1,1)
A partir de la esquina superior derecha
test:
    A = sheet.cell(row=3,column=1).value'''
   
Rango = sheet.calculate_dimension()
Rango = 'A2:' + Rango[3:]
Tuplas = sheet[Rango]

for i in range(len(Tuplas)):
    d.append(Tuplas[i][0].value)
    v.append(Tuplas[i][1].value)


plt.rcdefaults()
fig, ax = plt.subplots()

# Example data

y_pos = np.arange(len(d))


ax.barh(y_pos, v, align='center',
        color='green')
ax.set_yticklabels(d)
ax.invert_yaxis()  # labels read top-to-bottom
ax.set_xlabel('Performance')
ax.set_title('How fast do you want to go today?')