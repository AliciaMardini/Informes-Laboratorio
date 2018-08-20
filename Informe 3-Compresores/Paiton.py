# -*- coding: utf-8 -*-
"""
Created on Thu May 17 17:33:57 2018

@author: Alicia
"""

import numpy as np
import openpyxl as pyxl
import matplotlib.pyplot as plt


wb = pyxl.load_workbook('Compresores.xlsx', data_only=True)
sheet = wb.get_active_sheet()
'''Recordar que las celdas estan en coordenadas que comienzan en el (1,1)
A partir de la esquina superior derecha
test:
    A = sheet.cell(row=3,column=1).value'''
Matriz = []
Tuplas = sheet['A12:K34']
for i in range(len(Tuplas)):
    a = []
    for j in range(len(Tuplas[0])):
        a.append(float(Tuplas[i][j].value))
    Matriz.append(a)
Matriz = np.array(Matriz)
t = Matriz[:,0]
P1 = Matriz[:,1]
P2 = Matriz[:,2]
P3 = Matriz[:,3]
T0 = Matriz[:,4]
T1 = Matriz[:,5]
T2 = Matriz[:,6]
T3 = Matriz[:,7]
T4 = Matriz[:,8]
T5 = Matriz[:,9]
T6 = Matriz[:,10]


dpi = 400

plt.plot(t, P1, '--*', label='P1')
plt.plot(t, P2, '--*', label='P2')
plt.plot(t, P3, '--*', label='P3')
plt.ylabel(u'Presión [bar]')
plt.xlabel(u'Tiempo [min]')
plt.legend()
plt.savefig('Presiones.png', dpi=dpi, bbox_inches='tight')
plt.clf()

plt.plot(t, T0, '--*', label='T0')
plt.plot(t, T1, '--*', label='T1')
plt.plot(t, T2, '--*', label='T2')
plt.plot(t, T3, '--*', label='T3')
plt.plot(t, T4, '--*', label='T4')
plt.plot(t, T5, '--*', label='T5')
plt.plot(t, T6, '--*', label='T6')
plt.ylabel(u'Temperatura [°C]')
plt.xlabel(u'Tiempo [min]')
plt.legend(loc=2)
plt.savefig('Temperatura.png', dpi=dpi, bbox_inches='tight')
plt.clf()

#coeficiente politropico

def n(T1, T2, P1, P2):
    rel = np.log(T2 / T1) / np.log(P2 / P1)
    output = 1 / (1 - rel)
    return output


P0 = 1.01325 #[bar]
Ve = 500 #Volumen del estanque.
n0 = n(T0+273.15,T5+273.15,P0,P3+P0)
n1 = n(T0+273.15,T1+273.15,P0,P1+P0)
n2 = n(T2+273.15,T3+273.15,P1+P0,P2+P0)
n3 = n(T4+273.15,T5+273.15,P2+P0,P3+P0)


plt.plot(t, n1,'--*',  label='Compresor 1')
plt.plot(t, n2,'--*',  label='Compresor 2')
plt.plot(t[7:], n3[7:],'--*',  label='Compresor 3')
plt.ylabel(u'Coeficiente politrópico')
plt.xlabel(u'Tiempo [min]')
plt.legend()
plt.savefig('Coef.png', dpi=dpi, bbox_inches='tight')
plt.clf()

print 'n1 = ', np.mean(n1[:-3]), ' ', np.std(n1[:-3])
print 'n2 = ', np.mean(n2[:-3]), ' ', np.std(n2[:-3])
print 'n3 = ', np.mean(n3[7:-3]), ' ', np.std(n3[7:-3])    

N1 = np.mean(n1[:-3])
N2 = np.mean(n2[:-3])
N3 = np.mean(n3[7:-3])    

Q = Ve * np.diff(((T0+273.15)/(P0)) * (P3+P0)/(T6+273.15)) / np.diff(60 * t) 
indices = Q>0

plt.plot(t[:-1][indices], Q[indices], '--*')
plt.ylabel(u'Caudal [lt/s]')
plt.xlabel(u'tiempo [min]')
plt.savefig('CaudalTiempo.png', dpi=dpi, bbox_inches='tight')
plt.clf()

plt.plot(P3[:-1][indices], Q[indices], '--*')
plt.ylabel(u'Caudal [lt/s]')
plt.xlabel(u'Presión [bar]')
plt.savefig('CaudalPresion.png', dpi=dpi, bbox_inches='tight')
plt.clf()


#para el diagrama
p0 = P0 * np.ones_like(P1)
p1 = P1 + P0
p2 = P2 + P0
p3 = P3 + P0

t0 = T0 + 273.15
t1 = T1 + 273.15
t2 = T2 + 273.15
t3 = T3 + 273.15
t4 = T4 + 273.15
t5 = T5 + 273.15
t6 = T6 + 273.15

v0 = Ve * p3 / t6 * t0 / p0
v1 = Ve * p3 / t6 * t1 / p1
v2 = Ve * p3 / t6 * t2 / p1
v3 = Ve * p3 / t6 * t3 / p2
v4 = Ve * p3 / t6 * t4 / p2
v5 = Ve * p3 / t6 * t5 / p3
v6 = Ve * np.ones_like(v1)

M= np.array([v0, v1, v2, v3, v4, v5, v6]).T

# tuplas

A0 = (p0, v0)
A1 = (p1, v1)
A2 = (p1, v2)
A3 = (p2, v3)
A4 = (p2, v4)
A5 = (p3, v5)
A6 = (p3, v6)

Lista_de_puntos = [A0, A1, A2, A3, A4, A5, A6]

def ppath(Pi, Vi, Pf, Vf, n):
    v = np.linspace(Vi, Vf, 100)
    P = Pi * (Vi / v) ** n
    return (P, v)

def lpath(Pi, Vi, Vf):
    v = np.linspace(Vi, Vf, 50)
    P = Pi * np.ones_like(v)
    return (P, v)


def freim(i, filename):
    plt.plot(A6[1][i], A6[0][i],'*b')
    plt.plot(A5[1][i], A5[0][i],'*b')
    plt.plot(A4[1][i], A4[0][i],'*b')
    plt.plot(A3[1][i], A3[0][i],'*b')
    plt.plot(A2[1][i], A2[0][i],'*b')
    plt.plot(A1[1][i], A1[0][i],'*b')
    plt.plot(A0[1][i], A0[0][i],'*b')
    
    A65 = lpath(A6[0][i], A6[1][i], A5[1][i])
    plt.plot(A65[1], A65[0], '-b')
    
    A54 = ppath(A5[0][i], A5[1][i], A4[0][i], A4[1][i], n3[i])
    plt.plot(A54[1], A54[0], '-b')
    
    A43 = lpath(A4[0][i], A4[1][i], A3[1][i])
    plt.plot(A43[1], A43[0], '-b')
    
    A32 = ppath(A3[0][i], A3[1][i], A2[0][i], A2[1][i], n2[i])
    plt.plot(A32[1], A32[0], '-b')
    
    A21 = lpath(A2[0][i], A2[1][i], A1[1][i])
    plt.plot(A21[1], A21[0], '-b')
    
    A10 = ppath(A1[0][i], A1[1][i], A0[0][i], A0[1][i], n1[i])
    plt.plot(A10[1], A10[0], '-b')
    
    plt.xlabel(u'Volumen [lt]')
    plt.ylabel(u'Presión [bar]')
    plt.savefig(filename+'.png', dpi = dpi)
    plt.clf()

freim(3, 'tiempo' + str(int(t[3])))
freim(10, 'tiempo' + str(int(t[10])))
freim(20, 'tiempo' + str(int(t[20])))
freim(22, 'tiempo' + str(int(t[22])))

MATTTT = np.array([v0, v1, v2, v3, v4, v5, v6])

