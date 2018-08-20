# -*- coding: utf-8 -*-
"""
Script para realizar el laboratorio 1: Refrigeracion Mecanica
@author: Alicia
"""

import numpy as np
import openpyxl as pyxl
import matplotlib.pyplot as plt
import scipy.optimize as scop

wb = pyxl.load_workbook('Tablafreon12.xlsx', data_only=True)
sheet = wb.get_active_sheet()
'''Recordar que las celdas estan en coordenadas que comienzan en el (1,1)
A partir de la esquina superior derecha
test:
    A = sheet.cell(row=3,column=1).value'''
tk = 0

T = []
P = []
hl = []
hg = []
sl = []
sg = []

Rango = sheet.calculate_dimension()
Rango = 'A2:' + Rango[3:]
Tuplas = sheet[Rango]

for i in range(len(Tuplas)):
    T.append(Tuplas[i][0].value)
    P.append(Tuplas[i][1].value)
    hl.append(Tuplas[i][3].value)
    hg.append(Tuplas[i][4].value)
    sl.append(Tuplas[i][5].value)
    sg.append(Tuplas[i][6].value)

T = np.array(T) + tk
P = np.array(P)
hl = np.array(hl)
hg = np.array(hg)
sl = np.array(sl)
sg = np.array(sg)


plt.plot(sl,T, '-k')
plt.plot(sg,T,'-k')
plt.xlabel(u'Entropía [kJ/kgK] ')
plt.ylabel(u'Temperatura [°C]')

# Puntos 

# Puntos teoricos
st = np.array([sg[13],sg[13],sl[27],sl[27],sg[13]])
Tt = np.array([T[13],T[27],T[27],T[13],T[13]])

st2 = np.array([sg[13],sg[13],sg[27],sl[27],sg[13]])
Tt2 = np.array([T[13],T[27],T[27],T[13],T[13]])

# Puntos reales
sr2 = np.array([sg[13],sg[13],sg[27],sl[27],sl[27]+0.02,sl[27]+0.065,sl[27]+0.1,sg[13]])
Tr2 = np.array([T[13],T[27]+20,T[27],T[27],T[27]-25,T[15],T[13],T[13]])

sr = np.array([sg[13],sg[13],sg[27],sl[27],sl[27]+0.1,sg[13]])
Tr = np.array([T[13],T[27]+20,T[27],T[27],T[13],T[13]])

distx = 0.03
disty = 3
plt.plot(st,Tt,'-*b',label='Ciclo de Carnot')
plt.plot(sr,Tr, '*r',label='Ciclo Real')
plt.plot(sr2,Tr2, '--r' )

plt.text(st[0]+distx,Tt[0]+disty,u'1')
plt.text(st[1]+distx,Tt[1]+disty,u'2')
plt.text(sr[2]-1.5*distx,Tr[2]+disty,u'a')
plt.text(st[2]+distx,Tt[2]+disty,u'3')
plt.text(st[3]+distx,Tt[3]+disty,u'4')

plt.text(sr[1]+distx,Tr[1]+disty,'2\'')
plt.text(sr[4]+distx,Tr[4]+disty,'4\'')
plt.legend()
plt.savefig('DiagramaTs.png', dpi=600)
plt.clf()


# Grafico h-P
plt.semilogy(hl,P,'-b',label=u'Saturación')
plt.xlabel(u'Entalpía ' + r'$h$' + ' [kJ/kg]')
plt.ylabel(u'Presión ' + r'$P$' + ' [kPa]')
plt.plot(hg,P,'-b')
plt.ylim([100,4200])
plt.grid(True, which="both")

#Puntos
Pa = 686.466
Pb = 343.233
Entalpias = np.array([
        369.227, #entra cond
        np.interp(Pa,P,hg),
        359.486, #sale cond
        359.486, #entra evap
        365.191, #sale evap   
        ])
Entalpias2 = np.array([
        365.191,
        365.7,
        366.7,
        368,
        369.227
        ])
Presiones = np.array([
        Pa, #salida compresor
        Pa,
        Pa,
        Pb,
        Pb        
        ])
Presiones2 = np.array([
        Pb,
        0.5*(Pa+Pb) - 0.25*(Pa-Pb),
        0.5*(Pa+Pb),
        0.5*(Pa+Pb) + 0.25*(Pa-Pb),
        Pa
        ])

plt.plot(Entalpias, Presiones, '-*r',label='Ciclo del R12')
plt.plot(Entalpias2, Presiones2, '-r')
plt.legend()

dx = 3
dy = 10

plt.text(Entalpias[0]+dx, Presiones[0]+dy,'2')
plt.text(Entalpias[2]-3*dx, Presiones[2]+dy,'3')
plt.text(Entalpias[3]-0.5*dx, Presiones[3]-7*dy,'4')
plt.text(Entalpias[4]+0.5*dx, Presiones[4]-7*dy,'1')
plt.savefig('Ph.png',dpi=600)
plt.show()