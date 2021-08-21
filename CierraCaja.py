# -*- coding: utf-8 -*-
"""
Created on Tue Apr 20 19:02:53 2021

@author: horac
"""
#Pyinstaller tkcalendar --hidden-import babel.numbers
#
import tkinter
from tkcalendar import Calendar,DateEntry
import sqlite3
import pathlib
CPath = pathlib.Path().parent.absolute()

Caja=0.0

def Ventas_Hora(Hora_inicio, Hora_fin):
    S_Fecha = FCierre.get_date()
    conexion = sqlite3.connect(str(CPath)+"\FacturaLight.db")
    cursor = conexion.cursor()
    inicio=S_Fecha+" "+Hora_inicio
    fin=S_Fecha+" "+Hora_fin
    cursor.execute("SELECT SUM(Total), COUNT(Total) FROM Factura WHERE DtAdd BETWEEN  '{}' and '{}'".format(inicio,fin))
    LFacturas = cursor.fetchall()
    conexion.commit()
    conexion.close()
    return LFacturas
        
def Total_Caja():
    global Caja
    Ventas_m = Ventas_Hora("00:00:00","15:00:00")
    Ventas_t = Ventas_Hora("15:00:00","23:59:59")

    conexion = sqlite3.connect(str(CPath)+"\FacturaLight.db")
    cursor = conexion.cursor()
    S_Fecha = FCierre.get_date()
    cursor.execute("SELECT Emisor,SUM(Total) AS PV, COUNT(Total) FROM Factura WHERE DtAdd LIKE ? GROUP BY Emisor", (S_Fecha+"%",))
    LFacturas = cursor.fetchall()
    conexion.commit()
    conexion.close()

    import win32ui
    import win32con
    dc = win32ui.CreateDC()
    dc.CreatePrinterDC()
    dc.StartDoc('Cierre de caja')
    dc.StartPage()
    #Tamaño de letra y escalo de la pagina
    h=43
    Caja = Ventas_m[0][0]+Ventas_t[0][0]
    fontdata = {'name':'Calibri', 'height':h, 'weight':win32con.FW_NORMAL}
    font = win32ui.CreateFont(fontdata)
    dc.SelectObject(font)
    dc.TextOut(25,1*h,"Cierre del Dia %s" %(S_Fecha))
    dc.TextOut(25,2*h,"----------------------------------------")
    dc.TextOut(25,3*h,"Ventas del turno mañana $ %.2f" %(Ventas_m[0][0]))
    dc.TextOut(25,4*h,"Clientes del turno mañana %2d" %(Ventas_m[0][1]))
    dc.TextOut(25,5*h,"----------------------------------------")
    dc.TextOut(25,6*h,"Ventas del turno Tarde $ %.2f" %(Ventas_t[0][0]))
    dc.TextOut(25,7*h,"Clientes del turno mañana %2d" %(Ventas_t[0][1]))
    dc.TextOut(25,8*h,"----------------------------------------")
    dc.TextOut(25,9*h,"Ventas Totales $ %.2f" %(Caja))
    dc.TextOut(25,10*h,"Clientes Totales %s" %(Ventas_m[0][1] + Ventas_t[0][1]))
    dc.TextOut(25,11*h,"----------------------------------------")
    dc.TextOut(25,12*h,"Ventas x PV $")
    for i in range(0,len(LFacturas)):
        dc.TextOut(25,(13+i)*h,"PV %2d: $ %.2f Facturas:%2d" %( LFacturas[i][0] , LFacturas[i][1], LFacturas[i][2]))
    dc.EndPage()
    dc.EndDoc()
    
Ventana = tkinter.Tk()
Ventana.title("Cierre del Dia")
Ventana.resizable(False,False)
L_Fecha = tkinter.Label(Ventana,text ="Fecha de Cierre",font=("Calibri",20))
FCierre = Calendar(Ventana,date_pattern='yyyy-mm-dd')
#L_Fuente = tkinter.Label(Ventana,text ="Tamaño de fuente",font=("Calibri",10))
#E_fuente = tkinter.Entry(Ventana)
#Test = tkinter.Entry(Ventana)
def Impresion():
    Total_Caja()
    tkinter.messagebox.showinfo(title="Cierre del {}".format(FCierre.get_date()),message="Ventas totales $ {:.2f}".format(Caja),)
    
B_Imp = tkinter.Button(Ventana,text="Imprimir Cierre",width=30,height=2,command = lambda:Impresion())
L_Fecha.grid(row=0, column = 0 , columnspan=4, padx= 10)
FCierre.grid(row=1, column = 0 , columnspan=4, padx = 25)
B_Imp.grid(row=3, column = 0, columnspan=4,pady= 10)
#L_Fuente.grid(row=4, column = 0 , columnspan=2, padx = 10)
#E_fuente.grid(row=4, column = 2 , columnspan=2,pady= 10)
#E_fuente.insert(0, "50")
#Test.grid(row=4, column = 0, columnspan=4,pady=6)
Ventana.mainloop()
