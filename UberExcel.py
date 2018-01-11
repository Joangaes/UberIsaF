import sys
reload(sys)
sys.setdefaultencoding('utf-8')
import os
import glob
import csv
from xlsxwriter.workbook import Workbook
import smtplib
import time
import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font

def LetrasExcel(Numero):
    return {1:'A',2:'B',3:'C',4:'D',5:'E',6:'F',7:'G',8:'H',9:'I',10:'J',11:'K',12:'L',13:'M',14:'N',15:'O',16:'P',17:'Q',18:'R'}[Numero]

def  Formato(ws,ultima_fila):
    #Definimos los colores del formato previamente establecido por Isa F
    FontDate = Font(bold=True,size=11,name='Calibri',color='FFFFFF')
    FontStandard = Font(name='Calibri',color='FFFFFF')
    GreyBackground = PatternFill(fill_type='solid',start_color='CCCCCC',end_color='CCCCCC')
    Results = Font(bold=True)
    BlueBackground = PatternFill(fill_type='solid',start_color='01558b',end_color='01558b')
    #Definimos donde van a ser aplicados estos estilos
    FilaDate = ultima_fila+2
    Date = 'B' + str(FilaDate)
    DateCell=ws[Date]
    FilaRetencion = ultima_fila+3
    FilaComplementoPagado = ultima_fila+5
    FilaResultados = ultima_fila+6
    FilaSaldoFinal = ultima_fila+7
    FilaSaldoFinalAcumulado = ultima_fila+8
    SaldoFinalAcumulado = 'B' + str(FilaSaldoFinalAcumulado)
    SaldoFinalAcumuladoCell = ws[SaldoFinalAcumulado]
    #Aplicamos estilos
    #RangoTotal = 'Currency'
    for x in range(3,15):
        Letra = LetrasExcel(x)
        CoordDate = Letra+str(FilaDate)
        CeldaDate=[CoordDate]
        CoordRetencion = Letra+str(FilaRetencion)
        CeldaRetencion=ws[CoordRetencion]
        CeldaRetencion.fill=GreyBackground
        CoordFaltante=Letra+str(FilaRetencion)
        CeldaFaltante=ws[CoordFaltante]
        CoordComplemento = Letra+str(FilaComplementoPagado)
        CeldaComplemento=ws[CoordComplemento]
        CeldaComplemento.fill=GreyBackground
        CoordResultados = Letra+str(FilaResultados)
        CeldaResultados=ws[CoordResultados]
        CoordSaldoFinal = Letra+str(FilaResultados)
        CeldaSaldoFinal=ws[CoordSaldoFinal]
        CeldaSaldoFinal.font=Results
        #ws.cell(row=FilaRetencion,column=x).fill = GreyBackground
        #ws.cell(row=FilaComplementoPagado,column=x).fill = GreyBackground
        #ws.cell(row=FilaSaldoFinal,column=x).font=Results
    SaldoFinalAcumuladoCell.font = Results
    DateCell.font=FontDate
    DateCell.fill = BlueBackground

def CalculoFecha(ws,ultima_fila):
    Ultima_fecha = ws.cell(row=ultima_fila-5,column=2).value
    print(type(Ultima_fecha))
    semana = datetime.timedelta(days=7)
    Fecha_Final = Ultima_fecha+semana
    print(Fecha_Final)
    FilaDate = ultima_fila+2
    Date = 'B' + str(FilaDate)
    DateCell=ws[Date]
    ws[Date]=Fecha_Final
    FilaRetencion = ultima_fila+3
    Retencion = 'B' + str(FilaRetencion)
    ws[Retencion]='Retencion'
    FilaFaltante = ultima_fila+4
    Faltante = 'B' + str(FilaFaltante)
    ws[Faltante]='Faltante por pagar'
    FilaComplemento = ultima_fila+5
    Complemento = 'B' + str(FilaComplemento)
    ws[Complemento]='Complemento pagado'
    FilaTotal = ultima_fila+6
    Total = 'B' + str(FilaTotal)
    ws[Total]='Total Pagado'
    FilaSaldoFinal = ultima_fila+7
    SaldoFinal = 'B' + str(FilaSaldoFinal)
    ws[SaldoFinal]='SaldoFinalAcumulado'
    return(Fecha_Final)

def BusquedaID(id_unico,Pago_cargado,ws,ultima_fila,Fecha_Final):
    for x in range(3,15):
        IdentificadorExcel = ws.cell(row=3,column=x).value
        if(IdentificadorExcel == id_unico):
            Producto = ws.cell(row=7,column=x).value
            print('Llegue')
            print(type(str(Producto)))
            if(str(Producto)=='S'):
                Sedan(Pago_cargado,ws,ultima_fila,Fecha_Final,x)
                print('SiEntre')
            else:
                if(str(Producto)=='V'):
                    Versa(Pago_cargado,ws,ultima_fila,Fecha_Final,x)
                else:
                    if(str(Producto)=='SE'):
                        SinEnganche(Pago_cargado,ws,ultima_fila,Fecha_Final,x)
                    else:
                        if(str(Producto)=='ME'):
                            MedioEnganche(Pago_cargado,ws,ultima_fila,Fecha_Final,x)


def Sedan(Pago_cargado,ws,ultima_fila,Fecha_Final,columna):
    LetraCliente=LetrasExcel(columna)
    PagoCoord = LetraCliente+str(ultima_fila+3)
    print(PagoCoord)
    ws[PagoCoord]=float(Pago_cargado)


def Versa(Pago_cargado,ws,ultima_fila,Fecha_Final,columna):
    LetraCliente=LetrasExcel(columna)
    PagoCoord = LetraCliente+str(ultima_fila+3)
    print(PagoCoord)
    ws[PagoCoord]=float(Pago_cargado)


def SinEnganche(Pago_cargado,ws,ultima_fila,Fecha_Final,columna):
    LetraCliente=LetrasExcel(columna)
    PagoCoord = LetraCliente+str(ultima_fila+3)
    print(PagoCoord)
    ws[PagoCoord]=float(Pago_cargado)

def MedioEnganche(Pago_cargado,ws,ultima_fila,Fecha_Final,columna):
    LetraCliente=LetrasExcel(columna)
    PagoCoord = LetraCliente+str(ultima_fila+3)
    print(PagoCoord)
    ws[PagoCoord]=float(Pago_cargado)


wb=load_workbook('180105 Pagos Uber.xlsx')
ws = wb.active
ultima_fila = ws.max_row
execfile('csvtoexcel.py')
Formato(ws,ultima_fila)
Fecha_Final=CalculoFecha(ws,ultima_fila)
wb2=load_workbook('uber_to_arkafin.xlsx')
UberToArkafin=wb2.active
Ultima_Fila_Arkafin=UberToArkafin.max_row
for x in range(2,Ultima_Fila_Arkafin+1):
    id_unico = UberToArkafin.cell(row=x,column=2).value
    Pago_cargado = UberToArkafin.cell(row=x,column=4).value
    BusquedaID(id_unico,Pago_cargado,ws,ultima_fila,Fecha_Final)



wb.save('Prubes.xlsx')
