import sys
reload(sys)
sys.setdefaultencoding('utf-8')
import smtplib
import time
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
    RangoTotal = 'Currency'
    for x in (3,14):
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



wb=load_workbook('180105 Pagos Uber.xlsx')
ws = wb.active
ultima_fila = ws.max_row

Formato(ws,ultima_fila)
wb.save('Prubes.xlsx')
