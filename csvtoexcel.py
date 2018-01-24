import os
import glob
import csv
from xlsxwriter.workbook import Workbook


for csvfile in glob.glob(os.path.join("C:\Users\Mutuo Midgard\Box Sync\ISA F\UBER\Pagos Uber\PagosArkafin/*.csv")):
    workbook = Workbook(csvfile[:-4] + '.xlsx')
    worksheet = workbook.add_worksheet()
    with open(csvfile, 'rt') as f:
        reader = csv.reader(f)
        for r, row in enumerate(reader):
            for c, col in enumerate(row):
                worksheet.write(r, c, col)
    workbook.close()
    print(csvfile)
    extension = 'csv'
    os.chdir('C:\Users\Mutuo Midgard\Box Sync\ISA F\UBER\Pagos Uber\PagosArkafin/')
    VisorFile = [i for i in glob.glob('*.{}'.format(extension))]
    print VisorFile[0]
    os.rename('C:\Users\Mutuo Midgard\Box Sync\ISA F\UBER\Pagos Uber\PagosArkafin/'+VisorFile[0],'C:\Users\Mutuo Midgard\Box Sync\ISA F\UBER\Pagos Uber\PagosArkafin/Anteriores/'+VisorFile[0])
