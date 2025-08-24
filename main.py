import os
import openpyxl
from openpyxl.workbook import Workbook

directory = "pdf"
files = os.listdir(directory)
files_quantity = len(files)

if files_quantity == 0:
    raise Exception("No files found in the directory")

wb = Workbook()
ws = wb.active
ws.title = "Invoice Imports"

# Referenciando as colunas do Excel
ws['A1'] = "Invoice #"
ws['B1'] = "Date"
ws['C1'] = "File name"
ws['D1'] = "Status"

last_empty_line = 1
while ws["A" + str(last_empty_line)].value is not None:
    last_empty_line += 1



