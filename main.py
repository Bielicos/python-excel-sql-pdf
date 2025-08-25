import os
import openpyxl
from openpyxl.workbook import Workbook
import pdfplumber
import re
from datetime import datetime

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

# Para file
for file in files:
    # Abrir seu pdf
    with pdfplumber.open(directory + "/" + file) as pdf:
        first_page = pdf.pages[0]
        pdf_text = first_page.extract_text()

    # Instrucao Regex
    inv_number_re_pattern = r"INVOICE #(\d+)"
    inv_date_re_pattern = r"DATE: (\d{2}/\d{2}/\d{4})"

    # Procure o texto dentro do pdf baseado nessa instrucao regex
    match_number = re.search(inv_number_re_pattern, pdf_text)
    match_date = re.search(inv_date_re_pattern, pdf_text)

    # Se encontrou algo, ou seja, True :
    if match_number:
        # Se você encontrou, pega o primeiro numero para mim, que é o resultado que eu quero
        invoice_number = match_number.group(1)
        # Coloco o valor no excel
        ws["A" + str(last_empty_line)] = invoice_number
    else:
        ws["A" + str(last_empty_line)] = "None"

    if match_date:
        invoice_date = match_date.group(1)
        ws["B" + str(last_empty_line)] = invoice_date
    else:
        ws["B" + str(last_empty_line)] = "None"

    ws["C" + str(last_empty_line)] = file
    ws["D" + str(last_empty_line)] = "Completed"

    last_empty_line += 1

full_now = str(datetime.now()).replace(":", "-")
dot_index = full_now.index(".")
# é para ir até o indice, ou seja, remover os mini segundos
now = full_now[:dot_index]

wb.save("Invoices -" + str(now) + ".xlsx")