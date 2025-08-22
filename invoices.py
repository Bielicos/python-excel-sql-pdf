import os
# Bibliotêca que serve para interagir com sistema de ficheiros e ambiente do sistema operacional

directory = 'pdf_invoices'
# String que armazena o diretório "pdf_invoices" que está na raiz

files = os.listdir(directory)
# Retorna uma lista de Strings com os nomes de cada árquivo do diretório

files_quantity = len(files)
# Var que armazena a quantidade de árquivos do diretório

if files_quantity == 0:
    raise Exception('No files found in the directory')
# Se a quantidade for zero, será jogado uma exception


