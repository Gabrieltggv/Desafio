import os
from typing import List

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

directory: List[str] = os.listdir(".")
arquivos: List[str] = [i for i in directory if ".pdf" in i]
[print(a) for a in arquivos]
option: str = input("Qual o número da página a ser modificada: ")
for b in arquivos:
    if option in b:
        os.rename(b, "Página {} - Modificado.pdf".format(option))
        work = Workbook()
        sheet = work.worksheets[0]
        table = Worksheet.values
        sheet['A1'] = "Nome do documento"
        sheet['B1'] = "Status"
        sheet['A2'] = b
        sheet['B2'] = "Documento alterado"
        work.save("Relatório de execução.xlsx")
        