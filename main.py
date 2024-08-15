from docx import Document
from docx.shared import Pt
from num2words import num2words
import os
from copy import deepcopy 

# estrutura básica das colunas
default_columns = ["Número","Unid. Med","Descrição", "Marca","Modelo","Valor Unit","Valor Descrito","Qtde Unit","Total"]

library = "C:\\Users\\Flavinho\\Desktop\\function_breno\\teste\\table.docx"

def generate_table(products, extra_column = None):
    doc = Document()

    # adicionando coluna extra se houver
    columns = deepcopy(default_columns)
    if extra_column:
        columns.extend(extra_column)

    # criar tabela
    table = doc.add_table(rows=len(products)+1,cols=len(columns),style="Table Grid")

    # preencher o cabeçalho da tabela
    header_table = table.rows[0].cells
    for i,name_columns in enumerate(columns):
        header_table[i].text = name_columns

    # preencher os dados da tabela
    total_value = 0
    for i,product in enumerate(products):
        row_cell = table.rows[i+1].cells
        for j,column in enumerate(columns):
            value = product.get(column, '')
            row_cell[j].text = str(value)
            if column == "Total":
                total_value += float(product.get("Total",0))      

    # adicionar uma nova linha para o total
    total_row = table.add_row().cells

    # mesclar todas as colunas exceto a última
    total_row[0].merge(total_row[-1])

    # definit o texto do total na última linha
    total_table = num2words(total_value, lang="pt_BR").capitalize()
    total_row[0].text = f"TOTAL: {total_table}   |   R${total_value:.3f}"

    # salvar o documento
    doc.save(library)

# exemplo de uso

products = [
    {"Número": 1,"Unid. Med":"UN", "Descrição": "PC GAMER", "Marca": "Ryzhen", "Modelo": "AMD","Valor Unit": "3500,99", "Valor Descrito": "teste","Qtde Unit": 1, "Total": 3500.99,"Observações": "teste","Flavinho": "Fala"}
]

new_column = ['Flavinho']

generate_table(products,new_column)