from docx import Document
from docx.shared import Pt
from num2words import num2words
import os
from copy import deepcopy 

default_columns = ["Número", "Unid. Med", "Descrição", "Marca", "Modelo", "Valor Unit", "Valor Descrito", "Qtde Unit", "Total"]

def generate_table(products, file_name, library, extra_column=None):
    doc = Document()

    columns = deepcopy(default_columns)
    if extra_column:
        columns.extend(extra_column)

    table = doc.add_table(rows=len(products) + 1, cols=len(columns), style="Table Grid")

    header_table = table.rows[0].cells
    for i, name_columns in enumerate(columns):
        header_table[i].text = name_columns

    total_value = 0
    for i, product in enumerate(products):
        row_cell = table.rows[i + 1].cells
        for j, column in enumerate(columns):
            value = product.get(column, '')
            row_cell[j].text = str(value)
            if column == "Total":
                total_value += float(product.get("Total", 0))

    total_row = table.add_row().cells
    total_row[0].merge(total_row[-1])

    total_table = num2words(total_value, lang="pt_BR").capitalize()
    total_row[0].text = f"TOTAL: {total_table}   |   R${total_value:.2f}"

    file_path = os.path.join(library, file_name)
    doc.save(file_path)


products = [
    {"Número": 1, "Unid. Med": "UN", "Descrição": "PC GAMER", "Marca": "Ryzhen", "Modelo": "AMD", "Valor Unit": 35.00, "Valor Descrito": "teste", "Qtde Unit": 1, "Total": 3500.99},
    {"Número": 2, "Unid. Med": "UN", "Descrição": "Calculadora", "Marca": "Ryzhen", "Modelo": "AMD", "Valor Unit": 35.00, "Valor Descrito": "teste", "Qtde Unit": 70, "Total": 3500.99}
]


generate_table(products, 'tabela_produtos.docx', library="C:/GitHub/Python/gerar_word_tables")
