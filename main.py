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
            if column == "Valor Unit":
                valor_unit = product.get(column, 0)
            if column == "Qtde Unit":
                qtde_unit = product.get(column, 0)
            if column == "Total":
                value = float(valor_unit) * float(qtde_unit)
                row_cell[j].text = str(value)
                total_value += float(value)

    total_row = table.add_row().cells
    total_row[0].merge(total_row[-1])

    reais = int(total_value)
    centavos = int(round((total_value - reais) * 100))
    
    real_extensive = num2words(reais,lang="pt_BR",to='currency')
    centavos_extensive = num2words(centavos,lang="pt_BR",to='currency')

    join_extensive = f"{real_extensive} e {centavos_extensive}"

    total_row[0].text = f"TOTAL: {join_extensive}   |   R${total_value:.2f}"

    # Ajustar as margens dos parágrafos dentro das células
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:

                # Configurar indentação (margem) à esquerda e direita
                paragraph.paragraph_format.left_indent = Pt(1)
                paragraph.paragraph_format.right_indent = Pt(1)

                # Configurar espaço antes e depois do parágrafo
                paragraph.paragraph_format.space_before = Pt(1)
                paragraph.paragraph_format.space_after = Pt(1)

                # Configurar espaçamento entre linhas
                paragraph.paragraph_format.line_spacing = Pt(12)

    file_path = os.path.join(library, file_name)
    doc.save(file_path)

products = [
    {"Número": 1, "Unid. Med": "UN", "Descrição": "PC GAMER", "Marca": "Ryzhen", "Modelo": "AMD", "Valor Unit": 35.95, "Valor Descrito": "teste", "Qtde Unit": 1, "Total": 0},
    {"Número": 2, "Unid. Med": "UN", "Descrição": "Calculadora", "Marca": "Ryzhen", "Modelo": "AMD", "Valor Unit": 35.95, "Valor Descrito": "teste", "Qtde Unit": 70, "Total": 0}
]


generate_table(products, 'tabela_produtos.docx', library="C:/GitHub/Python/gerar_word_tables")
