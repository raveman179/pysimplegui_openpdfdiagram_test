import openpyxl
from pprint import pprint

diagram_number = openpyxl.load_workbook('./diagramNo.xlsx')

# cell = diagram_number.cell(column=2)
sheet = diagram_number['Sheet1']
cell = sheet['A1:A10']

pprint(cell.value)


# print(diagram_number['Sheet1']['B2'].value)