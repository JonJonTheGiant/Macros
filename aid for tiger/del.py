import json
from openpyxl import Workbook
import os

class makeit:
    def __init__(self, input_file):
        self.f = open(input_file)  # json_filename = str
        # noinspection PyTypeChecker
        self.json_data = json.load(self.f, parse_float=str)  # 'parse_float' argument reads all floats as strings
        self.f.close()

        variable_data = self.json_data['Variables']
        wb: object = Workbook()
        ws: object = wb.active  # ws: object is a worksheet
        ws.title = 'parameters'

        i = 1
        for parameter in variable_data:
            c1 = ws.cell(row=i, column=1)

            c1.value = parameter['QualifiedName'].replace('Parameters.', '')

            i += 1

        wb.save("ce_parameters.xlsx")

makeit(input_file='C:\VT\Satchel\in.json')