import xlrd
import random

class Config:
    def __init__(self, f = 'Times New Roman', t = 'template.docx', o = 'out', a = 'addressates', hol = 'holidays', c = 'congrats', x = '50', y = '50', w = '100',
                 h = '100'):
        self.font = f
        self.template = t
        self.out = o
        self.addressates = a
        self.holidays = hol
        self.congrats = c
        self.text_box_pos_x = x
        self.text_box_pos_y = y
        self.text_box_width = w
        self.text_box_height = h

    def __setitem__(self, key, value):
        match key:
            case 'font':
                self.font = value
            case 'template':
                self.template = value
            case 'out':
                self.out = value
            case 'addressates':
                self.addressates = value
            case 'holidays':
                self.holidays = value
            case 'congrats':
                self.congrats = value
            case 'text_box_pos_x':
                self.text_box_pos_x = value
            case 'text_box_pos_y':
                self.text_box_pos_y = value
            case 'text_box_width':
                self.text_box_width = value
            case 'text_box_height':
                self.text_box_height = value

def import_sheets(filename):
    data = xlrd.open_workbook(filename, formatting_info=True)
    sheets = dict()
    for sheet in data.sheets():
        sheets[sheet.name] = sheet
    return sheets


def import_config(sheet: xlrd.sheet.Sheet):
    config = Config()
    for i in range(0, sheet.nrows):
        config_field = sheet.cell(rowx=i, colx=0).value
