import os
import random
import datetime
import win32com.client as office

from load_data import import_sheets, import_config, import_sheet
from generate_gongratulation import generate_triad, generate_holiday, generate_congrat

excel_data = 'data.xls'
config_sheet_name = 'config'
addressates = 'addressates'

sheets = import_sheets(excel_data)
print(sheets)

config = import_config(sheets[config_sheet_name])
print(config.text_box_height, config.text_box_width)

congrats_list = config.congrats.split(",")
print(congrats_list)

inx1 = random.randint(0, len(congrats_list) - 1)
inx2 = random.randint(0, len(congrats_list) - 1)
inx3 = random.randint(0, len(congrats_list) - 1)

while (inx1 == inx2 or inx1 == inx3 or inx2 == inx3):
    inx1 = random.randint(0, len(congrats_list) - 1)
    inx2 = random.randint(0, len(congrats_list) - 1)
    inx3 = random.randint(0, len(congrats_list) - 1)

print(inx1, inx2, inx3)

congrats1 = import_sheet(sheets[congrats_list[inx1]])
congrats2 = import_sheet(sheets[congrats_list[inx2]])
congrats3 = import_sheet(sheets[congrats_list[inx3]])
print(congrats1, congrats2, congrats3)

holidays = import_sheet(sheets[config.holidays])
print(holidays)

triad = generate_triad(congrats1, congrats2, congrats3)
print(triad)

holiday = generate_holiday(holidays)
print(generate_congrat('Виктор', holiday, triad))

if not os.path.exists(config.out):
    os.mkdir(config.out)

word = office.gencache.EnsureDispatch('Word.Application')
for addressat in import_sheet(sheets[config.addressates]):
    holiday = generate_holiday(holidays)
    triad = generate_triad(congrats1, congrats2, congrats3)
    congrat = generate_congrat(addressat,holiday,triad)
    doc = word.Documents.Open(f'{os.getcwd()}\\{config.template}')

    try:
        textbox = doc.Shapes.AddTextbox(1, config.text_box_pos_x, config.text_box_pos_y,config.text_box_width,config.text_box_height)
        textbox.TextFrame.TextRange.Text = congrat
        textbox.TextFrame.MarginTop = 0
        textbox.TextFrame.MarginLeft = 0
        textbox.Fill.Visible = 0
        textbox.Line.Visible = 0



        doc.SaveAs2(f'{os.getcwd()}\\{config.out}\\{addressat}_congrat ({str(datetime.datetime.now()).replace(".",",").replace(":",",")}).docx')
        doc.Close()
    except BaseException as exception:
        doc.Close()
        raise Exception(exception)

word.Application.Quit()