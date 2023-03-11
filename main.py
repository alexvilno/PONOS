import os
import random
import datetime
import win32com.client as ponos

from load_data import load_data
from generate_gongratulation import generate_triad, generate_holiday, generate_congrats

excel_data = 'data.xls'
config_sheet_name = 'config'
addressates = 'addressates'

sheets, config, congrats1, congrats2, congrats3 = load_data(excel_data, config_sheet_name)

congratulations = generate_congrats()

if not os.path.exists(config.out):
    os.mkdir(config.out)

word = ponos.gencache.EnsureDispatch('Word.Application')
for addressat in import_sheet(sheets[config.addressates]):
    holiday = generate_holiday(holidays)
    triad = generate_triad(congrats1, congrats2, congrats3)
    congrat = generate_congrats(addressat, holiday, triad, )
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