"""
表示する画面の処理

"""

from planningsheet_parse import sheet_search
import PySimpleGUI as sg
import subprocess
import re

sheetsearch = sheet_search()

# ------pdfビュワーのパス
pdf_viwer_path = r"C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"

sg.theme('Dark Green 1')

layout = [
    [sg.Text('     受注No.(品質経歴書⑨)バーコードを読み込む', pad=((10,10),(20,20))), sg.InputText(key='order_num')],
    [sg.Button('図面を開く', size=(40,1), key='open_diagram'),sg.Button('閉じる', size=(40,1), key='Quit')]
]

window = sg.Window('生産計画、図面チェックプログラム(仮)', layout)

while True:
    event, values = window.read()
    print('イベント:', event ,', 値:',values) # 確認表示
    
    prod_name = sheetsearch.get_prodname(values['order_num'])

    # 生産計画シート内に読み込んだFS受注No.が存在しない場合、
    # ウィンドウを閉じる
    if prod_name == "n/a":    
        sg.Popup('受注No.が存在しません')
        break

    pdf_abs_path = sheetsearch.get_diagram_num(prod_name)
    if pdf_abs_path == "n/a":    
        sg.Popup('該当するPDF図面ファイルが存在しません')
        break

    if event == 'open_diagram':
        subprocess.Popen([pdf_viwer_path, pdf_abs_path], shell=True)

    if event == sg.WIN_CLOSED or event == 'Quit':
        break

window.close()