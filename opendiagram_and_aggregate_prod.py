"""
動作

1.受注No.のバーコードをスキャナで読み込んで、受注No.をフォームに入力する
2."生産計画シート"のFS受注No.の列から
"""

import PySimpleGUI as sg
import openpyxl
from os.path import basename

# ------生産計画シートのパス
production_schedule = ""

# ------図面No.<-->品名変換テーブルのパス
diagram_number = ""

# ------図面フォルダのパス

sg.theme('Dark Green 1')

layout = [
    [sg.Text('     受注No.(品質経歴書⑨)バーコードを読み込む', pad=((10,10),(20,20))), sg.InputText(key='order_num')],
    [sg.Button('図面を開く', size=(40,1), key='open_diagram'),sg.Button('閉じる', size=(40,1), key='Quit')]
]

window = sg.Window('生産計画、図面チェックプログラム(仮)', layout)

while True:
    event, values = window.read()
    print('イベント:', event ,', 値:',values) # 確認表示
    if event == sg.WIN_CLOSED or event == 'Quit':
        break

window.close()