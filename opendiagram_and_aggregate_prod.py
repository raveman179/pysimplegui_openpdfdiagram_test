"""
動作

1.受注No.のバーコードをスキャナで読み込んで、受注No.をフォームに入力する
2.図面を開くボタンを左クリック
3."生産計画シート"のFS受注No.の列から入力された受注Noを検索する。
    受注No.が存在するなら
        最後の列'済'欄にチェックして、
        品目名のセルに入っている値を取得。ー＞変数:prod_nameへ代入する。
    存在しないなら
        "受注No.が存在しません"とかエラーを返す。
4.図面フォルダのパス内から prod_name + '.pdf'と合致するファイルを開く。
    変換テーブルを検索して図面が存在するなら
        Popenで図面を開く
    何らかの理由(横型以外の品名等)で図面がフォルダ内に存在しない場合
        "図面ファイルが存在しません"エラーを返す。
"""

from planningsheet_parse import sheet_search
import PySimpleGUI as sg
from os.path import basename
import subprocess
import re

sheetsearch = sheet_search()

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
        subprocess.Popen(['start', pdf_abs_path], shell=True)

    if event == sg.WIN_CLOSED or event == 'Quit':
        break

window.close()