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
"""

import PySimpleGUI as sg
import openpyxl
from os.path import basename
import subprocess


sg.theme('Dark Green 1')

# ------生産計画シートのパス
prod_schedule_path = ""

# ------品名<-->図面No.変換テーブルのパス
pd_conv_table_path = ""

# ------図面フォルダのパス
diagram_dir_path = "C:\Users\Desktop\[図面の入っているフォルダ名]"

# ------変換テーブルで取得した図面No.
diagram_num = ""

# ------開こうとするPDF図面の絶対パス
pdf_abs_path = diagram_dir_path + diagram_num + ".pdf"


layout = [
    [sg.Text('     受注No.(品質経歴書⑨)バーコードを読み込む', pad=((10,10),(20,20))), sg.InputText(key='order_num')],
    [sg.Button('図面を開く', size=(40,1), key='open_diagram'),sg.Button('閉じる', size=(40,1), key='Quit')]
]

window = sg.Window('生産計画、図面チェックプログラム(仮)', layout)

while True:
    event, values = window.read()
    # print('イベント:', event ,', 値:',values) # 確認表示

# 3."生産計画シート"のFS受注No.の列から入力された受注Noを検索する。
#     受注No.が存在するなら
#         最後の列'済'欄にチェックして、
#         品目名のセルに入っている値を取得。ー＞変数:prod_nameへ代入する。
#     存在しないなら
#         "受注No.が存在しません"とかエラーを返す。


# 図面のフルパスを投げてpopen関数で図面ファイルを開く
    if event == 'open_diagram':


        
        subprocess.Popen(['start', pdf_abs_path], shell=True)





    if event == sg.WIN_CLOSED or event == 'Quit':
        break

window.close()