import os
import settings
import glob
import openpyxl
from openpyxl.styles import PatternFill
import modules

# 定数
BOOK_TITLE  = settings.book_title       # ファイル名
SHEET_TITLE = settings.sheet_title      # シート名
ROW_INI     = settings.row_ini          # 行の初期位置
COL_INI     = settings.col_ini          # 列の初期位置




# OSを判定し、ファイルパスの区切り文字を決定
if os.name == "nt":         # Windows
    delimiter = "¥"
elif os.name == 'posix':    # Mac or Linux
    delimiter = "/"

# エクセルファイルの作成
book        = openpyxl.Workbook()     # 新規Bookの作成
sheet       = book.worksheets[0]      # シート設定
sheet.title = SHEET_TITLE             # シートのタイトル設定

# 初期位置を設定
row = ROW_INI
col = COL_INI

# 比較用のリストを作成
pre_file = []

# 同階層以下のフォルダ/ファイル一覧を取得
f_list = glob.glob("**", recursive=True)

# 一覧を書き出す処理
for file in f_list:
    file = file.split(delimiter)    #取得した一覧をデリミタで分割し、リスト化

    # ファイルがdesktop.iniの場合は書き出さない
    if file[-1] != "desktop.ini":
        for i, f_name in enumerate(file):
            if len(pre_file) == 0 or i >= len(pre_file) or file[i] != pre_file[i]:
                sheet.cell(row=row, column=col).value = f_name

            col += 1    # 次の列に移動

        pre_file = file

        row += 1        # 次の行に移動
        col = COL_INI   # 列の位置を初期位置に移動

# 罫線を引く処理
modules.draw_boader(sheet, ROW_INI, sheet.max_row, COL_INI, sheet.max_column)

# 背景色
fill = PatternFill(patternType='solid', fgColor='ffffff')
for row in sheet:
    for cell in row:
        sheet[cell.coordinate].fill = fill

# 保存 & 終了
book.save(BOOK_TITLE)
book.close()

# TODO:書き出さない処理は別にする
# TODO: 自分自身はファイルリストに含めない
# TODO: oldファイルは含めない
# TODO: 一番最後の行を含めない
# TODO:モジュール化する
# TODO:モジュールを別ファイルに切り出す

