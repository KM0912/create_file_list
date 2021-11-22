import os
import settings
import glob
import openpyxl
from openpyxl.styles.borders import Border, Side
from openpyxl.styles import PatternFill

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
book = openpyxl.Workbook()      # 新規Bookの作成
sheet = book.worksheets[0]      # シート設定
sheet.title = SHEET_TITLE       # シートのタイトル設定

# 初期位置を設定
row = ROW_INI
col = COL_INI

# 比較用のリストを作成
pre_file = []

# 同階層以下のフォルダ/ファイル一覧を取得
files = glob.glob("**", recursive=True)

for file in files:
    file = file.split(delimiter)

    # ファイルがdesktop.iniの場合は書き出さない
    if file[-1] != "desktop.ini":
        for i, f_name in enumerate(file):
            if len(pre_file) == 0 or i >= len(pre_file) or file[i] != pre_file[i]:
                sheet.cell(row=row, column=col).value = f_name

            col += 1    # 次の列に移動

        pre_file = file

        row += 1        # 次の行に移動
        col = COL_INI   # 列の位置を初期位置に移動

print(sheet.max_row)


for row in range(ROW_INI, sheet.max_row + 1):
    flag = 0

    for col in range(COL_INI, 10):

        if (sheet.cell(row=row, column=col).value != None):
            sheet.cell(row=row, column=col).border = Border(top=Side(
                style='thin', color='000000'), left=Side(style='thin', color='000000'))

            if col != 1:
                for col_i in range(col-1, 0, -1):
                    sheet.cell(row=row, column=col_i).border = Border(
                        left=Side(style='thin', color='000000'))

            flag = 1

        elif (flag == 1):
            sheet.cell(row=row, column=col).border = Border(
                top=Side(style='thin', color='000000'))

# 背景色
fill = PatternFill(patternType='solid', fgColor='ffffff')
for row in sheet:
    for cell in row:
        sheet[cell.coordinate].fill = fill

# 保存 & 終了
book.save(BOOK_TITLE)
book.close()

# TODO: 自分自身はファイルリストに含めない
# TODO: oldファイルは含めない
# TODO: 一番最後の行を含めない
