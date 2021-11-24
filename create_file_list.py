import settings
import openpyxl
from openpyxl.styles import PatternFill
import modules

# 定数
BOOK_TITLE  = settings.book_title       # ファイル名
SHEET_TITLE = settings.sheet_title      # シート名
ROW_INI     = settings.row_ini          # 行の初期位置
COL_INI     = settings.col_ini          # 列の初期位置

# OSを判定し、ファイルパスの区切り文字を決定
delimiter = modules.get_delimiter()

# エクセルファイルの作成
book        = openpyxl.Workbook()     # 新規Bookの作成
sheet       = book.worksheets[0]      # シート設定
sheet.title = SHEET_TITLE             # シートのタイトル設定

# 初期位置を設定
row = ROW_INI
col = COL_INI

# 一覧を書き出す処理
modules.write_file_list(sheet, ROW_INI, COL_INI, delimiter)

# 罫線を引く処理
modules.draw_boader(sheet, ROW_INI, sheet.max_row, COL_INI, sheet.max_column)

# 背景色
for row in sheet:
    for cell in row:
        sheet[cell.coordinate].fill = PatternFill(patternType='solid', fgColor='ffffff')

# 保存 & 終了
book.save(BOOK_TITLE)
book.close()

# TODO:書き出さない処理は別にする
# TODO: 自分自身はファイルリストに含めない
# TODO: oldファイルは含めない
# TODO: 一番最後の行を含めない
# TODO:モジュールのリファクタリング

