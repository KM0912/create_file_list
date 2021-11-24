import os
import glob
from openpyxl.styles.borders import Border, Side

# フォルダ名/ファイル名が記入されている列より左側か右側かを判定するためのフラグの値
LEFT_CELL   = 0
RIGHT_CELL  = 1

# 罫線を引く処理
def draw_boader (sheet, start_row, end_row, start_col, end_col) :
    for row in range(start_row, end_row + 1):

        cell_flag = LEFT_CELL   # フォルダ名/ファイル名が記入されている列より左側か右側かを判定するためのフラグ
        for col in range(start_col, end_col + 1):

            # セルが空でなければ上罫線/左罫線を引く
            if (sheet.cell(row=row, column=col).value != None):
                sheet.cell(row=row, column=col).border = Border(top=Side(style='thin', color='000000'), left=Side(style='thin', color='000000'))

                #　空でないセルが１列目でなければ、左側のセルに左罫線を引いていく
                if col != 1:
                    for col_i in range(col-1, 0, -1):
                        sheet.cell(row=row, column=col_i).border = Border(left=Side(style='thin', color='000000'))

                cell_flag = RIGHT_CELL  # フォルダ名/ファイル名が記入されている列の罫線を引く処理が完了したので、フラグを更新

            # セルが空であり、フォルダ名/ファイル名が記入されている列より右側のセルの場合、上罫線のみを引く
            elif (cell_flag == RIGHT_CELL):
                sheet.cell(row=row, column=col).border = Border(top=Side(style='thin', color='000000'))

def get_delimiter () :
    if os.name == "nt":         # Windows
        return "¥"

    elif os.name == 'posix':    # Mac or Linux
        return "/"

def write_file_list (sheet, start_row, start_col, delimiter) :
    row = start_row
    col = start_col

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
            col = start_col # 列の位置を初期位置に移動