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

                #　１列目でなければ
                # if col != 1:
                #     for col_i in range(col-1, 0, -1):
                #         sheet.cell(row=row, column=col_i).border = Border(left=Side(style='thin', color='000000'))

                cell_flag = RIGHT_CELL  # フォルダ名/ファイル名が記入されている列の罫線を引く処理が完了したので、フラグを更新

            # セルが空であり、フォルダ名/ファイル名が記入されている列より右側のセルの場合、上罫線のみを引く
            elif (cell_flag == RIGHT_CELL):
                sheet.cell(row=row, column=col).border = Border(top=Side(style='thin', color='000000'))
