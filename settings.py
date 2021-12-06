# 定数
book_title = "./test.xlsx"      # ファイル名      
sheet_title = "file_list"       # シート名
start_row = 1                     # 行の初期位置
start_col = 1  

# 列幅
col_widht       = 4     # 列幅
end_col_widht   = 30    # 最後の列の列幅


# 書き出し対象外にするフォルダ名/ファイル名のリスト
exclusion_list = [
    "settings.py",          # 本ファイル
    "modules.py",           # 関数を定義しているファイル
    "create_file_list.py",  # メイン処理のファイル
    "desktop.ini",
]