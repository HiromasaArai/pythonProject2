import xlwings as xw

from edit_memo.create_memo.create_memo import create_memo
from edit_memo.new_create_memo.new_create_memo import new_create_memo
from edit_memo.print_memo.print_memo import print_memo


def excel_edit_exe(exe_str):
    # アクティブブックを取得
    wb = xw.books.active
    # 高速モード>>開始
    wb.app.calculation = "manual"
    wb.app.screen_updating = False

    # 引数に応じた関数を実行
    if exe_str == "new_create_memo":
        new_create_memo(wb)
    elif exe_str == "create_memo":
        create_memo(wb)
    elif exe_str == "print_memo_A4":
        print_memo(wb, "A4")
    elif exe_str == "print_memo_A5":
        print_memo(wb, "A5")

    # 高速モード>>終了
    wb.app.calculation = "automatic"
    wb.app.screen_updating = True
    