import xlwings as xw

from edit_memo.common import const
from edit_memo.common.util import sh_format, get_cell_range, excel_edit_start, excel_edit_end


def input_sh_format(wb: xw.main.Book):
    sh = wb.sheets(const.INPUT_SH_NAME)
    start_rg = sh.range("B3")
    last_rg = sh.range(sh.range("A2").end("down").row, 10)
    rg = sh.range(start_rg, last_rg)
    rg.clear_contents()
    # フォントネームを強制
    rg.font.name = "ＭＳ ゴシック"
    # Noの初期化 セルを見て値があればリセットした値を入れる。
    rg = sh.range("A3")
    rg_val = 1
    rg.value = rg_val
    while rg.offset(1, 0).value is not None:
        rg_val += 1
        rg = rg.offset(1, 0)
        rg.value = rg_val


def input_index_sh_format(wb: xw.main.Book):
    sh = wb.sheets(const.INPUT_INDEX_SH_NAME)
    get_cell_range(sh, "B6", "B5").clear()


def cover_sh_format(wb: xw.main.Book):
    sh = wb.sheets(const.COVER_SH_NAME)
    sh.range("B7").value = "メモ"
    sh.range("G18:I23").clear_contents()
    sh.range("G18:I23").clear_contents()
    sh.range("G37:I38").clear_contents()
    sh.range("B41:D42").clear_contents()
    sh.range("G41:I42").clear_contents()
