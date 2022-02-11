import xlwings as xw


def get_cell_range(sh: xw.main.Sheet, srn, ern):
    """
    連続するセル範囲を取得する
    :param sh: xlwingsのシートオブジェクト
    :param srn: start_rg_nm 始まりのセル
    :param ern: 終わりのセルの起点となるもの
    :return: セル範囲
    """
    start_rg = sh.range(srn)
    last_rg = sh.range(sh.range(ern).end("down").row, sh.range(ern).end("right").column)
    return sh.range(start_rg, last_rg)


def sh_format(sh: xw.main.Sheet):
    start_rg = sh.range("B3")
    last_rg = sh.range(sh.range("C2").end("down").row, sh.range("C2").end("right").column)
    rg = sh.range(start_rg, last_rg)
    rg.clear()


def excel_edit_start():
    # アクティブブックを取得
    wb = xw.books.active
    # 高速モード>>開始
    wb.app.calculation = "manual"
    wb.app.screen_updating = False
    return wb


def excel_edit_end(wb: xw.main.Book):
    # 高速モード>>終了
    wb.app.calculation = "automatic"
    wb.app.screen_updating = True


def or_chk_is_none(*args):
    """
    引数に一つでもNoneがあったらTrue
    :param args:
    :return:
    """
    is_result = False
    for arg in args:
        if arg is None:
            is_result = True
    return is_result
