import xlwings as xw


def get_cell_range(sh: xw.main.Sheet, srn, ern):
    # 連続するセル範囲を取得する
    # srn: str: start_rg_nm    セル範囲の始まり
    # ern: str: end_base_rg_nm セル範囲の終わり
    start_rg = sh.range(srn)
    last_rg = sh.range(sh.range(ern).end("down").row, sh.range(ern).end("right").column)
    return sh.range(start_rg, last_rg)


def sh_format(sh: xw.main.Sheet):
    start_rg = sh.range("B3")
    last_rg = sh.range(sh.range("C2").end("down").row, sh.range("C2").end("right").column)
    rg = sh.range(start_rg, last_rg)
    rg.clear()
