import xlwings as xw


def sh_format(sh: xw.main.Sheet):
    start_rg = sh.range("B3")
    last_rg = sh.range(sh.range("C2").end("down").row, sh.range("C2").end("right").column)
    rg = sh.range(start_rg, last_rg)
    rg.clear()
