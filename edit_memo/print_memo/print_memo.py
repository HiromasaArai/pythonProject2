import xlwings as xw

from edit_memo.common import const


def sh_page_setup(print_size, sh: xw.main.Sheet):
    sh.api.PageSetup.Zoom = False
    sh.api.PageSetup.FitToPagesWide = 1
    sh.api.PageSetup.FitToPagesTall = False
    sh.api.PageSetup.CenterHorizontally = True
    sh.api.PageSetup.PaperSize = print_size


def print_memo(wb: xw.main.Book, print_size):
    sh = wb.sheets(const.HYOUSHI_SH_NAME)
    sh_page_setup(print_size, sh)
    sh = wb.sheets(const.MOKUZI_SH_NAME)
    sh_page_setup(print_size, sh)
    sh = wb.sheets(const.NAIYOU_SH_NAME)
    sh_page_setup(print_size, sh)
    sh = wb.sheets(const.SAKUIN_SH_NAME)
    sh_page_setup(print_size, sh)
