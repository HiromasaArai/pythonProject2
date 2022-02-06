import xlwings as xw

from edit_memo.common import const


def sh_page_setup(print_size, sh: xw.main.Sheet):
    sh.api.PageSetup.Zoom = False
    sh.api.PageSetup.FitToPagesWide = 1
    sh.api.PageSetup.FitToPagesTall = False
    sh.api.PageSetup.CenterHorizontally = True
    sh.api.PageSetup.PaperSize = print_size


def print_memo(wb: xw.main.Book, print_size):
    sh = wb.sheets(const.COVER_SH_NAME)
    sh_page_setup(print_size, sh)
    sh = wb.sheets(const.TOC_SH_NAME)
    sh_page_setup(print_size, sh)
    sh = wb.sheets(const.CONTENTS_SH_NAME)
    sh_page_setup(print_size, sh)
    sh = wb.sheets(const.INDEX_SH_NAME)
    sh_page_setup(print_size, sh)
