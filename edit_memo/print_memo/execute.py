import xlwings as xw

from edit_memo.common import const
from edit_memo.common.util import excel_edit_start, excel_edit_end


def sh_page_setup(print_size, sh: xw.main.Sheet):
    sh.api.PageSetup.Zoom = False
    sh.api.PageSetup.FitToPagesWide = 1
    sh.api.PageSetup.FitToPagesTall = False
    sh.api.PageSetup.CenterHorizontally = True
    sh.api.PageSetup.PaperSize = print_size


def print_memo(arg_wb: xw.main.Book, print_size):
    # 印刷設定
    sh_cover = arg_wb.sheets(const.COVER_SH_NAME)
    sh_page_setup(print_size, sh_cover)
    sh_toc = arg_wb.sheets(const.TOC_SH_NAME)
    sh_page_setup(print_size, sh_toc)
    sh_contents = arg_wb.sheets(const.CONTENTS_SH_NAME)
    sh_page_setup(print_size, sh_contents)
    sh_index = arg_wb.sheets(const.INDEX_SH_NAME)
    sh_page_setup(print_size, sh_index)
    # 印刷
    sh_cover.api.PrintOut(Preview=False)
    sh_toc.api.PrintOut(Preview=False)
    sh_contents.api.PrintOut(Preview=False)
    sh_index.api.PrintOut(Preview=False)


if __name__ == '__main__':
    wb = excel_edit_start()
    # 9 == ペーパーサイズA4
    print_memo(wb, 9)
    excel_edit_end(wb)
