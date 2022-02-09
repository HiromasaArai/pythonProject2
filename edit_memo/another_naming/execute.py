from edit_memo.common.util import excel_edit_start, excel_edit_end


def another_naming(arg_wb):
    pass


if __name__ == '__main__':
    wb = excel_edit_start()
    another_naming(wb)
    excel_edit_end(wb)
