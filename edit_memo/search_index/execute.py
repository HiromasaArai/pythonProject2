from edit_memo.common import const
from edit_memo.common.util import excel_edit_start, excel_edit_end, get_cell_range


def search_index(arg_wb):
    sh = arg_wb.sheets(const.INPUT_INDEX_SH_NAME)
    rg = get_cell_range(sh, "B6", "B5")
    search_val = sh.range("G3").value
    index_data = rg.options(ndim=2).value
    msg_rg = sh.range("J1")
    is_being_val = False
    if search_val is not None:
        for i in range(rg.rows.count):
            if index_data[i][5] == search_val:
                sh.cells(3, 3).value = index_data[i][1]
                sh.cells(3, 4).value = index_data[i][2]
                sh.cells(3, 9).value = index_data[i][3]
                sh.cells(3, 10).value = index_data[i][4]
                is_being_val = True
                msg_rg.clear_contents()
                break

    if not is_being_val:
        sh.cells(3, 3).clear_contents()
        sh.cells(3, 4).clear_contents()
        sh.cells(3, 9).clear_contents()
        sh.cells(3, 10).clear_contents()
        msg_rg.value = "値が存在しませんでした。"


if __name__ == '__main__':
    wb = excel_edit_start()
    search_index(wb)
    excel_edit_end(wb)
