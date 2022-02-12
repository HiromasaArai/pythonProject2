from edit_memo.common import const
from edit_memo.common.util import sh_format, excel_edit_start, excel_edit_end
from edit_memo.create_memo.func import func_input_sh, func_toc_sh, func_contents, func_input_index, \
    func_index, func_cover


def create_memo(arg_wb):
    """
    シートごとの処理を分けている複合関数
    :param arg_wb:
    :return:
    """
    # 入力シートのデータを取得してソート（分類ごと）
    rg = func_input_sh(arg_wb)
    # rrc = rg.rows.count
    ish_array = sorted(rg.options(ndim=2).value, key=lambda x: x[3])
    # 各シート初期化 >>目次、内容、索引
    sh_format(arg_wb.sheets(const.TOC_SH_NAME))
    sh_format(arg_wb.sheets(const.CONTENTS_SH_NAME))
    sh_format(arg_wb.sheets(const.INDEX_SH_NAME))
    # 表紙シート入力
    func_cover(wb, ish_array)
    # 目次シート入力
    ish_array = func_toc_sh(arg_wb, ish_array)
    # 索引登録シート入力
    func_input_index(arg_wb, ish_array)
    # 内容シート入力
    func_contents(arg_wb, ish_array)
    # 索引シートの入力
    func_index(arg_wb)


if __name__ == '__main__':
    wb = excel_edit_start()
    create_memo(wb)
    excel_edit_end(wb)
