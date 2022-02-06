import datetime

import xlwings as xw

from edit_memo.common import const
from edit_memo.common.util import sh_format


def get_cell_range(sh: xw.main.Sheet, srn, ern):
    # 連続するセル範囲を取得する
    # srn: str: start_rg_nm
    # ern: str: end_base_rg_nm
    start_rg = sh.range(srn)
    last_rg = sh.range(sh.range(ern).end("down").row, sh.range(ern).end("right").column)
    return sh.range(start_rg, last_rg)


def func_input_sh(wb: xw.main.Book):
    sh = wb.sheets(const.INPUT_SH_NAME)
    rg = get_cell_range(sh, "A3", "C2")
    # rrc = rg.rows.count
    rg2 = sh.range("I3")
    for i in range(rg.rows.count):
        if rg2.offset(i, 0).value is None:
            rg2.offset(i, 0).value = datetime.datetime.now().date()

        if rg2.offset(i, 1).value is None:
            rg2.offset(i, 1).value = datetime.datetime.now().date()

    return rg


def func_toc_sh(wb: xw.main.Book, rrc, ish_array):
    sh = wb.sheets(const.TOC_SH_NAME)
    for i in range(rrc):
        # 分類
        sh.cells(i + 3, 2).value = ish_array[i][3]
        # 標語
        sh.cells(i + 3, 3).value = ish_array[i][1]
        # 作成日
        sh.cells(i + 3, 4).value = ish_array[i][8]
        # 更新日
        sh.cells(i + 3, 5).value = ish_array[i][9]
        # No
        sh.cells(i + 3, 6).value = i + 1
        # 状態
        # if (datetime.datetime.now().date() - ish_array[i][8]).days <= 7:
        #     sh.cells(i + 3, 7).value = "NEW"
        # 管理番号
        sh.cells(i + 3, 8).value = ish_array[i][0]
        # 項目番号をリストに追加
        ish_array[i].append(i + 1)

    rg = get_cell_range(sh, "B3", "B2")
    # 罫線を引く　それぞれ左、下、右、中
    rg.api.Borders(7).LineStyle = 1
    rg.api.Borders(9).LineStyle = 1
    rg.api.Borders(10).LineStyle = 1
    rg.api.Borders(11).LineStyle = 1
    rg.api.Borders(12).LineStyle = 1
    # フォントネームを強制
    rg.font.name = "ＭＳ ゴシック"
    return ish_array


def func_contents(wb: xw.main.Book, rrc, ish_array):
    sh = wb.sheets(const.CONTENTS_SH_NAME)
    for i in range(rrc):
        magic_num = i * 6 + 3
        # 項目記入
        sh.cells(magic_num + 0, 2).value = i + 1
        sh.cells(magic_num + 0, 3).value = "標語"
        sh.cells(magic_num + 0, 3).font.bold = True
        sh.cells(magic_num + 0, 4).value = ish_array[i][1]
        sh.cells(magic_num + 1, 3).value = "分類"
        sh.cells(magic_num + 1, 4).value = ish_array[i][3]
        sh.cells(magic_num + 2, 3).value = "事実"
        sh.cells(magic_num + 2, 4).value = ish_array[i][4]
        sh.cells(magic_num + 3, 3).value = "抽象"
        sh.cells(magic_num + 3, 4).value = ish_array[i][5]
        sh.cells(magic_num + 4, 3).value = "転用"
        sh.cells(magic_num + 4, 4).value = ish_array[i][6]
        sh.cells(magic_num + 5, 3).value = "補足"
        sh.cells(magic_num + 5, 4).value = ish_array[i][7]
        # 項目記入:記入日（更新日）
        sh.cells(magic_num + 0, 5).value = ish_array[i][9]
        # 罫線を引く1
        rg = sh.range((magic_num + 0, 3), (magic_num + 5, 4))
        rg.api.Borders(7).LineStyle = 1
        rg.api.Borders(11).LineStyle = 1
        rg.api.Borders(12).LineStyle = 1
        # 罫線を引く2
        rg = sh.range((magic_num + 0, 5), (magic_num + 5, 5))
        rg.api.Borders(11).LineStyle = 1
        rg.api.Borders(12).LineStyle = 1
        # 罫線を引く3
        rg = sh.range((magic_num + 0, 2), (magic_num + 5, 5))
        rg.api.Borders(7).LineStyle = 1
        rg.api.Borders(7).Weight = 3
        rg.api.Borders(9).LineStyle = 1
        rg.api.Borders(9).Weight = 3
        rg.api.Borders(10).LineStyle = 1
        rg.api.Borders(10).Weight = 3

    # フォントネームを強制
    rg = get_cell_range(sh, "B3", "C2")
    rg.font.name = "ＭＳ ゴシック"
    return None


def index_output(sh: xw.main.Sheet, ish_array, i, last_data_row):
    sh.cells(last_data_row, 2).value = last_data_row - 5
    sh.cells(last_data_row, 3).value = ish_array[i][10]
    sh.cells(last_data_row, 4).value = ish_array[i][3]
    sh.cells(last_data_row, 5).value = ish_array[i][1]
    sh.cells(last_data_row, 6).value = ish_array[i][2]
    sh.cells(last_data_row, 7).value = ish_array[i][0]
    return last_data_row + 1


def func_input_index(wb: xw.main.Book, rrc, ish_array):
    last_data_row = 6
    sh = wb.sheets(const.INPUT_INDEX_SH_NAME)
    if sh.range("B6").value is not None:
        last_data_row = sh.range("B5").end("down").row + 1
        rg = get_cell_range(sh, "B6", "B5")
        # 索引登録シートのデータをリストして取得
        i_index_array = rg.options(ndim=2).value
        # 入力シートのデータ数だけ繰り返す
        for i in range(rrc):
            # 索引登録シートのデータ数だけ繰り返す
            is_flag = True
            for i2 in range(rg.rows.count):
                term1 = ish_array[i][0] == i_index_array[i2][5]
                term2 = ish_array[i][3] == i_index_array[i2][2]
                term3 = ish_array[i][1] == i_index_array[i2][3]
                term4 = ish_array[i][2] == i_index_array[i2][4]
                if term1 and term2 and term3 and term4:
                    is_flag = False
            if is_flag:
                last_data_row = index_output(sh, ish_array, i, last_data_row)
    else:
        # 入力シートのデータ数だけ繰り返す
        for i in range(rrc):
            last_data_row = index_output(sh, ish_array, i, last_data_row)

    # 罫線を引く
    rg = get_cell_range(sh, "B6", "B5")
    rg.api.Borders(7).LineStyle = 1
    rg.api.Borders(9).LineStyle = 1
    rg.api.Borders(10).LineStyle = 1
    rg.api.Borders(11).LineStyle = 1
    rg.api.Borders(12).LineStyle = 1
    # フォントネームを強制
    rg.font.name = "ＭＳ ゴシック"


def func_index(wb: xw.main.Book):
    sh = wb.sheets(const.INPUT_INDEX_SH_NAME)
    rg = get_cell_range(sh, "B6", "B5")
    # 索引登録シートのデータをリストして取得し、「ヒョウゴ」項目でソート
    i_index_array = sorted(rg.options(ndim=2).value, key=lambda x: x[4])
    sh = wb.sheets(const.INDEX_SH_NAME)
    for i in range(rg.rows.count):
        # 項目入力
        sh.cells(i + 3, 2).value = i_index_array[i][3]
        sh.cells(i + 3, 3).value = i_index_array[i][4]
        sh.cells(i + 3, 4).value = i_index_array[i][2]
        sh.cells(i + 3, 5).value = i_index_array[i][1]
        sh.cells(i + 3, 6).value = i_index_array[i][5]

    # 罫線を引く
    rg = get_cell_range(sh, "B3", "B2")
    rg.api.Borders(7).LineStyle = 1
    rg.api.Borders(9).LineStyle = 1
    rg.api.Borders(10).LineStyle = 1
    rg.api.Borders(11).LineStyle = 1
    rg.api.Borders(12).LineStyle = 1
    # フォントネームを強制
    rg.font.name = "ＭＳ ゴシック"


def create_memo(wb: xw.main.Book):
    # 入力シートのデータを取得してソート（分類ごと）
    rg = func_input_sh(wb)
    rrc = rg.rows.count
    ish_array = sorted(rg.options(ndim=2).value, key=lambda x: x[3])
    # 各シート初期化 >>目次、内容、索引
    sh_format(wb.sheets(const.TOC_SH_NAME))
    sh_format(wb.sheets(const.CONTENTS_SH_NAME))
    sh_format(wb.sheets(const.INDEX_SH_NAME))
    # 目次シート入力
    ish_array = func_toc_sh(wb, rrc, ish_array)
    # 内容シート入力
    func_contents(wb, rrc, ish_array)
    # 索引登録シート入力
    func_input_index(wb, rrc, ish_array)
    # 索引シートの入力
    func_index(wb)
