import datetime

import xlwings as xw

from edit_memo.common import const
from edit_memo.common.util import get_cell_range


def func_input_sh(wb: xw.main.Book):
    """
    入力シート内データ取得に関する関数
    :param wb:
    :return:
    """
    sh = wb.sheets(const.INPUT_SH_NAME)
    rg = get_cell_range(sh, "A3", "C2")
    rg2 = sh.range("I3")
    for i in range(rg.rows.count):
        if rg2.offset(i, 0).value is None:
            rg2.offset(i, 0).value = datetime.datetime.now().date()

        if rg2.offset(i, 1).value is None:
            rg2.offset(i, 1).value = datetime.datetime.now().date()
    return rg


def func_toc_sh(wb: xw.main.Book, ish_array):
    """
    目次シート編集　及び取得した入力シートデータの並び替え
    :param wb:
    :param ish_array:
    :return:項番の付与された入力シートデータ（二次元配列）
    """
    sh = wb.sheets(const.TOC_SH_NAME)
    for i in range(len(ish_array)):
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


def func_cover(wb, ish_array):
    """
    表紙シート編集
    :param wb:
    :param ish_array:
    :return:
    """
    sh = wb.sheets(const.COVER_SH_NAME)
    last_update_date_rg = sh.range("G18")
    second_last_update_date_rg = sh.range("G20")
    third_last_update_date_rg = sh.range("G22")
    item_nm_rg = sh.range("G37")
    start_date_rg = sh.range("B41")
    end_update_date_rg = sh.range("G41")
    # 前々回更新日
    if second_last_update_date_rg.value is not None:
        third_last_update_date_rg.value = second_last_update_date_rg.value

    # 前回更新日
    if last_update_date_rg.value is not None:
        second_last_update_date_rg.value = last_update_date_rg.value
    # 最終更新日
    last_update_date_rg.value = datetime.datetime.now().strftime("%Y/%m/%d %T")
    # 項目数
    item_nm_rg.value = sorted(ish_array, key=lambda x: x[0], reverse=True)[0][0]
    # メモ作成開始日
    start_date_rg.value = sorted(ish_array, key=lambda x: x[8], reverse=False)[0][8]
    # メモ作成終了日
    end_update_date_rg.value = sorted(ish_array, key=lambda x: x[9], reverse=True)[0][9]


def func_contents(wb: xw.main.Book, ish_array):
    """
    内容シート編集
    :param wb:
    :param ish_array:
    :return: None
    """
    sh = wb.sheets(const.CONTENTS_SH_NAME)
    # 索引シートのデータを取得　i_index_array
    i_index_array = get_cell_range(wb.sheets(const.INPUT_INDEX_SH_NAME), "B6", "B5").options(ndim=2).value
    for i in range(len(ish_array)):
        magic_num = i * 7 + 3
        # 項目記入:標語
        sh.cells(magic_num + 0, 2).value = i + 1
        sh.cells(magic_num + 0, 3).value = "標語"
        sh.cells(magic_num + 0, 3).font.bold = True
        sh.cells(magic_num + 0, 4).value = ish_array[i][1]
        # 項目記入:別名
        synonym = ""
        for i_index in i_index_array:
            # 管理番号の一致判定と標語の不一致判定
            if i_index[5] == ish_array[i][0] and i_index[3] != ish_array[i][1]:
                if synonym == "":
                    synonym += i_index[3]
                else:
                    synonym += ", " + i_index[3]
        sh.cells(magic_num + 1, 3).value = "別名"
        if synonym != "":
            sh.cells(magic_num + 1, 4).value = synonym
        # 項目記入:分類
        sh.cells(magic_num + 2, 3).value = "分類"
        sh.cells(magic_num + 2, 4).value = ish_array[i][3]
        # 項目記入:事実
        sh.cells(magic_num + 3, 3).value = "事実"
        sh.cells(magic_num + 3, 4).value = ish_array[i][4]
        # 項目記入:抽象
        sh.cells(magic_num + 4, 3).value = "抽象"
        sh.cells(magic_num + 4, 4).value = ish_array[i][5]
        # 項目記入:転用
        sh.cells(magic_num + 5, 3).value = "転用"
        sh.cells(magic_num + 5, 4).value = ish_array[i][6]
        # 項目記入:補足
        sh.cells(magic_num + 6, 3).value = "補足"
        sh.cells(magic_num + 6, 4).value = ish_array[i][7]
        # 項目記入:記入日（更新日）
        sh.cells(magic_num + 0, 5).value = ish_array[i][9]
        # 折り返して表示
        sh.range(sh.cells(magic_num + 0, 4), sh.cells(magic_num + 6, 4)).api.WrapText = True
        # 罫線を引く1
        rg = sh.range((magic_num + 0, 3), (magic_num + 6, 4))
        rg.api.Borders(7).LineStyle = 1
        rg.api.Borders(11).LineStyle = 1
        rg.api.Borders(12).LineStyle = 1
        # 罫線を引く2
        rg = sh.range((magic_num + 0, 5), (magic_num + 6, 5))
        rg.api.Borders(11).LineStyle = 1
        rg.api.Borders(12).LineStyle = 1
        # 罫線を引く3
        rg = sh.range((magic_num + 0, 2), (magic_num + 6, 5))
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


def func_input_index(wb: xw.main.Book, ish_array):
    """
    索引登録シート初期設定及びデータ取得に関する関数
    :param wb:
    :param ish_array:
    :return:
    """
    last_data_row = 6
    sh = wb.sheets(const.INPUT_INDEX_SH_NAME)
    if sh.range("B6").value is not None:
        last_data_row = sh.range("B5").end("down").row + 1
        rg = get_cell_range(sh, "B6", "B5")
        # 索引登録シートのデータをリストして取得
        i_index_array = rg.options(ndim=2).value
        # 入力シートのデータ数だけ繰り返す
        for i in range(len(ish_array)):
            # 索引登録シートのデータ数だけ繰り返す
            is_match = False
            for i2 in range(rg.rows.count):
                # 値の一致確認：管理No
                term1 = ish_array[i][0] == i_index_array[i2][5]
                if term1:
                    # 管理Noが一致していれば目次Noと分類入れ替え
                    sh.range("C6").offset(i2, 0).value = ish_array[i][10]
                    sh.range("C6").offset(i2, 1).value = ish_array[i][3]
                    is_match = True
            # 索引登録シート内に一致しているデータがない場合は新規登録
            if not is_match:
                last_data_row = index_output(sh, ish_array, i, last_data_row)
    else:
        # 入力シートのデータ数だけ繰り返す
        for i in range(len(ish_array)):
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
    """
    索引シート編集
    :param wb:
    :return:
    """
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
