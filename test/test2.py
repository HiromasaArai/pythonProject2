import datetime

import xlwings as xw
from edit_memo.common import const


def test():
    wb = xw.books.active
    sh = wb.sheets(const.INPUT_SH_NAME)
    # <Range [【学習メモ_ver1.02】Java基礎知識_ver001.xlsm]入力!$A$2:$J$502>
    # rg = sh.range("A2").expand("table")
    # rg = rg[rg.count - 1]
    # <Range [【学習メモ】20220122103804.xlsm]入力!$B$1>
    rg = sh.range(sh.range("A2").end("down").row, sh.range("A2").end("right").column)

    print(type(wb))


def test2():
    # xlwings経由でアクティブブックを取得し、残りの処理をopenpyxlで実行してみる
    wb = xw.books.active
    sh = wb.sheets.active
    start_rg = sh.range("A3")
    last_rg = sh.range(sh.range("C2").end("down").row, sh.range("C2").end("right").column)
    rg = sh.range(start_rg, last_rg)
    my_array = rg.options(ndim=2).value
    print(rg.value)
    # 配列に格納した文字列を別の場所に出力
    sh = wb.sheets.add("sample")
    rrc = rg.rows.count
    rcc = rg.columns.count
    for i in range(rrc):
        for ii in range(rcc):
            print(my_array[i][ii])
            sh.cells(i + 1, ii + 1).value = my_array[i][ii]


def test3():
    # 入力シートのデータを取得
    wb = xw.books.active
    sh = wb.sheets(const.INPUT_SH_NAME)
    start_rg = sh.range("A3")
    last_rg = sh.range(sh.range("C2").end("down").row, sh.range("C2").end("right").column)
    rg = sh.range(start_rg, last_rg)
    my_array = rg.options(ndim=2).value
    # 配列に格納した文字列を別の場所に出力
    sh = wb.sheets.add("sample")
    rrc = rg.rows.count
    rcc = rg.columns.count
    # for i in range(rrc):
    #     for ii in range(rcc):
    #         sh.cells(i + 1, ii + 1).value = my_array[i][ii].lower()

    my_array = sorted(my_array, key=lambda x: x[2])

    for i in range(rrc):
        for ii in range(rcc):
            # print(my_array[i][ii])
            sh.cells(i + 1, ii + 1).value = my_array[i][ii]


def test4():
    dt = datetime.datetime.now().date()
    dt_add = datetime.timedelta(days=5)
    dt2 = dt + dt_add
    ans_dt = dt2 - dt

    print(ans_dt.days)


if __name__ == '__main__':
    test4()
