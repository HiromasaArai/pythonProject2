import datetime

import xlwings as xw


def test():
    # ws = xw.sheets.active
    wb = xw.books.active
    my_rg = wb.selection[0][0]
    # print(my_rg.address)
    my_rg.value = "test"


if __name__ == '__main__':
    test()
