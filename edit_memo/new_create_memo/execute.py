import datetime
import os

import xlwings as xw

from edit_memo.common import const
from edit_memo.common.util import sh_format, excel_edit_start, excel_edit_end
from edit_memo.new_create_memo.func import input_sh_format, input_index_sh_format, cover_sh_format


def new_create_memo(arg_wb: xw.main.Book):
    # 保存先フルパスを作成
    save_dir = os.path.dirname(arg_wb.fullname)
    save_name = f"【学習メモ】{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}.xlsm"
    save_fullname = os.path.join(save_dir, save_name)
    # 既存ブックの保存 & 別名で保存（コピーしたブックを開くことと同義）
    arg_wb.save()
    arg_wb.save(save_fullname)
    arg_wb = xw.books.active
    # 各種シート初期化:入力シート
    input_sh_format(arg_wb)
    # 各種シート初期化:索引登録シート
    input_index_sh_format(arg_wb)
    # 各種シート初期化:表紙シート
    cover_sh_format(arg_wb)
    # 各種シート初期化:目次、内容、索引
    sh_format(arg_wb.sheets(const.TOC_SH_NAME))
    sh_format(arg_wb.sheets(const.CONTENTS_SH_NAME))
    sh_format(arg_wb.sheets(const.INDEX_SH_NAME))
    arg_wb.sheets(const.COVER_SH_NAME).activate()
    arg_wb.save()


if __name__ == '__main__':
    wb = excel_edit_start()
    new_create_memo(wb)
    excel_edit_end(wb)
