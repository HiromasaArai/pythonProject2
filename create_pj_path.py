import os.path
import sys

if __name__ == '__main__':
    # コマンドラインから引数受け取り
    pj_path = sys.argv[1]
    pj_path_conf = pj_path + "/.venv/set_path.pth"
    # 設定ファイル = Pythonプロジェクトを指し示す設定ファイル
    # 設定ファイルがある場合は削除する
    if os.path.exists(pj_path_conf):
        os.remove(pj_path_conf)

    # 設定ファイルを作成し、そこにPythonプロジェクトへのフルパスを書き込む
    with open(pj_path_conf, mode="w") as f:
        f.write(pj_path)
