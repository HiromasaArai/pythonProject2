import configparser
import os.path

if __name__ == '__main__':
    """
    設定ファイル(Pythonプロジェクトを指し示す設定ファイル)を作成する
    """
    ini = configparser.ConfigParser()
    ini.read(r"C:\PythonPath.ini", "UTF-8")
    pj_path = ini["user_path"]["py_pj"]
    venv_nm = ini["user_path"]["venv"]
    pj_path_conf = f"{pj_path}/{venv_nm}/set_path.pth"
    # 設定ファイルがある場合は削除する
    if os.path.exists(pj_path_conf):
        os.remove(pj_path_conf)

    # 設定ファイルを作成し、そこにPythonプロジェクトへのフルパスを書き込む
    with open(pj_path_conf, mode="w") as f:
        f.write(pj_path)
