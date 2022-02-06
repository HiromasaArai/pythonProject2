Attribute VB_Name = "Main"
Option Explicit

' モジュール名はべた書き
' パッケージ名はiniファイルから取得する　get_ini_value()

Sub メモ更新()
    Dim py_name As String: py_name = "main_create_memo.py"
    Call run_python(get_ini_value("package", "edit_memo") & "\" & py_name)
End Sub

Sub 新規メモ作成()
    Dim py_name As String: py_name = "main_new_create_memo.py"
    Call run_python(get_ini_value("package", "edit_memo") & "\" & py_name)
End Sub
