Attribute VB_Name = "Main"
Option Explicit

' モジュール名はべた書き
' パッケージ名はiniファイルから取得する　get_ini_value()

Sub メモ更新()
    Dim py_name As String: py_name = "create_memo\execute.py"
    Call run_python(get_ini_value("package", "edit_memo") & "\" & py_name)
End Sub

Sub 新規メモ作成()
    Dim py_name As String: py_name = "new_create_memo\execute.py"
    Call run_python(get_ini_value("package", "edit_memo") & "\" & py_name)
End Sub

Sub 登録済み索引検索()
    Dim py_name As String: py_name = "search_index\execute.py"
    Call run_python(get_ini_value("package", "edit_memo") & "\" & py_name)
End Sub

Sub 索引登録()
    Dim py_name As String: py_name = "another_naming\execute.py"
    Call run_python(get_ini_value("package", "edit_memo") & "\" & py_name)
End Sub
