Attribute VB_Name = "executer"
Option Explicit

' iniファイル取得用
Declare PtrSafe Function GetPrivateProfileString Lib _
    "kernel32" Alias "GetPrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String _
) As Long

Function get_ini_value(ini_section As String, ini_key As String)
    Dim sPath                   '// INIファイルのパス
    Dim sValue      As String   '// 取得値
    Dim lSize       As Long     '// 取得値のサイズ
    Dim lRet        As Long     '// 戻り値
    
    lSize = 2000
    sPath = "C:\PythonPath.ini"

    ' 取得バッファを初期化
    sValue = Space(lSize)
    lRet = GetPrivateProfileString(ini_section, ini_key, "non", sValue, lSize, sPath)
    
    ' 値を返す
    get_ini_value = Trim(Left(sValue, InStr(sValue, Chr(0)) - 1))
End Function

Private Function file_check(py_venv As String, create_pj_path As String, py_fullname As String) As String
    Dim err_msg As String: err_msg = ""
    
    ' 実行したい仮想環境が存在しない場合は処理中断
    If Dir(py_venv) = "" Then
        err_msg = "有効化したい仮想環境[" & py_venv & "]が存在しません。"
    End If
    
    ' Python探索パスを設定するpyモジュール
    If Dir(create_pj_path) = "" Then
        err_msg = "有効化したい仮想環境[" & create_pj_path & "]が存在しません。"
    End If
    
    ' 実行したいpyモジュールが存在しない場合は処理中断
    If Dir(py_fullname) = "" Then
        err_msg = "実行したいpyファイル[" & py_fullname & "]が存在しません。"
    End If
    
    file_check = err_msg
End Function

' コマンドをアンパッサーと半角スペースで繋ぐ関数（配列を渡すとcmdを返す）
Private Function create_cmd(cmd_list As Variant) As String
    Dim cmd As String
    Dim i As Integer
    For i = 0 To UBound(cmd_list)
        If (i > 0) Then
            cmd = cmd & " & "
        End If
        cmd = cmd & cmd_list(i)
    Next i
    ' コマンドプロンプトに渡せるコマンドが1行しかないのでそのようにコマンドを整理している
    create_cmd = "cmd.exe /c " & cmd
End Function

' ▲▲▲▲メイン関数の支援関数▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
' ▼▼▼▼メイン関数▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

' コマンドプロンプト経由でPython実行
Sub run_python(py_name As String)
    Dim err_msg As String: err_msg = "予期しないエラーです。"
    On Error GoTo Catch
        ' 高速化処理
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        
        Dim wsh As Object: Set wsh = CreateObject("WScript.Shell")
        
        ' 実行したいpyファイルが存在するプロジェクトディレクトリ
        Dim py_pj As String: py_pj = get_ini_value("python_path", "py_pj")

        ' 実行したい仮想環境
        Dim py_venv As String
        py_venv = py_pj & "\" & get_ini_value("python_path", "venv") & "\" & get_ini_value("python_path", "venv_batch")
        
        ' 実行したいpyモジュール(Python探索パスを設定するpyモジュール)
        Dim create_pj_path As String: create_pj_path = py_pj & "\create_pj_path.py"
        
        ' 実行したいpyモジュール(フルパス)
        Dim py_fullname As String: py_fullname = py_pj & "\" & py_name
        
        ' エラーチェック（ファイルの存在確認）
        Dim chk_result As String: chk_result = file_check(py_venv, create_pj_path, py_fullname)
        If chk_result <> "" Then
            err_msg = chk_result
            GoTo Catch
        End If
        
        ' コマンドを作成及び実行1
        Dim cmd_list(1)
        cmd_list(0) = py_venv
        cmd_list(1) = "python " & create_pj_path
        ' create_cmd関数 : 配列に格納したコマンドをワンラインに整形する関数
        Call wsh.Run(create_cmd(cmd_list), vbHide, True)
        
        ' コマンドを作成及び実行2
        cmd_list(1) = "python " & py_fullname
        Call wsh.Run(create_cmd(cmd_list), vbHide, True)
        
        ' メモリ解放
        Set wsh = Nothing
        
        ' 正常終了
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        
        Exit Sub
Catch:
        ' 例外処理
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        
        MsgBox err_msg
End Sub
