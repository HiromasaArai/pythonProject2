Attribute VB_Name = "Main"
Option Explicit

' ���W���[�����ׂ͂�����
' �p�b�P�[�W����ini�t�@�C������擾����@get_ini_value()

Sub �����X�V()
    Dim py_name As String: py_name = "main_create_memo.py"
    Call run_python(get_ini_value("package", "edit_memo") & "\" & py_name)
End Sub

Sub �V�K�����쐬()
    Dim py_name As String: py_name = "main_new_create_memo.py"
    Call run_python(get_ini_value("package", "edit_memo") & "\" & py_name)
End Sub
