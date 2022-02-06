Attribute VB_Name = "executer"
Option Explicit

' ini�t�@�C���擾�p
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
    Dim sPath                   '// INI�t�@�C���̃p�X
    Dim sValue      As String   '// �擾�l
    Dim lSize       As Long     '// �擾�l�̃T�C�Y
    Dim lRet        As Long     '// �߂�l
    
    lSize = 2000
    sPath = "C:\PythonPath.ini"

    ' �擾�o�b�t�@��������
    sValue = Space(lSize)
    lRet = GetPrivateProfileString(ini_section, ini_key, "non", sValue, lSize, sPath)
    
    ' �l��Ԃ�
    get_ini_value = Trim(Left(sValue, InStr(sValue, Chr(0)) - 1))
End Function

Private Function file_check(py_venv As String, create_pj_path As String, py_fullname As String) As String
    Dim err_msg As String: err_msg = ""
    
    ' ���s���������z�������݂��Ȃ��ꍇ�͏������f
    If Dir(py_venv) = "" Then
        err_msg = "�L�������������z��[" & py_venv & "]�����݂��܂���B"
    End If
    
    ' Python�T���p�X��ݒ肷��py���W���[��
    If Dir(create_pj_path) = "" Then
        err_msg = "�L�������������z��[" & create_pj_path & "]�����݂��܂���B"
    End If
    
    ' ���s������py���W���[�������݂��Ȃ��ꍇ�͏������f
    If Dir(py_fullname) = "" Then
        err_msg = "���s������py�t�@�C��[" & py_fullname & "]�����݂��܂���B"
    End If
    
    file_check = err_msg
End Function

' �R�}���h���A���p�b�T�[�Ɣ��p�X�y�[�X�Ōq���֐��i�z���n����cmd��Ԃ��j
Private Function create_cmd(cmd_list As Variant) As String
    Dim cmd As String
    Dim i As Integer
    For i = 0 To UBound(cmd_list)
        If (i > 0) Then
            cmd = cmd & " & "
        End If
        cmd = cmd & cmd_list(i)
    Next i
    ' �R�}���h�v�����v�g�ɓn����R�}���h��1�s�����Ȃ��̂ł��̂悤�ɃR�}���h�𐮗����Ă���
    create_cmd = "cmd.exe /c " & cmd
End Function

' �����������C���֐��̎x���֐���������������������������������������������������������������������������������������������������������
' �����������C���֐���������������������������������������������������������������������������������������������������������������������������

' �R�}���h�v�����v�g�o�R��Python���s
Sub run_python(py_name As String)
    Dim err_msg As String: err_msg = "�\�����Ȃ��G���[�ł��B"
    On Error GoTo Catch
        ' ����������
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        
        Dim wsh As Object: Set wsh = CreateObject("WScript.Shell")
        
        ' ���s������py�t�@�C�������݂���v���W�F�N�g�f�B���N�g��
        Dim py_pj As String: py_pj = get_ini_value("python_path", "py_pj")

        ' ���s���������z��
        Dim py_venv As String
        py_venv = py_pj & "\" & get_ini_value("python_path", "venv") & "\" & get_ini_value("python_path", "venv_batch")
        
        ' ���s������py���W���[��(Python�T���p�X��ݒ肷��py���W���[��)
        Dim create_pj_path As String: create_pj_path = py_pj & "\create_pj_path.py"
        
        ' ���s������py���W���[��(�t���p�X)
        Dim py_fullname As String: py_fullname = py_pj & "\" & py_name
        
        ' �G���[�`�F�b�N�i�t�@�C���̑��݊m�F�j
        Dim chk_result As String: chk_result = file_check(py_venv, create_pj_path, py_fullname)
        If chk_result <> "" Then
            err_msg = chk_result
            GoTo Catch
        End If
        
        ' �R�}���h���쐬�y�ю��s1
        Dim cmd_list(1)
        cmd_list(0) = py_venv
        cmd_list(1) = "python " & create_pj_path
        ' create_cmd�֐� : �z��Ɋi�[�����R�}���h���������C���ɐ��`����֐�
        Call wsh.Run(create_cmd(cmd_list), vbHide, True)
        
        ' �R�}���h���쐬�y�ю��s2
        cmd_list(1) = "python " & py_fullname
        Call wsh.Run(create_cmd(cmd_list), vbHide, True)
        
        ' ���������
        Set wsh = Nothing
        
        ' ����I��
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        
        Exit Sub
Catch:
        ' ��O����
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        
        MsgBox err_msg
End Sub
