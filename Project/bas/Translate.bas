Attribute VB_Name = "Translate"
Option Explicit
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
' ע����Ϣ
Public Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long

'���ڻ�ȡ��windows��������������ʱ�䳤�ȣ����룩
Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Declare Function IsFileExist Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Boolean  '�ж��ļ��Ƿ����

Public Const vbIDE As Long = 0
Public Const vbEXE As Long = 1

Public G_Init As Boolean
Public G_AutoParam As Boolean



'��������������Զ�����һ���ⲿini�ļ����������޸����ini�ļ������޸�������Ϣ���������ࣨ�����ĵ�Ӣ�ģ��ȵ�...��
'���û������������벻Ҫ���������
Public Function Translate_Description(Description As String) As String
    If Len(Description) = 0 Then
        Translate_Description = Description
        Return
    End If

    Dim IniPathName As String
    IniPathName = App.Path & "\" & App.EXEName & ".ini"
    
    Dim StringToChange, StringChanged As String
    StringToChange = """"
    StringToChange = StringToChange & Description
    StringToChange = StringToChange & """"
    
    StringChanged = String(1024, 0)
    
    StringToChange = Replace(StringToChange, "=", "��")
    
    Call GetPrivateProfileString("language", StringToChange, "blank_value", StringChanged, 1024, IniPathName)
    If Not StrComp(StringChanged, "blank_value") Then
        Call WritePrivateProfileString("language", StringToChange, StringToChange, IniPathName)
    Else
        StringToChange = StringChanged
    End If
    
    StringToChange = Replace(StringToChange, "��", "=")
    StringToChange = Replace(StringToChange, """", "")
    Translate_Description = StringToChange
End Function



' �׳��쳣
Public Function Throw(ByVal Source As String, ByVal Description As String, Optional ByVal HelpFile As String, Optional ByVal HelpContext As Long)
    Dim lMsg As Long
    
    If App.LogMode = 0 Then
        Debug.Print Description
    Else
        ' ע����Ϣ��
        lMsg = RegisterWindowMessage(Description)
        ' �׳��쳣
        Err.Raise vbObjectError + lMsg, App.EXEName & "." & Source, Description, HelpFile, HelpContext
    End If
End Function
