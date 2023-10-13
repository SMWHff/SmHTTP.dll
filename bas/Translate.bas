Attribute VB_Name = "Translate"
Option Explicit
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
' 注册消息
Public Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long

'用于获取自windows启动以来经历的时间长度（毫秒）
Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Declare Function IsFileExist Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Boolean  '判断文件是否存在

Public Const vbIDE As Long = 0
Public Const vbEXE As Long = 1

Public G_Init As Boolean
Public G_AutoParam As Boolean



'下面这个函数会自动建立一个外部ini文件，您可以修改这个ini文件，以修改描述信息的语言种类（从中文到英文，等等...）
'如果没有特殊情况，请不要动这个函数
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
    
    StringToChange = Replace(StringToChange, "=", "〓")
    
    Call GetPrivateProfileString("language", StringToChange, "blank_value", StringChanged, 1024, IniPathName)
    If Not StrComp(StringChanged, "blank_value") Then
        Call WritePrivateProfileString("language", StringToChange, StringToChange, IniPathName)
    Else
        StringToChange = StringChanged
    End If
    
    StringToChange = Replace(StringToChange, "〓", "=")
    StringToChange = Replace(StringToChange, """", "")
    Translate_Description = StringToChange
End Function



' 抛出异常
Public Function Throw(ByVal Source As String, ByVal Description As String, Optional ByVal HelpFile As String, Optional ByVal HelpContext As Long)
    Dim lMsg As Long
    
    If App.LogMode = 0 Then
        Debug.Print Description
    Else
        ' 注册消息。
        lMsg = RegisterWindowMessage(Description)
        ' 抛出异常
        Err.Raise vbObjectError + lMsg, App.EXEName & "." & Source, Description, HelpFile, HelpContext
    End If
End Function
