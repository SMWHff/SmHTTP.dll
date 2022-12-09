WScript.Echo "�X�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�["
WScript.Echo "�U              ���ζ��Բ��v1.0 �汾                 �U"
WScript.Echo "�d�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�g"
WScript.Echo "�U    ����ʱ�䣺2022-04-01                            �U"
WScript.Echo "�U    ������̳��http://SMWHff.com                     �U"
WScript.Echo "�U    ģ�����ߣ������޺�                              �U"
WScript.Echo "�U    ���ߣѣѣ�1042207232                            �U"
WScript.Echo "�U    ������Ⱥ��624655641                             �U"
WScript.Echo "�d�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�g"
WScript.Echo "�U ģ��˵��������������֤ʵ�ʽ���Ƿ����Ԥ��         �U"
WScript.Echo "�U ���ι��ߣ�http://pan.baidu.com/s/1dESHf8X          �U"
WScript.Echo "�^�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�a"


'====Ԥ������������������====
if WScript.Arguments.Count > 1 Then
    If UCase(WScript.Arguments(0)) = "-P" Then
        Set Paraphrase = New VbsParaphrase
        WScript.echo Paraphrase.ParaphraseFile(WScript.Arguments(1))
        WScript.Quit
    End If
End If


'====ע������ϵͳ====
DLLName = "SmAssert"   '���Ӣ����
DLLPath = GetScriptDir() + "\"& DLLName &".dll"
If IsFileExist(DLLPath) = False Then DLLPath = Mid(DLLPath, 1, InStrRev(DLLPath, "\TestCase")-1) + "\"& DLLName &".dll"
If IsFileExist(DLLPath) Then
    SysDir = Left(WScript.FullName, InStrRev(WScript.FullName, "\")-1)
    CreateObject("WScript.Shell").Run SysDir & "\regsvr32.exe """ + DLLPath + """ /s", 0, True
Else
    MsgBox "������ǰ����Ŀ¼��δ�ҵ� "& DLLName &".dll �����", 16 + 4096, GetScriptName()
    WScript.Quit
End If


'====������������====
Set Plugin = New VbsQMPlugin
Set SmAssert = Plugin.SmAssert

Class VbsQMPlugin
    Private QM_SmAssert
    Private Sub class_Initialize()
        Set QM_SmAssert = New VbsSmAssert
    End Sub
    Public Property Get SmAssert
        Set SmAssert = QM_SmAssert()
    End Property
End Class
Class VbsSmAssert
    Private Var_SmAssert
    
    Private Sub class_Initialize()
        Set Var_SmAssert = Nothing
        Set Var_SmAssert = CreateObject("SMWH.SmAssert")
        If Var_SmAssert Is Nothing Then
            MsgBox "��ʼ��ʧ�ܣ����Ƚ� SmAssert.dll ���ע�ᵽϵͳ��", 16 + 4096, "SmAssert.vbs"
            WScript.Quit
        End If
    End Sub
    Private Sub Class_Terminate
        Set Var_SmAssert = Nothing
    End Sub

    '��ȡ�������
    Public Default Function GetSmAssert()
        Set GetSmAssert = Var_SmAssert
    End Function
End Class

'�������е�����������Ԥ����
Class VbsParaphrase
    Private quoted, comments, specialdim, code
    
    '��ʽ���ļ�
    Public Function ParaphraseFile(ByVal Path)
        Dim fso, GetDir, Name, tPath
        Set fso = CreateObject("Scripting.FileSystemObject")
        Call Paraphrase(fso.OpenTextFile(Path).ReadAll)
        '���洦���Ĵ���
        GetDir = fso.GetFile(Path).ParentFolder.Path
        Name = fso.GetFileName(fso.GetFile(Path))
        tPath = GetDir & "\" & Name & ".tmp"
        If fso.fileExists(tPath) Then fso.DeleteFile(tPath)
        fso.OpenTextFile(tPath, 2, True).Write(code)
        fso.GetFile(tPath).Attributes=2 '����
        ParaphraseFile = tPath
    End Function
    
    '����������Ԥ����
    Public Function Paraphrase(ByVal input)
        code = input
        Call GetQuoted()
        Call GetComments()
        Call GetSpecialDim()
             ReplaceZHWord()
        Call PutSpecialDim()
        Call PutComments()
        Call PutQuoted()
        Paraphrase = code
    End Function
    
    '���ַ����滻�� %[ quoted ]%
    Private Sub GetQuoted()
        Dim re
        Set re = New RegExp
        re.Global = True
        re.Pattern = """.*?"""
        Set quoted = re.Execute(code)
        code = re.Replace(code, "%["&"quoted"&"]%")
    End Sub

    '�� %[ quoted ]% �滻���ַ���
    Private Sub PutQuoted()
        Dim i
        For Each i In quoted
            code = Replace(code, "%["&"quoted"&"]%", i, 1, 1)
        Next
    End Sub

    '��ע���滻�� %[ comment ]%
    Private Sub GetComments()
        Dim re
        Set re = New RegExp
        re.Global = True
        re.Pattern = "'.*"
        Set comments = re.Execute(code)
        code = re.Replace(code, "%["&"comment"&"]%")
    End Sub

    '�� %[ comment ]% �滻��ע��
    Private Sub PutComments()
        Dim i
        For Each i In comments
            code = Replace(code, "%["&"comment"&"]%", i, 1, 1)
        Next
    End Sub
    
    '������������滻�� %[ specialdim ]%
    Private Sub GetSpecialDim()
        Dim re
        Set re = New RegExp
        re.Global = True
        re.Pattern = "\[.+?\]"
        Set specialdim = re.Execute(code)
        code = re.Replace(code, "%["&"specialdim"&"]%")
    End Sub

    '�� %[ specialdim ]% �滻�����������
    Private Sub PutSpecialDim()
        Dim i
        For Each i In specialdim
            code = Replace(code, "%["&"specialdim"&"]%", i, 1, 1)
        Next
    End Sub
    
    '�����ı�������������������������
    Private Sub ReplaceZHWord()
        Dim re
        Set re = New RegExp
        re.Global = True
        re.IgnoreCase = True
        re.MultiLine = True

        re.Pattern = "([^\s\.\(\)\=\,\b]*[\u4e00-\u9fa5]+[^\s\.\(\)\=\,\r\n\b]*)"
        code = re.Replace(code, "[$1]")
    End Sub
End Class





'================================�����ݺ�����================================
' �ӳ�
Public Function Delay(ms)
    WScript.Sleep ms
End Function


' �������
Public Function TracePrint(Text)
    WScript.Echo Text
End Function


' �˳���ǰ�ű�������
Public Function ExitScript()
    WScript.Quit
End Function


'��������ͼ��
Public Function CreaTray()
    If TypeName(SmAssert) = "SmAssert" Then CreaTray = SmAssert.Get_Plugin_Interpret_Template("CreaTray", -1, "C:\Windows\system32\cmd.exe")
End Function


'������������
Public Function Tips(Text)
    [����״̬] = Text
    If TypeName(SmAssert) = "SmAssert" Then Tips = SmAssert.Get_Plugin_Interpret_Template("TipsTray", Text)
End Function


'��������ͼ��
Public Function UnTray()
    If TypeName(SmAssert) = "SmAssert" Then UnTray = SmAssert.Get_Plugin_Interpret_Template("UnTray")
End Function


'���г���
Function RunApp(Path)
    dim ws
    Set ws = CreateObject("WScript.Shell")
    ws.Run Path
    Set ws = Nothing
End Function

'�ж��Ƿ�64λ����ϵͳ
Function Is64Bit()
    Dim strComputer, objWMIService, colItems, objItem, strSystemType
    strComputer = "."
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
    Set colItems = objWMIService.ExecQuery("Select * from Win32_ComputerSystem",,48)
    For Each objItem in colItems
        strSystemType = objItem.SystemType
    Next
    Is64Bit = (InStr(strSystemType, "x64") > 0)
End Function

'��ȡ��ǰ�ű�����Ŀ¼
Function GetScriptDir()
    GetScriptDir = Left(Wscript.ScriptFullName, InStrRev(Wscript.ScriptFullName, "\")-1)
End Function

'��ȡ��ǰ�ű��ļ���
Function GetScriptName()
    GetScriptName = Mid(Wscript.ScriptFullName, InStrRev(Wscript.ScriptFullName, "\")+1)
End Function

'�ж��ļ��Ƿ����
Function IsFileExist(Path)
    IsFileExist = Createobject("Scripting.FileSystemObject").fileExists(Path)
End Function

' ����
Sub assert(Expression, FailMessage)
    If Expression Then
    Else
        WScript.Echo "����ʧ�ܣ�" & FailMessage
        WScript.Quit
    End If
End Sub

'�Զ��л�Ϊ32λVBS������
Function Run32()
    If Is64Bit() = False Then
        Set WshShell = CreateObject("WScript.Shell")
        WshPath = WScript.FullName
        If InStr(1, WshPath, "system32", 1) > 0 Then
            WshPath = Replace(WshPath, "system32", "SysWOW64", 1, 1, 1)
            WshShell.Run WshPath & " " & """" & WScript.ScriptFullName & """", 10, False
            ExitScript
        End If
    End If
End Function