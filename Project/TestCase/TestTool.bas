Attribute VB_Name = "TestTool"
Option Explicit

Function GetFuncNameAll(ByVal cls_FilePath As String) As Variant
    Dim sFileContent    As String
    Dim FuncName        As Variant
    Dim i               As Long
    Dim ret             As Variant
    
    FuncName = Array()
    ' �ж��ļ��Ƿ����
    If Dir(cls_FilePath) = "" Then
        GetFuncNameAll = FuncName
    End If
    
    ' ��ȡ�ļ�����
    sFileContent = LoadFileContent(cls_FilePath)
    'Clipboard.SetText sFileContent
    'Debug.Print sFileContent
    
    If Len(sFileContent) > 0 Then
        ' ʹ��������ʽ�������еĺ�����
        Dim regex   As Object
        Dim matches As Object
        Set regex = CreateObject("VBScript.RegExp")
        regex.Global = True
        regex.IgnoreCase = True
        regex.MultiLine = True
        
        ' ȥ���ַ���
        regex.Pattern = """[^""]*"""
        sFileContent = regex.Replace(sFileContent, """{�ַ���}""")
        
        ' ȥ��ע��
        regex.Pattern = "'.*"
        sFileContent = regex.Replace(sFileContent, "'{ע��}")
        
        'Clipboard.SetText sFileContent
        'Debug.Print sFileContent
        regex.Pattern = "Public (Function|Sub) (test_\w+)\s*\(|\n (Function|Sub) (test_\w+)\s*\("
        Set matches = regex.Execute(sFileContent)
        ReDim FuncName(matches.Count - 1)
        For i = 0 To matches.Count - 1
            FuncName(i) = matches(i).SubMatches(1)
            'Debug.Print FuncName(i)
        Next
        Set matches = Nothing
        Set regex = Nothing
    End If
    GetFuncNameAll = FuncName
End Function


Function LoadFileContent(ByVal sFilePath As String) As String
    Dim BinFileNo As Long
    Dim FileLen As Long
    Dim FileByte() As Byte
    
    BinFileNo = FreeFile()      ' FreeFile ����һ�� Integer��������һ���ɹ��ļ���
    Open sFilePath For Binary As #BinFileNo
        FileLen = LOF(BinFileNo)
        ReDim FileByte(FileLen) As Byte
        Get #BinFileNo, , FileByte()
    Close #BinFileNo
    LoadFileContent = StrConv(FileByte(), vbUnicode)
End Function


Function Expression(ByVal actual As Variant, ByVal sq As String, ByVal expected As Variant) As Boolean
    Dim sc          As New ScriptControl
    Dim code        As String
    Dim sActual     As String
    Dim sExpected   As String
    
    On Error Resume Next
    sc.Language = "VBScript"
    sc.AddCode "dim x,y"
    
    ' ���� Null
    If IsNull(actual) Or IsNull(expected) Then
        sc.CodeObject.x = "<" & TypeName(actual) & ">"
        sc.CodeObject.y = "<" & TypeName(expected) & ">"
        sActual = sc.CodeObject.x
        sExpected = sc.CodeObject.y
        GoTo ���ʽ����
    End If
    
    ' ����ʵ�ʽ��
    If IsObject(actual) Then
        Set sc.CodeObject.x = actual
        sActual = "<" & TypeName(actual) & "@" & ObjPtr(actual) & ">"
    ElseIf IsArray(actual) Then
        sc.CodeObject.x = Join(actual, ",")
        sActual = "Array(" & sc.CodeObject.x & ")"
    Else
        sc.CodeObject.x = actual
        If VarType(actual) = vbString Then
            sActual = """" & actual & """"
        Else
            sActual = CStr(actual)
        End If
    End If
    
    ' ����Ԥ�ڽ��
    If IsObject(expected) Then
        Set sc.CodeObject.y = expected
        sExpected = "<" & TypeName(expected) & "@" & ObjPtr(expected) & ">"
    ElseIf IsArray(expected) Then
        sc.CodeObject.y = Join(expected, ",")
        sExpected = "Array(" & sc.CodeObject.y & ")"
    Else
        sc.CodeObject.y = expected
        If VarType(expected) = vbString Then
            sExpected = """" & expected & """"
        Else
            sExpected = CStr(expected)
        End If
    End If
    
���ʽ����:
    Select Case LCase(sq)
    Case "in"
        code = "InStr(1, y, x, vbTextCompare) > 0"
    Case "not in"
        code = "InStr(1, y, x, vbTextCompare) = 0"
    Case Else
        code = "x " & sq & " y"
    End Select
    Expression = sc.Eval(code)
    sc.Reset
    Set sc = Nothing
    
    If Expression = False Or Err Then
        Debug.Print , "���ʽ��������", sActual & vbTab & sq & vbTab & sExpected
    End If
End Function


Sub ClearImmediateWindow()
    Dim wsh As Object
    
    Set wsh = CreateObject("WScript.Shell")
    ' ���ý��㵽��������
    wsh.SendKeys "^g"
    ' ȫѡ
    wsh.SendKeys "^a"
    ' ɾ��
    wsh.SendKeys "{DEL}"
    Set wsh = Nothing
End Sub


'URL����
Public Function EscapeURL(ByVal url)
    Dim sc As ScriptControl  '��Ҫ���ù��� Microsoft Script Control

    Set sc = CreateObject("MSScriptControl.ScriptControl")
    sc.Language = "JScript"
    EscapeURL = sc.Run("encodeURIComponent", url)
    Set sc = Nothing
End Function


'�ַ�����
Public Function EscapeUnicode(ByVal text)
    Dim sc As ScriptControl  '��Ҫ���ù��� Microsoft Script Control

    Set sc = CreateObject("MSScriptControl.ScriptControl")
    sc.Language = "JScript"
    sc.AddCode "function encodeUnicode(text){var result=[];for(var i=0;i<text.length;i++){var c=text.charAt(i);if(c>'\u00FF'){result.push('\\u'+c.charCodeAt(0).toString(16).toUpperCase());}else{result.push(c);}}return result.join('');}"
    EscapeUnicode = sc.Run("encodeUnicode", text)
    Set sc = Nothing
End Function


'ȡ���м��ı�
Function GetMidS(ByVal str As String, ByVal StrHome As String, Optional ByVal StrEnd As String = vbNullString)
    Dim ret, arr1, arr2

    ret = ""
    arr1 = Split(str, StrHome, 2)
    If UBound(arr1) = 1 Then
        If Len(StrEnd) = 0 Then
            ret = arr1(1)
        Else
            arr2 = Split(arr1(1), StrEnd, 2)
            If UBound(arr2) = 1 Then
                ret = arr2(0)
            End If
        End If
    End If
    GetMidS = ret
End Function

' ����GUID
Function GetGUID()
    Dim objScriptContext As Object
    Dim Guid As String
    
    Set objScriptContext = CreateObject("Scriptlet.TypeLib")
    Guid = Mid(CreateObject("Scriptlet.TypeLib").Guid, 2, 36)
    GetGUID = Guid
End Function
