WScript.Echo "╔══════════════════════════╗"
WScript.Echo "║              神梦断言插件v1.0 版本                 ║"
WScript.Echo "╠══════════════════════════╣"
WScript.Echo "║    更新时间：2022-04-01                            ║"
WScript.Echo "║    神梦论坛：http://SMWHff.com                     ║"
WScript.Echo "║    模块作者：神梦无痕                              ║"
WScript.Echo "║    作者ＱＱ：1042207232                            ║"
WScript.Echo "║    交流①群：624655641                             ║"
WScript.Echo "╠══════════════════════════╣"
WScript.Echo "║ 模块说明：断言用于验证实际结果是否符合预期         ║"
WScript.Echo "║ 神梦工具：http://pan.baidu.com/s/1dESHf8X          ║"
WScript.Echo "╚══════════════════════════╝"


'====预处理代码里的中文命令====
if WScript.Arguments.Count > 1 Then
    If UCase(WScript.Arguments(0)) = "-P" Then
        Set Paraphrase = New VbsParaphrase
        WScript.echo Paraphrase.ParaphraseFile(WScript.Arguments(1))
        WScript.Quit
    End If
End If


'====注册插件到系统====
DLLName = "SmAssert"   '插件英文名
DLLPath = GetScriptDir() + "\"& DLLName &".dll"
If IsFileExist(DLLPath) = False Then DLLPath = Mid(DLLPath, 1, InStrRev(DLLPath, "\TestCase")-1) + "\"& DLLName &".dll"
If IsFileExist(DLLPath) Then
    SysDir = Left(WScript.FullName, InStrRev(WScript.FullName, "\")-1)
    CreateObject("WScript.Shell").Run SysDir & "\regsvr32.exe """ + DLLPath + """ /s", 0, True
Else
    MsgBox "出错，当前程序目录下未找到 "& DLLName &".dll 插件！", 16 + 4096, GetScriptName()
    WScript.Quit
End If


'====创建对象引用====
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
            MsgBox "初始化失败，请先将 SmAssert.dll 插件注册到系统！", 16 + 4096, "SmAssert.vbs"
            WScript.Quit
        End If
    End Sub
    Private Sub Class_Terminate
        Set Var_SmAssert = Nothing
    End Sub

    '获取插件对象
    Public Default Function GetSmAssert()
        Set GetSmAssert = Var_SmAssert
    End Function
End Class

'将代码中的中文名进行预处理
Class VbsParaphrase
    Private quoted, comments, specialdim, code
    
    '格式化文件
    Public Function ParaphraseFile(ByVal Path)
        Dim fso, GetDir, Name, tPath
        Set fso = CreateObject("Scripting.FileSystemObject")
        Call Paraphrase(fso.OpenTextFile(Path).ReadAll)
        '保存处理后的代码
        GetDir = fso.GetFile(Path).ParentFolder.Path
        Name = fso.GetFileName(fso.GetFile(Path))
        tPath = GetDir & "\" & Name & ".tmp"
        If fso.fileExists(tPath) Then fso.DeleteFile(tPath)
        fso.OpenTextFile(tPath, 2, True).Write(code)
        fso.GetFile(tPath).Attributes=2 '隐藏
        ParaphraseFile = tPath
    End Function
    
    '进行中文名预处理
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
    
    '将字符串替换成 %[ quoted ]%
    Private Sub GetQuoted()
        Dim re
        Set re = New RegExp
        re.Global = True
        re.Pattern = """.*?"""
        Set quoted = re.Execute(code)
        code = re.Replace(code, "%["&"quoted"&"]%")
    End Sub

    '将 %[ quoted ]% 替换回字符串
    Private Sub PutQuoted()
        Dim i
        For Each i In quoted
            code = Replace(code, "%["&"quoted"&"]%", i, 1, 1)
        Next
    End Sub

    '将注释替换成 %[ comment ]%
    Private Sub GetComments()
        Dim re
        Set re = New RegExp
        re.Global = True
        re.Pattern = "'.*"
        Set comments = re.Execute(code)
        code = re.Replace(code, "%["&"comment"&"]%")
    End Sub

    '将 %[ comment ]% 替换回注释
    Private Sub PutComments()
        Dim i
        For Each i In comments
            code = Replace(code, "%["&"comment"&"]%", i, 1, 1)
        Next
    End Sub
    
    '将特殊变量名替换成 %[ specialdim ]%
    Private Sub GetSpecialDim()
        Dim re
        Set re = New RegExp
        re.Global = True
        re.Pattern = "\[.+?\]"
        Set specialdim = re.Execute(code)
        code = re.Replace(code, "%["&"specialdim"&"]%")
    End Sub

    '将 %[ specialdim ]% 替换回特殊变量名
    Private Sub PutSpecialDim()
        Dim i
        For Each i In specialdim
            code = Replace(code, "%["&"specialdim"&"]%", i, 1, 1)
        Next
    End Sub
    
    '将中文变量、函数、参数加上中括号
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





'================================【兼容函数】================================
' 延迟
Public Function Delay(ms)
    WScript.Sleep ms
End Function


' 调试输出
Public Function TracePrint(Text)
    WScript.Echo Text
End Function


' 退出当前脚本的运行
Public Function ExitScript()
    WScript.Quit
End Function


'创建托盘图标
Public Function CreaTray()
    If TypeName(SmAssert) = "SmAssert" Then CreaTray = SmAssert.Get_Plugin_Interpret_Template("CreaTray", -1, "C:\Windows\system32\cmd.exe")
End Function


'设置托盘气泡
Public Function Tips(Text)
    [运行状态] = Text
    If TypeName(SmAssert) = "SmAssert" Then Tips = SmAssert.Get_Plugin_Interpret_Template("TipsTray", Text)
End Function


'销毁托盘图标
Public Function UnTray()
    If TypeName(SmAssert) = "SmAssert" Then UnTray = SmAssert.Get_Plugin_Interpret_Template("UnTray")
End Function


'运行程序
Function RunApp(Path)
    dim ws
    Set ws = CreateObject("WScript.Shell")
    ws.Run Path
    Set ws = Nothing
End Function

'判断是否64位操作系统
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

'获取当前脚本所在目录
Function GetScriptDir()
    GetScriptDir = Left(Wscript.ScriptFullName, InStrRev(Wscript.ScriptFullName, "\")-1)
End Function

'获取当前脚本文件名
Function GetScriptName()
    GetScriptName = Mid(Wscript.ScriptFullName, InStrRev(Wscript.ScriptFullName, "\")+1)
End Function

'判断文件是否存在
Function IsFileExist(Path)
    IsFileExist = Createobject("Scripting.FileSystemObject").fileExists(Path)
End Function

' 断言
Sub assert(Expression, FailMessage)
    If Expression Then
    Else
        WScript.Echo "断言失败，" & FailMessage
        WScript.Quit
    End If
End Sub

'自动切换为32位VBS解释器
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