:'↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓【重要！下面代码请勿修改】↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓
:On Error Resume Next
:Sub bat
    echo off & cls
    echo '>nul&Title 一键打包发行版
    echo '>nul&set cDir=%~dp0
    echo '>nul&set cDir=%cDir:~,-1%
    echo '>nul&for /f "delims=" %%i in ("%cDir%") do set ProjectName=%%~ni
    echo '>nul&set ProjectName=【插件】神梦HTTP请求
    echo '>nul&set SysDir=%SystemRoot%\System32
    echo '>nul&set SysBit=%PROCESSOR_ARCHITECTURE:~-2%
    echo '>nul&if %SysBit%==64 (set SysDir=%SystemRoot%\SysWOW64)
    echo '>nul&DEL /F /A /Q "%cDir%\Documents\SmHTTP_chm\神梦HTTP请求插件文档.chm"
    echo '>nul&"%cDir%\Documents\SmHTTP_chm\hhc.exe" "%cDir%\Documents\SmHTTP_chm\SmHTTP.hhp"
    echo '>nul&DEL /F /Q /S "%cDir%\Examples\SmHTTP.dll"
    echo '>nul&DEL /F /Q /S "%cDir%\Releases\*.zip"
    echo '>nul&xcopy "%cDir%\Documents\SmHTTP.html"         "%cDir%\Releases\%ProjectName%\"      /s /c /d /y
    echo '>nul&xcopy "%cDir%\Documents\SmHTTP_chm\*.chm"    "%cDir%\Releases\%ProjectName%\"      /s /c /d /y
    echo '>nul&xcopy "%cDir%\Examples\"                     "%cDir%\Releases\%ProjectName%\"      /s /c /d /y
    echo '>nul& copy "%cDir%\Releases\SmHTTP.dll"           "%cDir%\Releases\%ProjectName%\SmHTTP.dll"
    echo '>nul&%SysDir%\CScript.exe //nologo //E:vbscript "%~f0" "%ProjectName%" %*
    echo '>nul&move "%cDir%\%ProjectName%.zip" "%cDir%\Releases\"
    echo '>nul&explorer "%cDir%\Releases\"
    echo '>nul&echo 脚本已经停止运行
    echo '>nul&ping -n 5 127.0.0.1>nul
    Exit Sub
End Sub

REM 下面是VBS代码
Set fso = CreateObject("Scripting.FileSystemObject")
cd = fso.GetFile(wsh.ScriptFullName).ParentFolder.Path
ProjectName = WScript.Arguments(0)
sDir = cd & "\Releases\" & ProjectName
delDir = cd & "\Releases\" & ProjectName
ZipFile = cd & "\" & ProjectName &".zip"
fso.DeleteFolder ZipFile,True
Call Zip(sDir, ZipFile)
fso.DeleteFolder delDir,True
Set fso = Nothing
wsh.echo ""
wsh.echo "发布成功！"
wsh.echo ""


Sub Zip(ByVal mySourceDir, ByVal myZipFile)
    Dim fso,f,objShell,objTarget
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.GetExtensionName(myZipFile) <> "zip" Then
        Exit Sub
    ElseIf fso.FolderExists(mySourceDir) Then
        FType = "Folder"
    ElseIf fso.FileExists(mySourceDir) Then
        FType = "File"
        FileName = fso.GetFileName(mySourceDir)
        FolderPath = Left(mySourceDir, Len(mySourceDir) - Len(FileName))
    Else
        Exit Sub
    End If
    Set f = fso.CreateTextFile(myZipFile, True)
        f.Write "PK" & Chr(5) & Chr(6) & String(18, Chr(0))
        f.Close
    Set objShell = CreateObject("Shell.Application")
    Select Case Ftype
        Case "Folder"
            Set objSource = objShell.NameSpace(mySourceDir)
            Set objFolderItem = objSource.Items()
        Case "File"
            Set objSource = objShell.NameSpace(FolderPath)
            Set objFolderItem = objSource.ParseName(FileName)
    End Select
    Set objTarget = objShell.NameSpace(myZipFile)
    intOptions = 256
    objTarget.CopyHere objFolderItem, intOptions
    Do
        WScript.Sleep 1000
    Loop Until objTarget.Items.Count > 0
End Sub

Sub UnZip(ByVal myZipFile, ByVal myTargetDir)
    Dim fso,objShell,objSource,objTarget
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If NOT fso.FileExists(myZipFile) Then
        Exit Sub
    ElseIf fso.GetExtensionName(myZipFile) <> "zip" Then
        Exit Sub
    ElseIf NOT fso.FolderExists(myTargetDir) Then
        fso.CreateFolder(myTargetDir)
    End If
    Set objShell = CreateObject("Shell.Application")
    Set objSource = objShell.NameSpace(myZipFile)
    Set objFolderItem = objSource.Items()
    Set objTarget = objShell.NameSpace(myTargetDir)
    intOptions = 256
    objTarget.CopyHere objFolderItem, intOptions
End Sub