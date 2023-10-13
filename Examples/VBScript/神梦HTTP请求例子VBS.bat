:'↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓【重要！下面代码请勿修改】↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓
:'On Error Resume Next
:Sub bat
echo off & cls
echo '>nul&set SysDir=%SystemRoot%\System32
echo '>nul&set SysBit=%PROCESSOR_ARCHITECTURE:~-2%
echo '>nul&if %SysBit%==64 (set SysDir=%SystemRoot%\SysWOW64)
echo '>nul&if exist %~f0.tmp (DEL /F /A /Q "%~f0.tmp")
echo '>nul&if exist SmHTTP.vbs ( %SysDir%\CScript.exe //nologo //E:vbscript SmHTTP.vbs -P "%~f0">nul )
echo '>nul&if exist SmHTTP.vbs ( %SysDir%\CScript.exe //nologo //E:vbscript "%~f0.tmp" %* )
echo '>nul&if not exist SmHTTP.vbs ( echo 出错，未找到 SmHTTP.vbs 模块！ )
echo '>nul&echo 脚本已经停止运行 &pause>nul
Exit Sub :End Sub:Sub Import(P):Dim o,f,s:On Error Resume Next:Set o=CreateObject("Scripting.FileSystemObject"):Set f=o.OpenTextFile(P):s = f.ReadAll:f.Close:ExecuteGlobal s:End Sub:Set fso=CreateObject("Scripting.FileSystemObject"):If fso.fileExists(WScript.ScriptName) Then fso.DeleteFile(WScript.ScriptName)
'#================================================================
'#         神梦HTTP请求插件 SmHTTP.dll 演示 VBScript 例子
'#----------------------------------------------------------------
'#        【作者】：神梦无痕
'#        【ＱＱ】：1042207232
'#        【Ｑ群】：624655641
'#        【更新】：2022-03-27
'#----------------------------------------------------------------
'#  插件说明：用于HTTP协议的请求访问操作
'#----------------------------------------------------------------
'#  神梦工具：http://pan.baidu.com/s/1dESHf8X
'#================================================================
'↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑【重要！上面代码请勿修改】↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑


'导入【SmHTTP.vbs】模块--------------------------（从这开始就是VBS代码了）
Import "SmHTTP.vbs"


TracePrint("**********************【神梦HTTP请求插件 SmHTTP.dll 演示 VBScript 例子】**********************")

'判断插件版本
SmHTTP "1.0.0.0" = SmHTTP.Ver(), "出错，插件版本号不匹配！"




MsgBox "脚本执行完毕！", 4096
