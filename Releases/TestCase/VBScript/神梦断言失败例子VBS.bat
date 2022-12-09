:'↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓【重要！下面代码请勿修改】↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓
:'On Error Resume Next
:Sub bat
echo off & cls
echo '>nul&set SysDir=%SystemRoot%\System32
echo '>nul&set SysBit=%PROCESSOR_ARCHITECTURE:~-2%
echo '>nul&if %SysBit%==64 (set SysDir=%SystemRoot%\SysWOW64)
echo '>nul&if exist %~f0.tmp (DEL /F /A /Q "%~f0.tmp")
echo '>nul&if exist SmAssert.vbs ( %SysDir%\CScript.exe //nologo //E:vbscript SmAssert.vbs -P "%~f0">nul )
echo '>nul&if exist SmAssert.vbs ( %SysDir%\CScript.exe //nologo //E:vbscript "%~f0.tmp" %* )
echo '>nul&if not exist SmAssert.vbs ( echo 出错，未找到 SmAssert.vbs 模块！ )
echo '>nul&echo 脚本已经停止运行 &pause>nul
Exit Sub :End Sub:Sub Import(P):Dim o,f,s:On Error Resume Next:Set o=CreateObject("Scripting.FileSystemObject"):Set f=o.OpenTextFile(P):s = f.ReadAll:f.Close:ExecuteGlobal s:End Sub:Set fso=CreateObject("Scripting.FileSystemObject"):If fso.fileExists(WScript.ScriptName) Then fso.DeleteFile(WScript.ScriptName)
'#================================================================
'#         神梦填表插件 SmAssert.dll 演示 VBScript 断言失败例子
'#----------------------------------------------------------------
'#        【作者】：神梦无痕
'#        【ＱＱ】：1042207232
'#        【Ｑ群】：624655641
'#        【更新】：2022-03-27
'#----------------------------------------------------------------
'#  插件说明：断言用于验证实际结果是否符合预期
'#----------------------------------------------------------------
'#  神梦工具：http://pan.baidu.com/s/1dESHf8X
'#================================================================
'↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑【重要！上面代码请勿修改】↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑


'导入【SmAssert.vbs】模块--------------------------（从这开始就是VBS代码了）
Import "SmAssert.vbs"


TracePrint("**********************【神梦断言插件 SmAssert.dll 演示 VBScript 断言失败例子】**********************")

'判断插件版本
SmAssert "1.1.0.0" = SmAssert.Ver(), "出错，插件版本号不匹配！"

' 断言失败例子
SmAssert.Fail
SmAssert.IsTrue False
SmAssert.IsFalse True
SmAssert.IsEquals 1, 2
SmAssert.IsNotEquals 1, 1
SmAssert.IsContains "SMWH", "神梦科技|神梦无痕|神梦插件"
SmAssert.IsNotContains "神梦插件", "神梦科技|神梦无痕|神梦插件"
SmAssert.IsMatches "QQ:\d+", "作者：神梦无痕"
SmAssert.IsNotMatches "QQ:\d+", "QQ:1042207232"
SmAssert.IsBetween 1, 100, 1024
SmAssert.IsNotBetween 1, 100, 99
SmAssert.That Array(3.14, "SMWH"), "=", Array("SMWH")
SmAssert.That Null, "=", "Null"
SmAssert.That Empty, "=", "Empty"
SmAssert.That 1024, "=", 10240
SmAssert.That 1024, ">", 10000
SmAssert.That 1024, "<", 0.2048
SmAssert.That "SMWH", ">=", "SMWHff"
SmAssert.That "神梦无痕", "<=", "神梦"
SmAssert.That 0.1 + 0.2, "~=", 3
SmAssert.That 1 + 1, "<>", 2
SmAssert.That 1 + 1, "!=", 2
SmAssert.That "天使", "in", "每个人心中都住着[恶魔]"
SmAssert.That "魔鬼", "not in", "每个人心中都住着[魔鬼]"
SmAssert.That "自私", "in", Array("傲慢", "嫉妒", "暴怒", "懒惰", "贪婪", "暴食", "色欲")
SmAssert.That "傲慢", "not in", Array("傲慢", "嫉妒", "暴怒", "懒惰", "贪婪", "暴食", "色欲")
SmAssert.That Array("富强", "和谐", "爱国", "敬业", "友善", "团结"), "in", Array("富强", "民主", "文明", "和谐", "自由", "平等", "公正", "法制", "爱国", "敬业", "诚信", "友善")
SmAssert.That Array("平等", "公正"), "not in", Array("富强", "民主", "文明", "和谐", "自由", "平等", "公正", "法制", "爱国", "敬业", "诚信", "友善")
SmAssert.That SmAssert, "is", Nothing
SmAssert.That SmAssert, "not is", SmAssert

MsgBox "脚本执行完毕！", 4096
