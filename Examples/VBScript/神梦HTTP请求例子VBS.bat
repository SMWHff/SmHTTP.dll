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
Exit Sub :End Sub:Sub Import(P):Dim o,f,s:Set o=CreateObject("Scripting.FileSystemObject"):Set f=o.OpenTextFile(P):s = f.ReadAll:f.Close:ExecuteGlobal s:End Sub:Set fso=CreateObject("Scripting.FileSystemObject"):If fso.fileExists(WScript.ScriptName) Then fso.DeleteFile(WScript.ScriptName)
'#================================================================
'#         神梦HTTP请求插件 SmHTTP.dll 演示 VBScript 例子
'#----------------------------------------------------------------
'#        【作者】：神梦无痕
'#        【ＱＱ】：1042207232
'#        【Ｑ群】：624655641
'#        【更新】：2023-11-03
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
Assert "1.0.0.4" = SmHTTP.Ver(), "出错，插件版本号不匹配！"

' 定义变量
Dim user, pass, Data, Ret, Cookies, Headers

' 配置账号
user = "你的按键精灵论坛账号"
pass = "你的按键精灵论坛密码"

If user = "你的按键精灵论坛账号" Then user = Environ("AJ_USER")
If pass = "你的按键精灵论坛密码" Then pass = Environ("AJ_PASS")

Data = SmHTTP.Data( _
    "username", user, _
    "password", pass, _
    "question", "0", _
    "answer", "", _
    "templateid", "0", _
    "login", "", _
    "expires", "43200" _
)


' 开启自动识别参数模式
Call SmHTTP.SetAutoParamArray(True)


' 登录论坛账号
Ret = SmHTTP.HTTP_POST("http://bbs.anjian.com/login.aspx?referer=forumindex.aspx", Data)
' 判断是否登录成功
If InStr(Ret, user) = 0 Then
    MsgBox "出错，登录失败！", 16 + 4096, "报错！"
    EndScript
End If
Cookies = SmHTTP.GetCookies()


' 打卡签到
Data = SmHTTP.Data( _
    "signmessage", "签个到，每天心情都是美美哒~~按键精灵祝大家新年好运连连！！" _
)
Headers = SmHTTP.Headers( _
    "Referer", "http://bbs.anjian.com/" _
)
Ret = SmHTTP.HTTP_POST("http://bbs.anjian.com/addsignin.aspx?infloat=1&inajax=1", Data, Headers, Cookies)
If InStr(Ret, "恭喜您获取本日签到奖励") Or InStr(Ret, "你今天已经签到过了") Then ' 判断是否签到成功
    TracePrint "恭喜，您已完成签到任务！"
Else
    TracePrint Ret
    MsgBox "出错，签到失败！", 16 + 4096, "报错！"
    EndScript
End If


MsgBox "脚本执行完毕！", 4096
