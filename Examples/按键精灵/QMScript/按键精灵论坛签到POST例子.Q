[General]
SyntaxVersion=2
BeginHotkey=121
BeginHotkeyMod=0
PauseHotkey=0
PauseHotkeyMod=0
StopHotkey=123
StopHotkeyMod=0
RunOnce=1
EnableWindow=
MacroID=a1e646e4-30ea-47eb-8521-b836c22f40a9
Description=按键精灵论坛签到POST例子
Enable=1
AutoRun=0
[Repeat]
Type=0
Number=1
[SetupUI]
Type=2
QUI=
[Relative]
SetupOCXFile=
[Comment]

[Script]

' 导入神梦HTTP请求插件
Import "SmHTTP.dll"


' 判断插件版本号
Assert "1.0.0.3" = SmHTTP.Ver(), "出错，插件版本号不匹配！"


' 定义变量
Dim user, pass, Data, Ret, Cookies, Headers

' 配置账号
user = "你的按键精灵论坛账号"
pass = "你的按键精灵论坛密码"

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
End If



' 断言
Sub Assert(Expression, FailMessage)
    If Expression Then
    Else
        TracePrint "断言失败，" & FailMessage
        ExitScript
    End If
End Sub
