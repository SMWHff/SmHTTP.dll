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
MacroID=0893033f-17c8-4e4c-bf43-7882ca6c252c
Description=断言失败例子
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

' 导入神梦断言插件
Import "SmAssert.dll"


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
SmAssert.IsBetween 1, 100, 888
SmAssert.IsBetween 99, 0, 66
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