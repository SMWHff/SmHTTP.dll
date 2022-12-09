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
MacroID=cdb84bd5-30ea-47eb-8521-b836c22f40a8
Description=断言成功例子
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


' 判断插件版本号
SmAssert "1.1.0.0" = SmAssert.Ver(), "出错，插件版本号不匹配！"


' 断言成功例子
SmAssert.IsTrue True
SmAssert.IsFalse False
SmAssert.IsEquals 1, 1
SmAssert.IsNotEquals 1, 2
SmAssert.IsContains "神梦插件", "神梦科技|神梦无痕|神梦插件"
SmAssert.IsNotContains "SMWH", "神梦科技|神梦无痕|神梦插件"
SmAssert.IsMatches "QQ:\d+", "QQ:1042207232"
SmAssert.IsNotMatches "QQ:\d+", "作者：神梦无痕"
SmAssert.IsBetween 1, 100, 88
SmAssert.IsNotBetween 99, 0, 666
SmAssert.That Array(3.14, "SMWH"), "=", Array(3.14, "SMWH")
SmAssert.That Null, "=", Null
SmAssert.That Empty, "=", Empty
SmAssert.That 1024, "=", 1024
SmAssert.That 1024, ">", 1000
SmAssert.That 1024, "<", 2048
SmAssert.That "SMWHff", ">=", "SMWH"
SmAssert.That "神梦", "<=", "神梦无痕"
SmAssert.That 0.1 + 0.2, "~=", 0.3
SmAssert.That 1 + 1, "<>", 3
SmAssert.That 1 + 1, "!=", 4
SmAssert.That "天使", "In", "每个人心中都住着[天使]"
SmAssert.That "魔鬼", "Not In", "每个人心中都住着[天使]"
SmAssert.That "傲慢", "In", Array("傲慢", "嫉妒", "暴怒", "懒惰", "贪婪", "暴食", "色欲")
SmAssert.That "自私", "Not In", Array("傲慢", "嫉妒", "暴怒", "懒惰", "贪婪", "暴食", "色欲")
SmAssert.That Array("富强", "和谐", "爱国", "敬业", "友善"), "In", Array("富强", "民主", "文明", "和谐", "自由", "平等", "公正", "法制", "爱国", "敬业", "诚信", "友善")
SmAssert.That Array("可爱", "高富帅"), "Not In", Array("富强", "民主", "文明", "和谐", "自由", "平等", "公正", "法制", "爱国", "敬业", "诚信", "友善")
SmAssert.That SmAssert, "Is", SmAssert
SmAssert.That SmAssert, "Not Is", Nothing


