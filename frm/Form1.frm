VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   7455
   ClientTop       =   5115
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   Begin VB.CommandButton Command3 
      Caption         =   "失败（默认）"
      Height          =   735
      Left            =   1200
      TabIndex        =   2
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "失败（自定义）"
      Height          =   615
      Left            =   1200
      TabIndex        =   1
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "成功"
      Height          =   615
      Left            =   1200
      TabIndex        =   0
      Top             =   480
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Assert As SmAssert

Private Sub Form_Load()
    Set Assert = New SmAssert
End Sub


Private Sub Command1_Click()
    ' 成功例子
    'Set Assert = CreateObject("SMWH.Assert")
    Assert.Fail
    Assert.IsTrue True
    Assert.IsFalse False
    Assert.IsEquals 1, 1
    Assert.IsNotEquals 1, 2
    Assert.IsContains "神梦插件", "神梦科技|神梦无痕|神梦插件"
    Assert.IsNotContains "SMWH", "神梦科技|神梦无痕|神梦插件"
    Assert.IsMatches "QQ:\d+", "QQ:1042207232"
    Assert.IsNotMatches "QQ:\d+", "作者：神梦无痕"
    Assert.That Array(9.37, 7.6), "=", Array(9.37, 7.6)
    Assert.That Null, "=", Null
    Assert.That Empty, "=", Empty
    Assert.That 1024, "=", 1024
    Assert.That 1024, ">", 1000
    Assert.That 1024, "<", 2048
    Assert.That "SMWHff", ">=", "SMWH"
    Assert.That "神梦", "<=", "神梦无痕"
    Assert.That 0.1 + 0.2, "~=", 0.3
    Assert.That 1 + 1, "<>", 3
    Assert.That 1 + 1, "!=", 4
    Assert.That "天使", "in", "每个人心中都住着[天使]"
    Assert.That "魔鬼", "not in", "每个人心中都住着[天使]"
    Assert.That "傲慢", "in", Array("傲慢", "嫉妒", "暴怒", "懒惰", "贪婪", "暴食", "色欲")
    Assert.That Array("富强", "和谐", "爱国", "敬业", "友善"), "in", Array("富强", "民主", "文明", "和谐", "自由", "平等", "公正", "法制", "爱国", "敬业", "诚信", "友善")
    Assert.That Assert, "is", Assert
    Assert.That Assert, "not is", Nothing
    
    MsgBox "成功！", 4096
End Sub



Private Sub Command2_Click()
    'Assert.Fail "断言失败"
    Assert.IsTrue False, "断言失败，表达式不为真"
    Assert.IsFalse True, "断言失败，表达式不为假"
    Assert.IsEquals 1, 2, "断言失败，1≠2"
    Assert.IsNotEquals 1, 1, "断言失败，1=1"
    Assert.IsContains "神梦论坛", "神梦科技|神梦无痕|神梦插件", "断言失败，未找到“神梦论坛”"
    Assert.IsNotContains "神梦插件", "神梦科技|神梦无痕|神梦插件", "断言失败，已存在“神梦插件”"
    Assert.IsMatches "WX:\d+", "QQ:1042207232", "断言失败，未匹配“WX:\d+”"
    Assert.IsNotMatches "QQ:\d+", "QQ:1042207232", "断言失败，存在匹配“QQ:\d+”"
    Assert.That Array(0), "=", Array(1), "断言失败，Array(0) = Array(1)"
    Assert.That Null, "=", 1, "断言失败，Null = 1"
    Assert.That Empty, "=", 2, "断言失败，Empty = 2"
    Assert.That True, "=", False, "断言失败，True = False"
    Assert.That 1024, ">", 10000, "断言失败，1024 > 10000"
    Assert.That 1024, "<", 0.2048, "断言失败，1024 < 0.2048"
    Assert.That "SMWH", ">=", "SMWHff", "断言失败，“SMWH” >= “SMWHff”"
    Assert.That "神梦无痕", "<=", "神梦", "断言失败，“神梦无痕” <= “神梦”"
    Assert.That 0.1 + 0.2, "~=", 3, "断言失败，0.3 ~= 3"
    Assert.That 1 + 1, "<>", 2, "断言失败，2 <> 2"
    Assert.That 1 + 1, "!=", 2, "断言失败，2 != 2"
    Assert.That "天使", "in", "每个人心中都住着[恶魔]", "断言失败，“天使” In “每个人心中都住着[恶魔]”"
    Assert.That "魔鬼", "not in", "每个人心中都住着[魔鬼]", "断言失败，“魔鬼” Not In “每个人心中都住着[魔鬼]”"
    Assert.That "自私", "in", Array("傲慢", "嫉妒", "暴怒", "懒惰", "贪婪", "暴食", "色欲"), "断言失败，“自私” In Array(6)"
    Assert.That Array("富强", "和谐", "爱国", "敬业", "友善", "团结"), "in", Array("富强", "民主", "文明", "和谐", "自由", "平等", "公正", "法制", "爱国", "敬业", "诚信", "友善"), "断言失败，Array(5) In Array(11)"
    Assert.That Assert, "is", Nothing, "断言失败，Assert Is Nothing"
    Assert.That Assert, "not is", Assert, "断言失败，Assert Not Is Assert"
End Sub


Private Sub Command3_Click()
    ' 成功例子
    'Set Assert = CreateObject("SMWH.Assert")
    Set obj = New SmAssert
    Assert.Fail
    Assert.IsTrue False
    Assert.IsFalse True
    Assert.IsEquals 1, 2
    Assert.IsNotEquals 1, 1
    Assert.IsBetween 100, 10, 500
    Assert.IsNotBetween 10, 100, 50
    Assert.IsContains "SMWH", "神梦科技|神梦无痕|神梦插件"
    Assert.IsNotContains "神梦插件", "神梦科技|神梦无痕|神梦插件"
    Assert.IsMatches "QQ:\d+", "作者：神梦无痕"
    Assert.IsNotMatches "QQ:\d+", "QQ:1042207232"
    Assert.That Array(0), "=", Array(1)
    Assert.That Null, "=", "Null"
    Assert.That Empty, "=", "Empty"
    Assert.That 1024, "=", 10240
    Assert.That 1024, ">", 10000
    Assert.That 1024, "<", 0.2048
    Assert.That "SMWH", ">=", "SMWHff"
    Assert.That "神梦无痕", "<=", "神梦"
    Assert.That 0.1 + 0.2, "~=", 3
    Assert.That 1 + 1, "<>", 2
    Assert.That 1 + 1, "!=", 2
    Assert.That "天使", "in", "每个人心中都住着[恶魔]"
    Assert.That "魔鬼", "not in", "每个人心中都住着[魔鬼]"
    Assert.That "自私", "in", Array("傲慢", "嫉妒", "暴怒", "懒惰", "贪婪", "暴食", "色欲")
    Assert.That Array("富强", "和谐", "爱国", "敬业", "友善", "团结"), "in", Array("富强", "民主", "文明", "和谐", "自由", "平等", "公正", "法制", "爱国", "敬业", "诚信", "友善")
    Assert.That Assert, "is", obj
    Assert.That Assert, "not is", Assert
End Sub
