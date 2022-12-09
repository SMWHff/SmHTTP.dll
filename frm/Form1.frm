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
      Caption         =   "ʧ�ܣ�Ĭ�ϣ�"
      Height          =   735
      Left            =   1200
      TabIndex        =   2
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ʧ�ܣ��Զ��壩"
      Height          =   615
      Left            =   1200
      TabIndex        =   1
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�ɹ�"
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
    ' �ɹ�����
    'Set Assert = CreateObject("SMWH.Assert")
    Assert.Fail
    Assert.IsTrue True
    Assert.IsFalse False
    Assert.IsEquals 1, 1
    Assert.IsNotEquals 1, 2
    Assert.IsContains "���β��", "���οƼ�|�����޺�|���β��"
    Assert.IsNotContains "SMWH", "���οƼ�|�����޺�|���β��"
    Assert.IsMatches "QQ:\d+", "QQ:1042207232"
    Assert.IsNotMatches "QQ:\d+", "���ߣ������޺�"
    Assert.That Array(9.37, 7.6), "=", Array(9.37, 7.6)
    Assert.That Null, "=", Null
    Assert.That Empty, "=", Empty
    Assert.That 1024, "=", 1024
    Assert.That 1024, ">", 1000
    Assert.That 1024, "<", 2048
    Assert.That "SMWHff", ">=", "SMWH"
    Assert.That "����", "<=", "�����޺�"
    Assert.That 0.1 + 0.2, "~=", 0.3
    Assert.That 1 + 1, "<>", 3
    Assert.That 1 + 1, "!=", 4
    Assert.That "��ʹ", "in", "ÿ�������ж�ס��[��ʹ]"
    Assert.That "ħ��", "not in", "ÿ�������ж�ס��[��ʹ]"
    Assert.That "����", "in", Array("����", "����", "��ŭ", "����", "̰��", "��ʳ", "ɫ��")
    Assert.That Array("��ǿ", "��г", "����", "��ҵ", "����"), "in", Array("��ǿ", "����", "����", "��г", "����", "ƽ��", "����", "����", "����", "��ҵ", "����", "����")
    Assert.That Assert, "is", Assert
    Assert.That Assert, "not is", Nothing
    
    MsgBox "�ɹ���", 4096
End Sub



Private Sub Command2_Click()
    'Assert.Fail "����ʧ��"
    Assert.IsTrue False, "����ʧ�ܣ����ʽ��Ϊ��"
    Assert.IsFalse True, "����ʧ�ܣ����ʽ��Ϊ��"
    Assert.IsEquals 1, 2, "����ʧ�ܣ�1��2"
    Assert.IsNotEquals 1, 1, "����ʧ�ܣ�1=1"
    Assert.IsContains "������̳", "���οƼ�|�����޺�|���β��", "����ʧ�ܣ�δ�ҵ���������̳��"
    Assert.IsNotContains "���β��", "���οƼ�|�����޺�|���β��", "����ʧ�ܣ��Ѵ��ڡ����β����"
    Assert.IsMatches "WX:\d+", "QQ:1042207232", "����ʧ�ܣ�δƥ�䡰WX:\d+��"
    Assert.IsNotMatches "QQ:\d+", "QQ:1042207232", "����ʧ�ܣ�����ƥ�䡰QQ:\d+��"
    Assert.That Array(0), "=", Array(1), "����ʧ�ܣ�Array(0) = Array(1)"
    Assert.That Null, "=", 1, "����ʧ�ܣ�Null = 1"
    Assert.That Empty, "=", 2, "����ʧ�ܣ�Empty = 2"
    Assert.That True, "=", False, "����ʧ�ܣ�True = False"
    Assert.That 1024, ">", 10000, "����ʧ�ܣ�1024 > 10000"
    Assert.That 1024, "<", 0.2048, "����ʧ�ܣ�1024 < 0.2048"
    Assert.That "SMWH", ">=", "SMWHff", "����ʧ�ܣ���SMWH�� >= ��SMWHff��"
    Assert.That "�����޺�", "<=", "����", "����ʧ�ܣ��������޺ۡ� <= �����Ρ�"
    Assert.That 0.1 + 0.2, "~=", 3, "����ʧ�ܣ�0.3 ~= 3"
    Assert.That 1 + 1, "<>", 2, "����ʧ�ܣ�2 <> 2"
    Assert.That 1 + 1, "!=", 2, "����ʧ�ܣ�2 != 2"
    Assert.That "��ʹ", "in", "ÿ�������ж�ס��[��ħ]", "����ʧ�ܣ�����ʹ�� In ��ÿ�������ж�ס��[��ħ]��"
    Assert.That "ħ��", "not in", "ÿ�������ж�ס��[ħ��]", "����ʧ�ܣ���ħ�� Not In ��ÿ�������ж�ס��[ħ��]��"
    Assert.That "��˽", "in", Array("����", "����", "��ŭ", "����", "̰��", "��ʳ", "ɫ��"), "����ʧ�ܣ�����˽�� In Array(6)"
    Assert.That Array("��ǿ", "��г", "����", "��ҵ", "����", "�Ž�"), "in", Array("��ǿ", "����", "����", "��г", "����", "ƽ��", "����", "����", "����", "��ҵ", "����", "����"), "����ʧ�ܣ�Array(5) In Array(11)"
    Assert.That Assert, "is", Nothing, "����ʧ�ܣ�Assert Is Nothing"
    Assert.That Assert, "not is", Assert, "����ʧ�ܣ�Assert Not Is Assert"
End Sub


Private Sub Command3_Click()
    ' �ɹ�����
    'Set Assert = CreateObject("SMWH.Assert")
    Set obj = New SmAssert
    Assert.Fail
    Assert.IsTrue False
    Assert.IsFalse True
    Assert.IsEquals 1, 2
    Assert.IsNotEquals 1, 1
    Assert.IsBetween 100, 10, 500
    Assert.IsNotBetween 10, 100, 50
    Assert.IsContains "SMWH", "���οƼ�|�����޺�|���β��"
    Assert.IsNotContains "���β��", "���οƼ�|�����޺�|���β��"
    Assert.IsMatches "QQ:\d+", "���ߣ������޺�"
    Assert.IsNotMatches "QQ:\d+", "QQ:1042207232"
    Assert.That Array(0), "=", Array(1)
    Assert.That Null, "=", "Null"
    Assert.That Empty, "=", "Empty"
    Assert.That 1024, "=", 10240
    Assert.That 1024, ">", 10000
    Assert.That 1024, "<", 0.2048
    Assert.That "SMWH", ">=", "SMWHff"
    Assert.That "�����޺�", "<=", "����"
    Assert.That 0.1 + 0.2, "~=", 3
    Assert.That 1 + 1, "<>", 2
    Assert.That 1 + 1, "!=", 2
    Assert.That "��ʹ", "in", "ÿ�������ж�ס��[��ħ]"
    Assert.That "ħ��", "not in", "ÿ�������ж�ס��[ħ��]"
    Assert.That "��˽", "in", Array("����", "����", "��ŭ", "����", "̰��", "��ʳ", "ɫ��")
    Assert.That Array("��ǿ", "��г", "����", "��ҵ", "����", "�Ž�"), "in", Array("��ǿ", "����", "����", "��г", "����", "ƽ��", "����", "����", "����", "��ҵ", "����", "����")
    Assert.That Assert, "is", obj
    Assert.That Assert, "not is", Assert
End Sub
