:'��������������������������������������Ҫ��������������޸ġ���������������������������������������������
:'On Error Resume Next
:Sub bat
echo off & cls
echo '>nul&set SysDir=%SystemRoot%\System32
echo '>nul&set SysBit=%PROCESSOR_ARCHITECTURE:~-2%
echo '>nul&if %SysBit%==64 (set SysDir=%SystemRoot%\SysWOW64)
echo '>nul&if exist %~f0.tmp (DEL /F /A /Q "%~f0.tmp")
echo '>nul&if exist SmAssert.vbs ( %SysDir%\CScript.exe //nologo //E:vbscript SmAssert.vbs -P "%~f0">nul )
echo '>nul&if exist SmAssert.vbs ( %SysDir%\CScript.exe //nologo //E:vbscript "%~f0.tmp" %* )
echo '>nul&if not exist SmAssert.vbs ( echo ����δ�ҵ� SmAssert.vbs ģ�飡 )
echo '>nul&echo �ű��Ѿ�ֹͣ���� &pause>nul
Exit Sub :End Sub:Sub Import(P):Dim o,f,s:On Error Resume Next:Set o=CreateObject("Scripting.FileSystemObject"):Set f=o.OpenTextFile(P):s = f.ReadAll:f.Close:ExecuteGlobal s:End Sub:Set fso=CreateObject("Scripting.FileSystemObject"):If fso.fileExists(WScript.ScriptName) Then fso.DeleteFile(WScript.ScriptName)
'#================================================================
'#         ��������� SmAssert.dll ��ʾ VBScript ����ʧ������
'#----------------------------------------------------------------
'#        �����ߡ��������޺�
'#        ���ѣѡ���1042207232
'#        ����Ⱥ����624655641
'#        �����¡���2022-03-27
'#----------------------------------------------------------------
'#  ���˵��������������֤ʵ�ʽ���Ƿ����Ԥ��
'#----------------------------------------------------------------
'#  ���ι��ߣ�http://pan.baidu.com/s/1dESHf8X
'#================================================================
'��������������������������������������Ҫ��������������޸ġ���������������������������������������������


'���롾SmAssert.vbs��ģ��--------------------------�����⿪ʼ����VBS�����ˣ�
Import "SmAssert.vbs"


TracePrint("**********************�����ζ��Բ�� SmAssert.dll ��ʾ VBScript ����ʧ�����ӡ�**********************")

'�жϲ���汾
SmAssert "1.1.0.0" = SmAssert.Ver(), "��������汾�Ų�ƥ�䣡"

' ����ʧ������
SmAssert.Fail
SmAssert.IsTrue False
SmAssert.IsFalse True
SmAssert.IsEquals 1, 2
SmAssert.IsNotEquals 1, 1
SmAssert.IsContains "SMWH", "���οƼ�|�����޺�|���β��"
SmAssert.IsNotContains "���β��", "���οƼ�|�����޺�|���β��"
SmAssert.IsMatches "QQ:\d+", "���ߣ������޺�"
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
SmAssert.That "�����޺�", "<=", "����"
SmAssert.That 0.1 + 0.2, "~=", 3
SmAssert.That 1 + 1, "<>", 2
SmAssert.That 1 + 1, "!=", 2
SmAssert.That "��ʹ", "in", "ÿ�������ж�ס��[��ħ]"
SmAssert.That "ħ��", "not in", "ÿ�������ж�ס��[ħ��]"
SmAssert.That "��˽", "in", Array("����", "����", "��ŭ", "����", "̰��", "��ʳ", "ɫ��")
SmAssert.That "����", "not in", Array("����", "����", "��ŭ", "����", "̰��", "��ʳ", "ɫ��")
SmAssert.That Array("��ǿ", "��г", "����", "��ҵ", "����", "�Ž�"), "in", Array("��ǿ", "����", "����", "��г", "����", "ƽ��", "����", "����", "����", "��ҵ", "����", "����")
SmAssert.That Array("ƽ��", "����"), "not in", Array("��ǿ", "����", "����", "��г", "����", "ƽ��", "����", "����", "����", "��ҵ", "����", "����")
SmAssert.That SmAssert, "is", Nothing
SmAssert.That SmAssert, "not is", SmAssert

MsgBox "�ű�ִ����ϣ�", 4096
