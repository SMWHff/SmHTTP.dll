:'��������������������������������������Ҫ��������������޸ġ���������������������������������������������
:'On Error Resume Next
:Sub bat
echo off & cls
echo '>nul&set SysDir=%SystemRoot%\System32
echo '>nul&set SysBit=%PROCESSOR_ARCHITECTURE:~-2%
echo '>nul&if %SysBit%==64 (set SysDir=%SystemRoot%\SysWOW64)
echo '>nul&if exist %~f0.tmp (DEL /F /A /Q "%~f0.tmp")
echo '>nul&if exist SmHTTP.vbs ( %SysDir%\CScript.exe //nologo //E:vbscript SmHTTP.vbs -P "%~f0">nul )
echo '>nul&if exist SmHTTP.vbs ( %SysDir%\CScript.exe //nologo //E:vbscript "%~f0.tmp" %* )
echo '>nul&if not exist SmHTTP.vbs ( echo ����δ�ҵ� SmHTTP.vbs ģ�飡 )
echo '>nul&echo �ű��Ѿ�ֹͣ���� &pause>nul
Exit Sub :End Sub:Sub Import(P):Dim o,f,s:On Error Resume Next:Set o=CreateObject("Scripting.FileSystemObject"):Set f=o.OpenTextFile(P):s = f.ReadAll:f.Close:ExecuteGlobal s:End Sub:Set fso=CreateObject("Scripting.FileSystemObject"):If fso.fileExists(WScript.ScriptName) Then fso.DeleteFile(WScript.ScriptName)
'#================================================================
'#         ����HTTP������ SmHTTP.dll ��ʾ VBScript ����
'#----------------------------------------------------------------
'#        �����ߡ��������޺�
'#        ���ѣѡ���1042207232
'#        ����Ⱥ����624655641
'#        �����¡���2022-03-27
'#----------------------------------------------------------------
'#  ���˵��������HTTPЭ���������ʲ���
'#----------------------------------------------------------------
'#  ���ι��ߣ�http://pan.baidu.com/s/1dESHf8X
'#================================================================
'��������������������������������������Ҫ��������������޸ġ���������������������������������������������


'���롾SmHTTP.vbs��ģ��--------------------------�����⿪ʼ����VBS�����ˣ�
Import "SmHTTP.vbs"


TracePrint("**********************������HTTP������ SmHTTP.dll ��ʾ VBScript ���ӡ�**********************")

'�жϲ���汾
SmHTTP "1.0.0.0" = SmHTTP.Ver(), "��������汾�Ų�ƥ�䣡"




MsgBox "�ű�ִ����ϣ�", 4096
