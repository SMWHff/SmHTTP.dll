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
Exit Sub :End Sub:Sub Import(P):Dim o,f,s:Set o=CreateObject("Scripting.FileSystemObject"):Set f=o.OpenTextFile(P):s = f.ReadAll:f.Close:ExecuteGlobal s:End Sub:Set fso=CreateObject("Scripting.FileSystemObject"):If fso.fileExists(WScript.ScriptName) Then fso.DeleteFile(WScript.ScriptName)
'#================================================================
'#         ����HTTP������ SmHTTP.dll ��ʾ VBScript ����
'#----------------------------------------------------------------
'#        �����ߡ��������޺�
'#        ���ѣѡ���1042207232
'#        ����Ⱥ����624655641
'#        �����¡���2023-11-03
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
Assert "1.0.0.4" = SmHTTP.Ver(), "��������汾�Ų�ƥ�䣡"

' �������
Dim user, pass, Data, Ret, Cookies, Headers

' �����˺�
user = "��İ���������̳�˺�"
pass = "��İ���������̳����"

If user = "��İ���������̳�˺�" Then user = Environ("AJ_USER")
If pass = "��İ���������̳����" Then pass = Environ("AJ_PASS")

Data = SmHTTP.Data( _
    "username", user, _
    "password", pass, _
    "question", "0", _
    "answer", "", _
    "templateid", "0", _
    "login", "", _
    "expires", "43200" _
)


' �����Զ�ʶ�����ģʽ
Call SmHTTP.SetAutoParamArray(True)


' ��¼��̳�˺�
Ret = SmHTTP.HTTP_POST("http://bbs.anjian.com/login.aspx?referer=forumindex.aspx", Data)
' �ж��Ƿ��¼�ɹ�
If InStr(Ret, user) = 0 Then
    MsgBox "������¼ʧ�ܣ�", 16 + 4096, "����"
    EndScript
End If
Cookies = SmHTTP.GetCookies()


' ��ǩ��
Data = SmHTTP.Data( _
    "signmessage", "ǩ������ÿ�����鶼��������~~��������ף������������������" _
)
Headers = SmHTTP.Headers( _
    "Referer", "http://bbs.anjian.com/" _
)
Ret = SmHTTP.HTTP_POST("http://bbs.anjian.com/addsignin.aspx?infloat=1&inajax=1", Data, Headers, Cookies)
If InStr(Ret, "��ϲ����ȡ����ǩ������") Or InStr(Ret, "������Ѿ�ǩ������") Then ' �ж��Ƿ�ǩ���ɹ�
    TracePrint "��ϲ���������ǩ������"
Else
    TracePrint Ret
    MsgBox "����ǩ��ʧ�ܣ�", 16 + 4096, "����"
    EndScript
End If


MsgBox "�ű�ִ����ϣ�", 4096
