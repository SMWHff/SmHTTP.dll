VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 1  'vbSimpleBound
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Response"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Implements WinHttpRequest '�������� Microsoft WinHTTP Services, version 5.1




'[����] ���ô�����������Ϣ
Private Sub WinHttpRequest_SetProxy(ByVal ProxySetting As WinHttp.HTTPREQUEST_PROXY_SETTING, Optional ByVal ProxyServer As Variant, Optional ByVal BypassList As Variant)

End Sub


'[����] ��ƾ������Ϊ�� HTTP ������һ��ʹ�ã������Ǵ�������������ԭʼ��������
Private Sub WinHttpRequest_SetCredentials(ByVal UserName As String, ByVal Password As String, ByVal Flags As WinHttp.HTTPREQUEST_SETCREDENTIALS_FLAGS)

End Sub


'[����] ���� HTTP ��Դ�� HTTP ���ӡ�
Private Sub WinHttpRequest_Open(ByVal Method As String, ByVal URL As String, Optional ByVal Async As Variant)

End Sub


'[����] ���ӡ����Ļ�ɾ�� HTTP �����ͷ��
Private Sub WinHttpRequest_SetRequestHeader(ByVal Header As String, ByVal Value As String)

End Sub


'[����] ���� HTTP ��Ӧ��ͷ��
Private Function WinHttpRequest_GetResponseHeader(ByVal Header As String) As String

End Function


'[����] �������� HTTP ��Ӧ��ͷ��
Private Function WinHttpRequest_GetAllResponseHeaders() As String

End Function


'[����] �� HTTP �����͵� HTTP ��������
Private Sub WinHttpRequest_Send(Optional ByVal Body As Variant)

End Sub


'[����] �ȴ��첽���ͷ�����ɣ�����Ϊ��λ��ѡ��ʱֵ����
Private Function WinHttpRequest_WaitForResponse(Optional ByVal Timeout As Variant) As Boolean

End Function


'[����] ��ֹ WinHTTPSend ������
Private Sub WinHttpRequest_Abort()

End Sub


'[����] ָ������/���ղ����ĸ�����ʱ������Ժ���Ϊ��λ����
Private Sub WinHttpRequest_SetTimeouts(ByVal ResolveTimeout As Long, ByVal ConnectTimeout As Long, ByVal SendTimeout As Long, ByVal ReceiveTimeout As Long)

End Sub


'[����] ѡ��Ҫ���͵���ȫ���ı�����Э��Ŀͻ���֤�飬 (HTTPS) ��������
Private Sub WinHttpRequest_SetClientCertificate(ByVal ClientCertificate As String)

End Sub


'[����] ���õ�ǰ���Զ���¼���ԡ�
Private Sub WinHttpRequest_SetAutoLogonPolicy(ByVal AutoLogonPolicy As WinHttp.WinHttpRequestAutoLogonPolicy)

End Sub


'[����][ֻ��] �����һ����Ӧ�м��� HTTP ״̬���롣
Private Property Get WinHttpRequest_Status() As Long

End Property


'[����][ֻ��] ���� HTTP ״̬�ı���
Private Property Get WinHttpRequest_StatusText() As String

End Property


'[����][ֻ��] ����Ӧʵ�����ļ���Ϊ�ı���
Private Property Get WinHttpRequest_ResponseText() As String

End Property


'[����][ֻ��] ����Ӧʵ�����ļ���Ϊ�޷����ֽ����顣
Private Property Get WinHttpRequest_ResponseBody() As Variant

End Property


'[����][ֻ��] �� IStream ����ʽ������Ӧʵ�����ġ�
Private Property Get WinHttpRequest_ResponseStream() As Variant

End Property


'[����][��д] ���û���� Microsoft Windows HTTP Services (WinHTTP) ѡ��ֵ��
Private Property Let WinHttpRequest_Option(ByVal Options As WinHttp.WinHttpRequestOption, ByVal RHS As Variant)

End Property
Private Property Get WinHttpRequest_Option(ByVal Options As WinHttp.WinHttpRequestOption) As Variant

End Property
