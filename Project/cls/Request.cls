VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Request"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Option Explicit
'
'Enum HTTPREQUEST_SETCREDENTIALS_FLAGS
'    HTTPREQUEST_PROXYSETTING_DEFAULT = 0      ' Ĭ�ϴ������ã���Ч�� 2
'    HTTPREQUEST_PROXYSETTING_PRECONFIG = 0    ' ָʾӦ��ע�����ȡ�������ã���Ҫ���� Proxycfg.exe�������Ч�� 1��
'    HTTPREQUEST_PROXYSETTING_DIRECT = 1       ' ָʾӦֱ�ӷ������� HTTP �� HTTPS �������� ���û�д�������������ʹ�ô����
'    HTTPREQUEST_PROXYSETTING_PROXY = 2        ' ָ��������������varProxyServer ������ַ��varBypassList ������������
'End Enum
'
'Const HTTPREQUEST_SETCREDENTIALS_FOR_SERVER = 0     ' ƾ�ݽ����ݵ�������
'Const HTTPREQUEST_SETCREDENTIALS_FOR_PROXY = 1      ' ƾ�ݽ����ݸ�����
'
'Const SslErrorIgnoreFlags_UNKNOWN_CA_OR_UNTRUSTED_ROOT = &H100          ' δ֪֤��䷢���� ��CA�� �������εĸ�
'Const SslErrorIgnoreFlags_WRONG_USAGE = &H200                           ' �÷�����
'Const SslErrorIgnoreFlags_INVALID_CN = &H1000                           ' ������ ��CN�� ��Ч
'Const SslErrorIgnoreFlags_INVALID_DATE_OR_CERTIFICATE_EXPIRED = &H2000  ' ������Ч��֤���ѹ���
'Const SslErrorIgnoreFlags_ALL = &H3300                                  ' ����֤�����
'
'
'' ��ȡ��ҳ����
'Private Function GetCharset(ByVal Text As String, Optional ByVal HomeString As String = "charset=", Optional ByVal EndString As String = ";") As String
'    Dim StartIndex
'    Dim EndIndex
'    Dim Charset
'
'    ' ���ҿ�ʼ�Ӵ�λ��
'    StartIndex = InStr(1, Text, HomeString, vbTextCompare)
'    ' ���δ�ҵ����򷵻ؿ�
'    If StartIndex = 0 Then GoTo return_
'
'    ' ���ҽ����Ӵ�λ��
'    EndIndex = InStr(StartIndex + Len(HomeString) + 1, Text, EndString, vbTextCompare)
'    ' ���δ�ҵ�������ĸ��������Ϊ�����Ӵ�λ��
'    If EndIndex = 0 Then EndIndex = Len(Text)
'
'    ' �������λ�� - ��ʼλ�� - ��ʼ�Ӵ����� = 0���򷵻ؿ�
'    If EndIndex - StartIndex - Len(HomeString) = 0 Then GoTo return_
'
'    ' ���ؽ��
'    Charset = Mid(Text, StartIndex + Len(HomeString), EndIndex - StartIndex - Len(HomeString))
'return_:
'    GetCharset = Charset
'End Function
'
'
'' �����ȡ��ҳ����
'Private Function GetCharsetMatch(ByVal Text As String) As String
'    Dim re      As New RegExp
'    Dim Result  As String
'
'    re.IgnoreCase = True
'    re.pattern = "\bcharset=[""']?([^\s<>""']+)"
'    If re.Test(Text) Then
'        Result = re.Execute(Text).Item(0).SubMatches.Item(0)
'    End If
'    Set re = Nothing
'    GetCharsetMatch = Result
'End Function
'
'
'' �����ж������ļ���ʽ�Ƿ�������ļ�
'Private Function IsBytesByContentTypeMatch(ByVal Text As String) As Boolean
'    Dim re      As New RegExp
'    Dim Result  As String
'
'    re.IgnoreCase = True
'    re.pattern = "Content-Type:(image/.*|application/octet-stream)"
'    Result = re.Test(Text)
'    Set re = Nothing
'    IsBytesByContentTypeMatch = Result
'End Function
'
'
'Public Function WinHtpRequest(this As WinHttpRequest, _
'                            ByVal Method As String, _
'                            ByVal URL As String, _
'                            Optional ByVal Data As Variant, _
'                            Optional ByVal Headers As String, _
'                            Optional ByVal Cookies As String, _
'                            Optional ByVal Charset As String, _
'                            Optional ByVal Timeout As Long, _
'                            Optional ByVal Auth As String, _
'                            Optional ByVal Redirects As Boolean = True, _
'                            Optional ByVal Proxy As String, _
'                            Optional ByVal ProxyUser As String, _
'                            Optional ByVal ProxyPass As String, _
'                            Optional ByVal ProxyID As HTTPREQUEST_SETCREDENTIALS_FLAGS = HTTPREQUEST_PROXYSETTING_DIRECT, _
'                            Optional ByVal IsAsync As Boolean, _
'                            Optional ByRef Res_Status As Boolean, _
'                            Optional ByRef Res_Headers As String, _
'                            Optional ByRef Res_Cookies As String _
'                ) As Variant
'    Dim http                As WinHttpRequest
'    Dim ObjStream           As Stream
'    Dim UserPassB64         As String
'    Dim HeadersArr          As Variant
'    Dim Header              As Variant
'    Dim Heads               As Variant
'    Dim ContentType         As String
'    Dim ContentTypeCharset  As String
'    Dim ResBody()           As Byte
'    Dim Result              As Variant
'
'    ' ����̳�
'    If this Is Nothing Then
'        Set http = New WinHttpRequest
'    Else
'        Set http = this
'    End If
'
'    ' ���ó�ʱ
'    If Timeout = -1 Then
'        Timeout = -1    ' ���޵ȴ�
'    ElseIf Timeout < 1 Then
'        Timeout = 30000
'    Else
'        Timeout = Timeout * 1000
'    End If
'    http.SetTimeouts Timeout, Timeout, Timeout, Timeout
'
'
'    ' ������ַ
'    If Len(Proxy) <> 0 Then
'        http.SetProxy HTTPREQUEST_PROXYSETTING_PROXY, Proxy
'    End If
'
'    ' ������ַ
'    http.Open Method, URL, IsAsync
'
'    ' ���ú��Է�����֤�����
'    http.Option(WinHttpRequestOption_SslErrorIgnoreFlags) = SslErrorIgnoreFlags_ALL
'
'    ' �����Ƿ��ض���
'    http.Option(WinHttpRequestOption_EnableRedirects) = Redirects
'
'    ' ���ô�����������֤
'    If Len(ProxyUser) <> 0 And Len(ProxyPass) <> 0 Then
'        http.SetCredentials ProxyUser, ProxyPass, ProxyID
'        If ProxyID = HTTPREQUEST_SETCREDENTIALS_FOR_PROXY Then
'            UserPassB64 = Base64Encoder(ProxyUser & ":" & ProxyPass)
'            http.SetRequestHeader "Proxy-Authorization", "Basic " & UserPassB64
'        End If
'    End If
'
'    ' ���� Cookies
'    http.SetRequestHeader "Cookie", Cookies
'
'    ' ��������ͷ
'    HeadersArr = Split(Headers, vbCrLf)
'    For Each Header In HeadersArr
'        Heads = Split(Header, ":", 2)
'        If UBound(Heads) = 1 Then
'            http.SetRequestHeader Heads(0), Heads(1)
'        End If
'    Next
'
'    ' ����������
'    http.Send Data
'
'    ' ����첽���򲻵ȴ����ؽ��
'    If IsAsync Then GoTo return_
'
'    ' ����Э��ͷ
'    Res_Headers = http.GetAllResponseHeaders
'
'    ' ���� Cookies
'    If InStr(Res_Headers, "Cookie:") > 0 Then
'        Res_Cookies = http.GetResponseHeader("Cookie")
'    End If
'
'    ' ������ҳ����
'    If InStr(Res_Headers, "Content-Type:") > 0 Then
'        ContentType = http.GetResponseHeader("Content-Type")
'        If IsBytesByContentTypeMatch(ContentType) Then
'            ContentTypeCharset = "Byte()"
'        Else
'            ContentTypeCharset = GetCharset(ContentType, "charset=", ";")
'        End If
'    End If
'
'    ' ����״̬��
'    Res_Status = http.status
'
'    ' ������ҳ����
'    ResBody = http.ResponseBody
'    If Charset = "Byte()" Or ContentTypeCharset = "Byte()" Then
'        Result = ResBody
'    ElseIf Len(ContentTypeCharset) <> 0 Then
'        Result = http.ResponseText
'    Else
'        If Len(Charset) <> 0 Then
'            Charset = GetCharsetMatch(http.ResponseText)
'        End If
'        Set ObjStream = CreateObject("Adodb.Stream")
'        With ObjStream
'            .Type = 1
'            .Mode = 3
'            .Open
'            .write ResBody
'            .position = 0
'            .Type = 2
'            .Charset = Charset
'             Result = .ReadText
'            .Close
'        End With
'        Set ObjStream = Nothing
'    End If
'    Set http = this
'
'return_:
'    WinHtpRequest = Result
'End Function





