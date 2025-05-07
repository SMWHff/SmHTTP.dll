Attribute VB_Name = "XHR"
Option Explicit


' �����֤ö��
Public Enum HTTPREQUEST_AUTHORIZATION
    HTTP_AUTH_BASIC = 0     ' ������֤
    HTTP_AUTH_DIGEST = 2    ' ժҪ��֤
    HTTP_AUTH_FORMBASE = 3  ' ����֤
End Enum


Private Enum METHOD_OPTIONS
    METHOD_GET = 0
    METHOD_POST = 1
    METHOD_HEAD = 2
    METHOD_PUT = 3
    METHOD_OPTIONS = 4
    METHOD_DELETE = 5
    METHOD_TRACE = 6
    METHOD_CONNECT = 7
    METHOD_PATCH = 8
End Enum

Private Enum HTTPREQUEST_SETCREDENTIALS_FLAGS
    HTTPREQUEST_PROXYSETTING_DEFAULT = 0      ' Ĭ�ϴ������ã���Ч�� 2
    HTTPREQUEST_PROXYSETTING_PRECONFIG = 0    ' ָʾӦ��ע����ȡ�������ã���Ҫ���� Proxycfg.exe�������Ч�� 1��
    HTTPREQUEST_PROXYSETTING_DIRECT = 1       ' ָʾӦֱ�ӷ������� HTTP �� HTTPS �������� ���û�д������������ʹ�ô����
    HTTPREQUEST_PROXYSETTING_PROXY = 2        ' ָ�������������varProxyServer �����ַ��varBypassList ������������
End Enum

Private Const HTTPREQUEST_SETCREDENTIALS_FOR_SERVER = 0     ' ƾ�ݽ����ݵ�������
Private Const HTTPREQUEST_SETCREDENTIALS_FOR_PROXY = 1      ' ƾ�ݽ����ݸ�����

Private Const SslErrorIgnoreFlags_UNKNOWN_CA_OR_UNTRUSTED_ROOT = &H100          ' δ֪֤��䷢���� ��CA�� �������εĸ�
Private Const SslErrorIgnoreFlags_WRONG_USAGE = &H200                           ' �÷�����
Private Const SslErrorIgnoreFlags_INVALID_CN = &H1000                           ' ������ ��CN�� ��Ч
Private Const SslErrorIgnoreFlags_INVALID_DATE_OR_CERTIFICATE_EXPIRED = &H2000  ' ������Ч��֤���ѹ���
Private Const SslErrorIgnoreFlags_ALL = &H3300                                  ' ����֤�����


' �����ȡ��ҳ����
Private Function GetCharsetMatch(ByVal Text As String) As String
    Dim re      As New RegExp
    Dim Result  As String
    
    re.IgnoreCase = True
    re.pattern = "\bcharset=[""']?([^\s<>""']+)"
    If re.Test(Text) Then
        Result = re.Execute(Text).Item(0).SubMatches.Item(0)
    End If
    Set re = Nothing
    GetCharsetMatch = Result
End Function


' �����ж������ļ���ʽ�Ƿ�������ļ�
Private Function IsBytesByContentTypeMatch(ByVal Text As String) As Boolean
    Dim re      As New RegExp
    Dim Result  As String
    
    re.IgnoreCase = True
    re.pattern = "image/.*|audio/.*|application/octet-stream"
    Result = re.Test(Text)
    Set re = Nothing
    IsBytesByContentTypeMatch = Result
End Function


' ���ɴ�������������֤��ַ
Public Function Gen_Proxies(ByVal Proxy As String, ByVal ProxyUser As String, ByVal ProxyPass As String)
    Dim Result As String
    
    If Len(ProxyUser) <> 0 And Len(ProxyPass) <> 0 Then
        Result = ProxyUser & ":" & ProxyPass & "@"
    End If
    Result = Result & Proxy
End Function


' ��ȡ����Э��ͷ�е� Cookies
Public Function FetchCookies(ByVal AllHeaders As String) As String
    Dim Result      As String
    Dim HeadersArr  As Variant
    Dim Header      As Variant
    
    If InStr(AllHeaders, "Set-Cookie:") > 0 Then
        HeadersArr = Split(AllHeaders, vbCrLf)
        For Each Header In HeadersArr
            If Left(Header, 11) = "Set-Cookie:" Then
                If InStr(Header, ";") > 0 Then
                    Result = Result & Trim(MidStr(Header, "Set-Cookie:", ";")) & "; "
                Else
                    Result = Result & Trim(Replace(Header, "Set-Cookie:", "")) & "; "
                End If
            End If
        Next
    End If
    FetchCookies = Result
End Function


' �ϲ����� Cookies
Public Function MergeUpdateCookies(ByVal oldCookies As String, ByVal NwCookies As String) As String
    Dim Dict        As Dictionary   ' �������� Microsoft Scripting Runtime
    Dim re          As RegExp       ' �������� Microsoft VBScript Regular Expressions 5.5
    Dim Matchs      As Object
    Dim SubMatchs   As Object
    Dim mCookies    As String
    Dim i           As Long
    Dim key         As String
    Dim value       As String
    Dim Cookies()   As Variant
    Dim DictKeys    As Variant
    Dim DictItems   As Variant
    
    ' �����ϲ��¾�Cookies
    mCookies = NwCookies & IIf(Len(oldCookies) <> 0 And Len(NwCookies) <> 0, "; ", "") & oldCookies
    
    Set re = New RegExp
    re.IgnoreCase = True
    re.Global = True
    re.pattern = "([^=; ]+)=([^; ]+)"
    If Not re.Test(mCookies) Then GoTo return_over
    Set Dict = New Dictionary
    Set Matchs = re.Execute(mCookies)
    For i = 0 To Matchs.Count - 1
        Set SubMatchs = Matchs.Item(i).SubMatches
        If Not SubMatchs.Count = 2 Then GoTo continue
        key = SubMatchs.Item(0)
        value = SubMatchs.Item(1)
        If Not Dict.Exists(key) = False Then GoTo continue
        If Len(value) <> 0 And value <> "deleted" Then
            Dict(key) = value
        End If
continue:
    Next
    DictKeys = Dict.keys
    DictItems = Dict.Items
    ReDim Cookies(Dict.Count - 1)
    For i = 0 To Dict.Count - 1
        Cookies(i) = DictKeys(i) & "=" & DictItems(i)
    Next
    Set SubMatchs = Nothing
    Set Matchs = Nothing
    Set Dict = Nothing
return_over:
    Set re = Nothing
    MergeUpdateCookies = Join(Cookies, "; ")
End Function



Public Function WinHtpRequest(ByRef this As WinHttpRequest, _
                                ByVal Method As String, _
                                ByVal URL As String, _
                                Optional ByVal Data As Variant, _
                                Optional ByVal Headers As String, _
                                Optional ByVal Cookies As String, _
                                Optional ByVal Charset As String, _
                                Optional ByVal Timeout As Long, _
                                Optional ByVal Auth As String, _
                                Optional ByVal Redirects As Boolean = True, _
                                Optional ByVal Proxy As String, _
                                Optional ByVal ProxyUser As String, _
                                Optional ByVal ProxyPass As String, _
                                Optional ByVal AuthType As String, _
                                Optional ByVal AuthUser As String, _
                                Optional ByVal AuthPass As String, _
                                Optional ByVal IsAsync As Boolean, _
                                Optional ByVal CompleteHeaders As Boolean = True, _
                                Optional ByVal CompleteCookies As Boolean = True, _
                                Optional ByRef Res_Status As Long, _
                                Optional ByRef Res_Headers As String, _
                                Optional ByRef Res_Cookies As String, _
                                Optional ByRef Res_Body As Variant _
                            ) As Variant
                            
    Dim http                As WinHttpRequest
    Dim ObjStream           As stream
    Dim UserPassB64         As String
    Dim HeadersArr          As Variant
    Dim Header              As Variant
    Dim Heads               As Variant
    Dim ContentType         As String
    Dim ContentTypeCharset  As String
    Dim ContentEncoding     As String
    Dim IsGzip              As Boolean
    Dim ResBody             As Variant
    Dim Result              As Variant
    
Begin:
    On Error Resume Next
    ' ����̳�
    If this Is Nothing Then
        Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    Else
        Set http = this
    End If

    ' ���ó�ʱ
    If Timeout = -1 Or LCase(Charset) = "byte()" Then
        Timeout = -1    ' ���޵ȴ�
    ElseIf Timeout < 1 Then
        Timeout = 30000 ' Ĭ��30��
    Else
        Timeout = Timeout * 1000
    End If
    http.SetTimeouts Timeout, Timeout, Timeout, Timeout
    

    ' �����ַ
    If Len(Proxy) <> 0 Then
        http.SetProxy HTTPREQUEST_PROXYSETTING_PROXY, Proxy
    End If
    
    ' ������ַ
    'Debug.Print "��URL��=", URL
    http.Open Method, URL, IsAsync
    
    ' ���ú��Է�����֤�����
    http.Option(WinHttpRequestOption_SslErrorIgnoreFlags) = SslErrorIgnoreFlags_ALL
    
    ' �����Ƿ��ض���
    http.Option(WinHttpRequestOption_EnableRedirects) = Redirects
    
    ' ���ô��������֤��Ϣ
    If Len(ProxyUser) <> 0 And Len(ProxyPass) <> 0 Then
        http.SetCredentials ProxyUser, ProxyPass, HTTPREQUEST_SETCREDENTIALS_FOR_PROXY
    End If
    
    ' ���ô��������֤��Ϣ
    If Len(AuthUser) <> 0 And Len(AuthPass) <> 0 Then
        Select Case UCase(AuthType)
        Case "BASIC"    ' ������֤
            http.SetCredentials AuthUser, AuthPass, HTTP_AUTH_BASIC
        Case "DIGEST"   ' ժҪ��֤
            Headers = Headers & vbCrLf & "WWW-Authenticate: DIGEST ժҪ��Ϣ"
        Case "FORMBASE" ' ����֤
        Case "BEARER"      ' OAuth �� JWT ��Ȩ
            Headers = Headers & vbCrLf & "Authorization: Bearer ��Ȩ��Ϣ"
        End Select
    End If
    
    
    ' �Ƿ�ȫ��ҪЭ��ͷ
    If CompleteHeaders Then
        If InStr(1, Headers, "Accept:", 1) = 0 Then
            Headers = Headers & vbCrLf & "Accept: */*"
        End If
        If InStr(1, Headers, "Accept-Encoding:", 1) = 0 Then
            Headers = Headers & vbCrLf & "Accept-Encoding: identity"  ' ǿ�Ʒ���������δѹ��������
        End If
        If InStr(1, Headers, "Accept-Language:", 1) = 0 Then
            Headers = Headers & vbCrLf & "Accept-Language: zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6"
        End If
        If InStr(1, Headers, "Cache-Control:", 1) = 0 Then
            Headers = Headers & vbCrLf & "Cache-Control: no-cache"
        End If
        If InStr(1, Headers, "Referer:", 1) = 0 Then
            Headers = Headers & vbCrLf & "Referer: " & URL
        End If
        If InStr(1, Headers, "Host:", 1) = 0 Then
            Headers = Headers & vbCrLf & "Host: " & Split(URL, "/")(2)
        End If
        If InStr(1, Headers, "User-Agent:", 1) = 0 Then
            Headers = Headers & vbCrLf & "User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36 Edg/119.0.0.0"
        End If
        If InStr(1, Headers, "Content-Type:", 1) = 0 Then
            Headers = Headers & vbCrLf & "Content-Type: application/x-www-form-urlencoded"
        End If
        If InStr(1, Headers, "Content-Length:", 1) = 0 And Len(Data) > 0 Then
            If VarType(Data) = vbArray + vbByte Then
                Headers = Headers & vbCrLf & "Content-Length: " & UBound(Data)
            Else
                Headers = Headers & vbCrLf & "Content-Length: " & Len(Data)
            End If
        End If
    End If
    'Debug.Print "��Headers����", CompleteHeaders, Headers
    
    ' ��������ͷ
    HeadersArr = Split(Headers, vbCrLf)
    For Each Header In HeadersArr
        Heads = Split(Header, ":", 2)
        If UBound(Heads) = 1 Then
            http.SetRequestHeader Heads(0), Heads(1)
        End If
    Next
    
    ' ���� Cookies
    If Len(Cookies) <> 0 Then
        http.SetRequestHeader "Cookie", Cookies
    End If
    
    ' ����������
    'Debug.Print "��Data��=", Data
    http.Send Data
    
    ' ����첽���򲻵ȴ����ؽ��
    If IsAsync Then
        GoTo return_
    End If
    
    ' ����Э��ͷ
    Res_Headers = http.GetAllResponseHeaders

    
    ' ���� Cookies
    Res_Cookies = FetchCookies(Res_Headers)
    
    ' �Ƿ��Զ��ϲ�����Cookie
    If CompleteCookies And Len(Res_Cookies) <> 0 Then
        Res_Cookies = MergeUpdateCookies(Cookies, Res_Cookies)
        'Debug.Print "Res_Cookies=", Res_Cookies
    End If
    
    ' ������ҳ����
    If InStr(Res_Headers, "Content-Type:") > 0 Then
        ContentType = http.GetResponseHeader("Content-Type")
        If IsBytesByContentTypeMatch(ContentType) Then
            ContentTypeCharset = "Byte()"
        Else
            ContentTypeCharset = GetCharsetMatch(ContentType)
        End If
    End If
    
    ' ����ѹ����ʽ
    If InStr(Res_Headers, "Content-Encoding:") > 0 Then
        ContentEncoding = http.GetResponseHeader("Content-Encoding")
        If InStr(ContentEncoding, "gzip") > 0 Then
            IsGzip = True
        End If
    End If
    
    
    ' ����״̬��
    Res_Status = http.Status
    
    ' ���ض���������
    Res_Body = http.ResponseBody
    
    
    ' ������ҳ����
    If InStr("HEAD", Method) = 0 Then
        ResBody = http.ResponseBody
    End If
    If UCase(Charset) = "ANSI" Then Charset = ""
    If LCase(Charset) = "byte()" Or ContentTypeCharset = "Byte()" Then
        Result = ResBody
    ElseIf Len(ContentTypeCharset) <> 0 And Len(Charset) = 0 And IsGzip = False Then
        Result = http.ResponseText
    End If
    If Len(Result) = 0 And Len(ResBody) > 0 Then
        ' ���û��ָ�����룬���Զ�����ҳԴ���л�ȡ�������ȡʧ�ܣ���Ĭ�� UTF-8
        If Len(Charset) = 0 Then Charset = GetCharsetMatch(StrConv(ResBody, vbUnicode))
        If Len(Charset) = 0 Then Charset = "UTF-8"
        If Left(LCase(Charset), 5) = "file|" Then
            ' ���浽�ļ�
            Dim ReqStream       As stream   '�������� Microsoft ActiveX Data Objects 2.8 Libary
            Dim bufferSize      As Long
            Dim bytesRead       As Long
            Dim buffer          As Variant
            Dim FilePath        As String
            
            FilePath = Mid(Charset, 6)
            Set ReqStream = http.ResponseStream
            Set ObjStream = CreateObject("ADODB.Stream")
            ObjStream.Type = 1 ' Binary
            ObjStream.Open
            ' �� ResponseStream ��ȡ���ݲ�д�� outputStream
            bufferSize = 2048 ' ���磬ÿ�ζ�ȡ 2048 �ֽ�
            Do
                buffer = ReqStream.Read(bufferSize)
                bytesRead = LenB(buffer)
                If bytesRead > 0 Then
                    ObjStream.Write buffer
                End If
            Loop While bytesRead = bufferSize
            ObjStream.SaveToFile FilePath, 2 ' 2 = Overwrite
            ObjStream.Close
            Set ObjStream = Nothing
            Set ReqStream = Nothing
            Result = IIf(IsFileExist(FilePath), 1, 0)
        Else
            ' �����ı�����
            Set ObjStream = CreateObject("Adodb.Stream")
            With ObjStream
                .Type = 1
                .Mode = 3
                .Open
                .Write ResBody
                .position = 0
                .Type = 2
                .Charset = Charset
                 Result = .ReadText
                .Close
            End With
            Set ObjStream = Nothing
        End If
    End If
    Set http = Nothing
    
return_:
    'If Err Then Debug.Print Err.Description
    WinHtpRequest = Result
End Function


