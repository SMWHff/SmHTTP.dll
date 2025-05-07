Attribute VB_Name = "XHR"
Option Explicit


' 身份认证枚举
Public Enum HTTPREQUEST_AUTHORIZATION
    HTTP_AUTH_BASIC = 0     ' 基本认证
    HTTP_AUTH_DIGEST = 2    ' 摘要认证
    HTTP_AUTH_FORMBASE = 3  ' 表单认证
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
    HTTPREQUEST_PROXYSETTING_DEFAULT = 0      ' 默认代理设置，等效于 2
    HTTPREQUEST_PROXYSETTING_PRECONFIG = 0    ' 指示应从注册表获取代理设置（需要运行 Proxycfg.exe，否则等效于 1）
    HTTPREQUEST_PROXYSETTING_DIRECT = 1       ' 指示应直接访问所有 HTTP 和 HTTPS 服务器。 如果没有代理服务器，请使用此命令。
    HTTPREQUEST_PROXYSETTING_PROXY = 2        ' 指定代理服务器（varProxyServer 代理地址，varBypassList 域名黑名单）
End Enum

Private Const HTTPREQUEST_SETCREDENTIALS_FOR_SERVER = 0     ' 凭据将传递到服务器
Private Const HTTPREQUEST_SETCREDENTIALS_FOR_PROXY = 1      ' 凭据将传递给代理

Private Const SslErrorIgnoreFlags_UNKNOWN_CA_OR_UNTRUSTED_ROOT = &H100          ' 未知证书颁发机构 （CA） 或不受信任的根
Private Const SslErrorIgnoreFlags_WRONG_USAGE = &H200                           ' 用法错误
Private Const SslErrorIgnoreFlags_INVALID_CN = &H1000                           ' 公用名 （CN） 无效
Private Const SslErrorIgnoreFlags_INVALID_DATE_OR_CERTIFICATE_EXPIRED = &H2000  ' 日期无效或证书已过期
Private Const SslErrorIgnoreFlags_ALL = &H3300                                  ' 所有证书错误


' 正则获取网页编码
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


' 正则判断网络文件格式是否二进制文件
Private Function IsBytesByContentTypeMatch(ByVal Text As String) As Boolean
    Dim re      As New RegExp
    Dim Result  As String
    
    re.IgnoreCase = True
    re.pattern = "image/.*|audio/.*|application/octet-stream"
    Result = re.Test(Text)
    Set re = Nothing
    IsBytesByContentTypeMatch = Result
End Function


' 生成代理服务器身份认证地址
Public Function Gen_Proxies(ByVal Proxy As String, ByVal ProxyUser As String, ByVal ProxyPass As String)
    Dim Result As String
    
    If Len(ProxyUser) <> 0 And Len(ProxyPass) <> 0 Then
        Result = ProxyUser & ":" & ProxyPass & "@"
    End If
    Result = Result & Proxy
End Function


' 提取返回协议头中的 Cookies
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


' 合并更新 Cookies
Public Function MergeUpdateCookies(ByVal oldCookies As String, ByVal NwCookies As String) As String
    Dim Dict        As Dictionary   ' 工程引用 Microsoft Scripting Runtime
    Dim re          As RegExp       ' 工程引用 Microsoft VBScript Regular Expressions 5.5
    Dim Matchs      As Object
    Dim SubMatchs   As Object
    Dim mCookies    As String
    Dim i           As Long
    Dim key         As String
    Dim value       As String
    Dim Cookies()   As Variant
    Dim DictKeys    As Variant
    Dim DictItems   As Variant
    
    ' 初步合并新旧Cookies
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
    ' 对象继承
    If this Is Nothing Then
        Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    Else
        Set http = this
    End If

    ' 设置超时
    If Timeout = -1 Or LCase(Charset) = "byte()" Then
        Timeout = -1    ' 无限等待
    ElseIf Timeout < 1 Then
        Timeout = 30000 ' 默认30秒
    Else
        Timeout = Timeout * 1000
    End If
    http.SetTimeouts Timeout, Timeout, Timeout, Timeout
    

    ' 代理地址
    If Len(Proxy) <> 0 Then
        http.SetProxy HTTPREQUEST_PROXYSETTING_PROXY, Proxy
    End If
    
    ' 访问网址
    'Debug.Print "【URL】=", URL
    http.Open Method, URL, IsAsync
    
    ' 设置忽略服务器证书错误
    http.Option(WinHttpRequestOption_SslErrorIgnoreFlags) = SslErrorIgnoreFlags_ALL
    
    ' 设置是否重定向
    http.Option(WinHttpRequestOption_EnableRedirects) = Redirects
    
    ' 设置代理身份认证信息
    If Len(ProxyUser) <> 0 And Len(ProxyPass) <> 0 Then
        http.SetCredentials ProxyUser, ProxyPass, HTTPREQUEST_SETCREDENTIALS_FOR_PROXY
    End If
    
    ' 设置代理身份认证信息
    If Len(AuthUser) <> 0 And Len(AuthPass) <> 0 Then
        Select Case UCase(AuthType)
        Case "BASIC"    ' 基本认证
            http.SetCredentials AuthUser, AuthPass, HTTP_AUTH_BASIC
        Case "DIGEST"   ' 摘要认证
            Headers = Headers & vbCrLf & "WWW-Authenticate: DIGEST 摘要信息"
        Case "FORMBASE" ' 表单认证
        Case "BEARER"      ' OAuth 和 JWT 授权
            Headers = Headers & vbCrLf & "Authorization: Bearer 授权信息"
        End Select
    End If
    
    
    ' 是否补全必要协议头
    If CompleteHeaders Then
        If InStr(1, Headers, "Accept:", 1) = 0 Then
            Headers = Headers & vbCrLf & "Accept: */*"
        End If
        If InStr(1, Headers, "Accept-Encoding:", 1) = 0 Then
            Headers = Headers & vbCrLf & "Accept-Encoding: identity"  ' 强制服务器返回未压缩的内容
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
    'Debug.Print "【Headers】：", CompleteHeaders, Headers
    
    ' 设置请求头
    HeadersArr = Split(Headers, vbCrLf)
    For Each Header In HeadersArr
        Heads = Split(Header, ":", 2)
        If UBound(Heads) = 1 Then
            http.SetRequestHeader Heads(0), Heads(1)
        End If
    Next
    
    ' 设置 Cookies
    If Len(Cookies) <> 0 Then
        http.SetRequestHeader "Cookie", Cookies
    End If
    
    ' 发送请求体
    'Debug.Print "【Data】=", Data
    http.Send Data
    
    ' 如果异步，则不等待返回结果
    If IsAsync Then
        GoTo return_
    End If
    
    ' 返回协议头
    Res_Headers = http.GetAllResponseHeaders

    
    ' 返回 Cookies
    Res_Cookies = FetchCookies(Res_Headers)
    
    ' 是否自动合并更新Cookie
    If CompleteCookies And Len(Res_Cookies) <> 0 Then
        Res_Cookies = MergeUpdateCookies(Cookies, Res_Cookies)
        'Debug.Print "Res_Cookies=", Res_Cookies
    End If
    
    ' 返回网页编码
    If InStr(Res_Headers, "Content-Type:") > 0 Then
        ContentType = http.GetResponseHeader("Content-Type")
        If IsBytesByContentTypeMatch(ContentType) Then
            ContentTypeCharset = "Byte()"
        Else
            ContentTypeCharset = GetCharsetMatch(ContentType)
        End If
    End If
    
    ' 返回压缩格式
    If InStr(Res_Headers, "Content-Encoding:") > 0 Then
        ContentEncoding = http.GetResponseHeader("Content-Encoding")
        If InStr(ContentEncoding, "gzip") > 0 Then
            IsGzip = True
        End If
    End If
    
    
    ' 返回状态码
    Res_Status = http.Status
    
    ' 返回二进制内容
    Res_Body = http.ResponseBody
    
    
    ' 返回网页内容
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
        ' 如果没有指定编码，则自动从网页源码中获取，如果获取失败，则默认 UTF-8
        If Len(Charset) = 0 Then Charset = GetCharsetMatch(StrConv(ResBody, vbUnicode))
        If Len(Charset) = 0 Then Charset = "UTF-8"
        If Left(LCase(Charset), 5) = "file|" Then
            ' 保存到文件
            Dim ReqStream       As stream   '工程引用 Microsoft ActiveX Data Objects 2.8 Libary
            Dim bufferSize      As Long
            Dim bytesRead       As Long
            Dim buffer          As Variant
            Dim FilePath        As String
            
            FilePath = Mid(Charset, 6)
            Set ReqStream = http.ResponseStream
            Set ObjStream = CreateObject("ADODB.Stream")
            ObjStream.Type = 1 ' Binary
            ObjStream.Open
            ' 从 ResponseStream 读取数据并写入 outputStream
            bufferSize = 2048 ' 例如，每次读取 2048 字节
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
            ' 返回文本内容
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


