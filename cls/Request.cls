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
'    HTTPREQUEST_PROXYSETTING_DEFAULT = 0      ' 默认代理设置，等效于 2
'    HTTPREQUEST_PROXYSETTING_PRECONFIG = 0    ' 指示应从注册表获取代理设置（需要运行 Proxycfg.exe，否则等效于 1）
'    HTTPREQUEST_PROXYSETTING_DIRECT = 1       ' 指示应直接访问所有 HTTP 和 HTTPS 服务器。 如果没有代理服务器，请使用此命令。
'    HTTPREQUEST_PROXYSETTING_PROXY = 2        ' 指定代理服务器（varProxyServer 代理地址，varBypassList 域名黑名单）
'End Enum
'
'Const HTTPREQUEST_SETCREDENTIALS_FOR_SERVER = 0     ' 凭据将传递到服务器
'Const HTTPREQUEST_SETCREDENTIALS_FOR_PROXY = 1      ' 凭据将传递给代理
'
'Const SslErrorIgnoreFlags_UNKNOWN_CA_OR_UNTRUSTED_ROOT = &H100          ' 未知证书颁发机构 （CA） 或不受信任的根
'Const SslErrorIgnoreFlags_WRONG_USAGE = &H200                           ' 用法错误
'Const SslErrorIgnoreFlags_INVALID_CN = &H1000                           ' 公用名 （CN） 无效
'Const SslErrorIgnoreFlags_INVALID_DATE_OR_CERTIFICATE_EXPIRED = &H2000  ' 日期无效或证书已过期
'Const SslErrorIgnoreFlags_ALL = &H3300                                  ' 所有证书错误
'
'
'' 获取网页编码
'Private Function GetCharset(ByVal Text As String, Optional ByVal HomeString As String = "charset=", Optional ByVal EndString As String = ";") As String
'    Dim StartIndex
'    Dim EndIndex
'    Dim Charset
'
'    ' 查找开始子串位置
'    StartIndex = InStr(1, Text, HomeString, vbTextCompare)
'    ' 如果未找到，则返回空
'    If StartIndex = 0 Then GoTo return_
'
'    ' 查找结束子串位置
'    EndIndex = InStr(StartIndex + Len(HomeString) + 1, Text, EndString, vbTextCompare)
'    ' 如果未找到，则以母串长度作为结束子串位置
'    If EndIndex = 0 Then EndIndex = Len(Text)
'
'    ' 如果结束位置 - 开始位置 - 开始子串长度 = 0，则返回空
'    If EndIndex - StartIndex - Len(HomeString) = 0 Then GoTo return_
'
'    ' 返回结果
'    Charset = Mid(Text, StartIndex + Len(HomeString), EndIndex - StartIndex - Len(HomeString))
'return_:
'    GetCharset = Charset
'End Function
'
'
'' 正则获取网页编码
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
'' 正则判断网络文件格式是否二进制文件
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
'    ' 对象继承
'    If this Is Nothing Then
'        Set http = New WinHttpRequest
'    Else
'        Set http = this
'    End If
'
'    ' 设置超时
'    If Timeout = -1 Then
'        Timeout = -1    ' 无限等待
'    ElseIf Timeout < 1 Then
'        Timeout = 30000
'    Else
'        Timeout = Timeout * 1000
'    End If
'    http.SetTimeouts Timeout, Timeout, Timeout, Timeout
'
'
'    ' 代理地址
'    If Len(Proxy) <> 0 Then
'        http.SetProxy HTTPREQUEST_PROXYSETTING_PROXY, Proxy
'    End If
'
'    ' 访问网址
'    http.Open Method, URL, IsAsync
'
'    ' 设置忽略服务器证书错误
'    http.Option(WinHttpRequestOption_SslErrorIgnoreFlags) = SslErrorIgnoreFlags_ALL
'
'    ' 设置是否重定向
'    http.Option(WinHttpRequestOption_EnableRedirects) = Redirects
'
'    ' 设置代理服务器认证
'    If Len(ProxyUser) <> 0 And Len(ProxyPass) <> 0 Then
'        http.SetCredentials ProxyUser, ProxyPass, ProxyID
'        If ProxyID = HTTPREQUEST_SETCREDENTIALS_FOR_PROXY Then
'            UserPassB64 = Base64Encoder(ProxyUser & ":" & ProxyPass)
'            http.SetRequestHeader "Proxy-Authorization", "Basic " & UserPassB64
'        End If
'    End If
'
'    ' 设置 Cookies
'    http.SetRequestHeader "Cookie", Cookies
'
'    ' 设置请求头
'    HeadersArr = Split(Headers, vbCrLf)
'    For Each Header In HeadersArr
'        Heads = Split(Header, ":", 2)
'        If UBound(Heads) = 1 Then
'            http.SetRequestHeader Heads(0), Heads(1)
'        End If
'    Next
'
'    ' 发送请求体
'    http.Send Data
'
'    ' 如果异步，则不等待返回结果
'    If IsAsync Then GoTo return_
'
'    ' 返回协议头
'    Res_Headers = http.GetAllResponseHeaders
'
'    ' 返回 Cookies
'    If InStr(Res_Headers, "Cookie:") > 0 Then
'        Res_Cookies = http.GetResponseHeader("Cookie")
'    End If
'
'    ' 返回网页编码
'    If InStr(Res_Headers, "Content-Type:") > 0 Then
'        ContentType = http.GetResponseHeader("Content-Type")
'        If IsBytesByContentTypeMatch(ContentType) Then
'            ContentTypeCharset = "Byte()"
'        Else
'            ContentTypeCharset = GetCharset(ContentType, "charset=", ";")
'        End If
'    End If
'
'    ' 返回状态码
'    Res_Status = http.status
'
'    ' 返回网页内容
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






