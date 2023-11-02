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

Implements WinHttpRequest '工程引用 Microsoft WinHTTP Services, version 5.1




'[方法] 设置代理服务器信息
Private Sub WinHttpRequest_SetProxy(ByVal ProxySetting As WinHttp.HTTPREQUEST_PROXY_SETTING, Optional ByVal ProxyServer As Variant, Optional ByVal BypassList As Variant)

End Sub


'[方法] 将凭据设置为与 HTTP 服务器一起使用，无论是代理服务器还是原始服务器。
Private Sub WinHttpRequest_SetCredentials(ByVal UserName As String, ByVal Password As String, ByVal Flags As WinHttp.HTTPREQUEST_SETCREDENTIALS_FLAGS)

End Sub


'[方法] 打开与 HTTP 资源的 HTTP 连接。
Private Sub WinHttpRequest_Open(ByVal Method As String, ByVal URL As String, Optional ByVal Async As Variant)

End Sub


'[方法] 添加、更改或删除 HTTP 请求标头。
Private Sub WinHttpRequest_SetRequestHeader(ByVal Header As String, ByVal Value As String)

End Sub


'[方法] 检索 HTTP 响应标头。
Private Function WinHttpRequest_GetResponseHeader(ByVal Header As String) As String

End Function


'[方法] 检索所有 HTTP 响应标头。
Private Function WinHttpRequest_GetAllResponseHeaders() As String

End Function


'[方法] 将 HTTP 请求发送到 HTTP 服务器。
Private Sub WinHttpRequest_Send(Optional ByVal Body As Variant)

End Sub


'[方法] 等待异步发送方法完成（以秒为单位可选超时值）。
Private Function WinHttpRequest_WaitForResponse(Optional ByVal Timeout As Variant) As Boolean

End Function


'[方法] 中止 WinHTTPSend 方法。
Private Sub WinHttpRequest_Abort()

End Sub


'[方法] 指定发送/接收操作的各个超时组件（以毫秒为单位）。
Private Sub WinHttpRequest_SetTimeouts(ByVal ResolveTimeout As Long, ByVal ConnectTimeout As Long, ByVal SendTimeout As Long, ByVal ReceiveTimeout As Long)

End Sub


'[方法] 选择要发送到安全超文本传输协议的客户端证书， (HTTPS) 服务器。
Private Sub WinHttpRequest_SetClientCertificate(ByVal ClientCertificate As String)

End Sub


'[方法] 设置当前的自动登录策略。
Private Sub WinHttpRequest_SetAutoLogonPolicy(ByVal AutoLogonPolicy As WinHttp.WinHttpRequestAutoLogonPolicy)

End Sub


'[属性][只读] 从最后一个响应中检索 HTTP 状态代码。
Private Property Get WinHttpRequest_Status() As Long

End Property


'[属性][只读] 检索 HTTP 状态文本。
Private Property Get WinHttpRequest_StatusText() As String

End Property


'[属性][只读] 将响应实体正文检索为文本。
Private Property Get WinHttpRequest_ResponseText() As String

End Property


'[属性][只读] 将响应实体正文检索为无符号字节数组。
Private Property Get WinHttpRequest_ResponseBody() As Variant

End Property


'[属性][只读] 以 IStream 的形式检索响应实体正文。
Private Property Get WinHttpRequest_ResponseStream() As Variant

End Property


'[属性][读写] 设置或检索 Microsoft Windows HTTP Services (WinHTTP) 选项值。
Private Property Let WinHttpRequest_Option(ByVal Options As WinHttp.WinHttpRequestOption, ByVal RHS As Variant)

End Property
Private Property Get WinHttpRequest_Option(ByVal Options As WinHttp.WinHttpRequestOption) As Variant

End Property

