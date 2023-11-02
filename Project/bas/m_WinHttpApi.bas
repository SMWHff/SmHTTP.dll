Attribute VB_Name = "m_WinHttpAPI"
Option Explicit
Public Declare Function WinHttpCheckPlatform Lib "Winhttp.dll" () As Boolean
Public Declare Function WinHttpCrackUrl Lib "Winhttp.dll" (ByVal pwszUrl As String, ByVal dwUrlLength As Long, ByVal dwFlags As Long, ByRef lpUrlComponents As URL_COMPONENTS) As Boolean
Public Declare Function WinHttpOpen Lib "Winhttp.dll" (ByVal pwszUserAgent As Long, ByVal dwAccessType As Long, ByRef pwszProxyName As Any, ByVal pwszProxyBypass As Long, ByVal dwFlags As Long) As Long
Public Declare Function WinHttpConnect Lib "Winhttp.dll" (ByVal hSession As Long, ByRef pswzServerName As Any, ByVal nServerPort As Long, ByVal dwReserved As Long) As Long
Public Declare Function WinHttpOpenRequest Lib "Winhttp.dll" (ByVal hConnect As Long, ByRef pwszVerb As Any, ByRef pwszObjectName As Any, ByVal pwszVersion As Long, ByVal pwszReferrer As Long, ByVal ppwszAcceptTypes As Long, ByVal dwFlags As Long) As Long
Public Declare Function WinHttpCloseHandle Lib "Winhttp.dll" (ByVal hInternet As Long) As Long
Public Declare Function WinHttpSetTimeouts Lib "Winhttp.dll" (ByVal hInternet As Long, ByVal dwResolveTimeout As Long, ByVal dwConnectTimeout As Long, ByVal dwSendTimeout As Long, ByVal dwReceiveTimeout As Long) As Boolean
Public Declare Function WinHttpSetCredentials Lib "Winhttp.dll" (ByVal hRequest As Long, ByVal AuthTargets As Long, ByVal AuthScheme As Long, ByRef pwszUserName As Any, ByRef pwszPassword As Any, ByVal pAuthParams As Long) As Boolean
Public Declare Function WinHttpSetOption Lib "Winhttp.dll" (ByVal hInternet As Long, ByVal dwOption As Long, ByRef lpBuffer As Any, ByVal dwBufferLength As Long) As Boolean
Public Declare Function WinHttpAddRequestHeaders Lib "Winhttp.dll" (ByVal hRequest As Long, pwszHeaders As Any, ByVal dwHeadersLength As Long, ByVal dwModifiers As Long) As Boolean
Public Declare Function WinHttpSendRequest Lib "Winhttp.dll" (ByVal hRequest As Long, ByVal pwszHeaders As Long, ByVal dwHeadersLength As Long, ByRef lpOptional As Any, ByVal dwOptionalLength As Long, ByVal dwTotalLength As Long, ByVal dwContext As Long) As Boolean
Public Declare Function WinHttpReceiveResponse Lib "Winhttp.dll" (ByVal hRequest As Long, ByVal lpReserved As Long) As Boolean
Public Declare Function WinHttpQueryDataAvailable Lib "Winhttp.dll" (ByVal hRequest As Long, ByRef lpdwNumberOfBytesAvailable As Long) As Boolean
Public Declare Function WinHttpReadData Lib "Winhttp.dll" (ByVal hRequest As Long, ByRef lpBuffer As Any, ByVal dwNumberOfBytesToRead As Long, ByRef lpdwNumberOfBytesRead As Long) As Boolean
Public Declare Function WinHttpQueryHeaders Lib "Winhttp.dll" (ByVal hRequest As Long, ByVal dwInfoLevel As Long, ByVal pwszName As Long, ByRef lpBuffer As Any, ByRef lpdwBufferLength As Long, ByRef lpdwIndex As Long) As Boolean


Public Type URL_COMPONENTS
    dwStructSize        As Long         ' 本结构长度，注：60
    lpszScheme          As String * 128 ' 协议类型
    dwSchemeLength      As Long         ' 协议类型缓冲区长度
    nScheme             As Integer      ' 链接类型，1=http，2=https（INTERNET_SCHEME_HTTP=1，INTERNET_SCHEME_HTTPS=2）
    lpszHostName        As String * 128 ' 主机域名(Host)
    dwHostNameLength    As Long         ' 主机域名缓冲区长度
    nPort               As Integer      ' 端口
    lpszUserName        As String * 128 ' 帐号
    dwUserNameLength    As Long         ' 帐号缓冲区长度
    lpszPassword        As String * 128 ' 密码
    dwPasswordLength    As Long         ' 密码缓冲区长度
    lpszUrlPath         As String * 128 ' 路径(页面地址)
    dwUrlPathLength     As Long         ' 路径缓冲区长度
    lpszExtraInfo       As String * 128 ' 额外信息，比如“?”或“#”之后的参数字符串。
    dwExtraInfoLength   As Long         ' 额外信息长度
End Type
'
