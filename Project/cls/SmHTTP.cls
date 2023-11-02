VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SmHTTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'--------------------------------------------------------------------------
'   您可以在这里添加自己的插件函数，插件的制作和使用方法如下：
'   第一步：在下面添加插件函数，一个插件可以添加多个函数。
'   第二步：请修改下面的Get_Plugin_Description函数，加入插件和插件函数的说明信息，帮助信息会显示在按键精灵里方便使用
'   第三步：请修改Class Module（类模块）的名字，尽量用一些独特的名字避免和别人的插件名字冲突(默认是MyPluginName)
'   第四步：生成插件DLL！将DLL文件放到按键精灵的PLUGIN目录下，然后启动按键精灵测试您制作的插件功能是否正常
'   对插件有任何问题请发邮件给我们： hi@vrbrothers.com
'--------------------------------------------------------------------------

Private cls_WinHttp     As WinHttpRequest
Private cls_AutoParam   As Boolean
Private cls_Status      As Long
Private cls_Stream      As Variant
Private cls_Headers     As String
Private cls_Cookies     As String
Private cls_ContentType As String
Private cls_Boundary    As String

Private 调试开关 As Boolean





'========================================【插件基本信息】========================================
' 获取插件版本号
Public Function Ver() As String
Attribute Ver.VB_Description = "获取插件版本号"
    Ver = App.Major & "." & App.Minor & ".0." & App.Revision
End Function


'取插件路径
Public Function GetBasePath() As String
Attribute GetBasePath.VB_Description = "取插件路径"
    GetBasePath = App.Path & "\" & App.EXEName & ".dll"
End Function


'取对象ID
Public Function GetID() As String
Attribute GetID.VB_Description = "取对象ID"
    GetID = ObjPtr(Me)
End Function


'获取插件命令的最后错误
Public Function GetLastError() As String
Attribute GetLastError.VB_Description = "获取插件命令的最后错误"
    GetLastError = Err.Description
End Function


' 设置自动识别对应传入参数
Public Function SetAutoParamArray(ByVal Enable As Variant) As Long
Attribute SetAutoParamArray.VB_Description = "设置自动识别对应传入参数"
    G_AutoParam = Enable
    SetAutoParamArray = 1
    
    ' 百度统计代码
    'If G_Init = False Then
        G_Init = True
        On Error Resume Next
        Dim Baidu_Tongji_Http As WinHttpRequest
        Dim Args As String
        Args = "v=1.3.0&cc=1&ck=1&cl=" & Utils.GetBitsPerPel() & "-bit&ds=" & Utils.GetScreenXY() & "&ln=" & Utils.GetBrowserLanguage() & "&et=0&ep=0&fl=0&ja=" & Abs(Utils.IsJavaInstalled()) & "&u=" & Utils.EscapeURL("https://www.SmHTTP.com/") & "&tt=" & Utils.EscapeURL("神梦HTTP请求插件")
        'Debug.Print Args
        Set Baidu_Tongji_Http = CreateObject("WinHttp.WinHttpRequest.5.1")
        Baidu_Tongji_Http.Open "GET", "https://hm.baidu.com/hm.gif?si=212f0c3f981a440699efc710f3043368&" & Args, False
        Baidu_Tongji_Http.SetRequestHeader "Referer", "https://www.SmHTTP.com/"
        Baidu_Tongji_Http.SetRequestHeader "User-Agent", Utils.GetUserAgent()
        Baidu_Tongji_Http.SetRequestHeader "Cookie", "HMACCOUNT=" & Utils.GetHDDSN() & Utils.GetBaseBoard() & Utils.GetBIOS()
        Baidu_Tongji_Http.Send
        Set Baidu_Tongji_Http = Nothing
    'End If
End Function


' 构造协议头
Public Function Headers(ParamArray Args()) As String
Attribute Headers.VB_Description = "构造协议头"
    Dim i       As Long
    Dim Result  As Variant
    
    Result = Array()
    While (i < UBound(Args))
        ReDim Preserve Result(UBound(Result) + 1)
        Result(UBound(Result)) = Args(i) & ":" & Args(i + 1)
        i = i + 2
    Wend
    Headers = Join(Result, vbCrLf)
End Function


' 构造Cookie
Public Function Cookies(ParamArray Args()) As String
Attribute Cookies.VB_Description = "构造Cookie"
    Dim i       As Long
    Dim Result  As Variant
    
    Result = Array()
    While (i < UBound(Args))
        ReDim Preserve Result(UBound(Result) + 1)
        Result(UBound(Result)) = Args(i) & "=" & Args(i + 1)
        i = i + 2
    Wend
    Cookies = Join(Result, ";")
End Function
 
 
' 构造常用主体
Public Function Data(ParamArray Args()) As String
Attribute Data.VB_Description = "构造常用主体"
    Dim i       As Long
    Dim Result  As Variant
    
    cls_ContentType = "Content-Type: application/x-www-form-urlencoded"
    Result = Array()
    While (i < UBound(Args))
        ReDim Preserve Result(UBound(Result) + 1)
        Result(UBound(Result)) = EscapeURL(Args(i)) & "=" & EscapeURL(Args(i + 1))
        i = i + 2
    Wend
    Data = Join(Result, "&")
End Function


' 构造表单主体
'   --分隔符(boundary)[换行]
'   Content-Disposition: form-data; name="参数名"[换行]
'   [换行]
'   参数值 [换行]
'   --分隔符(boundary)[换行]
'   Content-Disposition: form-data; name="图片名"; filename="图片文件名"[换行]
'   Content-Type: 类型[换行]
'   [换行]
'   图片文件的二进制内容 [换行]
'   --分隔符(boundary)--
Public Function FormData(ParamArray Args()) As Variant
    Dim boundary    As String
    Dim i           As Long
    Dim head        As String
    Dim Data        As String
    Dim Body        As String
    Dim bin()       As Byte
    Dim key         As Variant
    Dim value       As Variant
    Dim ValueBuff() As Byte
    Dim fileName    As String
    Dim fileType    As String
    Dim Result      As Variant
    Dim IsBinary    As Boolean
    
    boundary = "WebKitFormBoundary" & IIf(App.LogMode = vbIDE, "SmHTTP", GetUnixTime()) & "SMWHff"
    cls_ContentType = "Content-Type:multipart/form-data; boundary=" & boundary & vbCrLf
    Result = Array()
    While (i < UBound(Args))
        ' Debug.Print "i < UBound(Args)=", i, UBound(Args)
        key = Args(i): i = i + 1
        value = Args(i): i = i + 1
        fileName = ""
        fileType = ""
        If Left(key, 1) = "@" Then
            IsBinary = True
            key = Mid(key, 2)
            If VarType(value) = vbArray + vbByte Then
                ValueBuff() = value
                fileName = key
            ElseIf IsFileExist(value) Then
                ValueBuff() = Utils.ReadBin(value)
                fileName = Mid(value, InStrRev(value, "\") + 1)
            End If
            fileType = "application/octet-stream"
            If InStr(Args(i), "/") > 0 Then fileType = Args(i): i = i + 1
        Else
            ValueBuff() = Utils.ToBin(value)
        End If
        ReDim Preserve Result(UBound(Result) + 1)
        Result(UBound(Result)) = Array(key, ValueBuff(), fileName, fileType)
    Wend
    
    head = "--" & boundary & vbCrLf
    For i = 0 To UBound(Result)
        key = Result(i)(0)
        ValueBuff() = Result(i)(1)
        fileName = Result(i)(2)
        fileType = Result(i)(3)
        Data = head & "Content-Disposition: form-data; name=""" & key & """"
        If Len(fileName) <> 0 Then
            Data = Data & "; filename=""" & fileName & """" & vbCrLf
        Else
            Data = Data & vbCrLf
        End If
        If Len(fileType) <> 0 Then
            Data = Data & "Content-Type: " & fileType & vbCrLf
        End If
        Data = Data & vbCrLf
        If IsBinary Then
            Call Utils.BinCat(bin(), Utils.ToBin(Data))
            Call Utils.BinCat(bin(), ValueBuff)
            Call Utils.BinCat(bin(), Utils.ToBin(vbCrLf))
        Else
            Body = Body & Data & StrConv(ValueBuff(), vbUnicode) & vbCrLf
        End If
    Next
    
    If IsBinary Then
        Result = Utils.BinCat(bin(), Utils.ToBin("--" & boundary & "--" & vbCrLf))
    Else
        Result = Body & "--" & boundary & "--" & vbCrLf
    End If
    
    FormData = Result
End Function


' 构造JSON主体
Public Function JsonData(ParamArray Args()) As String
Attribute JsonData.VB_Description = "构造JSON主体"
    Dim sc      As ScriptControl  '需要引用工程 Microsoft Script Control
    Dim i       As Long
    Dim key     As String
    Dim value   As Variant
    
    cls_ContentType = "Content-Type:application/json"
    Set sc = CreateObject("MSScriptControl.ScriptControl")
    sc.Language = "JScript"
    sc.AddCode LoadResData(101, 6) '载入js库 json2
    sc.AddCode "var $={}; function SmSET(str){Temp = str; return 'Temp';}"
    sc.AddCode "function JSArray(vbarray){Temp = new VBArray(vbarray).toArray(); return 'Temp';}"
    While (i < UBound(Args))
        key = Args(i)
        value = Args(i + 1)
        If IsArray(value) Then
            value = sc.Run("JSArray", value)
        ElseIf IsObject(value) Then
            value = """" & TypeName(value) & """"
        ElseIf VarType(value) = vbEmpty Then
            value = "''"
        ElseIf VarType(value) = vbArray + vbByte Then
            value = """字节集={...}"""
        ElseIf VarType(value) = vbString Then
            If Left(value, 1) = "{" And Right(value, 1) = "}" Then
                value = value
            ElseIf Left(value, 1) = "[" And Right(value, 1) = "]" Then
                value = value
            Else
                If Len(value) = 0 Then value = ""
                value = sc.Run("SmSET", value)
            End If
        ElseIf VarType(value) = vbBoolean Then
            value = IIf(value, "true", "false")
        ElseIf VarType(value) = vbDate Then
            value = """" & value & """"
        ElseIf IsNumeric(value) Then
            value = value
        Else
            value = "null"
        End If
        Call sc.Eval("$['" & key & "']=" & value)
        i = i + 2
    Wend
    JsonData = sc.Eval("JSON.stringify($)||'{}'")
    Set sc = Nothing
End Function


' 获取返回状态
Public Function GetStatus() As Long
Attribute GetStatus.VB_Description = "获取返回状态"
    If cls_WinHttp Is Nothing Then
        GetStatus = cls_Status
    Else
        GetStatus = cls_WinHttp.Status
    End If
End Function


' 获取响应的二进制数据
Public Function GetStream(Optional ByVal Path As String = "") As Byte()
    If cls_WinHttp Is Nothing Then
        GetStream = cls_Stream
    Else
        GetStream = cls_WinHttp.ResponseBody
    End If
    
    If Len(Path) > 0 And VarType(GetStream) = vbArray + vbByte Then
        Dim Handle As Long
        Handle = FreeFile() '获得文件的句柄
        Open Path For Binary As #Handle
            Put #Handle, , GetStream()
        Close #Handle
    End If
    cls_Stream = Empty
End Function


' 获取返回所有协议头
Public Function GetHeaders() As String
Attribute GetHeaders.VB_Description = "获取返回协议头"
    If cls_WinHttp Is Nothing Then
        GetHeaders = cls_Headers
    Else
        GetHeaders = cls_WinHttp.GetAllResponseHeaders
    End If
End Function


' 获取返回所有Cookies
Public Function GetCookies() As String
Attribute GetCookies.VB_Description = "获取返回Cookies"
    If cls_WinHttp Is Nothing Then
        GetCookies = cls_Cookies
    Else
        GetCookies = XHR.FetchCookies(cls_WinHttp.GetAllResponseHeaders)
    End If
End Function


' 获取指定协议头的值
Public Function GetHeaderByName(ByVal Headers As String, ByVal Name As String) As String
    Dim Ret     As String
    Dim Arr()   As String
    
    Const LIMIT = 2
    
    If Len(Name) > 0 Then
        Arr = Split(vbCrLf & Headers & vbCrLf, vbCrLf & Name & ": ", LIMIT, vbTextCompare)
        If UBound(Arr) = 1 Then
            Ret = Split(Arr(1), vbCrLf)(0)
        End If
    End If
    GetHeaderByName = Ret
End Function


' 获取指定Cookies的值
Public Function GetCookieByName(ByVal Cookies As String, ByVal Name As String) As String
    Dim Ret     As String
    Dim Arr()   As String
    
    Const SEMICOLON = "; "
    Const LIMIT = 2

    If Len(Name) > 0 Then
        Arr = Split(SEMICOLON & Cookies & SEMICOLON, SEMICOLON & Name & "=", LIMIT, vbTextCompare)
        If UBound(Arr) = 1 Then
            Ret = Split(Arr(1), SEMICOLON)(0)
        End If
    End If
    GetCookieByName = Ret
End Function


' 合并更新Cookies
Public Function UpdateCookies(ByVal oldCookies As String, ByVal NewCookies As String) As String
Attribute UpdateCookies.VB_Description = "合并更新Cookies"
    UpdateCookies = XHR.MergeUpdateCookies(oldCookies, NewCookies)
End Function


' JSON字符串转字典对象
Public Function JSON_toObject(ByVal sJSON As String) As Object
Attribute JSON_toObject.VB_Description = "JSON字符串转对象"
    Set JSON_toObject = JSON.parse(sJSON)
End Function


' 字典对象转JSON字符串
Public Function JSON_toString(ByVal oJSON As Object) As String
Attribute JSON_toString.VB_Description = "JSON对象转字符串"
    If Not oJSON Is Nothing Then
        JSON_toString = JSON.toString(oJSON)
    End If
End Function


' 解析JSON取值
Public Function GetJSON(ByVal sJSON As String, ByVal key As String) As Variant
Attribute GetJSON.VB_Description = "解析JSON取值"
    GetJSON = Utils.Fun_GetJSON(sJSON, key)
End Function



'========================================【同步请求函数】========================================
' Get请求
Public Function HTTP_GET(ByVal URL As String, Optional ByVal Headers As Variant = "", Optional ByVal Cookies As Variant = "", Optional ByVal Charset As Variant = "", Optional ByVal Timeout As Variant = 0, Optional ByVal Redirects As Variant = True, Optional ByVal Proxy As Variant = "", Optional ByVal ProxyUser As Variant = "", Optional ByVal ProxyPass As Variant = "", Optional ByVal AuthType As Variant = "", Optional ByVal AuthUser As Variant = "", Optional ByVal AuthPass As Variant = "") As Variant
Attribute HTTP_GET.VB_Description = "Get请求"
    If G_AutoParam Then Proxy = Gen_Proxies(Proxy, ProxyUser, ProxyPass): ProxyUser = "": ProxyPass = ""
    HTTP_GET = HTTP_Request("GET", URL, "", Headers, Cookies, Charset, Timeout, Redirects, Proxy, ProxyUser, ProxyPass, AuthType, AuthUser, AuthPass, "同步")
End Function

' Post请求
Public Function HTTP_POST(ByVal URL As String, Optional ByVal Data As Variant = "", Optional ByVal Headers As Variant = "", Optional ByVal Cookies As Variant = "", Optional ByVal Charset As Variant = "", Optional ByVal Timeout As Variant = 0, Optional ByVal Redirects As Variant = True, Optional ByVal Proxy As Variant = "", Optional ByVal ProxyUser As Variant = "", Optional ByVal ProxyPass As Variant = "", Optional ByVal AuthType As Variant = "", Optional ByVal AuthUser As Variant = "", Optional ByVal AuthPass As Variant = "") As Variant
Attribute HTTP_POST.VB_Description = "Post请求"
    If G_AutoParam Then Proxy = Gen_Proxies(Proxy, ProxyUser, ProxyPass): ProxyUser = "": ProxyPass = ""
    HTTP_POST = HTTP_Request("POST", URL, Data, Headers, Cookies, Charset, Timeout, Redirects, Proxy, ProxyUser, ProxyPass, AuthType, AuthUser, AuthPass, "同步")
End Function

' Head请求
Public Function HTTP_HEAD(ByVal URL As String, Optional ByVal Headers As Variant = "", Optional ByVal Cookies As Variant = "", Optional ByVal Charset As Variant = "", Optional ByVal Timeout As Variant = 0, Optional ByVal Redirects As Variant = True, Optional ByVal Proxy As Variant = "", Optional ByVal ProxyUser As Variant = "", Optional ByVal ProxyPass As Variant = "", Optional ByVal AuthType As Variant = "", Optional ByVal AuthUser As Variant = "", Optional ByVal AuthPass As Variant = "") As Variant
Attribute HTTP_HEAD.VB_Description = "Head请求"
    If G_AutoParam Then Proxy = Gen_Proxies(Proxy, ProxyUser, ProxyPass): ProxyUser = "": ProxyPass = ""
    HTTP_HEAD = HTTP_Request("HEAD", URL, "", Headers, Cookies, Charset, Timeout, Redirects, Proxy, ProxyUser, ProxyPass, AuthType, AuthUser, AuthPass, "同步")
End Function

' Options请求
Public Function HTTP_OPTIONS(ByVal URL As String, Optional ByVal Headers As Variant = "", Optional ByVal Cookies As Variant = "", Optional ByVal Charset As Variant = "", Optional ByVal Timeout As Variant = 0, Optional ByVal Redirects As Variant = True, Optional ByVal Proxy As Variant = "", Optional ByVal ProxyUser As Variant = "", Optional ByVal ProxyPass As Variant = "", Optional ByVal AuthType As Variant = "", Optional ByVal AuthUser As Variant = "", Optional ByVal AuthPass As Variant = "") As Variant
Attribute HTTP_OPTIONS.VB_Description = "Options请求"
    If G_AutoParam Then Proxy = Gen_Proxies(Proxy, ProxyUser, ProxyPass): ProxyUser = "": ProxyPass = ""
    HTTP_OPTIONS = HTTP_Request("OPTIONS", URL, "", Headers, Cookies, Charset, Timeout, Redirects, Proxy, ProxyUser, ProxyPass, AuthType, AuthUser, AuthPass, "同步")
End Function

' Patch请求
Public Function HTTP_PATCH(ByVal URL As String, Optional ByVal Data As Variant = "", Optional ByVal Headers As Variant = "", Optional ByVal Cookies As Variant = "", Optional ByVal Charset As Variant = "", Optional ByVal Timeout As Variant = 0, Optional ByVal Redirects As Variant = True, Optional ByVal Proxy As Variant = "", Optional ByVal ProxyUser As Variant = "", Optional ByVal ProxyPass As Variant = "", Optional ByVal AuthType As Variant = "", Optional ByVal AuthUser As Variant = "", Optional ByVal AuthPass As Variant = "") As Variant
Attribute HTTP_PATCH.VB_Description = "Patch请求"
    If G_AutoParam Then Proxy = Gen_Proxies(Proxy, ProxyUser, ProxyPass): ProxyUser = "": ProxyPass = ""
    HTTP_PATCH = HTTP_Request("PATCH", URL, Data, Headers, Cookies, Charset, Timeout, Redirects, Proxy, ProxyUser, ProxyPass, AuthType, AuthUser, AuthPass, "同步")
End Function

' Put请求
Public Function HTTP_PUT(ByVal URL As String, Optional ByVal Data As Variant = "", Optional ByVal Headers As Variant = "", Optional ByVal Cookies As Variant = "", Optional ByVal Charset As Variant = "", Optional ByVal Timeout As Variant = 0, Optional ByVal Redirects As Variant = True, Optional ByVal Proxy As Variant = "", Optional ByVal ProxyUser As Variant = "", Optional ByVal ProxyPass As Variant = "", Optional ByVal AuthType As Variant = "", Optional ByVal AuthUser As Variant = "", Optional ByVal AuthPass As Variant = "") As Variant
Attribute HTTP_PUT.VB_Description = "Put请求"
    If G_AutoParam Then Proxy = Gen_Proxies(Proxy, ProxyUser, ProxyPass): ProxyUser = "": ProxyPass = ""
    HTTP_PUT = HTTP_Request("PUT", URL, Data, Headers, Cookies, Charset, Timeout, Redirects, Proxy, ProxyUser, ProxyPass, AuthType, AuthUser, AuthPass, "同步")
End Function

' Delete请求
Public Function HTTP_DELETE(ByVal URL As String, Optional ByVal Data As Variant = "", Optional ByVal Headers As Variant = "", Optional ByVal Cookies As Variant = "", Optional ByVal Charset As Variant = "", Optional ByVal Timeout As Variant = 0, Optional ByVal Redirects As Variant = True, Optional ByVal Proxy As Variant = "", Optional ByVal ProxyUser As Variant = "", Optional ByVal ProxyPass As Variant = "", Optional ByVal AuthType As Variant = "", Optional ByVal AuthUser As Variant = "", Optional ByVal AuthPass As Variant = "") As Variant
Attribute HTTP_DELETE.VB_Description = "Delete请求"
    If G_AutoParam Then Proxy = Gen_Proxies(Proxy, ProxyUser, ProxyPass): ProxyUser = "": ProxyPass = ""
    HTTP_DELETE = HTTP_Request("DELETE", URL, Data, Headers, Cookies, Charset, Timeout, Redirects, Proxy, ProxyUser, ProxyPass, AuthType, AuthUser, AuthPass, "同步")
End Function

' Trace请求
Public Function HTTP_TRACE(ByVal URL As String, Optional ByVal Headers As Variant = "", Optional ByVal Cookies As Variant = "", Optional ByVal Charset As Variant = "", Optional ByVal Timeout As Variant = 0, Optional ByVal Redirects As Variant = True, Optional ByVal Proxy As Variant = "", Optional ByVal ProxyUser As Variant = "", Optional ByVal ProxyPass As Variant = "", Optional ByVal AuthType As Variant = "", Optional ByVal AuthUser As Variant = "", Optional ByVal AuthPass As Variant = "") As Variant
    If G_AutoParam Then Proxy = Gen_Proxies(Proxy, ProxyUser, ProxyPass): ProxyUser = "": ProxyPass = ""
    HTTP_TRACE = HTTP_Request("TRACE", URL, "", Headers, Cookies, Charset, Timeout, Redirects, Proxy, ProxyUser, ProxyPass, AuthType, AuthUser, AuthPass, "同步")
End Function


' HTTP请求
Public Function HTTP_Request(ParamArray Args()) As Variant
Attribute HTTP_Request.VB_Description = "HTTP请求"
Attribute HTTP_Request.VB_UserMemId = 0
    Dim Ret         As Variant
    Dim Har         As String
    Dim i           As Long
    Dim v           As Variant
    Dim Method      As Variant
    Dim URL         As Variant
    Dim Data        As Variant
    Dim Headers     As Variant
    Dim Cookies     As Variant
    Dim Charset     As Variant
    Dim Timeout     As Variant
    Dim Redirects   As Variant
    Dim Proxies     As Variant
    Dim Proxy       As Variant
    Dim ProxyUser   As Variant
    Dim ProxyPass   As Variant
    Dim UserPass    As Variant
    Dim AuthType    As Variant
    Dim AuthType_i  As Variant
    Dim AuthUser    As Variant
    Dim AuthPass    As Variant
    Dim SyncAsync   As Variant
    Dim CompHeaders As Variant
    Dim CompCookies As Variant
    
    
    If UBound(Args) = -1 Then
        HTTP_Request = "SmHTTP"
        Exit Function
    ElseIf UBound(Args) = 0 Then
        ' {
        '   "Request":{
        '       "url": "",
        '       "method": "",
        '       "charset": "UTF-8",
        '       "cookies": {},
        '       "headers": {},
        '       "queries": {},
        '       "data": {},
        '       "json": {},
        '   },
        '   "Response":{
        '       "status": 200,
        '       "statusText": "OK",
        '       "cookies": "",
        '       "headers": "",
        '       "text": "",
        '       "body": {},
        '       "json": {},
        '   }
        ' }
        Har = Args(0)
        If VarType(Har) = vbString Then
            If Left(Har, 1) = "{" And Right(Har, 1) = "}" Then
                Dim sc As ScriptControl
                Set sc = CreateObject("MSScriptControl.ScriptControl")
                sc.Language = "JScript"
                sc.AddCode "var $ = " & Har & ";"
                Method = sc.Eval("$.method|''")
                URL = sc.Eval("$.url|''")
                Data = sc.Eval("$.data|''")
                Headers = sc.Eval("$.headers|''")
                Cookies = sc.Eval("$.cookies|''")
                Charset = sc.Eval("$.charset|''")
                Timeout = sc.Eval("$.timeout|30")
                Redirects = sc.Eval("$.redirects|true")
                Proxy = sc.Eval("$.proxy|''")
                ProxyUser = sc.Eval("$.proxyUser|''")
                ProxyPass = sc.Eval("$.proxyPass|''")
                AuthUser = sc.Eval("$.authUser|''")
                AuthPass = sc.Eval("$.authPass|''")
                AuthType = sc.Eval("$.authType|''")
                SyncAsync = (sc.Eval("$.syncAsync|'同步'") = "同步")
                sc.Reset
                Set sc = Nothing
                If G_AutoParam Then
                    Proxies = Array()
                    If Utils.ArgumentsMatch(Proxy, vbString, "^([^:]+:[^@]+)?@?(\d+\.\d+\.\d+\.\d+)(:\d{1,6})?$", Result:=Proxies) Then
                        ' 代理地址格式【用户名:密码@IP:端口号】
                        If Len(Proxies(0)) <> 0 Then
                            UserPass = Split(Proxies(0), ":")
                            If UBound(UserPass) >= 0 Then AuthUser = UserPass(0)
                            If UBound(UserPass) >= 1 Then AuthPass = UserPass(1)
                        End If
                        Proxy = Proxies(1) & Proxies(2)
                    End If
                End If
            End If
        End If
    ElseIf G_AutoParam Then
        Proxies = Array()
        While i <= UBound(Args)
            v = Args(i)
            Select Case True
            Case Utils.ArgumentsMatch(v, vbString, "^(GET|POST|HEAD|PUT|OPTIONS|DELETE|TRACE|CONNECT|PATCH)$", SavaArg:=Method)
                Method = v
            Case Utils.ArgumentsMatch(v, vbString, "^(http|https)://.*$", SavaArg:=URL)
                URL = v
            Case Utils.ArgumentsMatch(v, vbString, "^(ANSI|ASCII|UTF-8|UTF-16|Unicode|GB\d+|GBK|BIG5|Byte\(\))$", SavaArg:=Charset)
                Charset = v
            Case Utils.ArgumentsMatch(v, vbDecimal, "x >= -1", SavaArg:=Timeout)
                Timeout = v
            Case Utils.ArgumentsMatch(v, vbBoolean, SavaArg:=Redirects)
                Redirects = v
            Case Utils.ArgumentsMatch(v, vbString, "^([^:]+:[^@]+)?@?(\d+\.\d+\.\d+\.\d+)(:\d{1,6})?$", SavaArg:=Proxy, Result:=Proxies)
                ' 代理地址格式【用户名:密码@IP:端口号】
                If Len(Proxies(0)) <> 0 Then
                    UserPass = Split(Proxies(0), ":")
                    If UBound(UserPass) >= 0 Then ProxyUser = UserPass(0)
                    If UBound(UserPass) >= 1 Then ProxyPass = UserPass(1)
                End If
                Proxy = Proxies(1) & Proxies(2)
            Case Utils.ArgumentsMatch(v, vbString, "^(BASIC|DIGEST|FORMBASE)$", SavaArg:=AuthType)
                AuthType = v
                If i + 1 <= UBound(Args) Then AuthUser = Args(i + 1)
                If i + 2 <= UBound(Args) Then AuthPass = Args(i + 2)
                i = i + 2
            Case Utils.ArgumentsMatch(v, vbString, "^(同步|异步)$", SavaArg:=SyncAsync)
                SyncAsync = (v = "异步")
            Case Utils.ArgumentsMatch(v, vbString, "^(开启|关闭)补全$", SavaArg:=CompHeaders)
                CompHeaders = (v <> "关闭补全")
            Case Utils.ArgumentsMatch(v, vbString, "^(开启|关闭)合并$", SavaArg:=CompCookies)
                CompCookies = (v <> "关闭合并")
            Case Utils.ArgumentsMatch(v, vbString, "^([^=\r\n]+=[^;\r\n]+;)+[^=\r\n]+=[^;\r\n]+[;\s]*$", SavaArg:=Cookies)
                Cookies = v
            Case Utils.ArgumentsMatch(v, vbString, "(Accept|Accept-Encoding|Accept-Language|Cache-Control|Connection|Content-Type|Content-Length|Cookie|Host|Referer|Origin|User-Agent|X-Requested-With):.*", SavaArg:=Headers)
                Headers = v
            Case Utils.ArgumentsMatch(v, vbVariant, SavaArg:=Data)
                Data = v
            End Select
            i = i + 1
        Wend
    Else
        If UBound(Args) >= 0 Then Method = Args(0)
        If UBound(Args) >= 1 Then URL = Args(1)
        If UBound(Args) >= 2 Then Data = Args(2)
        If UBound(Args) >= 3 Then Headers = Args(3)
        If UBound(Args) >= 4 Then Cookies = Args(4)
        If UBound(Args) >= 5 Then Charset = Args(5)
        If UBound(Args) >= 6 Then Timeout = Args(6)
        If UBound(Args) >= 7 Then Redirects = Args(7)
        If UBound(Args) >= 8 Then Proxy = Args(8)
        If UBound(Args) >= 9 Then ProxyUser = Args(9)
        If UBound(Args) >= 10 Then ProxyPass = Args(10)
        If UBound(Args) >= 11 Then AuthType = Args(11)
        If UBound(Args) >= 12 Then AuthUser = Args(12)
        If UBound(Args) >= 13 Then AuthPass = Args(13)
        If UBound(Args) >= 14 Then SyncAsync = IIf(VarType(Args(14)) = vbBoolean Or IsNumeric(Args(14)), Args(14) <> 0, Args(14) = "异步")
        If UBound(Args) >= 15 Then CompHeaders = IIf(VarType(Args(15)) = vbBoolean Or IsNumeric(Args(15)), Args(15) <> 0, Args(15) <> "关闭补全")
        If UBound(Args) >= 16 Then CompCookies = IIf(VarType(Args(16)) = vbBoolean Or IsNumeric(Args(16)), Args(16) <> 0, Args(16) <> "关闭合并")
    End If
    
    ' 构造主体用的数据类型
    If InStr(1, Headers, "Content-Type:", 1) = 0 Then
        Headers = Headers & vbCrLf & cls_ContentType
    End If
    
    ' 设置参数默认值
    If IsEmpty(Timeout) Then Timeout = 30
    If IsEmpty(Redirects) Then Redirects = True
    If IsEmpty(SyncAsync) Then SyncAsync = False
    If IsEmpty(CompHeaders) Then CompHeaders = True
    If IsEmpty(CompCookies) Then CompCookies = True
    
    
    If 调试开关 Then
        Debug.Print "Method：", Method
        Debug.Print "URL：", URL
        Debug.Print "Data：", IIf(Len(Data) > 1024, TypeName(Data) & "(" & Len(Data) & ")", Data)
        Debug.Print "Headers：", Headers
        Debug.Print "Cookies：", Cookies
        Debug.Print "Charset：", Charset
        Debug.Print "Timeout：", Timeout
        Debug.Print "Redirects：", Redirects
        Debug.Print "Proxy：", Proxy
        Debug.Print "ProxyUser：", ProxyUser
        Debug.Print "ProxyPass：", ProxyPass
        Debug.Print "AuthType：", AuthType
        Debug.Print "AuthUser：", AuthUser
        Debug.Print "AuthPass：", AuthPass
        Debug.Print "SyncAsync：", SyncAsync
        Debug.Print "CompHeaders：", CompHeaders
        Debug.Print "CompCookies：", CompCookies
    End If
    
    HTTP_Request = XHR.WinHtpRequest(this:=cls_WinHttp, _
        Method:=Method, _
        URL:=URL, _
        Data:=Data, _
        Headers:=Headers, _
        Cookies:=Cookies, _
        Charset:=Charset, _
        Timeout:=Timeout, _
        Redirects:=Redirects, _
        Proxy:=Proxy, _
        ProxyUser:=ProxyUser, _
        ProxyPass:=ProxyPass, _
        AuthType:=AuthType, _
        AuthUser:=AuthUser, _
        AuthPass:=AuthPass, _
        IsAsync:=SyncAsync, _
        CompleteHeaders:=CompHeaders, _
        CompleteCookies:=CompCookies, _
        Res_Status:=cls_Status, _
        Res_Headers:=cls_Headers, _
        Res_Cookies:=cls_Cookies, _
        Res_Body:=cls_Stream _
    )
End Function



'========================================【实际应用函数】========================================
' HTTP下载
Public Function HTTP_DownLoad(ByVal URL As String, ByVal SavePath As String) As Long
    Dim ResBody As Variant
    
    ' 判断在 IDE 中
    If App.LogMode = 0 Then On Error Resume Next
    
    ResBody = XHR.WinHtpRequest(this:=Nothing, Method:="GET", URL:=URL, Charset:="Byte()")
    HTTP_DownLoad = SaveFile(SavePath, ResBody)
    On Error GoTo 0
    cls_Stream = Empty
End Function


' HTTP下载加强版
Private Function HTTP_DownLoadEx(ByVal URLs As String, Optional ByVal SavePaths As String, Optional ByVal Method As String = "GET", Optional ByVal Data As Variant = "", Optional ByVal Headers As Variant = "", Optional ByVal Cookies As Variant = "", Optional ByVal Charset As Variant = "", Optional ByVal Timeout As Variant = 30, Optional ByVal Redirects As Variant = True, Optional ByVal Proxy As Variant = "", Optional ByVal ProxyUser As Variant = "", Optional ByVal ProxyPass As Variant = "", Optional ByVal AuthType As Variant = "", Optional ByVal AuthUser As Variant = "", Optional ByVal AuthPass As Variant = "") As Variant
    Dim http()      As Object
    Dim URLArr()    As String
    Dim PathArr()   As String
    Dim i           As Long
    Dim n           As Long
    Dim t           As Long
    Dim errRet      As String
    Dim Ret()       As Variant
    
    If App.LogMode = 0 Then On Error Resume Next
    
    URLArr = Split(URLs, "|")
    PathArr = Split(SavePaths, "|")
    
    If UBound(URLArr) < UBound(PathArr) Then
        ReDim Preserve URLArr(UBound(PathArr))
    Else
        ReDim Preserve PathArr(UBound(URLArr))
    End If
    
    ' 批量创建 HTTP 对象
    ReDim http(UBound(URLArr))
    For i = 0 To UBound(URLArr)
        Set http(i) = New WinHttpRequest
        Call XHR.WinHtpRequest(this:=http(i), Method:=Method, URL:=URLArr(i), Data:=Data, Headers:=Headers, Cookies:=Cookies, Charset:="Byte()", Redirects:=Redirects, Proxy:=Proxy, ProxyUser:=ProxyUser, AuthType:=AuthType, AuthUser:=AuthUser, AuthPass:=AuthPass, IsAsync:=True)
    Next
    ' 异步等待响应内容
    ReDim Ret(UBound(URLArr))
    t = Timer(): n = 0
    Do
        For i = 0 To UBound(URLArr)
            If http(i).WaitForResponse(1) Then
                If http(i).Status = 200 And Len(Ret(i)) = 0 Then
                    n = n + 1
                    Ret(i) = SaveFile(PathArr(i), http(i).ResponseBody)
                    If Ret(i) = 0 Then
                        errRet = errRet & URLArr(i) & "," & PathArr(i) & "|"
                    End If
                End If
            End If
        Next
        DoEvents
    Loop Until Timer() - t > Timeout And Timeout > -1 Or n = UBound(URLArr) + 1
    ' 批量销毁 HTTP 对象
    For i = 0 To UBound(URLArr)
        Set http(i) = Nothing
    Next
    On Error GoTo 0
    If Len(errRet) > 0 Then errRet = Left(errRet, Len(errRet) - 1)
    HTTP_DownLoadEx = errRet
End Function





'========================================【插件帮助信息】========================================

'下面这个函数为您的插件提供帮助信息，请只修改和添加里面的描述信息，而不要修改函数本身
Public Function Get_Plugin_Description(ByVal ItemName As String) As String
    Dim Description_Text As String
    Description_Text = ""
    Select Case ItemName
    Case ""
        Description_Text = "神梦HTTP请求插件v" & App.Major & "." & App.Minor & ".0." & App.Revision & " 【QQ群：624655641】"
    Case "Ver"
        Description_Text = "取插件版本()【版本号$】"
    Case "GetBasePath"
        Description_Text = "取插件路径()【返回路径$】"
    Case "GetID"
        Description_Text = "取对象ID()【对象ID$】"
    Case "GetLastError"
        Description_Text = "获取插件命令的最后错误()【错误内容$】"
    Case "SetAutoParamArray"
        Description_Text = "设置自动识别对应传入参数(是否开启|)【是否成功|】"
    Case "Headers"
        Description_Text = "构造协议头(参数列表)【协议头$】"
    Case "Cookies"
        Description_Text = "构造Cookies(参数列表)【Cookies$】"
    Case "Data"
        Description_Text = "构造URL主体(参数列表)【URL主体$】"
    Case "FormData"
        Description_Text = "构造表单主体(参数列表)【表单主体$】"
    Case "JsonData"
        Description_Text = "构造JSON主体(参数列表)【JSON主体$】"
    Case "GetStatus"
        Description_Text = "获取返回状态()【状态&】"
    Case "GetStream"
        Description_Text = "获取返回的二进制内容([保存路径$])【返回字节集()】"
    Case "GetHeaders"
        Description_Text = "获取返回所有协议头()【返回协议头$】"
    Case "GetCookies"
        Description_Text = "获取返回所有Cookies()【返回Cookies$】"
    Case "GetHeaderByName"
        Description_Text = "获取指定协议头的值(所有协议头$, 键名$)【返回键值$】"
    Case "GetCookieByName"
        Description_Text = "获取指定协议头的值(所有Cookies$, 键名$)【返回Cookie$】"
    Case "UpdateCookies"
        Description_Text = "合并更新Cookies(旧Cookies$, 新Cookies$)【合并Cookies$】"
    Case "JSON_toObject"
        Description_Text = "JSON字符串转字典对象(JSON字符串$)【字典对象.】"
    Case "JSON_toString"
        Description_Text = "字典对象转JSON字符串(字典对象.)【JSON字符串$】"
    Case "GetJSON"
        Description_Text = "解析JSON取值(JSON字符串$, 键名$)【值】"
    Case "HTTP_GET"
        Description_Text = "Get请求(网址$, [请求协议头=〓一行一个用换行符隔开$], [请求Cookies$],［网页编码=ANSI|UTF-8|UTF-16|UNICODE|GB2312|GBK|GB18030|BIG5|ASCII|Byte()〓返回字节集$], [超时秒数=30|-1〓无限等待&], [是否重定向=True〓开启|False〓禁止|], [代理IP=〓格式为 1.1.1.1:80$], [代理用户名$], [代理密码$], [认证类型=BASIC〓基本认证|DIGEST〓摘要认证$], [认证账号$], [认证密码$])【网页源码$】"
    Case "HTTP_POST"
        Description_Text = "Post请求(网址$, [请求主体], [请求协议头=〓一行一个用换行符隔开$], [请求Cookies$],［网页编码=ANSI|UTF-8|UTF-16|UNICODE|GB2312|GBK|GB18030|BIG5|ASCII|Byte()〓返回字节集$], [超时秒数=30|-1〓无限等待&], [是否重定向=True〓开启|False〓禁止|], [代理IP=〓格式为 1.1.1.1:80$], [代理用户名$], [代理密码$], [认证类型=BASIC〓基本认证|DIGEST〓摘要认证$], [认证账号$], [认证密码$])"
    Case "HTTP_HEAD"
        Description_Text = "Head请求(网址$, [请求协议头=〓一行一个用换行符隔开$], [请求Cookies$],［网页编码=ANSI|UTF-8|UTF-16|UNICODE|GB2312|GBK|GB18030|BIG5|ASCII|Byte()〓返回字节集$], [超时秒数=30|-1〓无限等待&], [是否重定向=True〓开启|False〓禁止|], [代理IP=〓格式为 1.1.1.1:80$], [代理用户名$], [代理密码$], [认证类型=BASIC〓基本认证|DIGEST〓摘要认证$], [认证账号$], [认证密码$])"
    Case "HTTP_OPTIONS"
        Description_Text = "Options请求(网址$, [请求协议头=〓一行一个用换行符隔开$], [请求Cookies$],［网页编码=ANSI|UTF-8|UTF-16|UNICODE|GB2312|GBK|GB18030|BIG5|ASCII|Byte()〓返回字节集$], [超时秒数=30|-1〓无限等待&], [是否重定向=True〓开启|False〓禁止|], [代理IP=〓格式为 1.1.1.1:80$], [代理用户名$], [代理密码$], [认证类型=BASIC〓基本认证|DIGEST〓摘要认证$], [认证账号$], [认证密码$])【网页源码$】"
    Case "HTTP_PATCH"
        Description_Text = "Patch请求(网址$, [请求主体], [请求协议头=〓一行一个用换行符隔开$], [请求Cookies$],［网页编码=ANSI|UTF-8|UTF-16|UNICODE|GB2312|GBK|GB18030|BIG5|ASCII|Byte()〓返回字节集$], [超时秒数=30|-1〓无限等待&], [是否重定向=True〓开启|False〓禁止|], [代理IP=〓格式为 1.1.1.1:80$], [代理用户名$], [代理密码$], [认证类型=BASIC〓基本认证|DIGEST〓摘要认证$], [认证账号$], [认证密码$])【网页源码$】"
    Case "HTTP_PUT"
        Description_Text = "Put请求(网址$, [请求主体], [请求协议头=〓一行一个用换行符隔开$], [请求Cookies$],［网页编码=ANSI|UTF-8|UTF-16|UNICODE|GB2312|GBK|GB18030|BIG5|ASCII|Byte()〓返回字节集$], [超时秒数=30|-1〓无限等待&], [是否重定向=True〓开启|False〓禁止|], [代理IP=〓格式为 1.1.1.1:80$], [代理用户名$], [代理密码$], [认证类型=BASIC〓基本认证|DIGEST〓摘要认证$], [认证账号$], [认证密码$])【网页源码$】"
    Case "HTTP_DELETE"
        Description_Text = "Delete请求(网址$, [请求主体], [请求协议头=〓一行一个用换行符隔开$], [请求Cookies$],［网页编码=ANSI|UTF-8|UTF-16|UNICODE|GB2312|GBK|GB18030|BIG5|ASCII|Byte()〓返回字节集$], [超时秒数=30|-1〓无限等待&], [是否重定向=True〓开启|False〓禁止|], [代理IP=〓格式为 1.1.1.1:80$], [代理用户名$], [代理密码$], [认证类型=BASIC〓基本认证|DIGEST〓摘要认证$], [认证账号$], [认证密码$])【网页源码$】"
    Case "HTTP_TRACE"
        Description_Text = "Trace请求(网址$, [请求协议头=〓一行一个用换行符隔开$], [请求Cookies$],［网页编码=ANSI|UTF-8|UTF-16|UNICODE|GB2312|GBK|GB18030|BIG5|ASCII|Byte()〓返回字节集$], [超时秒数=30|-1〓无限等待&], [是否重定向=True〓开启|False〓禁止|], [代理IP=〓格式为 1.1.1.1:80$], [代理用户名$], [代理密码$], [认证类型=BASIC〓基本认证|DIGEST〓摘要认证$], [认证账号$], [认证密码$])【网页源码$】"
    Case "HTTP_Request"
        Description_Text = "HTTP请求(网址$, [访问方法=GET|POST|HEAD|PUT|OPTIONS|DELETE|TRACE|CONNECT|PATCH$], [请求主体], [请求协议头=〓一行一个用换行符隔开$], [请求Cookies$],［网页编码=ANSI|UTF-8|UTF-16|UNICODE|GB2312|GBK|GB18030|BIG5|ASCII|Byte()〓返回字节集$], [超时秒数=30|-1〓无限等待&], [是否重定向=True〓开启|False〓禁止|], [代理地址=〓格式为 user:pwd@1.1.1.1:80$], [认证类型=BASIC〓基本认证|DIGEST〓摘要认证$], [认证账号$], [认证密码$], [等待方式=同步|异步$], [自动补全协议头=开启补全|关闭补全$], [合并更新Cookies=开启合并|关闭合并$])【网页源码$】"
    Case "HTTP_DownLoad"
        Description_Text = "HTTP下载(下载地址$, 保存路径$)【是否成功|】"
    
    '------------------------------------------------
    '为您的每个插件函数建立一个Case，就可以在按键精灵里显示插件的使用说明
    '------------------------------------------------
    End Select
    Get_Plugin_Description = Translate_Description(Description_Text)
End Function



'下面这个函数为您的插件提供显示模板信息，请只修改和添加里面的描述信息，而不要修改函数本身
'描述信息中的$1 $2 ... 等内容再显示时将用第一个参数、第二个参数、... 代替
Public Function Get_Plugin_Interpret_Template(ByVal ItemName As String) As String
    Dim Description_Text As String
    Description_Text = ""
    Select Case ItemName
    Case ""
        Description_Text = "HTTP请求"
    Case "Ver"
        Description_Text = "取插件版本"
    Case "GetBasePath"
        Description_Text = "取插件路径"
    Case "GetID"
        Description_Text = "取对象ID"
    Case "GetLastError"
        Description_Text = "获取插件命令的最后错误"
    Case "SetAutoParamArray"
        Description_Text = "设置自动识别对应传入参数 $1"
    Case "Headers"
        Description_Text = "构造协议头 $1: $2" & vbCrLf & "$3: $4"
    Case "Cookies"
        Description_Text = "构造Cookie $1=$2;$3=$4;"
    Case "Data"
        Description_Text = "构造常用主体 $1=$2&$3=$4"
    Case "FormData"
        Description_Text = "构造表单主体 $1"
    Case "JsonData"
        Description_Text = "构造JSON主体 {""$1"":""$2"", ""$3"":""$4""}"
    Case "GetStatus"
        Description_Text = "获取返回状态"
    Case "GetStream"
        Description_Text = "获取返回的二进制内容，保存到路径 $1"
    Case "GetHeaders"
        Description_Text = "获取返回所有协议头"
    Case "GetCookies"
        Description_Text = "获取返回所有Cookies"
    Case "GetHeaderByName"
        Description_Text = "从协议头 $1 中获取 $2 的值"
    Case "GetCookieByName"
        Description_Text = "从Cookies $1 中获取 $2 的值"
    Case "UpdateCookies"
        Description_Text = "将 $1 和 $2 两份 Cookies 进行合并更新"
    Case "JSON_toObject"
        Description_Text = "将JSON字符串 $1 转为字典对象"
    Case "JSON_toString"
        Description_Text = "将字典对象 $1 转为JSON字符串"
    Case "GetJSON"
        Description_Text = "解析JSON字符串 $1 并获取键名 $2 的值"
    Case "HTTP_GET"
        Description_Text = "用 Get 方法访问 $1 并带上请求头 $2 和请求Cookes $3 重定向为 $6 通过代理IP $7@$8:$9 通过 $10 身份认证 $11:12 在 $5 秒内返回 $4 编码格式的网页内容"
    Case "HTTP_POST"
        Description_Text = "用 Post 方法请求 $1 发送 $2 并带上请求头 $3 和请求Cookes $4 重定向为 $7 通过代理IP $8@$9:$10 通过 $11 身份认证 $12:13 在 $6 秒内返回 $5 编码格式的网页内容"
    Case "HTTP_HEAD"
        Description_Text = "用 Head 方法请求 $1 并带上请求头 $2 和请求Cookes $3 重定向为 $6 通过代理IP $7@$8:$9 通过 $10 身份认证 $11:12 在 $5 秒内返回协议头内容"
    Case "HTTP_OPTIONS"
        Description_Text = "用 Options 方法请求 $1 并带上请求头 $2 和请求Cookes $3 重定向为 $6 通过代理IP $7@$8:$9 通过 $10 身份认证 $11:12 在 $5 秒内返回 $4 编码格式的网页内容"
    Case "HTTP_PATCH"
        Description_Text = "用 Patch 方法请求 $1 发送 $2 并带上请求头 $3 和请求Cookes $4 重定向为 $7 通过代理IP $8@$9:$10 通过 $11 身份认证 $12:13 在 $6 秒内返回 $5 编码格式的网页内容"
    Case "HTTP_PUT"
        Description_Text = "用 Put 方法请求 $1 发送 $2 并带上请求头 $3 和请求Cookes $4 重定向为 $7 通过代理IP $8@$9:$10 通过 $11 身份认证 $12:13 在 $6 秒内返回 $5 编码格式的网页内容"
    Case "HTTP_DELETE"
        Description_Text = "用 Delete 方法请求 $1 发送 $2 并带上请求头 $3 和请求Cookes $4 重定向为 $7 通过代理IP $8@$9:$10 通过 $11 身份认证 $12:13 在 $6 秒内返回 $5 编码格式的网页内容"
    Case "HTTP_TRACE"
        Description_Text = "用 Trace 方法请求 $1 发送 $2 并带上请求头 $3 和请求Cookes $4 重定向为 $7 通过代理IP $8@$9:$10 通过 $11 身份认证 $12:13 在 $6 秒内返回 $5 编码格式的请求提交内容"
    Case "HTTP_Request"
        Description_Text = "用 $1 方法请求 $2 发送 $3 并带上请求头 $4 和请求Cookes $5 重定向为 $8 通过代理IP $9@$10$11 通过 $12 身份认证 $13:$14 在 $7 秒内返回 $6 编码格式的网页内容"
    Case "HTTP_DownLoad"
        Description_Text = "下载 $1 并保存到 $2"
    '------------------------------------------------
    '为您的每个插件函数建立一个Case，就可以在按键精灵里显示插件的使用说明
    '------------------------------------------------
    End Select
    Get_Plugin_Interpret_Template = Translate_Description(Description_Text)
End Function
