VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTestSmHTTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ----------[变量定义]----------
Dim SmHTTP As New SmHTTP

Dim FnIndex As Long
Const FnCount As Long = 32


' 运行测试服务器
Public Sub test_run_server()
    Dim ws, ProceName, Proce, x
    
    Proce = False
    ProceName = "go-httpbin-win.exe" '判断的进程
    Set ws = CreateObject("wscript.shell")
    For Each x In GetObject("winmgmts:").instancesof("win32_process")
        If UCase(x.name) = UCase(ProceName) Then
            Proce = True
            Exit For
        End If
    Next
    Set ws = Nothing
    If Not Proce Then
        Shell "./Project/res/" & ProceName
    End If
End Sub


' 测试插件版本号
Public Sub test_ver()
    Dim ver: ver = SmHTTP.ver()
    
    ' 断言
    Debug.Assert Expression(ver, "=", "1.0.0.4")
End Sub


' 测试插件的路径
Public Sub test_getbasepath()
    Dim Path: Path = SmHTTP.GetBasePath()
    
    ' 断言
    Debug.Assert Expression(Path, "=", "E:\AppData\Roaming\GitHub\【VisualBasic - 开发项目】SmHTTP\Project\SmHTTP.dll")
End Sub


' 测试插件对象ID
Public Sub test_getid()
    Dim id: id = SmHTTP.GetID()
    
    ' 断言
    Debug.Assert Expression(id, ">", 0)
End Sub


' 测试设置开启自动识别对应传入参数
Public Sub test_set_auto_param_array_on()
    Dim ret: ret = SmHTTP.SetAutoParamArray(True)
    
    ' 断言
    Debug.Assert Expression(ret, "=", 1)
End Sub


' 测试设置关闭自动识别对应传入参数
Public Sub test_set_auto_param_array_off()
    Dim ret: ret = SmHTTP.SetAutoParamArray(False)
    
    ' 断言
    Debug.Assert Expression(ret, "=", 1)
End Sub


' 测试构造请求头
Public Sub test_headers()
    Dim Headers: Headers = SmHTTP.Headers( _
        "Accept", "*/*", _
        "Accept-Language", "zh-CN,zh;q=0.8", _
        "Host", "https://bbs.anjian.com", _
        "User-Agent", "Mozilla/4.0 (compatible; MSIE 9.0; Windows NT 6.1)", _
        "Content-Type", "application/x-www-form-urlencoded" _
    )
    
    ' 断言
    Debug.Assert Expression(Headers, "=", "Accept:*/*" & vbCrLf & _
        "Accept-Language:zh-CN,zh;q=0.8" & vbCrLf & _
        "Host:https://bbs.anjian.com" & vbCrLf & _
        "User-Agent:Mozilla/4.0 (compatible; MSIE 9.0; Windows NT 6.1)" & vbCrLf & _
        "Content-Type:application/x-www-form-urlencoded")
End Sub


' 构造Cookies
Public Sub test_cookies()
    Dim ret: ret = SmHTTP.Cookies( _
        "_ga", "GA1.2.1206281266.1647004488", _
        "BAIDUID_BFESS", "0F068EE7974C72C13A37B02D9855DD1C:SL=0:NR=10:FG=1", _
        "ZFY", "O7YgtvLEvTKeDmPLbV8Nbwq3xYhFAOP9m9A:BTtT0AkQ:C" _
    )
    
    ' 断言
    Debug.Assert Expression(ret, "=", "_ga=GA1.2.1206281266.1647004488;BAIDUID_BFESS=0F068EE7974C72C13A37B02D9855DD1C:SL=0:NR=10:FG=1;ZFY=O7YgtvLEvTKeDmPLbV8Nbwq3xYhFAOP9m9A:BTtT0AkQ:C")
End Sub

' 构造 URLData 请求体
Public Sub test_data()
    Dim ret: ret = SmHTTP.data( _
        "username", "SMWH", _
        "password", "123456" _
    )
    
    ' 断言
    Debug.Assert Expression(ret, "=", "username=SMWH&password=123456")
End Sub

' 构造 form-data 请求体
Public Sub test_form_data()
    Dim ret: ret = SmHTTP.formdata( _
        "username", "SMWH", _
        "password", "123456" _
    )
    
    ' 断言
    Debug.Assert Expression(ret, "=", "--WebKitFormBoundarySmHTTPSMWHff" & vbCrLf & _
                      "Content-Disposition: form-data; name=""username""" & vbCrLf & _
                      "" & vbCrLf & _
                      "SMWH" & vbCrLf & _
                      "--WebKitFormBoundarySmHTTPSMWHff" & vbCrLf & _
                      "Content-Disposition: form-data; name=""password""" & vbCrLf & _
                      "" & vbCrLf & _
                      "123456" & vbCrLf & _
                      "--WebKitFormBoundarySmHTTPSMWHff--" & vbCrLf)
End Sub

' 构造 JSON 请求体
Public Sub test_json_data()
    Dim json: json = SmHTTP.JsonData( _
        "empty", Empty, _
        "null", Null, _
        "int", 123, _
        "float", 3.14, _
        "bool", True, _
        "str", "神梦无痕""1042207232""", _
        "array", Array(1, 3.14, True, Null, "按键精灵") _
    )
    
    ' 断言
    Debug.Assert Expression(json, "=", "{""empty"":"""",""null"":null,""int"":123,""float"":3.14,""bool"":true,""str"":""\u795E\u68A6\u65E0\u75D5\""1042207232\"""",""array"":[1,3.14,true,null,""\u6309\u952E\u7CBE\u7075""]}")
End Sub


' 测试 GET 请求
Public Sub test_http_get()
    Dim ret: ret = SmHTTP.HTTP_GET("http://localhost:8080/get")
    Dim status: status = SmHTTP.GetStatus()
    
    ' 断言
    Debug.Assert Expression(status, "=", 200)
End Sub


' 测试 GET 请求API余额
Public Sub test_http_get_api_balance()
    Dim Headers: Headers = SmHTTP.Headers( _
        "api-id", "19682662733224024298926531509058411934", _
        "api-key", "3WKZFAP1FWRETY3C283XPLAABFH2EH" _
    )
    Dim ret: ret = SmHTTP.HTTP_GET("https://dp33c1xwtxufgln-db202204160934.adb.us-phoenix-1.oraclecloudapps.com/ords/tardigrade_api/api/v1/balance", Headers)
    Debug.Print ret
    Dim status: status = SmHTTP.GetStatus()
    
    ' 断言
    Debug.Assert Expression(status, "=", 200)
End Sub


' 测试 POST 请求
Public Sub test_http_post()
    Dim ret: ret = SmHTTP.HTTP_POST("http://localhost:8080/post", "username=SMWH&password=123456", "Content-Type: application/x-www-form-urlencoded")
    Dim status: status = SmHTTP.GetStatus()
    
    ' 断言
    Debug.Assert Expression(status, "=", 200)
End Sub


' 测试 HEAD 请求
Public Sub test_http_head()
    Dim ret: ret = SmHTTP.HTTP_HEAD("http://localhost:8080/head")
    Dim status: status = SmHTTP.GetStatus()
    
    ' 断言
    Debug.Assert Expression(status, "=", 200)
End Sub


' 测试 OPTIONS 请求
Public Sub test_http_options()
    Dim ret: ret = SmHTTP.HTTP_OPTIONS("http://localhost:8080/options")
    Dim status: status = SmHTTP.GetStatus()
    
    ' 断言
    Debug.Assert Expression(status, "=", 200)
End Sub


' 测试 PATCH 请求
Public Sub test_http_patch()
    Dim ret: ret = SmHTTP.HTTP_PATCH("http://localhost:8080/patch")
    Dim status: status = SmHTTP.GetStatus()
    
    ' 断言
    Debug.Assert Expression(status, "=", 200)
End Sub


' 测试 PUT 请求
Public Sub test_http_put()
    Dim ret: ret = SmHTTP.HTTP_PUT("http://localhost:8080/put")
    Dim status: status = SmHTTP.GetStatus()
    
    ' 断言
    Debug.Assert Expression(status, "=", 200)
End Sub


' 测试 DELETE 请求
Public Sub test_http_delete()
    Dim ret: ret = SmHTTP.HTTP_DELETE("http://localhost:8080/delete")
    Dim status: status = SmHTTP.GetStatus()
    
    ' 断言
    Debug.Assert Expression(status, "=", 200)
End Sub


' 测试 Request 请求
Public Sub test_http_request()
    Call test_set_auto_param_array_on
    Dim ret: ret = SmHTTP.HTTP_Request("GET", "https://www.bing.com/ipv6test/test?FORM=MONITR", "UTF-8")
    Dim status: status = SmHTTP.GetStatus()
    Call test_set_auto_param_array_off
    
    ' 断言
    Debug.Assert Expression(status, "=", 200)
End Sub


' 测试 Request 请求，无参数情况
Public Sub test_http_request_not_args()
    Dim ret: ret = SmHTTP.HTTP_Request()
    
    ' 断言
    Debug.Assert Expression(ret, "=", "SmHTTP")
End Sub


' 测试 Request 请求，参数匹配测试
Public Sub test_http_request_args_match()
    Call test_set_auto_param_array_on
    Dim ret: ret = SmHTTP.HTTP_Request("http://localhost:8080/post", "POST")
    Dim url: url = SmHTTP.GetJSON(ret, "url")
    Call test_set_auto_param_array_off
    
    ' 断言
    Debug.Assert Expression(url, "=", "http://localhost:8080/post")
End Sub


' 测试 Request 请求，异常情况
Public Sub test_http_request_invalid_url()
    Call test_set_auto_param_array_on
    On Error Resume Next
    Dim ret: ret = SmHTTP.HTTP_Request("GET", "invalid_url")
    On Error GoTo 0
    Call test_set_auto_param_array_off

    ' 断言
    Debug.Assert Expression(IsEmpty(ret), "=", True)
End Sub


' 测试 Request 请求，性能测试
Public Sub test_http_request_largeData()
    Call test_set_auto_param_array_on
    Dim data: data = String(10000, "a")
    Dim ret: ret = SmHTTP.HTTP_Request("POST", "http://localhost:8080/post", data)
    Call test_set_auto_param_array_off

    ' 断言
    Debug.Assert Expression(ret, "=", "error parsing request body: invalid character 'a' looking for beginning of value" & vbLf)
End Sub


' 测试 Request 请求，超时与重定向测试
' 注意：这需要实际的测试服务器或服务来模拟长时间响应和重定向。
Public Sub test_http_request_slow_response()
'    Call test_set_auto_param_array_on
'    Dim Ret: Ret = SmHTTP.HTTP_Request("GET", "http://slow-response.com", 1)
'    Call test_set_auto_param_array_off
'
'    ' 断言
'    Debug.Assert Expression(IsEmpty(Ret), "=", True)
End Sub


' 测试代理IP
Public Sub test_http_proxy()
'    Dim proxy: proxy = "136.243.90.203:80"
'    Call test_set_auto_param_array_on
'    Dim Ret: Ret = SmHTTP.HTTP_Request("GET", "http://www.bathome.net/s/ip.php", proxy)
'    Call test_set_auto_param_array_off
'
'    ' 断言
'    Debug.Assert Expression(Ret, "=", Split(proxy, ":")(0))
End Sub


' 测试代理IP(带身份认证)
Public Sub test_http_proxy_auth()
'    Call test_set_auto_param_array_on
'    Dim Ret: Ret = SmHTTP.HTTP_Request("GET", "http://www.bathome.net/s/ip.php", "112.5.56.2:9091")
'    Debug.Assert Ret = "112.5.56.2"
'    Call test_set_auto_param_array_off
End Sub


' 测试身份认证
Public Sub test_http_auth_basic()
    Call test_set_auto_param_array_on
    Dim ret: ret = SmHTTP.HTTP_GET("https://ssr3.scrape.center/", "BASIC", "admin", "admin")
    Dim status: status = SmHTTP.GetStatus()
    Call test_set_auto_param_array_off
    
    ' 断言
    Debug.Assert Expression(status, "=", 200)
End Sub


' 测试返回JSON
Public Sub test_http_ret_json()
    Call test_set_auto_param_array_on
    Dim url: url = "https://extreme-ip-lookup.com/json/?tdsourcetag=s_pcqq_aiomsg"
    Dim ret: ret = SmHTTP.HTTP_Request("GET", url)
    Dim message: message = SmHTTP.GetJSON(ret, "status")
    Call test_set_auto_param_array_off
    
    ' 断言
    Debug.Assert Expression(message, "=", "success")
End Sub


' 测试上传文件
Public Sub test_http_upload_file()
    Dim data: data = SmHTTP.formdata( _
        "@file", App.Path & "\..\Logo.png", "image/png", _
        "username", "SMWH", _
        "password", "123456" _
    )
    Call test_set_auto_param_array_on
    Dim ret: ret = SmHTTP.HTTP_POST("http://localhost:8080/post", data)
    Dim username: username = SmHTTP.GetJSON(ret, "form.username[0]")
    Dim password: password = SmHTTP.GetJSON(ret, "form.password[0]")
    Call test_set_auto_param_array_off
    
    ' 断言
    Debug.Assert Expression(username, "=", "SMWH")
    Debug.Assert Expression(password, "=", "123456")
End Sub


' 测试上传图片到SM.MS
Public Sub test_http_upload_file_smms()
    Dim token: token = Environ("SMMS_Authorization")
    Dim data: data = SmHTTP.formdata( _
        "@smfile", App.Path & "\..\Logo.png", "image/png", _
        "format", "json" _
    )
    Dim Headers: Headers = SmHTTP.Headers( _
        "Authorization", token _
    )
    Dim ret: ret = SmHTTP.HTTP_POST("https://smms.app/api/v2/upload", data, Headers)
    ' Debug.Print ret
    Dim success: success = SmHTTP.GetJSON(ret, "success")
    Dim filename: filename = SmHTTP.GetJSON(ret, "data.filename")
    Dim hash: hash = SmHTTP.GetJSON(ret, "data.hash")
    Call SmHTTP.HTTP_GET("https://smms.app/api/v2/delete/" & hash, Headers) ' 删除上传图片

    ' 断言
    Debug.Assert Expression(success, "=", True)
    Debug.Assert Expression(filename, "=", "Logo.png")
End Sub


' 测试上传图片到imgbb
Public Sub test_http_upload_file_imgbb()
    Dim ret: ret = SmHTTP.HTTP_GET("https://imgbb.com/")
    Dim url: url = TestTool.GetMidS(ret, "PF.obj.config.json_api=""", """;")
    Dim token: token = TestTool.GetMidS(ret, "PF.obj.config.auth_token=""", """;")
    Dim timestamp: timestamp = DateDiff("s", "1970-1-1 0:0:0", DateAdd("h", -8, Now)) & Right(CLng(Timer() * 1000), 3)
    Dim formdata: formdata = SmHTTP.formdata( _
        "@source", App.Path & "\..\Logo.png", "image/png", _
        "type", "file", _
        "action", "upload", _
        "timestamp", timestamp, _
        "auth_token", token _
    )
    Call test_set_auto_param_array_on
    Dim json: json = SmHTTP.HTTP_POST(url, formdata)
    ' Debug.Print json
    Dim code: code = SmHTTP.GetJSON(json, "status_code")
    Dim imgName: imgName = SmHTTP.GetJSON(json, "image.name")
    Dim imgURL: imgURL = SmHTTP.GetJSON(json, "image.url")
    Call test_set_auto_param_array_off

    ' 断言
    Debug.Assert Expression(code, "=", 200)
    Debug.Assert Expression(imgName, "=", "Logo")
    'Debug.Assert Expression(imgURL, "=", "https://i.ibb.co/PMhrKBd/Logo.png")
End Sub


' 测试上传图片到SunPics
Public Sub test_http_upload_file_SunPics()
    Dim timestamp: timestamp = DateDiff("s", "1970-1-1 0:0:0", DateAdd("h", -8, Now)) & Right(CLng(Timer() * 1000), 3)
    Dim formdata: formdata = SmHTTP.formdata( _
        "@file", App.Path & "\..\Logo.png", "image/png", _
        "name", "Logo.png", _
        "uuid", timestamp _
    )
    Dim ret: ret = SmHTTP.HTTP_POST("https://pics.sunbangyan.cn/application/upload.php", formdata)
    ' Debug.Print ret
    Dim code: code = SmHTTP.GetJSON(ret, "code")
    Dim srcName: srcName = SmHTTP.GetJSON(ret, "srcName")

    ' 断言
    Debug.Assert Expression(code, "=", 200)
    Debug.Assert Expression(srcName, "=", "Logo")
End Sub


' 测试上传图片到upload
Public Sub test_http_upload_file_upload()
'    Dim formdata: formdata = SmHTTP.formdata( _
'        "@uploaded_file[]", "C:\Users\SMWH\Pictures\Saved Pictures\jkansjxasd202374.ico", "image/png" _
'    )
'    Debug.Print formdata
'    Dim headers: headers = SmHTTP.headers( _
'        "origin", "https://upload.cc", _
'        "pragma", "no-cache", _
'        "referer", "https://upload.cc/", _
'        "X-Requested-With", "XMLHttpRequest" _
'    )
'    Dim ret: ret = SmHTTP.HTTP_POST("https://upload.cc/image_upload", formdata, headers)
'    Debug.Print ret
'    Dim code: code = SmHTTP.GetJSON(ret, "code")
'    Dim name: name = SmHTTP.GetJSON(ret, "success_image[0].name")
'
'    ' 断言
'    Debug.Assert Expression(code, "=", 100)
'    Debug.Assert Expression(name, "=", "Logo.png")
End Sub


' 测试按键精灵论坛签到
Public Sub test_bbs_anjian_signin()
    Dim ret, Cookies, Headers
    Dim user: user = Environ("AJ_USER")
    Dim pass: pass = Environ("AJ_PASS")
    'Debug.Print user, pass
    Dim data: data = SmHTTP.data( _
        "username", user, _
        "password", pass, _
        "question", "0", _
        "answer", "", _
        "templateid", "0", _
        "login", "", _
        "expires", "43200" _
    )
    'SmHTTP.调试开关 = True
    Call test_set_auto_param_array_on
    ' 登录论坛账号
    ret = SmHTTP.HTTP_POST("http://bbs.anjian.com/login.aspx?referer=forumindex.aspx", data)
    Debug.Assert InStr(ret, user)  ' 判断是否登录成功
    Cookies = SmHTTP.getCookies()
    ' 打卡签到
    data = SmHTTP.data( _
        "signmessage", "签个到，每天心情都是美美哒~~按键精灵祝大家新年好运连连！！" _
    )
    Headers = SmHTTP.Headers( _
        "Referer", "http://bbs.anjian.com/" _
    )
    ret = SmHTTP.HTTP_POST("http://bbs.anjian.com/addsignin.aspx?infloat=1&inajax=1", data, Headers, Cookies)
    Call test_set_auto_param_array_off
    
    ' 断言
    Debug.Assert Expression(InStr(ret, "恭喜您获取本日签到奖励") Or InStr(ret, "你今天已经签到过了"), ">", 0)  ' 判断是否签到成功
End Sub


' 测试强制服务器返回未压缩的内容
Public Sub test_http_ret_not_gzip()
    Call test_set_auto_param_array_on
    Dim json: json = SmHTTP.JsonData( _
        "detail_enid", "obyrmnqGdwxkXW6Yo3JelBz2D5ZO8a", _
        "note_type", 2, _
        "only_replied", False, _
        "page", 1, _
        "page_count", 20, _
        "sort_by", "create", _
        "source_type", 65 _
    )
    Dim Headers: Headers = SmHTTP.Headers( _
        "Accept-Encoding", "identity" _
    )
    Dim ret: ret = SmHTTP.HTTP_POST("https://www.dedao.cn/pc/ledgers/notes/article_comment_list", json, Headers)
    Dim message: message = SmHTTP.GetJSON(ret, "h.c")
    Call test_set_auto_param_array_off
    
    ' 断言
    Debug.Assert Expression(message, "=", 0)
End Sub


' 测试smos-cdn
Public Sub test_http_ret_smos_cdn()
    Dim json: json = SmHTTP.HTTP_GET("https://cname-china.vercel-dns.com/Ver/SmHTTP.json", "Host: smos-cdn.vercel.app")
    Dim name: name = SmHTTP.GetJSON(json, "name")
    
    ' 断言
    Debug.Assert Expression(name, "=", "神梦HTTP请求插件")
End Sub


' 测试获取QQ昵称
Public Sub test_get_qq_nick_name()
    Dim qq: qq = "1042207232"
    Call test_set_auto_param_array_on
    Dim ret: ret = SmHTTP.HTTP_GET("https://users.qzone.qq.com/fcg-bin/cgi_get_portrait.fcg?uins=" & qq, "UTF-8")
    Dim name: name = SmHTTP.GetJSON(ret, "[" + CStr(qq) + "][6]")
    Call test_set_auto_param_array_off

    ' 断言
    Debug.Assert Expression(EscapeURL(name), "=", "%EF%BF%BD%EF%BF%BD%EF%BF%BD%EF%BF%BD%EF%BF%BD%DE%BA%EF%BF%BD")
End Sub


' 测试获取QQ头像链接
Public Sub test_download_qq_avatar_url()
    Dim qq: qq = 1042207232
    Call test_set_auto_param_array_on
    Dim ret: ret = SmHTTP.HTTP_GET("https://users.qzone.qq.com/fcg-bin/cgi_get_portrait.fcg?uins=" & qq, "UTF-8")
    Dim img_src: img_src = SmHTTP.GetJSON(ret, "[" + CStr(qq) + "][0]")
    Call test_set_auto_param_array_off

    ' 断言
    Debug.Assert Expression(img_src, "=", "http://qlogo1.store.qq.com/qzone/1042207232/1042207232/100")
End Sub


' 测试下载QQ头像
Public Sub test_download_qq_qlogo()
    Dim qq: qq = 1042207232
    Call test_set_auto_param_array_on
    Dim ret: ret = SmHTTP.HTTP_GET("http://q1.qlogo.cn/g?b=qq&s=100&nk=" & qq)
    Call test_set_auto_param_array_off
    
    ' 断言
    Debug.Assert Expression(TypeName(ret), "=", "Byte()")
    Debug.Assert Expression(LenB(ret), "=", 16893)
End Sub



' 测试百度翻译(英译中)
Public Sub test_Translate_Baidu()
'    Dim enStr: enStr = "I Love You"
'    Dim timestamp: timestamp = DateDiff("s", "1970-1-1 0:0:0", DateAdd("h", -8, Now)) & Right(CLng(Timer() * 1000), 3)
'    Dim ret: ret = SmHTTP.HTTP_GET("https://www.baidu.com/")
'    Dim L: L = InStr(1, ret, "var s_domain = {", vbTextCompare): Debug.Assert L > 0
'    Dim r: r = InStr(L, ret, "};", vbTextCompare): Debug.Assert r > 0
'    Dim s_domain: s_domain = Mid(ret, L, r - L + 1)
'    Dim sensearch: sensearch = SmHTTP.GetJSON(s_domain, "ssllist['sensearch.baidu.com']")
'    ret = SmHTTP.HTTP_GET("http://" & sensearch & "/sensearch/selecttext?cb=jQuery_Fun_" & timestamp & "&q=" & enStr & "&_=" & timestamp)
'    Dim errno: errno = SmHTTP.GetJSON(ret, "errno")
'    Dim result: result = SmHTTP.GetJSON(ret, "data.result")
'
'    ' 断言
'    Debug.Assert Expression(errno, "=", 0)
'    Debug.Assert Expression(result, "=", "我爱你")
End Sub


' 测试微软翻译(英译中)
' 英文：en
' 中文：zh-Hans
Public Sub test_Translate_Microsoft()
    Dim enStr: enStr = "I Love You"
    Dim authCode: authCode = SmHTTP.HTTP_GET("https://edge.microsoft.com/translate/auth")
    Dim json: json = "[" & SmHTTP.JsonData( _
        "Text", enStr _
    ) & "]"
    Dim Headers: Headers = SmHTTP.Headers( _
        "Authorization", "Bearer " & authCode _
    )
    Dim result: result = SmHTTP.HTTP_POST("https://api-edge.cognitive.microsofttranslator.com/translate?from=&to=zh-Hans&api-version=3.0&includeSentenceLength=true", json, Headers)
    Dim status: status = SmHTTP.GetStatus()
    If status = 200 Then result = SmHTTP.GetJSON(result, "[0].translations[0].text")
    
    ' 断言
    Debug.Assert Expression(status, "=", 200)
    Debug.Assert Expression(result, "=", "我爱你")
End Sub


' 测试阿里翻译(英译中)
' 英文：en
' 中文：zh
Public Sub test_Translate_AliBaba()
    Dim enStr: enStr = "I Love You"
    Dim ret: ret = SmHTTP.HTTP_GET("https://translate.alibaba.com/api/translate/csrftoken")
    Dim ali_uuid: ali_uuid = SmHTTP.GetJSON(ret, "token")
    Dim dataName: dataName = SmHTTP.GetJSON(ret, "parameterName")
    Dim headerName: headerName = SmHTTP.GetJSON(ret, "headerName")
    Dim fromData: fromData = SmHTTP.formdata( _
        "srcLang", "auto", _
        "tgtLang", "zh", _
        "domain", "general", _
        "query", enStr, _
        dataName, ali_uuid _
    )
    Dim Headers: Headers = SmHTTP.Headers( _
        "accept", "application/json, text/plain, */*", _
        "Origin", "https://translate.alibaba.com", _
        "Referer", "https://translate.alibaba.com/", _
        headerName, ali_uuid, _
        "sec-fetch-site", "same-origin" _
    )
    Dim result: result = SmHTTP.HTTP_POST("https://translate.alibaba.com/api/translate/text", fromData, Headers)
    Dim status: status = SmHTTP.GetStatus()
    If status = 200 Then result = SmHTTP.GetJSON(result, "data.translateText")

    ' 断言
    Debug.Assert Expression(status, "=", 200)
    Debug.Assert Expression(result, "=", "我爱你")
End Sub


' 测试腾讯交互翻译(英译中)
' 英文：en
' 中文：zh
Public Sub test_Translate_TencentQQ()
    Dim enStr: enStr = "I Love You"
    Dim json: json = SmHTTP.JsonData( _
        "header", SmHTTP.JsonData( _
            "fn", "auto_translation" _
        ), _
        "type", "plain", _
        "model_category", "normal", _
        "text_domain", "general", _
        "source", SmHTTP.JsonData( _
            "lang", "auto", _
            "text_list", Array(enStr) _
        ), _
        "target", SmHTTP.JsonData( _
            "lang", "zh" _
        ) _
    )
    Dim Headers: Headers = SmHTTP.Headers( _
        "Host", "transmart.qq.com", _
        "Origin", "https://transmart.qq.com", _
        "Referer", "https://transmart.qq.com/" _
    )
    Dim result: result = SmHTTP.HTTP_POST("https://transmart.qq.com/api/imt", json, Headers)
    Dim status: status = SmHTTP.GetStatus()
    If status = 200 Then result = SmHTTP.GetJSON(result, "auto_translation[0]")
    
    ' 断言
    Debug.Assert Expression(status, "=", 200)
    Debug.Assert Expression(result, "=", "我爱你")
End Sub


' 测试 DeepL 翻译(英译中)
' 英文：EN
' 中文：ZH
Public Sub test_Translate_DeepL()
    Dim enStr: enStr = "I Love You"
    Dim from: from = "auto"
    Dim lang: lang = "ZH"
    Dim deepl_id: Randomize: deepl_id = 10000 * Round(10000 * Rnd)
    Dim r: r = DateDiff("s", "01/01/1970 00:00:00", Now())
    Dim n: n = 1
    Dim iCount: iCount = Len(enStr) - Len(Replace(enStr, "i", "")): n = n + iCount
    Dim timestamp: timestamp = r + (n - r Mod n)
    Dim jobsJSON: jobsJSON = "[" & _
        SmHTTP.JsonData( _
            "kind", "default", _
            "sentences", "[" & _
                SmHTTP.JsonData( _
                    "text", enStr, _
                    "id", 1, _
                    "prefix", "" _
                ) & _
            "]", _
            "raw_en_context_before", Array(), _
            "raw_en_context_after", Array(), _
            "preferred_num_beams", 4 _
        ) & _
    "]"
    Dim langJSON: langJSON = SmHTTP.JsonData( _
        "target_lang", lang, _
        "preference", SmHTTP.JsonData( _
            "weight", SmHTTP.JsonData( _
                "DE", 0.18427, _
                "EN", IIf(from = "EN", 9.93878, 5.90582), _
                "ES", 0.13236, "FR", 0.16311, "IT", 0.11621, "JA", 0.17963, "NL", 0.1865, "PL", 0.11549, "PT", 0.10159, "RU", 0.10577, _
                "ZH", IIf(from = "ZH", 9.93878, 5.90582), _
                "BG", 0.07468, "CS", 0.09005, "DA", 0.08567, "EL", 0.07069, "ET", 0.0836, "FI", 0.09628, "HU", 0.08731, "LT", 0.07119, _
                "LV", 0.06866, "RO", 0.07842, "SK", 0.07497, "SL", 0.08492, "SV", 0.10275, "TR", 0.07728, "ID", 0.09161, "UK", 0.08573, "KO", 0.04671, _
                "NB", 0.05511 _
            ), _
            "default", "default" _
        ), _
        "source_lang_user_selected", from _
    )
    Dim json: json = SmHTTP.JsonData( _
        "jsonrpc", "2.0", _
        "method", "LMT_handle_jobs", _
        "params", SmHTTP.JsonData( _
            "jobs", jobsJSON, _
            "lang", langJSON, _
            "priority", 1, _
            "commonJobParams", SmHTTP.JsonData( _
                "mode", "translate", _
                "textType", "plaintext", _
                "browserType", 1 _
            ), _
            "timestamp", timestamp _
        ), _
        "id", deepl_id _
    )
    json = Replace(json, "hod"":""", IIf((deepl_id + 3) Mod 13 = 0 Or (deepl_id + 5) Mod 29 = 0, "hod"" : """, "hod"": """))
    'Debug.Print json
    Dim Headers: Headers = SmHTTP.Headers( _
        "authority", "www2.deepl.com", _
        "Origin", "https://www.deepl.com", _
        "Referer", "https://www.deepl.com/" _
    )
    Dim result: result = SmHTTP.HTTP_POST("https://www2.deepl.com/jsonrpc", json, Headers)
    'Debug.Print result
    Dim status: status = SmHTTP.GetStatus()
    If status = 200 Then result = SmHTTP.GetJSON(result, "result.translations[0].beams[0].sentences[0].text")

    ' 断言
    'Debug.Assert Expression(status, "=", 200)
    'Debug.Assert Expression(result, "=", "我爱你")
End Sub


' 测试有道翻译(英译中)
' 英文：EN
' 中文：ZH_CN
Public Sub test_Translate_YoudaoMobileWeb()
    Dim enStr: enStr = "I Love You"
    Dim from: from = "EN"
    Dim lang: lang = "ZH_CN"
    Dim data: data = SmHTTP.data( _
        "inputtext", enStr, _
        "type", from & "2" & lang _
    )
    Dim Headers: Headers = SmHTTP.Headers( _
        "Origin", "https://m.youdao.com", _
        "Referer", "https://m.youdao.com/translate/" _
    )
    Dim result: result = SmHTTP.HTTP_POST("https://m.youdao.com/translate", data, Headers, , "UTF-8")
    'Debug.Print result
    Dim htmlDoc: Set htmlDoc = CreateObject("htmlfile")
    htmlDoc.Write result
    Dim status: status = SmHTTP.GetStatus()
    If status = 200 Then result = Trim(htmlDoc.getElementById("translateResult").getElementsByTagName("li")(0).innerText)
    Set htmlDoc = Nothing
    
    ' 断言
    Debug.Assert Expression(status, "=", 200)
    Debug.Assert Expression(result, "=", "我爱你")
End Sub


' 测试搜狗翻译(英译中)
' 英文：en
' 中文：zh-CHS
Public Sub test_Translate_SogouWeb()
    Dim enStr: enStr = "I Love You"
    Dim from: from = "en"
    Dim lang: lang = "zh-CHS"
    Dim auth: auth = SmHTTP.HTTP_GET("https://fanyi.sogou.com")
    Dim Cookies: Cookies = SmHTTP.getCookies()
    Dim re: Set re = New RegExp
    re.Pattern = "secretCode"":(\d+)"
    Dim secretCode: secretCode = re.Execute(auth)(0).SubMatches(0)
    re.Pattern = "uuid\"":""(.*?)"""
    Dim uuid: uuid = re.Execute(auth)(0).SubMatches(0)
    Set re = Nothing
    Dim sign: sign = TestCryptoAPI.MD5Hash(from & lang & enStr & secretCode)
    Dim data: data = SmHTTP.data( _
        "keyword", enStr, _
        "transfrom", from, _
        "transto", lang, _
        "model", "general" _
    )
    Dim json: json = SmHTTP.JsonData( _
        "from", from, _
        "to", lang, _
        "text", enStr, _
        "client", "pc", _
        "fr", "browser_pc", _
        "needQc", 1, _
        "s", sign, _
        "uuid", uuid, _
        "exchange", False _
    )
    Dim Headers: Headers = SmHTTP.Headers( _
        "Origin", "https://fanyi.sogou.com", _
        "Referer", "https://fanyi.sogou.com/text?" & data, _
        "Accept", "application/json, text/plain, */*", _
        "Pragma", "no-cache", _
        "Cache-Control", "no-cache" _
    )
    Dim result: result = SmHTTP.HTTP_POST("https://fanyi.sogou.com/api/transpc/text/result", json, Headers, Cookies)
    ' Debug.Print result
    Dim status: status = SmHTTP.GetStatus()
    If status = 200 Then result = SmHTTP.GetJSON(result, "data.translate.dit")

    ' 断言
    Debug.Assert Expression(status, "=", 200)
    Debug.Assert Expression(result, "=", "我爱你")
End Sub


' 百度统计
Public Sub test_Baidu_tongji()
    Dim data: data = SmHTTP.data( _
        "cc", 1, _
        "ck", 1, _
        "cl", "32-bit", _
        "ds", "1024x1024", _
        "et", 0, _
        "ep", 0, _
        "fl", "11.0" _
    )
    Dim Headers: Headers = SmHTTP.Headers( _
        "Referer", "https://www.SmHTTP.com/", _
        "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/118.0.0.0 Safari/537.36 Edg/118.0.2088.46 VB/6.0", _
        "Cookie", "HMACCOUNT=SMWHff-VB6-GUID" _
    )
    Dim url: url = "https://hm.baidu.com/hm.gif?si=212f0c3f981a440699efc710f3043368&" & data
    Call SmHTTP.HTTP_GET(url, Headers)
    Dim status: status = SmHTTP.GetStatus()
    
    ' 断言
    Debug.Assert Expression(status, "=", 200)
End Sub


' 测试下载文件成功
Public Sub test_http_download()
    Dim url: url = "https://bbs.anjian.com/templates/default/Images/logo1.png"
    Dim Path: Path = Environ("TEMP") & "\logo1.png"
    Dim ret: ret = SmHTTP.HTTP_DownLoad(url, Path)

    ' 断言
    Debug.Assert Expression(ret, "=", 1)
End Sub


' 测试下载文件，无效的URL
Public Sub test_http_download_invalid_url()
    Dim url: url = "http://invalid.url"
    Dim Path: Path = Environ("TEMP") & "\file.txt"
    Dim ret: ret = SmHTTP.HTTP_DownLoad(url, Path)

    ' 断言
    Debug.Assert Expression(ret, "=", 0)
End Sub


' 测试下载文件，空的URL
Public Sub test_http_download_empty_url()
    Dim url: url = ""
    Dim Path: Path = Environ("TEMP") & "\EmptyURLFile.txt"
    Dim ret: ret = SmHTTP.HTTP_DownLoad(url, Path)

    ' 断言
    Debug.Assert Expression(ret, "=", 0)
End Sub


' 测试下载文件，无效的保存路径
Public Sub test_http_download_invalid_path()
    Dim url: url = "https://bbs.anjian.com/templates/default/Images/logo1.png"
    Dim Path: Path = "X:\Invalid\Path\file.txt"
    Dim ret: ret = SmHTTP.HTTP_DownLoad(url, Path)

    ' 断言
    Debug.Assert Expression(ret, "=", 0)
End Sub


' 测试下载文件，空的保存路径
Public Sub test_http_download_empty_path()
    Dim url: url = "https://bbs.anjian.com/templates/default/Images/logo1.png"
    Dim Path: Path = ""
    Dim ret: ret = SmHTTP.HTTP_DownLoad(url, Path)

    ' 断言
    Debug.Assert Expression(ret, "=", 0)
End Sub


' 测试下载文件加强版
Public Sub test_http_downloadEx()
    Dim url: url = "https://bbs.anjian.com/templates/default/Images/logo1.png|http://res2.mobileanjian.com/Advert/20230531/139355dc1da4409f80f852b6523f8051.png"
    Dim Path: Path = Environ("TEMP") & "\logo1.png|" & Environ("TEMP") & "\139355dc1da4409f80f852b6523f8051.png"
    Dim ret: ret = SmHTTP.HTTP_DownLoadEx(url, Path)

    ' 断言
    Debug.Assert Expression(ret, "=", "")
End Sub


' 测试下载文件加强版，无效的URL
Public Sub test_http_downloadEx_invalid_url()
    Dim url: url = "https://bbs.anjian.com/templates/default/Images/logo1.png|http://res2.mobileanjian.com/Advert/20230531/139355dc1da4409f80f852b6523f8051.png|http://invalid.url"
    Dim Path: Path = Environ("TEMP") & "\logo1.png|" & Environ("TEMP") & "\139355dc1da4409f80f852b6523f8051.png|" & Environ("TEMP") & "\invalid.url"
    Dim ret: ret = SmHTTP.HTTP_DownLoadEx(url, Path)

    ' 断言
    Debug.Assert Expression(ret, "=", "http://invalid.url," & Environ("TEMP") & "\invalid.url")
End Sub


' 测试下载文件加强版，空的URL
Public Sub test_http_downloadEx_empty_url()
    Dim url: url = "https://bbs.anjian.com/templates/default/Images/logo1.png|http://res2.mobileanjian.com/Advert/20230531/139355dc1da4409f80f852b6523f8051.png"
    Dim Path: Path = Environ("TEMP") & "\logo1.png|" & Environ("TEMP") & "\139355dc1da4409f80f852b6523f8051.png|" & Environ("TEMP") & "\EmptyURLFileEx.txt"
    Dim ret: ret = SmHTTP.HTTP_DownLoadEx(url, Path)

    ' 断言
    Debug.Assert Expression(ret, "=", "," & Environ("TEMP") & "\EmptyURLFileEx.txt")
End Sub


' 测试下载文件加强版，无效的保存路径
Public Sub test_http_downloadEx_invalid_path()
    Dim url: url = "https://bbs.anjian.com/templates/default/Images/logo1.png|http://res2.mobileanjian.com/Advert/20230531/139355dc1da4409f80f852b6523f8051.png|http://files.anjian.com/avatars/avatars/upload/002/69/99/98_avatar_middle.jpg"
    Dim Path: Path = Environ("TEMP") & "\logo1.png|" & Environ("TEMP") & "\139355dc1da4409f80f852b6523f8051.png|X:\Invalid\Path\fileEx.txt"
    Dim ret: ret = SmHTTP.HTTP_DownLoadEx(url, Path)

    ' 断言
    Debug.Assert Expression(ret, "=", "http://files.anjian.com/avatars/avatars/upload/002/69/99/98_avatar_middle.jpg,X:\Invalid\Path\fileEx.txt")
End Sub


' 测试下载文件加强版，空的保存路径
Public Sub test_http_downloadEx_empty_path()
    Dim url: url = "https://bbs.anjian.com/templates/default/Images/logo1.png|http://res2.mobileanjian.com/Advert/20230531/139355dc1da4409f80f852b6523f8051.png|http://files.anjian.com/avatars/avatars/upload/002/69/99/98_avatar_middle.jpg"
    Dim Path: Path = Environ("TEMP") & "\logo1.png|" & Environ("TEMP") & "\139355dc1da4409f80f852b6523f8051.png|"
    Dim ret: ret = SmHTTP.HTTP_DownLoadEx(url, Path)

    ' 断言
    Debug.Assert Expression(ret, "=", "http://files.anjian.com/avatars/avatars/upload/002/69/99/98_avatar_middle.jpg,")
End Sub


' 测试文本转原神语音，并下载语音文件
Public Sub test_http_download_TTS_yuanshen()
    Dim text: text = "没想到没想到，胡桃我现在力如泉涌，心似雄鹰，非常有自信哦。"
    Dim json: json = SmHTTP.JsonData( _
        "data", Array( _
            text, _
            "胡桃_ZH", _
            0.2, _
            0.6, _
            0.8, _
            1, _
            "ZH", _
            Null, _
            "Happy", _
            "Text prompt", _
            "", _
            0.7 _
        ), _
        "event_data", Null, _
        "fn_index", 0, _
        "session_hash", TestTool.GetGUID() _
    )
    Dim Path: Path = Environ("TEMP") & "\胡桃.wav"
    Dim ret: ret = SmHTTP.HTTP_POST("https://v2.genshinvoice.top/run/predict", json)
    'Debug.Print ret
    Dim status, name: status = SmHTTP.GetStatus()
    If status = 200 Then status = SmHTTP.GetJSON(ret, "data[0]")
    If status = "Success" Then name = SmHTTP.GetJSON(ret, "data[1].name")
    ret = SmHTTP.HTTP_DownLoad("https://v2.genshinvoice.top/file=" & name, Path)
    Dim pRet: pRet = PlayWAV(Path): Call StopWAV

    ' 断言
    Debug.Assert Expression(pRet, "=", True)
    Debug.Assert Expression(ret, "=", 1)
End Sub


' 测试多角色文本转原神语音，并下载语音文件
Public Sub test_http_downloadEx_TTS_yuanshen()
    Dim urlArr(1)
    Dim json: json = SmHTTP.JsonData( _
        "data", Array( _
            "西风骑士团，「火花骑士」，可莉，前来报到！…呃—后面该说什么词来着？可莉背不下来啦...", _
            "可莉_ZH", _
            0.2, _
            0.6, _
            0.8, _
            1, _
            "ZH", _
            Null, _
            "Happy", _
            "Text prompt", _
            "", _
            0.7 _
        ), _
        "event_data", Null, _
        "fn_index", 0, _
        "session_hash", TestTool.GetGUID() _
    )
    Dim ret: ret = SmHTTP.HTTP_POST("https://v2.genshinvoice.top/run/predict", json)
    If InStr(ret, "Success") > 0 Then urlArr(0) = "https://v2.genshinvoice.top/file=" & SmHTTP.GetJSON(ret, "data[1].name")
    json = SmHTTP.JsonData( _
        "data", Array( _
            "嘿嘿，荒泷天下第一斗到此一游！ 哈哈哈哈…咳…", _
            "荒泷一斗_ZH", _
            0.2, _
            0.6, _
            0.8, _
            1, _
            "ZH", _
            Null, _
            "Happy", _
            "Text prompt", _
            "", _
            0.7 _
        ), _
        "event_data", Null, _
        "fn_index", 0, _
        "session_hash", TestTool.GetGUID() _
    )
    ret = SmHTTP.HTTP_POST("https://v2.genshinvoice.top/run/predict", json)
    If InStr(ret, "Success") > 0 Then urlArr(1) = "https://v2.genshinvoice.top/file=" & SmHTTP.GetJSON(ret, "data[1].name")
    
    Dim url: url = Join(urlArr, "|")
    Dim Path: Path = Environ("TEMP") & "\可莉.wav|" & Environ("TEMP") & "\荒泷一斗.wav"
    Dim pathArr: pathArr = Split(Path, "|")
    ret = SmHTTP.HTTP_DownLoadEx(url, Path, "GET")
    Dim pRet
    pRet = PlayWAV(pathArr(0)): Call StopWAV
    pRet = PlayWAV(pathArr(1)): Call StopWAV

    ' 断言
    Debug.Assert Expression(VarType(ret), "=", vbString)
    Debug.Assert Expression(ret, "=", "")
End Sub



' 测试小说多角色文本转原神语音，并下载语音文件
Public Sub test_http_downloadEx_TTS_yuanshen_ebook()
    Dim i, sp(8), json, ret, url, pRet
    Dim 赵倩倩: 赵倩倩 = "荧_ZH"
    Dim 医生: 医生 = "凯亚_ZH"
    Dim 电话: 电话 = "凝光_ZH"
    sp(0) = Array( _
        "我妈什么时候知道自己脑子里有良性肿瘤？", _
        赵倩倩, 0.2, 0.6, 0.8, 1, "ZH", Null, "Astounded", "Text prompt", "", 0.7 _
    )
    sp(1) = Array( _
        "大概半年前，汪女士视线模糊，看东西不清楚，所以到医院就诊。我们建议尽早手术切除，但汪女士想保守治疗，不愿意冒险。手术费加住院费，最多五六万。做手术是最佳方案。", _
        医生, 0.2, 0.6, 0.8, 1, "ZH", Null, "Earnest", "Text prompt", "", 0.7 _
    )
    sp(2) = Array( _
        "……变成恶性肿瘤会怎么样？", _
        赵倩倩, 0.2, 0.6, 0.8, 1, "ZH", Null, "Astounded", "Text prompt", "", 0.7 _
    )
    sp(3) = Array( _
        "即便切除，也有可能多次复发。有病要趁早医，这两天做手术来得及。", _
        医生, 0.2, 0.6, 0.8, 1, "ZH", Null, "Earnest", "Text prompt", "", 0.7 _
    )
    sp(4) = Array( _
        "我妈醒来后，我会跟她商量。那手术的事等病人苏醒后再说。", _
        赵倩倩, 0.2, 0.6, 0.8, 1, "ZH", Null, "sad", "Text prompt", "", 0.7 _
    )
    sp(5) = Array( _
        "麻烦您待会儿去柜台补交住院费用。", _
        医生, 0.2, 0.6, 0.8, 1, "ZH", Null, "Earnest", "Text prompt", "", 0.7 _
    )
    sp(6) = Array( _
        "好。钱包里只有101.5块，我哪来的钱交住院费用？", _
        赵倩倩, 0.2, 0.6, 0.8, 1, "ZH", Null, "sad", "Text prompt", "", 0.7 _
    )
    sp(7) = Array( _
        "您好，您拨打的用户已关机，请稍后再拨……", _
        电话, 0.2, 0.6, 0.8, 1, "ZH", Null, "emotionless", "Text prompt", "", 0.7 _
    )
    sp(8) = Array( _
        "年终奖到账，我买了理财。我借钱交首付，这会儿背了一屁股债，没有闲钱。", _
        赵倩倩, 0.2, 0.6, 0.8, 1, "ZH", Null, "sad", "Text prompt", "", 0.7 _
    )
    Dim pathArr()
    ReDim pathArr(UBound(sp))
    For i = 0 To UBound(sp)
        json = SmHTTP.JsonData( _
            "data", sp(i), _
            "event_data", Null, _
            "fn_index", 0, _
            "session_hash", TestTool.GetGUID() _
        )
        ret = SmHTTP.HTTP_POST("https://v2.genshinvoice.top/run/predict", json)
        'Debug.Print ret
        If InStr(ret, "Success") > 0 Then url = "https://v2.genshinvoice.top/file=" & SmHTTP.GetJSON(ret, "data[1].name")
        pathArr(i) = Environ("TEMP") & "\" & i & ".wav"
        ret = SmHTTP.HTTP_DownLoad(url, pathArr(i))
    Next
    For i = 0 To UBound(sp)
        pRet = PlayWAV(pathArr(i)): Call StopWAV
    Next

    ' 断言
    Debug.Assert Expression(ret, "=", 1)
End Sub
