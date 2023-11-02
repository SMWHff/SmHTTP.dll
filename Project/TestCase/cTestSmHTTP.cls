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
    Debug.Assert Expression(ver, "=", "1.0.0.2")
End Sub


' 测试插件的路径
Public Sub test_getbasepath()
    Dim path: path = SmHTTP.GetBasePath()
    
    ' 断言
    Debug.Assert Expression(path, "=", "E:\AppData\Roaming\GitHub\【VisualBasic - 开发项目】SmHTTP\Project\SmHTTP.dll")
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
    Dim headers: headers = SmHTTP.headers( _
        "Accept", "*/*", _
        "Accept-Language", "zh-CN,zh;q=0.8", _
        "Host", "https://bbs.anjian.com", _
        "User-Agent", "Mozilla/4.0 (compatible; MSIE 9.0; Windows NT 6.1)", _
        "Content-Type", "application/x-www-form-urlencoded" _
    )
    
    ' 断言
    Debug.Assert Expression(headers, "=", "Accept:*/*" & vbCrLf & _
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
    Dim ret: ret = SmHTTP.Data( _
        "username", "SMWH", _
        "password", "123456" _
    )
    
    ' 断言
    Debug.Assert Expression(ret, "=", "username=SMWH&password=123456")
End Sub

' 构造 form-data 请求体
Public Sub test_form_data()
    Dim ret: ret = SmHTTP.FormData( _
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
    Debug.Assert Expression(json, "=", "{""empty"":"""",""null"":null,""int"":123,""float"":3.14,""bool"":true,""str"":""神梦无痕\""1042207232\"""",""array"":[1,3.14,true,null,""按键精灵""]}")
End Sub


' 测试 GET 请求
Public Sub test_http_get()
    Dim ret: ret = SmHTTP.HTTP_GET("http://localhost:8080/get")
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
    Dim Data: Data = String(10000, "a")
    Dim ret: ret = SmHTTP.HTTP_Request("POST", "http://localhost:8080/post", Data)
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
    Dim Data: Data = SmHTTP.FormData( _
        "@file", "C:\Users\SMWH\Pictures\Saved Pictures\纸飞机.png", "image/png", _
        "username", "SMWH", _
        "password", "123456" _
    )
    Call test_set_auto_param_array_on
    Dim ret: ret = SmHTTP.HTTP_POST("http://localhost:8080/post", Data)
    Dim username: username = SmHTTP.GetJSON(ret, "form.username[0]")
    Dim password: password = SmHTTP.GetJSON(ret, "form.password[0]")
    Call test_set_auto_param_array_off
    
    ' 断言
    Debug.Assert Expression(username, "=", "SMWH")
    Debug.Assert Expression(password, "=", "123456")
End Sub


' 测试按键精灵论坛签到
Public Sub test_bbs_anjian_signin()
    Dim ret, Cookies, headers
    Dim user: user = Environ("AJ_USER")
    Dim pass: pass = Environ("AJ_PASS")
    'Debug.Print user, pass
    Dim Data: Data = SmHTTP.Data( _
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
    ret = SmHTTP.HTTP_POST("http://bbs.anjian.com/login.aspx?referer=forumindex.aspx", Data)
    Debug.Assert InStr(ret, user)  ' 判断是否登录成功
    Cookies = SmHTTP.Getcookies()
    ' 打卡签到
    Data = SmHTTP.Data( _
        "signmessage", "签个到，每天心情都是美美哒~~按键精灵祝大家新年好运连连！！" _
    )
    headers = SmHTTP.headers( _
        "Referer", "http://bbs.anjian.com/" _
    )
    ret = SmHTTP.HTTP_POST("http://bbs.anjian.com/addsignin.aspx?infloat=1&inajax=1", Data, headers, Cookies)
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
    Dim headers: headers = SmHTTP.headers( _
        "Accept-Encoding", "identity" _
    )
    Dim ret: ret = SmHTTP.HTTP_POST("https://www.dedao.cn/pc/ledgers/notes/article_comment_list", json, headers)
    Dim message: message = SmHTTP.GetJSON(ret, "h.c")
    Call test_set_auto_param_array_off
    
    ' 断言
    Debug.Assert Expression(message, "=", 0)
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
    Debug.Assert Expression(LenB(ret), "=", 15427)
End Sub



' 测试百度翻译(英译中)
Public Sub test_Baidu_Translate()
    Dim enStr: enStr = "I Love You"
    Dim timestamp: timestamp = DateDiff("s", "1970-1-1 0:0:0", DateAdd("h", -8, Now)) & Right(CLng(Timer() * 1000), 3)
    Dim ret: ret = SmHTTP.HTTP_GET("https://www.baidu.com/")
    Dim L: L = InStr(1, ret, "var s_domain = {", vbTextCompare): Debug.Assert L > 0
    Dim R: R = InStr(L, ret, "};", vbTextCompare): Debug.Assert R > 0
    Dim s_domain: s_domain = Mid(ret, L, R - L + 1)
    Dim sensearch: sensearch = SmHTTP.GetJSON(s_domain, "ssllist['sensearch.baidu.com']")
    ret = SmHTTP.HTTP_GET("http://" & sensearch & "/sensearch/selecttext?cb=jQuery_Fun_" & timestamp & "&q=" & enStr & "&_=" & timestamp)
    Dim errno: errno = SmHTTP.GetJSON(ret, "errno")
    Dim result: result = SmHTTP.GetJSON(ret, "data.result")
    
    ' 断言
    Debug.Assert Expression(errno, "=", 0)
    Debug.Assert Expression(result, "=", "我爱你")
End Sub


' 百度统计
Public Sub test_Baidu_tongji()
    Dim Data: Data = SmHTTP.Data( _
        "cc", 1, _
        "ck", 1, _
        "cl", "32-bit", _
        "ds", "1024x1024", _
        "et", 0, _
        "ep", 0, _
        "fl", "11.0" _
    )
    Dim headers: headers = SmHTTP.headers( _
        "Referer", "https://www.SmHTTP.com/", _
        "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/118.0.0.0 Safari/537.36 Edg/118.0.2088.46 VB/6.0", _
        "Cookie", "HMACCOUNT=SMWHff-VB6-GUID" _
    )
    Dim url: url = "https://hm.baidu.com/hm.gif?si=212f0c3f981a440699efc710f3043368&" & Data
    Call SmHTTP.HTTP_GET(url, headers)
    Dim status: status = SmHTTP.GetStatus()
    
    ' 断言
    Debug.Assert Expression(status, "=", 200)
End Sub


' 测试下载文件成功
Public Sub test_http_download()
    Dim url: url = "https://bbs.anjian.com/templates/default/Images/logo1.png"
    Dim path: path = Environ("TEMP") & "\logo1.png"
    Dim ret: ret = SmHTTP.HTTP_DownLoad(url, path)

    ' 断言
    Debug.Assert Expression(ret, "=", 1)
End Sub


' 测试下载文件，无效的URL
Public Sub test_http_download_invalid_url()
    Dim url: url = "http://invalid.url"
    Dim path: path = Environ("TEMP") & "\file.txt"
    Dim ret: ret = SmHTTP.HTTP_DownLoad(url, path)

    ' 断言
    Debug.Assert Expression(ret, "=", 0)
End Sub


' 测试下载文件，空的URL
Public Sub test_http_download_empty_url()
    Dim url: url = ""
    Dim path: path = Environ("TEMP") & "\EmptyURLFile.txt"
    Dim ret: ret = SmHTTP.HTTP_DownLoad(url, path)

    ' 断言
    Debug.Assert Expression(ret, "=", 0)
End Sub


' 测试下载文件，无效的保存路径
Public Sub test_http_download_invalid_path()
    Dim url: url = "https://bbs.anjian.com/templates/default/Images/logo1.png"
    Dim path: path = "X:\Invalid\Path\file.txt"
    Dim ret: ret = SmHTTP.HTTP_DownLoad(url, path)

    ' 断言
    Debug.Assert Expression(ret, "=", 0)
End Sub


' 测试下载文件，空的保存路径
Public Sub test_http_download_empty_path()
    Dim url: url = "https://bbs.anjian.com/templates/default/Images/logo1.png"
    Dim path: path = ""
    Dim ret: ret = SmHTTP.HTTP_DownLoad(url, path)

    ' 断言
    Debug.Assert Expression(ret, "=", 0)
End Sub


'' 测试下载文件加强版
'Public Sub test_http_downloadEx()
'    Dim url: url = "https://bbs.anjian.com/templates/default/Images/logo1.png|http://res2.mobileanjian.com/Advert/20230531/139355dc1da4409f80f852b6523f8051.png"
'    Dim path: path = Environ("TEMP") & "\logo1.png|" & Environ("TEMP") & "\139355dc1da4409f80f852b6523f8051.png"
'    Dim ret: ret = SmHTTP.HTTP_DownLoadEx(url, path)
'
'    ' 断言
'    Debug.Assert Expression(ret, "=", "")
'End Sub
'
'
'' 测试下载文件加强版，无效的URL
'Public Sub test_http_downloadEx_invalid_url()
'    Dim url: url = "https://bbs.anjian.com/templates/default/Images/logo1.png|http://res2.mobileanjian.com/Advert/20230531/139355dc1da4409f80f852b6523f8051.png|http://invalid.url"
'    Dim path: path = Environ("TEMP") & "\logo1.png|" & Environ("TEMP") & "\139355dc1da4409f80f852b6523f8051.png|" & Environ("TEMP") & "\invalid.url"
'    Dim ret: ret = SmHTTP.HTTP_DownLoadEx(url, path)
'
'    ' 断言
'    Debug.Assert Expression(ret, "=", "http://invalid.url," & Environ("TEMP") & "\invalid.url")
'End Sub
'
'
'' 测试下载文件加强版，空的URL
'Public Sub test_http_downloadEx_empty_url()
'    Dim url: url = "https://bbs.anjian.com/templates/default/Images/logo1.png|http://res2.mobileanjian.com/Advert/20230531/139355dc1da4409f80f852b6523f8051.png"
'    Dim path: path = Environ("TEMP") & "\logo1.png|" & Environ("TEMP") & "\139355dc1da4409f80f852b6523f8051.png|" & Environ("TEMP") & "\EmptyURLFileEx.txt"
'    Dim ret: ret = SmHTTP.HTTP_DownLoadEx(url, path)
'
'    ' 断言
'    Debug.Assert Expression(ret, "=", "," & Environ("TEMP") & "\EmptyURLFileEx.txt")
'End Sub
'
'
'' 测试下载文件加强版，无效的保存路径
'Public Sub test_http_downloadEx_invalid_path()
'    Dim url: url = "https://bbs.anjian.com/templates/default/Images/logo1.png|http://res2.mobileanjian.com/Advert/20230531/139355dc1da4409f80f852b6523f8051.png|http://files.anjian.com/avatars/avatars/upload/002/69/99/98_avatar_middle.jpg"
'    Dim path: path = Environ("TEMP") & "\logo1.png|" & Environ("TEMP") & "\139355dc1da4409f80f852b6523f8051.png|X:\Invalid\Path\fileEx.txt"
'    Dim ret: ret = SmHTTP.HTTP_DownLoadEx(url, path)
'
'    ' 断言
'    Debug.Assert Expression(ret, "=", "http://files.anjian.com/avatars/avatars/upload/002/69/99/98_avatar_middle.jpg,X:\Invalid\Path\fileEx.txt")
'End Sub
'
'
'' 测试下载文件加强版，空的保存路径
'Public Sub test_http_downloadEx_empty_path()
'    Dim url: url = "https://bbs.anjian.com/templates/default/Images/logo1.png|http://res2.mobileanjian.com/Advert/20230531/139355dc1da4409f80f852b6523f8051.png|http://files.anjian.com/avatars/avatars/upload/002/69/99/98_avatar_middle.jpg"
'    Dim path: path = Environ("TEMP") & "\logo1.png|" & Environ("TEMP") & "\139355dc1da4409f80f852b6523f8051.png|"
'    Dim ret: ret = SmHTTP.HTTP_DownLoadEx(url, path)
'
'    ' 断言
'    Debug.Assert Expression(ret, "=", "http://files.anjian.com/avatars/avatars/upload/002/69/99/98_avatar_middle.jpg,")
'End Sub
