Attribute VB_Name = "TestCase"
'Option Explicit
' ----------[变量定义]----------
Dim SmHTTP As New SmHTTP


' 运行测试服务器
Sub test_run_server()
    Dim ws, ProceName, Proce
    Proce = False
    ProceName = "go-httpbin-win.exe" '判断的进程
    Set ws = CreateObject("wscript.shell")
    For Each x In GetObject("winmgmts:").instancesof("win32_process")
        If UCase(x.Name) = UCase(ProceName) Then
            Proce = True
            Exit For
        End If
    Next
    Set ws = Nothing
    If Not Proce Then
        Shell ProceName
    End If
End Sub


' 测试插件初始化
Sub test_init()
'    Debug.Assert SmHTTP.Init() = 1
End Sub


' 测试插件版本号
 Sub test_ver()
    'Debug.Assert SmHTTP.ver() = "0.0.0.16"
End Sub


' 测试插件的路径
Sub test_getbasepath()
    Debug.Assert SmHTTP.GetBasePath() = "E:\AppData\Roaming\GitHub\【VisualBasic - 开发项目】SmHTTP\SmHTTP.dll"
End Sub


' 测试插件对象ID
Sub test_getid()
    Debug.Assert SmHTTP.GetID() > 0
End Sub


' 测试设置开启自动识别对应传入参数
Sub test_set_auto_param_array_on()
    Debug.Assert SmHTTP.SetAutoParamArray(True) = 1
End Sub


' 测试设置关闭自动识别对应传入参数
Sub test_set_auto_param_array_off()
    Debug.Assert SmHTTP.SetAutoParamArray(False) = 1
End Sub


' 测试构造请求头
Sub test_headers()
    Debug.Assert SmHTTP.Headers( _
        "Accept", "*/*", _
        "Accept-Language", "zh-CN,zh;q=0.8", _
        "Host", "https://bbs.anjian.com", _
        "User-Agent", "Mozilla/4.0 (compatible; MSIE 9.0; Windows NT 6.1)", _
        "Content-Type", "application/x-www-form-urlencoded" _
    ) = "Accept:*/*" & vbCrLf & _
        "Accept-Language:zh-CN,zh;q=0.8" & vbCrLf & _
        "Host:https://bbs.anjian.com" & vbCrLf & _
        "User-Agent:Mozilla/4.0 (compatible; MSIE 9.0; Windows NT 6.1)" & vbCrLf & _
        "Content-Type:application/x-www-form-urlencoded"
End Sub


' 构造Cookies
Sub test_cookies()
    Dim Ret: Ret = SmHTTP.Cookies( _
        "_ga", "GA1.2.1206281266.1647004488", _
        "BAIDUID_BFESS", "0F068EE7974C72C13A37B02D9855DD1C:SL=0:NR=10:FG=1", _
        "ZFY", "O7YgtvLEvTKeDmPLbV8Nbwq3xYhFAOP9m9A:BTtT0AkQ:C" _
    )
    Debug.Assert Ret = "_ga=GA1.2.1206281266.1647004488;BAIDUID_BFESS=0F068EE7974C72C13A37B02D9855DD1C:SL=0:NR=10:FG=1;ZFY=O7YgtvLEvTKeDmPLbV8Nbwq3xYhFAOP9m9A:BTtT0AkQ:C"
End Sub

' 构造 URLData 请求体
Sub test_data()
    Dim Ret: Ret = SmHTTP.Data( _
        "username", "SMWH", _
        "password", "123456" _
    )
    Debug.Assert Ret = "username=SMWH&password=123456"
End Sub

' 构造 form-data 请求体
Sub test_form_data()
    Dim Ret: Ret = SmHTTP.FormData( _
        "username", "SMWH", _
        "password", "123456" _
    )
    Debug.Assert Ret = "--WebKitFormBoundarySmHTTPSMWHff" & vbCrLf & _
                      "Content-Disposition: form-data; name=""username""" & vbCrLf & _
                      "" & vbCrLf & _
                      "SMWH" & vbCrLf & _
                      "--WebKitFormBoundarySmHTTPSMWHff" & vbCrLf & _
                      "Content-Disposition: form-data; name=""password""" & vbCrLf & _
                      "" & vbCrLf & _
                      "123456" & vbCrLf & _
                      "--WebKitFormBoundarySmHTTPSMWHff--" & vbCrLf
End Sub

' 构造 JSON 请求体
Sub test_json_data()
    Debug.Assert SmHTTP.JSONData( _
        "Empty", Empty, _
        "null", Null, _
        "int", 123, _
        "float", 3.14, _
        "bool", True, _
        "str", "神梦无痕""1042207232""", _
        "array", Array(1, 3.14, True, Null, "按键精灵") _
    ) = "{""Empty"":null,""null"":null,""int"":123,""float"":3.14,""bool"":true,""str"":""神梦无痕\""1042207232\"""",""array"":[1,3.14,true,null,""按键精灵""]}"
End Sub


' 测试 GET 请求
Sub test_http_get()
    Dim Ret: Ret = SmHTTP.HTTP_GET("http://127.0.0.1:8080/get")
    Debug.Assert SmHTTP.GetStatus() = 200
End Sub


' 测试 POST 请求
Sub test_http_post()
    Dim Ret: Ret = SmHTTP.HTTP_POST("http://127.0.0.1:8080/post", "username=SMWH&password=123456", "Content-Type: application/x-www-form-urlencoded")
    Debug.Assert SmHTTP.GetStatus() = 200
End Sub


' 测试 HEAD 请求
Sub test_http_head()
    Dim Ret: Ret = SmHTTP.HTTP_HEAD("http://127.0.0.1:8080/head")
    Debug.Assert SmHTTP.GetStatus() = 200
End Sub


' 测试 OPTIONS 请求
Sub test_http_options()
    Dim Ret: Ret = SmHTTP.HTTP_OPTIONS("http://127.0.0.1:8080/options")
    Debug.Assert SmHTTP.GetStatus() = 200
End Sub


' 测试 PATCH 请求
Sub test_http_patch()
    Dim Ret: Ret = SmHTTP.HTTP_PATCH("http://127.0.0.1:8080/patch")
    Debug.Assert SmHTTP.GetStatus() = 200
End Sub


' 测试 PUT 请求
Sub test_http_put()
    Dim Ret: Ret = SmHTTP.HTTP_PUT("http://127.0.0.1:8080/put")
    Debug.Assert SmHTTP.GetStatus() = 200
End Sub


' 测试 DELETE 请求
Sub test_http_delete()
    Dim Ret: Ret = SmHTTP.HTTP_DELETE("http://127.0.0.1:8080/delete")
    Debug.Assert SmHTTP.GetStatus() = 200
End Sub


' 测试 Request 请求
Sub test_http_request()
    Call test_set_auto_param_array_on
    Dim Ret: Ret = SmHTTP.HTTP_Request("GET", "https://www.bing.com/ipv6test/test?FORM=MONITR", "UTF-8")
    Debug.Print SmHTTP.GetCookieByName(SmHTTP.Getcookies(), "_SS")
    Call test_set_auto_param_array_off
End Sub


' 测试代理IP
Sub test_http_proxy()
    Call test_set_auto_param_array_on
    Dim Ret: Ret = SmHTTP.HTTP_Request("GET", "http://www.bathome.net/s/ip.php", "120.196.186.248:9091")
    'Debug.Assert Ret = "120.196.186.248"
    Call test_set_auto_param_array_off
End Sub


' 测试代理IP(带身份认证)
Sub test_http_proxy_auth()
'    Call test_set_auto_param_array_on
'    Dim Ret: Ret = SmHTTP.HTTP_Request("GET", "http://www.bathome.net/s/ip.php", "112.5.56.2:9091")
'    Debug.Assert Ret = "112.5.56.2"
'    Call test_set_auto_param_array_off
End Sub


' 测试身份认证
Sub test_http_auth_basic()
    Call test_set_auto_param_array_on
    Dim Ret: Ret = SmHTTP.HTTP_GET("https://ssr3.scrape.center/", "BASIC", "admin", "admin")
    Debug.Assert SmHTTP.GetStatus() = 200
    Call test_set_auto_param_array_off
End Sub


' 测试返回JSON
Sub test_http_ret_json()
    Call test_set_auto_param_array_on
    Dim URL: URL = "https://shenzhen.1200.com.cn/api/sale/querySecondHouse?cityId=11&pageSize=30&orderBy=DEFAULT&pageIndex=1&showAppreciateFlag=1"
    Dim Ret: Ret = SmHTTP.HTTP_Request("GET", URL)
    Dim message: message = SmHTTP.GetJSON(Ret, "message")
    Debug.Assert message = "操作成功"
    Call test_set_auto_param_array_off
End Sub


' 测试上传文件
Sub test_http_upload_file()
    Dim URL: URL = "http://127.0.0.1:8080/post"
    Dim Data: Data = SmHTTP.FormData( _
        "@file", "C:\Users\SMWH\Pictures\Saved Pictures\纸飞机.png", "image/png", _
        "username", "SMWH", _
        "password", "123456" _
    )
    Call test_set_auto_param_array_on
    Dim Ret: Ret = SmHTTP.HTTP_POST(URL, Data)
    Dim username: username = SmHTTP.GetJSON(Ret, "form.username[0]")
    Dim password: password = SmHTTP.GetJSON(Ret, "form.password[0]")
    Debug.Assert username = "SMWH"
    Debug.Assert password = "123456"
    Call test_set_auto_param_array_off
End Sub


' 测试按键精灵论坛签到
Sub test_bbs_anjian_signin()
    Dim Ret, Cookies, Headers
    Dim user: user = Environ("AJ_USER")
    Dim pass: pass = Environ("AJ_PASS")
    Debug.Print user, pass
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
    Ret = SmHTTP.HTTP_POST("http://bbs.anjian.com/login.aspx?referer=forumindex.aspx", Data)
    Debug.Assert InStr(Ret, user)  ' 判断是否登录成功
    Cookies = SmHTTP.Getcookies()
    ' 打卡签到
    Data = SmHTTP.Data( _
        "signmessage", "签个到，每天心情都是美美哒~~按键精灵祝大家新年好运连连！！" _
    )
    Headers = SmHTTP.Headers( _
        "Referer", "http://bbs.anjian.com/" _
    )
    Ret = SmHTTP.HTTP_POST("http://bbs.anjian.com/addsignin.aspx?infloat=1&inajax=1", Data, Headers, Cookies)
    Debug.Assert InStr(Ret, "恭喜您获取本日签到奖励") Or InStr(Ret, "你今天已经签到过了")  ' 判断是否签到成功
    Call test_set_auto_param_array_off
End Sub


' 测试强制服务器返回未压缩的内容
Sub test_http_ret_not_gzip()
'    Call test_set_auto_param_array_on
'    Dim params: params = SmHTTP.Data( _
'        "date", "", _
'        "lotCode", "10037" _
'    )
'    Dim Headers: Headers = SmHTTP.Headers( _
'        "Accept-Encoding", "identity" _
'    )
'    Dim Ret: Ret = SmHTTP.HTTP_GET("https://1680688kai.co/api/pks/getPksHistoryList.do?" & params, Headers)
'    Debug.Print Ret
'    Dim message: message = SmHTTP.GetJSON(Ret, "message")
'    Debug.Assert message = "操作成功"
'    Call test_set_auto_param_array_off
End Sub


' 测试获取QQ昵称
Sub test_get_qq_nick_name()
'    Dim qq: qq = 1042207232
'    Call test_set_auto_param_array_on
'    Dim Ret: Ret = SmHTTP.HTTP_GET("https://r.qzone.qq.com/fcg-bin/cgi_get_portrait.fcg?uins=" & qq, "GBK")
'    Dim name: name = SmHTTP.GetJSON(Ret, "[" + CStr(qq) + "][6]")
'    Debug.Assert name = "神梦无痕"
'    Call test_set_auto_param_array_off
End Sub


' 测试下载QQ头像
Sub test_download_qq_avatar()
'    Dim qq: qq = 1042207232
'    Call test_set_auto_param_array_on
'    Dim Ret: Ret = SmHTTP.HTTP_GET("https://r.qzone.qq.com/fcg-bin/cgi_get_portrait.fcg?uins=" & qq, "GBK")
'    Dim img_src: img_src = SmHTTP.GetJSON(Ret, "[" + CStr(qq) + "][0]")
'    Ret = SmHTTP.HTTP_GET(img_src)
'    Debug.Assert TypeName(Ret) = "Byte()"
'    Debug.Assert Len(Ret) = 7942
'    Call test_set_auto_param_array_off
End Sub


' 测试百度翻译(英译中)
Sub test_Baidu_Translate()
    Dim enStr: enStr = "I Love You"
    Dim timestamp: timestamp = DateDiff("s", "1970-1-1 0:0:0", DateAdd("h", -8, Now)) & Right(CLng(Timer() * 1000), 3)
    Dim Ret: Ret = SmHTTP.HTTP_GET("https://www.baidu.com/")
    Dim L: L = InStr(1, Ret, "var s_domain = {", vbTextCompare): Debug.Assert L > 0
    Dim R: R = InStr(L, Ret, "};", vbTextCompare): Debug.Assert R > 0
    Dim s_domain: s_domain = Mid(Ret, L, R - L + 1)
    Dim sensearch: sensearch = SmHTTP.GetJSON(s_domain, "ssllist['sensearch.baidu.com']")
    Ret = SmHTTP.HTTP_GET("http://" & sensearch & "/sensearch/selecttext?cb=jQuery_Fun_" & timestamp & "&q=" & enStr & "&_=" & timestamp)
    Debug.Assert SmHTTP.GetJSON(Ret, "errno") = 0
    Debug.Assert SmHTTP.GetJSON(Ret, "data.result") = "我爱你"
End Sub


' 百度统计
Sub test_Baidu_tongji()
    Dim Data: Data = SmHTTP.Data( _
        "cc", 1, _
        "ck", 1, _
        "cl", "32-bit", _
        "ds", "1024*1024", _
        "et", 0, _
        "ep", 0, _
        "fl", "11.0" _
    )
End Sub
