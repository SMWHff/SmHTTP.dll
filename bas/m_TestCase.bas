Attribute VB_Name = "TestCase"
'Option Explicit
' ----------[��������]----------
Dim SmHTTP As New SmHTTP

' ���Բ����ʼ��
Sub test_init()
    Debug.Assert SmHTTP.Init() = 1
End Sub


' ���Բ���汾��
 Sub test_ver()
    Debug.Assert SmHTTP.ver() = "0.0.0.11"
End Sub


' ���Բ����·��
Sub test_getbasepath()
    Debug.Assert SmHTTP.GetBasePath() = "E:\AppData\Roaming\GitHub\��VisualBasic - ������Ŀ��SmHTTP\SmHTTP.dll"
End Sub


' ���Բ������ID
Sub test_getid()
    Debug.Assert SmHTTP.GetID() > 0
End Sub


' �������ÿ����Զ�ʶ���Ӧ�������
Sub test_set_auto_param_array_on()
    Debug.Assert SmHTTP.SetAutoParamArray(True) = 1
End Sub


' �������ùر��Զ�ʶ���Ӧ�������
Sub test_set_auto_param_array_off()
    Debug.Assert SmHTTP.SetAutoParamArray(False) = 1
End Sub


' ���Թ�������ͷ
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


' ����Cookies
Sub test_cookies()
    Dim Ret: Ret = SmHTTP.Cookies( _
        "_ga", "GA1.2.1206281266.1647004488", _
        "BAIDUID_BFESS", "0F068EE7974C72C13A37B02D9855DD1C:SL=0:NR=10:FG=1", _
        "ZFY", "O7YgtvLEvTKeDmPLbV8Nbwq3xYhFAOP9m9A:BTtT0AkQ:C" _
    )
    Debug.Assert Ret = "_ga=GA1.2.1206281266.1647004488;BAIDUID_BFESS=0F068EE7974C72C13A37B02D9855DD1C:SL=0:NR=10:FG=1;ZFY=O7YgtvLEvTKeDmPLbV8Nbwq3xYhFAOP9m9A:BTtT0AkQ:C"
End Sub

' ���� URLData ������
Sub test_data()
    Dim Ret: Ret = SmHTTP.Data( _
        "username", "SMWH", _
        "password", "123456" _
    )
    Debug.Assert Ret = "username=SMWH&password=123456"
End Sub

' ���� form-data ������
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

' ���� JSON ������
Sub test_json_data()
    Debug.Assert SmHTTP.JSONData( _
        "Empty", Empty, _
        "null", Null, _
        "int", 123, _
        "float", 3.14, _
        "bool", True, _
        "str", "�����޺�""1042207232""", _
        "array", Array(1, 3.14, True, Null, "��������") _
    ) = "{""Empty"":null,""null"":null,""int"":123,""float"":3.14,""bool"":true,""str"":""�����޺�\""1042207232\"""",""array"":[1,3.14,true,null,""��������""]}"
End Sub


' ���� GET ����
Sub test_http_get()
    Dim Ret: Ret = SmHTTP.HTTP_GET("http://127.0.0.1:8080/get")
    Debug.Assert SmHTTP.GetStatus() = 200
End Sub


' ���� POST ����
Sub test_http_post()
    Dim Ret: Ret = SmHTTP.HTTP_POST("http://127.0.0.1:8080/post", "username=SMWH&password=123456", "Content-Type: application/x-www-form-urlencoded")
    Debug.Assert SmHTTP.GetStatus() = 200
End Sub


' ���� HEAD ����
Sub test_http_head()
    Dim Ret: Ret = SmHTTP.HTTP_HEAD("http://127.0.0.1:8080/head")
    Debug.Assert SmHTTP.GetStatus() = 200
End Sub


' ���� OPTIONS ����
Sub test_http_options()
    Dim Ret: Ret = SmHTTP.HTTP_OPTIONS("http://127.0.0.1:8080/options")
    Debug.Assert SmHTTP.GetStatus() = 200
End Sub


' ���� PATCH ����
Sub test_http_patch()
    Dim Ret: Ret = SmHTTP.HTTP_PATCH("http://127.0.0.1:8080/patch")
    Debug.Assert SmHTTP.GetStatus() = 200
End Sub


' ���� PUT ����
Sub test_http_put()
    Dim Ret: Ret = SmHTTP.HTTP_PUT("http://127.0.0.1:8080/put")
    Debug.Assert SmHTTP.GetStatus() = 200
End Sub


' ���� DELETE ����
Sub test_http_delete()
    Dim Ret: Ret = SmHTTP.HTTP_DELETE("http://127.0.0.1:8080/delete")
    Debug.Assert SmHTTP.GetStatus() = 200
End Sub


' ���� Request ����
Sub test_http_request()
    Call test_set_auto_param_array_on
    Dim Ret: Ret = SmHTTP.HTTP_Request("GET", "http://www.baidu.com", "UTF-8")
    Call test_set_auto_param_array_off
End Sub


' ���Դ���IP
Sub test_http_proxy()
    Call test_set_auto_param_array_on
    Dim Ret: Ret = SmHTTP.HTTP_Request("GET", "http://www.bathome.net/s/ip.php", "120.196.186.248:9091")
    Debug.Assert Ret = "120.196.186.248"
    Call test_set_auto_param_array_off
End Sub


' ���Է���JSON
Sub test_http_ret_json()
    Call test_set_auto_param_array_on
    Dim URL: URL = "https://shenzhen.1200.com.cn/api/sale/querySecondHouse?cityId=11&pageSize=30&orderBy=DEFAULT&pageIndex=1&showAppreciateFlag=1"
    Dim Ret: Ret = SmHTTP.HTTP_Request("GET", URL)
    Dim message: message = SmHTTP.GetJSON(Ret, "message")
    Debug.Assert message = "�����ɹ�"
    Call test_set_auto_param_array_off
End Sub


' �����ϴ��ļ�
Sub test_http_upload_file()
    Dim URL: URL = "http://127.0.0.1:8080/post"
    Dim Data: Data = SmHTTP.FormData( _
        "@file", "C:\Users\SMWH\Pictures\Saved Pictures\ֽ�ɻ�.png", "image/png", _
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


' ���԰���������̳ǩ��
Sub test_bbs_anjian_signin()
    Dim Ret, Cookies, Headers
    Dim user: user = Environ("AJ_USER")
    Dim pass: pass = Environ("AJ_PASS")
    Dim Data: Data = SmHTTP.Data( _
        "username", user, _
        "password", pass, _
        "question", "0", _
        "answer", "", _
        "templateid", "0", _
        "login", "", _
        "expires", "43200" _
    )
    'SmHTTP.���Կ��� = True
    Call test_set_auto_param_array_on
    ' ��¼��̳�˺�
    Ret = SmHTTP.HTTP_POST("http://bbs.anjian.com/login.aspx?referer=forumindex.aspx", Data)
    Debug.Assert InStr(Ret, user)  ' �ж��Ƿ��¼�ɹ�
    Cookies = SmHTTP.GetCookies()
    ' ��ǩ��
    Data = SmHTTP.Data( _
        "signmessage", "ǩ������ÿ�����鶼��������~~��������ף������������������" _
    )
    Headers = SmHTTP.Headers( _
        "Referer", "http://bbs.anjian.com/" _
    )
    Ret = SmHTTP.HTTP_POST("http://bbs.anjian.com/addsignin.aspx?infloat=1&inajax=1", Data, Headers, Cookies)
    Debug.Assert InStr(Ret, "��ϲ����ȡ����ǩ������") Or InStr(Ret, "������Ѿ�ǩ������")  ' �ж��Ƿ�ǩ���ɹ�
    Call test_set_auto_param_array_off
End Sub


' ����ǿ�Ʒ���������δѹ��������
Sub test_http_ret_not_gzip()
    Call test_set_auto_param_array_on
    Dim params: params = SmHTTP.Data( _
        "date", "", _
        "lotCode", "10037" _
    )
    Dim Headers: Headers = SmHTTP.Headers( _
        "Accept-Encoding", "identity" _
    )
    Dim Ret: Ret = SmHTTP.HTTP_GET("https://1680688kai.co/api/pks/getPksHistoryList.do?" & params, Headers)
    Dim message: message = SmHTTP.GetJSON(Ret, "message")
    Debug.Assert message = "�����ɹ�"
    Call test_set_auto_param_array_off
End Sub
