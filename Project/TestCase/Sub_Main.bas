Attribute VB_Name = "Sub_Main"
Dim TestSmHTTP As New cTestSmHTTP

Sub Main()
    Dim testNameAll As Variant
    Dim i           As Long
    Dim testName    As String
    Dim result      As String
    
    'Call CallByName(TestSmHTTP, "test_http_downloadEx_TTS_yuanshen_ebook", VbMethod): End
    
    ' ִ�в�������
    testNameAll = TestTool.GetFuncNameAll(App.Path & "\TestCase\cTestSmHTTP.cls")
    For i = 0 To UBound(testNameAll)
        testName = testNameAll(i)
        Debug.Print "[" & Format$(i / UBound(testNameAll), "00%") & "]", testName
        Call CallByName(TestSmHTTP, testName, VbMethod)
    Next
    
    Debug.Print "�ű��Ѿ�ֹͣ����"
End Sub


