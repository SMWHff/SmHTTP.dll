Attribute VB_Name = "Sub_Main"
Dim TestSmHTTP As New cTestSmHTTP

Sub Main()
    Dim testNameAll As Variant
    Dim i           As Long
    Dim testName    As String
    Dim result      As String
    
'    Call CallByName(TestSmHTTP, "test_base64_decode_str", VbMethod): End
    
    ' 执行测试用例
    testNameAll = TestTool.GetFuncNameAll(App.Path & "\TestCase\cTestSmHTTP.cls")
    For i = 0 To UBound(testNameAll)
        testName = testNameAll(i)
        If InStr(testName, "skip") > 0 Then
            Debug.Print "[" & Format$(i / UBound(testNameAll), "00%") & "]", "[SKIP]", testName
        Else
            Call CallByName(TestSmHTTP, testName, VbMethod)
            Debug.Print "[" & Format$(i / UBound(testNameAll), "00%") & "]", "[PASS]", testName
        End If
    Next
    Debug.Print "脚本已经停止运行"
End Sub


