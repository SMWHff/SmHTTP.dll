Attribute VB_Name = "m_Assert"
Option Explicit
' ==========【API声明】==========
' 注册消息
Private Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long


' ----------[变量定义]----------
Dim Comparator As New Comparator


' 断言
Public Function AssertThat(ByVal Expected As Variant, ByVal Operator As String, ByVal Expression As Variant, Optional ByVal FailMessage As String) As Long
    If CallByName(Comparator, Comparator.OperatorToProcName(Operator), VbMethod, Expected, Expression) Then
        AssertThat = 1
    Else
        Call AssertionError("AssertThat()", FailMessage, Expected, Operator, Expression)
    End If
End Function


' 断言异常错误
Private Function AssertionError(ByVal Source As String, ByVal Message As String, Optional ByVal a As Variant = vbNullString, Optional ByVal b As Variant = vbNullString, Optional ByVal c As Variant = vbNullString)
    Call Throw(Source, AssertException(Message, a, b, c))
End Function


' 异常处理
Private Function AssertException(ByVal Message As String, ByVal a As Variant, ByVal b As Variant, ByVal c As Variant) As String
    Dim a_Ptr As Long
    Dim c_Ptr As Long
    
    If Len(Message) = 0 Then
        '------[ a ]------
        If IsObject(a) Then
            a = "<" & TypeName(a) & "@" & ObjPtr(a) & ">"
        ElseIf IsNull(a) Or IsEmpty(a) Then
            a = TypeName(a)
        ElseIf IsArray(a) Then
            a = "[" & Join(a, ", ") & "]"
        ElseIf VarType(a) = vbString And StrPtr(a) <> 0 Then
            a = "“" & a & "”"
        Else
            a = a
        End If
        '------[ b ]------
        Select Case LCase(b)
        Case "in"
            b = "In"
        Case "not in"
            b = "Not In"
        Case "is"
            b = "Is"
        Case "not is"
            b = "Not Is"
        Case "not is"
            b = "Not Is"
        End Select
        '------[ c ]------
        If IsObject(c) Then
            c = "<" & TypeName(c) & "@" & ObjPtr(c) & ">"
        ElseIf IsNull(a) Or IsEmpty(a) Then
            a = TypeName(a)
        ElseIf IsArray(c) Then
            c = "[" & Join(c, ", ") & "]"
        ElseIf InStr(LCase(b), "between") > 0 Then
            c = c
        ElseIf VarType(c) = vbString And StrPtr(c) <> 0 Then
            c = "“" & c & "”"
        Else
            c = c
        End If
        Message = "出错，表达式不成立：" & a & " " & b & " " & c
    End If
    AssertException = Message
End Function


' 抛出异常
Private Function Throw(ByVal Source As String, ByVal Description As String, Optional ByVal HelpFile As String, Optional ByVal HelpContext As Long)
    Dim lMsg As Long
    
    If App.LogMode = 0 Then
        Debug.Print Description
    Else
        ' 注册消息。
        lMsg = RegisterWindowMessage(Description)
        ' 抛出异常
        Err.Raise vbObjectError + lMsg, App.EXEName & "." & Source, Description, HelpFile, HelpContext
    End If
End Function
