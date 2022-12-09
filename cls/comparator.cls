VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "comparator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
' ----------[常量定义]----------
Const vbSpace = " "
Const vbNone = ""


Public Function OperatorToProcName(ByRef Operator As String) As String
    Dim op As String

    op = Replace(LCase(Operator), vbSpace, vbNone)
    Select Case op
    Case "=", "＝"
        Operator = "="
        OperatorToProcName = "Equal"
    Case "!=", "<>", "≠"
        Operator = "≠"
        OperatorToProcName = "NotEqual"
    Case "==", "≡"
        Operator = "≡"
        OperatorToProcName = "Equivalent"
    Case ">", "＞"
        Operator = "＞"
        OperatorToProcName = "Greater"
    Case "<", "＜"
        Operator = "＜"
        OperatorToProcName = "Less"
    Case ">=", "≥"
        Operator = "≥"
        OperatorToProcName = "GreaterEqual"
    Case "<=", "≤"
        Operator = "≤"
        OperatorToProcName = "LessEqual"
    Case "~=", "≈"
        Operator = "≈"
        OperatorToProcName = "Approximately"
    Case "in"
        Operator = "In"
        OperatorToProcName = "Contains"
    Case "notin"
        Operator = "Not In"
        OperatorToProcName = "NotContains"
    Case "is"
        Operator = "Is"
        OperatorToProcName = "SameInstance"
    Case "notis"
        Operator = "Not Is"
        OperatorToProcName = "NotSameInstance"
    Case "match"
        Operator = "Match"
        OperatorToProcName = "Matches"
    Case "notmatch"
        Operator = "Not Match"
        OperatorToProcName = "NotMatches"
    Case Else
        Operator = TitleCase(Operator)
        OperatorToProcName = Operator
    End Select
End Function



' 等于 =
Public Function Equal(ByVal Expected As Variant, ByVal Expression As Variant, Optional ByRef FailMessage As String, Optional IsEqv As Boolean) As Boolean
    If IsArray(Expected) And IsArray(Expression) Then
        If IsEqv Then
            Equal = Array_Eqv_Array(Expected, Expression)
        Else
            Equal = Array_Eq_Array(Expected, Expression)
        End If
    ElseIf IsNull(Expected) Then
        Equal = IsNull(Expression)
    Else
        Equal = (Expected = Expression)
    End If
End Function



' 不等于 <>
Public Function NotEqual(ByVal Expected As Variant, ByVal Expression As Variant, Optional ByRef FailMessage As String) As Boolean
    NotEqual = Not Equal(Expected, Expression)
End Function



' 全等于 ==
Public Function Equivalent(ByVal Expected As Variant, ByVal Expression As Variant, Optional ByRef FailMessage As String) As Boolean
    If VarType(Expected) = VarType(Expression) Then
        Equivalent = Equal(Expected, Expression, IsEqv:=True)
    End If
End Function



' 大于 >
Public Function Greater(ByVal Expected As Variant, ByVal Expression As Variant, Optional ByRef FailMessage As String) As Boolean
    Greater = (Expected > Expression)
End Function



' 小于 <
Public Function Less(ByVal Expected As Variant, ByVal Expression As Variant, Optional ByRef FailMessage As String) As Boolean
    Less = (Expected < Expression)
End Function



' 大于等于 >=
Public Function GreaterEqual(ByVal Expected As Variant, ByVal Expression As Variant, Optional ByRef FailMessage As String) As Boolean
    GreaterEqual = (Expected >= Expression)
End Function



' 小于等于 <=
Public Function LessEqual(ByVal Expected As Variant, ByVal Expression As Variant, Optional ByRef FailMessage As String) As Boolean
    LessEqual = (Expected <= Expression)
End Function



' 约等于 ~=
Public Function Approximately(Expected As Variant, Expression As Variant, Optional ByRef FailMessage As String) As Boolean
    If IsNumeric(Expected) And IsNumeric(Expression) Then
        Approximately = ((1 / 10 ^ -5) * Abs(Expected - Expression) <= 1)
    Else
        Approximately = (StrComp(Expected, Expression, 1) = 0)
    End If
End Function



' 包含 in
Public Function Contains(Value As Variant, Expression As Variant, Optional ByRef FailMessage As String) As Boolean
    If IsArray(Value) And IsArray(Expression) Then
        Contains = Array_In_Array(Value, Expression)
    ElseIf IsArray(Expression) Then
        Contains = Value_In_Array(Value, Expression)
    ElseIf (VarType(Value) = vbString Or IsNumeric(Value)) And (VarType(Expression) = vbString Or IsNumeric(Expression)) Then
        Contains = Value_In_String(Value, Expression)
    End If
End Function



' 不包含 not in
Public Function NotContains(Value As Variant, Expression As Variant, Optional ByRef FailMessage As String) As Boolean
    NotContains = Not Contains(Value, Expression)
End Function



' 比较等于 is
Public Function SameInstance(Expected As Variant, Expression As Variant, Optional ByRef FailMessage As String) As Boolean
    If VarType(Expected) = VarType(Expression) Then
        If IsNull(Expected) Then
            SameInstance = IsNull(Expression)
        ElseIf IsEmpty(Expected) Then
            SameInstance = IsEmpty(Expression)
        ElseIf IsObject(Expected) And IsObject(Expression) Then
            SameInstance = Object_Is_Object(Expected, Expression)
        End If
    End If
End Function



' 比较不等于 not is
Public Function NotSameInstance(Expected As Variant, Expression As Variant, Optional ByRef FailMessage As String) As Boolean
    NotSameInstance = Not SameInstance(Expected, Expression)
End Function



' 正则匹配 match
Public Function Matches(ByVal Pattern As String, ByVal Expression As String, Optional ByRef FailMessage As String) As Boolean
    Dim RegEx As RegExp
    
    Set RegEx = New RegExp
    RegEx.Pattern = Pattern
    Matches = RegEx.Test(Expression)
    Set RegEx = Nothing
End Function



' 正则不匹配 not match
Public Function NotMatches(ByVal Pattern As String, ByVal Expression As String, Optional ByRef FailMessage As String) As Boolean
    NotMatches = Not Matches(Pattern, Expression, FailMessage)
End Function



' 两者之间 between
Public Function Between(ByVal NumberStart As Double, NumberEnd As Double, ByVal Expression As String, Optional ByRef FailMessage As String) As Boolean
    Dim Min As Double
    Dim Max As Double
    
    If NumberStart <= NumberEnd Then
        Min = NumberStart
        Max = NumberEnd
    Else
        Min = NumberEnd
        Max = NumberStart
    End If
    Between = (Min <= Expression And Expression <= Max)
End Function



' 非两者之间 not between
Public Function NotBetween(ByVal NumberStart As Double, NumberEnd As Double, ByVal Expression As String, Optional ByRef FailMessage As String) As Boolean
    NotBetween = Not Between(NumberStart, NumberEnd, Expression, FailMessage)
End Function







'========================================【内部函数】========================================

' 判断是否包含文本
Private Function Value_In_String(ByVal Value As Variant, ByVal InputString As Variant) As Boolean
    If InStr(InputString, Value) > 0 Then
        Value_In_String = True
    Else
        Value_In_String = False
    End If
End Function

' 判断数组是否包含元素
Private Function Value_In_Array(ByVal Value As Variant, ByVal InputArray As Variant) As Boolean
    If UBound(Filter(InputArray, Value)) > -1 Then
        Value_In_Array = True
    Else
        Value_In_Array = False
    End If
End Function

' 判断数组是否包含数组
Private Function Array_In_Array(ByVal ValArray As Variant, ByVal InputArray As Variant) As Boolean
    Dim i As Long
    
    Array_In_Array = True
    For i = 0 To UBound(ValArray)
        If UBound(Filter(InputArray, ValArray(i))) = -1 Then
            Array_In_Array = False
            Exit For
        End If
    Next
End Function


' 判断数组是否等于数组
Private Function Array_Eq_Array(ByVal ValArray As Variant, ByVal InputArray As Variant) As Boolean
    Dim i As Long
    
    Array_Eq_Array = True
    If UBound(ValArray) <> UBound(InputArray) Then Array_Eq_Array = False: Exit Function
    For i = 0 To UBound(ValArray)
        If IsObject(ValArray(i)) And IsObject(InputArray(i)) Then
            If Not ValArray(i) Is InputArray(i) Then
                Array_Eq_Array = False
                Exit For
            End If
        ElseIf IsObject(ValArray(i)) Or IsObject(InputArray(i)) Then
            Array_Eq_Array = False
            Exit For
        ElseIf ValArray(i) <> InputArray(i) Then
            Array_Eq_Array = False
            Exit For
        End If
    Next
End Function


' 判断数组是否全等于数组
Private Function Array_Eqv_Array(ByVal ValArray As Variant, ByVal InputArray As Variant) As Boolean
    Dim i As Long
    
    Array_Eqv_Array = True
    If UBound(ValArray) <> UBound(InputArray) Then Exit Function
    For i = 0 To UBound(ValArray)
        If VarType(ValArray(i)) <> VarType(InputArray(i)) Then
            Array_Eqv_Array = False
            Exit For
        ElseIf IsObject(ValArray(i)) And IsObject(InputArray(i)) Then
            If Not ValArray(i) Is InputArray(i) Then
                Array_Eqv_Array = False
                Exit For
            End If
        ElseIf ValArray(i) <> InputArray(i) Then
            Array_Eqv_Array = False
            Exit For
        End If
    Next
End Function


' 判断对象引用是否相同
Private Function Object_Is_Object(ByVal Object As Variant, ByVal InputObject As Variant) As Boolean
    If Object Is InputObject Then
        Object_Is_Object = True
    Else
        Object_Is_Object = False
    End If
End Function


' 单词首字母大写
Private Function TitleCase(ByVal word As String) As String
    Dim RegEx As RegExp
    Dim Matches As Object
    Dim sResult As String
    Dim w As Variant
    
    word = LCase(word)
    Set RegEx = New RegExp
    RegEx.Global = True
    RegEx.Pattern = "( |^)[a-z]"
    Set Matches = RegEx.Execute(word)
    sResult = RegEx.Replace(word, "%[words]%")
    For Each w In Matches
        sResult = Replace(sResult, "%[words]%", UCase(w), 1, 1)
    Next
    Set Matches = Nothing
    Set RegEx = Nothing
    TitleCase = sResult
End Function
