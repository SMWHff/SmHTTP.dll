Attribute VB_Name = "Encoding"
Option Explicit


' Base64编码
Public Function Encoding_Base64(ByRef Data() As Byte) As String
    Dim objXML  As MSXML2.DOMDocument
    Dim objNode As MSXML2.IXMLDOMElement

    ' 创建XML对象
    Set objXML = New MSXML2.DOMDocument
    
    ' 创建一个元素
    Set objNode = objXML.createElement("b64")

    ' 数据类型转换
    objNode.dataType = "bin.base64"
    
    ' 字符串转换为字节
    'objNode.nodeTypedValue = StrConv(Data, vbFromUnicode)
    objNode.nodeTypedValue = Data

    ' 获取编码后的Base64字符串
    Encoding_Base64 = objNode.Text

    ' 清理对象
    Set objNode = Nothing
    Set objXML = Nothing
End Function


' Base64解码
Public Function Encoding_UnBase64(ByRef B64 As String) As Byte()
    Dim objXML  As MSXML2.DOMDocument
    Dim objNode As MSXML2.IXMLDOMElement

    ' 创建 XML 对象
    Set objXML = New MSXML2.DOMDocument
    
    ' 创建一个元素
    Set objNode = objXML.createElement("b64")

    ' 设置数据类型
    objNode.dataType = "bin.base64"

    ' 设置要解码的文本
    objNode.Text = B64

    ' 执行解码并转换回字符串
    'Encoding_UnBase64 = StrConv(objNode.nodeTypedValue, vbUnicode)
    Encoding_UnBase64 = objNode.nodeTypedValue

    ' 清理对象
    Set objNode = Nothing
    Set objXML = Nothing
End Function


'URL编码
Public Function Encoding_URL(ByVal URL As String) As String
    Dim sc As ScriptControl  '需要引用工程 Microsoft Script Control

    Set sc = CreateObject("MSScriptControl.ScriptControl")
    sc.Language = "JScript"
    Encoding_URL = sc.Run("encodeURIComponent", URL)
    Set sc = Nothing
End Function


'URL解码
Public Function Encoding_UnURL(ByVal URLCode As String) As String
    Dim sc As ScriptControl  '需要引用工程 Microsoft Script Control

    Set sc = CreateObject("MSScriptControl.ScriptControl")
    sc.Language = "JScript"
    Encoding_UnURL = sc.Run("decodeURIComponent", URLCode)
    Set sc = Nothing
End Function


'Hex编码
Public Function Encoding_Hex(ByRef Data() As Byte) As String
    Dim xml     As DOMDocument      '引用工程 Microsoft XML v3.0
    Dim node    As IXMLDOMElement
    Dim sHex    As String

    Set xml = CreateObject("Microsoft.XMLDOM")
    Set node = xml.createElement("binary")
    node.dataType = "bin.hex"
    node.nodeTypedValue = Data
    sHex = UCase(node.Text)
    Set node = Nothing
    Set xml = Nothing
    Encoding_Hex = sHex
End Function


'Hex解码
Public Function Encoding_UnHex(ByVal strHex As String) As Byte()
    Dim MD As DOMDocument       '引用工程 Microsoft XML v3.0
    Dim node As IXMLDOMElement
    Dim bin() As Byte

    Set MD = CreateObject("Msxml2.DOMDocument")
    Set node = MD.createElement("binary")
    node.dataType = "bin.hex"
    node.Text = strHex
    bin() = node.nodeTypedValue
    Set node = Nothing
    Set MD = Nothing
    Encoding_UnHex = bin()
End Function


' 字符编码转化
Function Encoding_Convert(ByRef inputBin() As Byte, ByVal toCharset As String) As Byte()
    Dim stream      As Object   '工程引用 Microsoft ActiveX Data Objects 2.8 Libary
    Dim Result()    As Byte

    ' 创建 ADODB.Stream 对象
    Set stream = CreateObject("ADODB.Stream")
    With stream
        .Type = 1 ' adTypeBinary
        .Open
        .Write inputBin
        .position = 0
        .Charset = toCharset
        Result = .Read
        .Close
    End With
    Set stream = Nothing
    Encoding_Convert = Result
End Function

