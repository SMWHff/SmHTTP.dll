Attribute VB_Name = "Encoding"
Option Explicit


' Base64����
Public Function Encoding_Base64(ByRef Data() As Byte) As String
    Dim objXML  As MSXML2.DOMDocument
    Dim objNode As MSXML2.IXMLDOMElement

    ' ����XML����
    Set objXML = New MSXML2.DOMDocument
    
    ' ����һ��Ԫ��
    Set objNode = objXML.createElement("b64")

    ' ��������ת��
    objNode.dataType = "bin.base64"
    
    ' �ַ���ת��Ϊ�ֽ�
    'objNode.nodeTypedValue = StrConv(Data, vbFromUnicode)
    objNode.nodeTypedValue = Data

    ' ��ȡ������Base64�ַ���
    Encoding_Base64 = objNode.Text

    ' �������
    Set objNode = Nothing
    Set objXML = Nothing
End Function


' Base64����
Public Function Encoding_UnBase64(ByRef B64 As String) As Byte()
    Dim objXML  As MSXML2.DOMDocument
    Dim objNode As MSXML2.IXMLDOMElement

    ' ���� XML ����
    Set objXML = New MSXML2.DOMDocument
    
    ' ����һ��Ԫ��
    Set objNode = objXML.createElement("b64")

    ' ������������
    objNode.dataType = "bin.base64"

    ' ����Ҫ������ı�
    objNode.Text = B64

    ' ִ�н��벢ת�����ַ���
    'Encoding_UnBase64 = StrConv(objNode.nodeTypedValue, vbUnicode)
    Encoding_UnBase64 = objNode.nodeTypedValue

    ' �������
    Set objNode = Nothing
    Set objXML = Nothing
End Function


'URL����
Public Function Encoding_URL(ByVal URL As String) As String
    Dim sc As ScriptControl  '��Ҫ���ù��� Microsoft Script Control

    Set sc = CreateObject("MSScriptControl.ScriptControl")
    sc.Language = "JScript"
    Encoding_URL = sc.Run("encodeURIComponent", URL)
    Set sc = Nothing
End Function


'URL����
Public Function Encoding_UnURL(ByVal URLCode As String) As String
    Dim sc As ScriptControl  '��Ҫ���ù��� Microsoft Script Control

    Set sc = CreateObject("MSScriptControl.ScriptControl")
    sc.Language = "JScript"
    Encoding_UnURL = sc.Run("decodeURIComponent", URLCode)
    Set sc = Nothing
End Function


'Hex����
Public Function Encoding_Hex(ByRef Data() As Byte) As String
    Dim xml     As DOMDocument      '���ù��� Microsoft XML v3.0
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


'Hex����
Public Function Encoding_UnHex(ByVal strHex As String) As Byte()
    Dim MD As DOMDocument       '���ù��� Microsoft XML v3.0
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


' �ַ�����ת��
Function Encoding_Convert(ByRef inputBin() As Byte, ByVal toCharset As String) As Byte()
    Dim stream      As Object   '�������� Microsoft ActiveX Data Objects 2.8 Libary
    Dim Result()    As Byte

    ' ���� ADODB.Stream ����
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

