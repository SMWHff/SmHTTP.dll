Attribute VB_Name = "Utils"
Option Explicit
Private Declare Function SafeArrayGetDim Lib "oleaut32.dll" (ByRef saArray() As Any) As Long 'API判断数组为空或没有初始化
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function MymachineC Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function GetUserDefaultLCID Lib "kernel32" () As Long
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
'获取当前程序全路径
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
'取自身进程命令行参数
Private Declare Function GetCommandLine Lib "kernel32" Alias "GetCommandLineA" () As Long
''拷贝内存数据
'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
'取字符串长度
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long

Private Const LOCALE_SLANGDISPLAYNAME = &H6F



' 不定数量参数正则匹配（传入参数，参数类型，匹配模式，保存参数, 返回匹配结果）
'1、参数类型 = vbDecimal，则匹配模式为条件表达式（传入的参数用 x 表示）
'2、参数类型 = vbString， 则匹配模式为正则表达式
'3、参数类型 = vbVariant，则与所有类型都匹配
Public Function ArgumentsMatch(value As Variant, ByVal vType As VbVarType, Optional ByVal pattern As String, Optional ByVal SavaArg As Variant, Optional ByRef Result As Variant) As Boolean
    Dim re As New RegExp        '工程引用 Microsoft VBScript Regular Expressions 5.5
    Dim SubMatchs As SubMatches
    Dim valType As VbVarType
    Dim i As Long
    Dim Res As Boolean
    
    valType = VarType(value)
    If Not IsEmpty(SavaArg) Then GoTo return_
    If vType = vbVariant Then Res = True: GoTo return_
    If vType = vbDecimal And valType >= 2 And valType <= 5 Then
        Dim StringCalc As New ScriptControl
        StringCalc.Language = "VBScript"
        StringCalc.AddCode "x = " & value
        Res = StringCalc.Eval(pattern)
        Set StringCalc = Nothing
        GoTo return_
    End If
    If valType <> vType Then GoTo return_
    Res = True
    If vType = vbString Then
        re.IgnoreCase = True
        re.pattern = pattern
        Res = re.Test(value)
        If Res And IsMissing(Result) Then
            Set SubMatchs = re.Execute(value).Item(0).SubMatches
            If SubMatchs.Count > 0 Then
                ReDim Result(SubMatchs.Count - 1)
                For i = 0 To SubMatchs.Count - 1
                    Result(i) = SubMatchs.Item(i)
                Next
            End If
        End If
        Set re = Nothing
    End If
return_:
    ArgumentsMatch = Res
End Function


'Base64编码
Public Function Base64Encoder(ByRef Strs As Variant) As String
    Dim Buf() As Byte, str() As Byte
    Dim Length As Long, mods As Long
    Const B64_CHAR_DICT = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="

    On Error GoTo over
    If VarType(Strs) = vbArray + vbByte Then
        str() = Strs
    Else
        str() = StrConv(Strs, vbFromUnicode)
    End If
    mods = (UBound(str) + 1) Mod 3   '除以3的余数
    Length = UBound(str) + 1 - mods
    ReDim Buf(Length / 3 * 4 + IIf(mods <> 0, 4, 0) - 1)
    Dim i As Long
    For i = 0 To Length - 1 Step 3
        Buf(i / 3 * 4) = (str(i) And &HFC) / &H4
        Buf(i / 3 * 4 + 1) = (str(i) And &H3) * &H10 + (str(i + 1) And &HF0) / &H10
        Buf(i / 3 * 4 + 2) = (str(i + 1) And &HF) * &H4 + (str(i + 2) And &HC0) / &H40
        Buf(i / 3 * 4 + 3) = str(i + 2) And &H3F
    Next
    If mods = 1 Then
        Buf(Length / 3 * 4) = (str(Length) And &HFC) / &H4
        Buf(Length / 3 * 4 + 1) = (str(Length) And &H3) * &H10
        Buf(Length / 3 * 4 + 2) = 64
        Buf(Length / 3 * 4 + 3) = 64
    ElseIf mods = 2 Then
        Buf(Length / 3 * 4) = (str(Length) And &HFC) / &H4
        Buf(Length / 3 * 4 + 1) = (str(Length) And &H3) * &H10 + (str(Length + 1) And &HF0) / &H10
        Buf(Length / 3 * 4 + 2) = (str(Length + 1) And &HF) * &H4
        Buf(Length / 3 * 4 + 3) = 64
    End If
    For i = 0 To UBound(Buf)
        Base64Encoder = Base64Encoder + Mid(B64_CHAR_DICT, Buf(i) + 1, 1)
    Next
over:
End Function


'Base64解码
Public Function Base64Decoder(ByVal B64 As String, Optional ByVal IsByte As Boolean = False) As Variant
    Dim OutStr() As Byte
    Dim i As Long, j As Long
    Const B64_CHAR_DICT = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="

    On Error GoTo over
    If InStr(1, B64, "=") <> 0 Then B64 = Left(B64, InStr(1, B64, "=") - 1)     '判断Base64真实长度,除去补位
    Dim Length As Long, mods As Long
    mods = Len(B64) Mod 4
    Length = Len(B64) - mods
    ReDim OutStr(Length / 4 * 3 - 1 + Switch(mods = 0, 0, mods = 2, 1, mods = 3, 2))
    For i = 1 To Length Step 4
        Dim Buf(3) As Byte
        For j = 0 To 3
            Buf(j) = InStr(1, B64_CHAR_DICT, Mid(B64, i + j, 1)) - 1            '根据字符的位置取得索引值
        Next
        OutStr((i - 1) / 4 * 3) = Buf(0) * &H4 + (Buf(1) And &H30) / &H10
        OutStr((i - 1) / 4 * 3 + 1) = (Buf(1) And &HF) * &H10 + (Buf(2) And &H3C) / &H4
        OutStr((i - 1) / 4 * 3 + 2) = (Buf(2) And &H3) * &H40 + Buf(3)
    Next
    If mods = 2 Then
        OutStr(Length / 4 * 3) = (InStr(1, B64_CHAR_DICT, Mid(B64, Length + 1, 1)) - 1) * &H4 + ((InStr(1, B64_CHAR_DICT, Mid(B64, Length + 2, 1)) - 1) And &H30) / 16
    ElseIf mods = 3 Then
        OutStr(Length / 4 * 3) = (InStr(1, B64_CHAR_DICT, Mid(B64, Length + 1, 1)) - 1) * &H4 + ((InStr(1, B64_CHAR_DICT, Mid(B64, Length + 2, 1)) - 1) And &H30) / 16
        OutStr(Length / 4 * 3 + 1) = ((InStr(1, B64_CHAR_DICT, Mid(B64, Length + 2, 1)) - 1) And &HF) * &H10 + ((InStr(1, B64_CHAR_DICT, Mid(B64, Length + 3, 1)) - 1) And &H3C) / &H4
    End If
    If IsByte Then
        Base64Decoder = OutStr                                                       '读取解码结果
    Else
        Base64Decoder = StrConv(OutStr, vbUnicode)
    End If
over:
End Function


'取出中间文本
Function MidStr(ByVal str As String, ByVal StrHome As String, Optional ByVal StrEnd As String = vbNullString)
    Dim Ret, arr1, arr2

    Ret = ""
    arr1 = Split(str, StrHome, 2)
    If UBound(arr1) = 1 Then
        If Len(StrEnd) = 0 Then
            Ret = arr1(1)
        Else
            arr2 = Split(arr1(1), StrEnd, 2)
            If UBound(arr2) = 1 Then
                Ret = arr2(0)
            End If
        End If
    End If
    MidStr = Ret
End Function



'URL编码
Public Function EscapeURL(ByVal URL)
    Dim sc As ScriptControl  '需要引用工程 Microsoft Script Control

    Set sc = CreateObject("MSScriptControl.ScriptControl")
    sc.Language = "JScript"
    EscapeURL = sc.Run("encodeURIComponent", URL)
    Set sc = Nothing
End Function

'URL解码
Public Function UnEscapeURL(ByVal URLCode)
    Dim sc As ScriptControl  '需要引用工程 Microsoft Script Control

    Set sc = CreateObject("MSScriptControl.ScriptControl")
    sc.Language = "JScript"
    UnEscapeURL = sc.Run("decodeURIComponent", URLCode)
    Set sc = Nothing
End Function


'取现行时间戳
Public Function GetUnixTime(Optional ByVal IsDec As Boolean) As String
    GetUnixTime = DateDiff("s", "1970-1-1 0:0:0", DateAdd("h", -8, Now)) & IIf(IsDec, "", Right(GetTickCount(), 3))
End Function


' 拼接字节集
Public Function BinCat(ByRef BinA() As Byte, ByRef BinB() As Byte) As Byte()
    Dim Count       As Long
    Dim Length      As Long
    
    If SafeArrayGetDim(BinA) <> 0 Then Count = UBound(BinA) Else Count = -1
    If SafeArrayGetDim(BinB) <> 0 Then Length = UBound(BinB) Else Length = -1
    ReDim Preserve BinA(Count + Length + 1) '重新定义新字节数组
    Call CopyMemory(BinA(Count + 1), BinB(0), Length + 1)
    BinCat = BinA
End Function


' 拼接字节集加强
Public Function BinCatEx(ParamArray Args()) As Byte()
    Dim ArrBin()    As Variant
    Dim i           As Long
    
    ReDim ArrBin(UBound(Args))
    For i = 0 To UBound(Args)
        ArrBin(i) = Args(i)
    Next
    BinCatEx = JoinBin(ArrBin)
End Function


' 合并字节集
Public Function JoinBin(ByRef ArrBin() As Variant) As Byte()
    Dim i           As Long
    Dim Count       As Long
    Dim Length      As Long
    Dim TempBuff()  As Byte
    Dim Result()    As Byte
    
    If UBound(ArrBin) >= 0 Then Result = ArrBin(0)
    For i = 1 To UBound(ArrBin)
        TempBuff = ArrBin(i)
        If SafeArrayGetDim(Result) <> 0 Then Count = UBound(Result) Else Count = -1
        If SafeArrayGetDim(TempBuff) <> 0 Then Length = UBound(TempBuff) Else Length = -1
        ReDim Preserve Result(Count + Length + 1) '重新定义新字节数组
        Call CopyMemory(Result(Count + 1), TempBuff(0), Length + 1)
    Next
    JoinBin = Result
End Function


' 到字节集
Public Function ToBin(ByVal str As String) As Byte()
    ToBin = StrConv(str, vbFromUnicode)
End Function


' 读入字节集
Public Function ReadBin(ByVal Path As String) As Byte()
    Dim nFile       As Long
    Dim FileLen     As Long
    Dim FileByte()  As Byte
    
    nFile = FreeFile()      ' FreeFile 返回一个 Integer，代表下一个可供文件号
    Open Path For Binary As #nFile
        FileLen = LOF(nFile) - 1
        ReDim FileByte(FileLen) As Byte
        Get #nFile, , FileByte()
    Close #nFile
    ReadBin = FileByte()
End Function


'读入文件字节集
Public Function File_ReadByte(ByVal Path As String) As Byte()
    Dim ADO As stream  '工程引用 Microsoft ActiveX Data Objects 2.8 Libary

    Set ADO = CreateObject("ADODB.Stream")
    ADO.Type = 1
    ADO.Open
    ADO.LoadFromFile Path
    File_ReadByte = ADO.Read
    ADO.Close
    Set ADO = Nothing
End Function

'写出文件
Public Function File_WriteByte(ByVal Path As String, ByVal Bytes) As Long
    Dim ADO As stream   '工程引用 Microsoft ActiveX Data Objects 2.8 Libary
    Dim Ret As Long

    If VarType(Bytes) = 8209 Then
        If Len(Bytes) > 0 Then
            Set ADO = CreateObject("ADODB.Stream")
            ADO.Type = 1
            ADO.Mode = 3
            ADO.Open
            ADO.Write Bytes
            ADO.SaveToFile Path, 2
            ADO.Close
            Set ADO = Nothing
            If Dir(Path) <> "" And Len(Path) > 0 Then
                Ret = 1
            End If
        End If
    End If
    File_WriteByte = Ret
End Function

' 保存字节集到文件
Public Function SaveFile(ByVal Path As String, ByVal Data As Variant) As Long
    Dim Handle  As Long
    Dim bin()   As Byte
    Dim Ret     As Long
    
    If Len(Path) > 0 And VarType(Data) = vbArray + vbByte Then
        bin() = Data
        Handle = FreeFile() '获得文件的句柄
        Open Path For Binary As #Handle
            Put #Handle, , bin()
        Close #Handle
        If Len(Dir(Path)) > 0 Then Ret = 1
    End If
    SaveFile = Ret
End Function

'字节集连接
Public Function Concat_Byte(ByRef Bin1() As Byte, ByRef Bin2() As Byte) As Byte()
    Dim ADO As stream  '引用工程 Microsoft ActiveX Data Objects 6.1 Libary
    Dim bin() As Byte

    Set ADO = CreateObject("ADODB.Stream")
    ADO.Type = 1
    ADO.Open
    ADO.Write Bin1
    ADO.Write Bin2
    ADO.Position = 0
    bin = ADO.Read
    ADO.Close
    Set ADO = Nothing
    Concat_Byte = bin
End Function


'字节集连接
Public Function Concat_ByteByArray(ByRef Args As Variant) As Byte()
    Dim ADO     As stream  '引用工程 Microsoft ActiveX Data Objects 6.1 Libary
    Dim v       As Variant
    Dim bin()   As Byte

    Set ADO = CreateObject("ADODB.Stream")
    ADO.Type = 1
    ADO.Open
    For Each v In Args
        If VarType(v) = vbArray + vbByte Then
            ADO.Write v
        End If
    Next
    ADO.Position = 0
    bin = ADO.Read
    ADO.Close
    Set ADO = Nothing
    Concat_ByteByArray = bin
End Function


Function CBytes(str)
    Dim MD, node, i, StrH
    Set MD = CreateObject("Msxml2.DOMDocument")
    Set node = MD.createElement("binary")
    node.dataType = "bin.hex"
    For i = 1 To Len(str)
        StrH = StrH & Right("0" + Hex(Asc(Mid(str, i, 1))), 2)
    Next
    node.Text = StrH
    CBytes = node.nodeTypedValue
    Set node = Nothing
    Set MD = Nothing
End Function

'字节集到16进制
Public Function T_BinToHex_XML(ByRef Bytes) As String
    Dim xml As DOMDocument      '引用工程 Microsoft XML v3.0
    Dim node As IXMLDOMElement
    Dim sHex As String

    On Error Resume Next
    Set xml = CreateObject("Microsoft.XMLDOM")
    Set node = xml.createElement("binary")
    node.dataType = "bin.hex"
    node.nodeTypedValue = Bytes
    sHex = UCase(node.Text)
    Set node = Nothing
    Set xml = Nothing
    T_BinToHex_XML = sHex
End Function

'16进制到字节集
Public Function T_HexToBin_XML(ByVal HexStr As String) As Byte()
    Dim MD As DOMDocument       '引用工程 Microsoft XML v3.0
    Dim node As IXMLDOMElement
    Dim bin() As Byte

    On Error Resume Next
    Set MD = CreateObject("Msxml2.DOMDocument")
    Set node = MD.createElement("binary")
    node.dataType = "bin.hex"
    node.Text = HexStr
    bin() = node.nodeTypedValue
    Set node = Nothing
    Set MD = Nothing
    T_HexToBin_XML = bin()
End Function


'字节集到16进制 - 纯算法
Public Function T_BinToHex(ByRef Bytes) As String
    Dim iLen As Long
    Dim ibyte As Long
    Dim high As Long
    Dim low As Long
    Dim Buff() As Byte
    Dim Buff_len As Long
    Dim i As Long, j As Long
    Dim sHex As String

    If VarType(Bytes) = vbArray + vbByte Then
        iLen = UBound(Bytes)
        Buff_len = (iLen + 1) * 2 - 1
        ReDim Buff(Buff_len) As Byte
        For j = 0 To Buff_len Step 2
            ibyte = Bytes(i)
            i = i + 1
            If ibyte > 15 Then
                high = (ibyte / 2 ^ 4) And 15
                If high > 9 Then
                    Buff(j) = high + 55
                Else
                    Buff(j) = high + 48
                End If
            Else
                Buff(j) = high + 48
            End If
            low = ibyte And 15
            If low > 9 Then
                Buff(j + 1) = low + 55
            Else
                Buff(j + 1) = low + 48
            End If
        Next
        sHex = StrConv(Buff, vbUnicode)
    End If
    T_BinToHex = sHex
End Function


'16进制到字节集 - 纯算法
Public Function T_HexToBin(ByVal HexStr As String) As Byte()
    Dim Bytes() As Byte
    Dim Buff() As Byte
    Dim iLen As Long
    Dim p1 As Long
    Dim i As Long
    Dim byte1 As Long
    Dim byte2 As Long
    
    Bytes = StrConv(HexStr, vbFromUnicode)
    iLen = UBound(Bytes)
    If iLen And 1 = 1 Then
        iLen = iLen + 1
        ReDim Preserve Bytes(iLen) As Byte
        Bytes(iLen) = Bytes(iLen - 1)
        Bytes(iLen - 1) = 48
    End If
    ReDim Buff(iLen / 2 ^ 1) As Byte
    For p1 = 0 To iLen Step 2
        byte1 = Bytes(p1)
        byte2 = Bytes(p1 + 1)
        If byte1 > 96 Then
            byte1 = byte1 - 87
        ElseIf byte1 > 64 Then
            byte1 = byte1 - 55
        Else
            byte1 = byte1 - 48
        End If
        If byte2 > 96 Then
            byte2 = byte2 - 87
        ElseIf byte2 > 64 Then
            byte2 = byte2 - 55
        Else
            byte2 = byte2 - 48
        End If
        Buff(i) = byte1 * 2 ^ 4 + byte2
        i = i + 1
    Next
    T_HexToBin = Buff()
End Function


'获取网络时间
Function T_GetNetTime() As String
    Dim http As WinHttpRequest  '引用工程 Microsoft WinHTTP Services, version 5.1
    Dim sRet As String
    Dim sDate As String
    Dim i As Long
    Dim IPArr()
    
    On Error Resume Next
    IPArr = Array("223.5.5.5", "223.6.6.6", "119.29.29.98", "114.55.27.46")
    
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.SetTimeouts 100, 100, 100, 100
    For i = 0 To UBound(IPArr)
        http.Open "HEAD", "http://" & IPArr(i), False
        http.Send
        If http.Status = 200 Then
            sDate = http.GetResponseHeader("Date") '只取时间这段
            sRet = DateAdd("h", 8, CDate(Mid(sDate, 5, Len(sDate) - 8))) '转换时间格式
            sRet = Format$(sRet, "yyyy-mm-dd hh:mm:ss")
            Exit For
        End If
    Next
    Set http = Nothing
    T_GetNetTime = sRet
End Function


'解析JSON取值
Public Function Fun_GetJSON(ByVal JSONStr As String, ByVal key As String) As Variant
    Dim sc      As ScriptControl  '需要引用工程 Microsoft Script Control
    Dim L1      As Long
    Dim L2      As Long
    Dim sJSON   As String
    Dim temp    As String
    Dim Ret     As Variant

    ' On Error Resume Next
    L1 = InStr(JSONStr, "{"): L2 = InStrRev(JSONStr, "}")
    If L1 > 0 And L2 > L1 Then
        sJSON = Mid(JSONStr, L1, L2 - L1 + 1)
    End If
    L1 = InStr(JSONStr, "["): L2 = InStrRev(JSONStr, "]")
    If L1 > 0 And L2 > L1 Then
        temp = Mid(JSONStr, L1, L2 - L1 + 1)
        If Len(temp) > Len(sJSON) Then sJSON = temp
    End If
    If Len(sJSON) > 0 Then
        If Left(key, 1) <> "[" And key <> "" Then key = "." & key
        Set sc = CreateObject("MSScriptControl.ScriptControl")
        sc.Language = "JScript"
        sc.AddCode "var $ = eval(" & sJSON & ");"
        If key <> "" And sJSON <> "[]" And sJSON <> "{}" Then
            Ret = sc.Eval("$" & key)
            If IsNumeric(Ret) And Left(Ret, 1) = "." Then
                Ret = "0" & Ret
            End If
        End If
        sc.Reset
        Set sc = Nothing
    End If
    Fun_GetJSON = Ret
End Function


' 计算文件的 MD5
Public Function Fun_MD5_File(ByVal file_name As String) As String
    Dim wi          As Object
    Dim file_hash   As Object
    Dim hash_value  As String
    Dim i           As Long
    
    Set wi = CreateObject("WindowsInstaller.Installer")
    Set file_hash = wi.FileHash(file_name, 0)
    hash_value = ""
    For i = 1 To file_hash.FieldCount
        hash_value = hash_value & BigEndianHex(file_hash.IntegerData(i))
    Next
    Fun_MD5_File = hash_value
    Set file_hash = Nothing
    Set wi = Nothing
End Function
Private Function BigEndianHex(ByVal iInt As Long) As String
    Dim Result  As String
    Dim b1      As Long
    Dim b2      As Long
    Dim b3      As Long
    Dim b4      As Long
    
    Result = Hex(iInt)
    b1 = Mid(Result, 7, 2)
    b2 = Mid(Result, 5, 2)
    b3 = Mid(Result, 3, 2)
    b4 = Mid(Result, 1, 2)
    BigEndianHex = b1 & b2 & b3 & b4
End Function


' 获取硬盘序列号
Public Function GetHDDSN()
    Dim Ret, 硬盘序列号, Maxlen, Sysflag As Long: Dim VolName, FsysName As String
    Ret = MymachineC("c:\", VolName, 256, 硬盘序列号, Maxlen, Sysflag, FsysName, 256)
    GetHDDSN = Hex(硬盘序列号) & Hex(Sysflag)
End Function


' 获取主板序列号
Public Function GetBaseBoard()
    Dim objWMIService   As Object
    Dim colItems        As Object
    Dim objItem         As Variant
    Dim Ret             As String
    
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colItems = objWMIService.ExecQuery("Select * from Win32_BaseBoard")
    For Each objItem In colItems
        Ret = objItem.SerialNumber
    Next
    Set colItems = Nothing
    Set objWMIService = Nothing
    GetBaseBoard = Ret
End Function



' 获取BIOS序列号
Public Function GetBIOS()
    Dim objWMIService   As Object
    Dim colItems        As Object
    Dim objItem         As Variant
    Dim Ret             As String
    
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colItems = objWMIService.ExecQuery("Select * from Win32_BIOS")
    For Each objItem In colItems
        Ret = objItem.SerialNumber
    Next
    Set colItems = Nothing
    Set objWMIService = Nothing
    GetBIOS = Ret
End Function


' 获取IE浏览器UserAgent
Public Function GetUserAgent() As String
    Dim UserAgent       As String
    Dim HTMLDocument    As HTMLDocument ' 引用 Microsoft HTML Object Library
    Dim IE              As Object       ' 引用 Microsoft Internet Controls
    Dim WSH             As Object       ' 引用 Microsoft Windows Script Host Object Model

    Set HTMLDocument = CreateObject("htmlfile")
    HTMLDocument.Open
    UserAgent = HTMLDocument.parentWindow.navigator.UserAgent
    Set HTMLDocument = Nothing
    
    If Len(UserAgent) = 0 Then
        Set IE = CreateObject("InternetExplorer.Application")
        IE.Visible = 0
        IE.navigate "about:blank"
        Do While IE.Busy
            DoEvents
        Loop
        UserAgent = IE.document.parentWindow.navigator.UserAgent
        IE.Quit
    End If
    GetUserAgent = UserAgent
End Function


' 获取屏幕颜色深度
Public Function GetBitsPerPel() As Long
    Dim objWMIService   As Object
    Dim objItems        As Object
    Dim objItem         As Object
    Dim Bits            As Long
    
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set objItems = objWMIService.ExecQuery("Select * from Win32_DisplayConfiguration")
    For Each objItem In objItems
        Bits = objItem.BitsPerPel
        If Bits Then Exit For
    Next
    Set objItems = Nothing
    Set objWMIService = Nothing
    GetBitsPerPel = Bits
End Function


' 获取屏幕分辨率
Public Function GetScreenXY() As String
    Dim ScRX            As Long
    Dim ScRY            As Long
    
    ScRX = Screen.Width \ Screen.TwipsPerPixelX
    ScRY = Screen.Height \ Screen.TwipsPerPixelY
    GetScreenXY = ScRX & "x" & ScRY
End Function


' 判断是否支持JAVA
Public Function IsJavaInstalled() As Boolean
    Dim CLASSPATH As String
    Dim JAVA_HOME As String
    
    IsJavaInstalled = False
    CLASSPATH = Environ("CLASSPATH")
    JAVA_HOME = Environ("JAVA_HOME")
    
    If Len(CLASSPATH) <> 0 And Len(JAVA_HOME) <> 0 Then
        If Len(Dir(JAVA_HOME & "\bin\java.exe")) <> 0 Then
            IsJavaInstalled = True
        End If
    End If
End Function


' 获取浏览器语言环境
Public Function GetBrowserLanguage() As String
    Dim Language        As String
    Dim HTMLDocument    As HTMLDocument ' 引用 Microsoft HTML Object Library
    Dim IE              As Object       ' 引用 Microsoft Internet Controls
    
    Set HTMLDocument = CreateObject("htmlfile")
    HTMLDocument.Open
    Language = HTMLDocument.parentWindow.navigator.userLanguage
    Set HTMLDocument = Nothing
    
    If Len(Language) = 0 Then
        Set IE = CreateObject("InternetExplorer.Application")
        IE.Visible = 0
        IE.navigate "about:blank"
        Do While IE.Busy
            DoEvents
        Loop
        Language = IE.document.parentWindow.navigator.Language
        IE.Quit
    End If
    GetBrowserLanguage = LCase(Language)
End Function


' 获取当前系统语言环境
Public Function GetLanguageLocale() As String
    Dim lngLCID         As Long
    Dim strLocaleInfo   As String
    Dim lngResult       As Long

    ' 获取用户的默认语言环境
    lngLCID = GetUserDefaultLCID()

    ' 设置缓冲区大小
    strLocaleInfo = Space(255)

    ' 获取语言的显示名称
    lngResult = GetLocaleInfo(lngLCID, LOCALE_SLANGDISPLAYNAME, strLocaleInfo, 255)
    
    If lngResult > 0 Then
        GetLanguageLocale = Trim(Left$(strLocaleInfo, InStr(strLocaleInfo, vbNullChar) - 1))
    Else
        GetLanguageLocale = "Unknown"
    End If
End Function


'取出中间文本
Function T_MidS(ByVal str As String, ByVal StrHome As String, Optional ByVal StrEnd As String = vbNullString)
    Dim Ret, arr1, arr2

    Ret = ""
    arr1 = Split(str, StrHome, 2)
    If UBound(arr1) = 1 Then
        If Len(StrEnd) = 0 Then
            Ret = arr1(1)
        Else
            arr2 = Split(arr1(1), StrEnd, 2)
            If UBound(arr2) = 1 Then
                Ret = arr2(0)
            End If
        End If
    End If
    T_MidS = Ret
End Function


'获取自身命令行参数
Function GetCommLine() As String
    Dim RetStr As Long, sLen As Long
    Dim buffer As String

    RetStr = GetCommandLine()
    sLen = lstrlen(RetStr)
    If sLen > 0 Then
        GetCommLine = Space$(sLen)
        CopyMemory ByVal GetCommLine, ByVal RetStr, sLen
    End If
End Function


'获取自身执行文件名
Public Function GetMeExeName() As String
    Dim lRet As Long
    Dim Path As String
    Dim sRet As String

    Path = String(255, 0)
    lRet = GetModuleFileName(0, Path, 255)
    Path = Left(Path, lRet)
    sRet = Mid(Path, InStrRev(Path, "\") + 1)
    If LCase(sRet) = "runner.exe" Or LCase(sRet) = "runnerlua.exe" Then
        Path = T_MidS(GetCommLine(), """", """")
        sRet = Mid(Path, InStrRev(Path, "\") + 1)
    End If
    GetMeExeName = sRet
End Function
