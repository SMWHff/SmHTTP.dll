Attribute VB_Name = "TestSound"
Option Explicit

' 导入 PlaySound API 函数
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundW" (ByVal pszSound As String, ByVal hMod As Long, ByVal fdwSound As Long) As Long
' 声明 mciSendString API 函数
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long



' PlaySound 函数的标志参数常量
Public Const SND_SYNC = &H0         ' 播放同步声音
Public Const SND_ASYNC = &H1        ' 播放异步声音
Public Const SND_FILENAME = &H20000 ' 声音参数是文件名
Public Const SND_NODEFAULT = &H2    ' 没有默认声音时无声音播放
Public Const SND_LOOP = &H8         ' 循环播放声音，直到下一次调用 PlaySound
Public Const SND_NOSTOP = &H10      ' 即使已有声音在播放，也不停止当前声音


' 播放指定的 WAV 文件
Public Function PlayWAVFile(ByVal strSoundFile As String, Optional ByVal bAsync As Boolean = True) As Long
    Dim flags As Long
    flags = SND_FILENAME Or SND_NODEFAULT
    If bAsync Then
        flags = flags Or SND_ASYNC
    Else
        flags = flags Or SND_SYNC
    End If
    
    ' 播放 WAV 文件
    PlayWAVFile = PlaySound(strSoundFile, 0&, flags)
    Debug.Print PlayWAVFile
End Function


' 使用 mciSendString 来播放 MP3
Public Sub PlayMP3(ByVal mp3FileName As String)
    Dim sCommand As String
    Dim iReturn As Integer

    ' 打开 MP3 文件
    sCommand = "open """ & mp3FileName & """ type mpegvideo alias mp3"
    iReturn = mciSendString(sCommand, "", 0, 0)

    ' 播放 MP3 文件
    sCommand = "play mp3"
    iReturn = mciSendString(sCommand, "", 0, 0)
End Sub


' 停止并关闭 MP3
Public Sub StopMP3()
    Dim sCommand As String
    Dim iReturn As Integer

    ' 停止播放
    sCommand = "stop mp3"
    iReturn = mciSendString(sCommand, "", 0, 0)

    ' 关闭设备
    sCommand = "close mp3"
    iReturn = mciSendString(sCommand, "", 0, 0)
End Sub


' 播放 WAV 文件
Public Function PlayWAV(ByVal wavFileName As String) As Boolean
    Dim sCommand As String
    Dim iReturn As Integer
    Dim bRet    As Boolean

    ' 打开 WAV 文件
    sCommand = "open """ & wavFileName & """ alias myWAV"
    iReturn = mciSendString(sCommand, "", 0, 0)
    If iReturn <> 0 Then
        PlayWAV = False
        Debug.Print "出错，打开 WAV 文件失败！"
        Exit Function
    End If

    ' 播放 WAV 文件
    sCommand = "play myWAV wait"
    iReturn = mciSendString(sCommand, "", 0, 0)
    If iReturn <> 0 Then
        ' 关闭设备，因为打开了但没能播放
        mciSendString "close myWAV", "", 0, 0
        PlayWAV = False
        Debug.Print "出错，播放 WAV 文件失败！"
    Else
        PlayWAV = True
    End If
End Function


' 停止并关闭 WAV
Public Function StopWAV()
    Dim sCommand As String

    ' 停止播放
    sCommand = "stop myWAV"
    Call mciSendString(sCommand, "", 0, 0)

    ' 关闭设备
    sCommand = "close myWAV"
    Call mciSendString(sCommand, "", 0, 0)
End Function



Sub PlayDTMFTone(ByVal key As String)
    Dim freq1 As Long, freq2 As Long
    
    Select Case key
        Case "1"
            freq1 = 697: freq2 = 1209
        Case "2"
            freq1 = 697: freq2 = 1336
        Case "3"
            freq1 = 697: freq2 = 1477
        Case "4"
            freq1 = 770: freq2 = 1209
        Case "5"
            freq1 = 770: freq2 = 1336
        Case "6"
            freq1 = 770: freq2 = 1477
        Case "7"
            freq1 = 852: freq2 = 1209
        Case "8"
            freq1 = 852: freq2 = 1336
        Case "9"
            freq1 = 852: freq2 = 1477
        Case "0"
            freq1 = 941: freq2 = 1336
        Case Else
            freq1 = 0: freq2 = 0
    End Select
    
    ' 播放每个按键的声音，每个频率持续100毫秒
    If freq1 > 0 And freq2 > 0 Then
        Beep freq1, 100
        Beep freq2, 100
    End If
End Sub

Sub PlayPhoneNumber(ByVal phoneNumber As String)
    Dim i As Integer
    For i = 1 To Len(phoneNumber)
        PlayDTMFTone (Mid$(phoneNumber, i, 1))
        ' 每个按键之间暂停一小段时间，例如100毫秒
        Beep 0, 100
    Next i
End Sub
