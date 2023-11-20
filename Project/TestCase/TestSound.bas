Attribute VB_Name = "TestSound"
Option Explicit

' ���� PlaySound API ����
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundW" (ByVal pszSound As String, ByVal hMod As Long, ByVal fdwSound As Long) As Long
' ���� mciSendString API ����
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long



' PlaySound �����ı�־��������
Public Const SND_SYNC = &H0         ' ����ͬ������
Public Const SND_ASYNC = &H1        ' �����첽����
Public Const SND_FILENAME = &H20000 ' �����������ļ���
Public Const SND_NODEFAULT = &H2    ' û��Ĭ������ʱ����������
Public Const SND_LOOP = &H8         ' ѭ������������ֱ����һ�ε��� PlaySound
Public Const SND_NOSTOP = &H10      ' ��ʹ���������ڲ��ţ�Ҳ��ֹͣ��ǰ����


' ����ָ���� WAV �ļ�
Public Function PlayWAVFile(ByVal strSoundFile As String, Optional ByVal bAsync As Boolean = True) As Long
    Dim flags As Long
    flags = SND_FILENAME Or SND_NODEFAULT
    If bAsync Then
        flags = flags Or SND_ASYNC
    Else
        flags = flags Or SND_SYNC
    End If
    
    ' ���� WAV �ļ�
    PlayWAVFile = PlaySound(strSoundFile, 0&, flags)
    Debug.Print PlayWAVFile
End Function


' ʹ�� mciSendString ������ MP3
Public Sub PlayMP3(ByVal mp3FileName As String)
    Dim sCommand As String
    Dim iReturn As Integer

    ' �� MP3 �ļ�
    sCommand = "open """ & mp3FileName & """ type mpegvideo alias mp3"
    iReturn = mciSendString(sCommand, "", 0, 0)

    ' ���� MP3 �ļ�
    sCommand = "play mp3"
    iReturn = mciSendString(sCommand, "", 0, 0)
End Sub


' ֹͣ���ر� MP3
Public Sub StopMP3()
    Dim sCommand As String
    Dim iReturn As Integer

    ' ֹͣ����
    sCommand = "stop mp3"
    iReturn = mciSendString(sCommand, "", 0, 0)

    ' �ر��豸
    sCommand = "close mp3"
    iReturn = mciSendString(sCommand, "", 0, 0)
End Sub


' ���� WAV �ļ�
Public Function PlayWAV(ByVal wavFileName As String) As Boolean
    Dim sCommand As String
    Dim iReturn As Integer
    Dim bRet    As Boolean

    ' �� WAV �ļ�
    sCommand = "open """ & wavFileName & """ alias myWAV"
    iReturn = mciSendString(sCommand, "", 0, 0)
    If iReturn <> 0 Then
        PlayWAV = False
        Debug.Print "������ WAV �ļ�ʧ�ܣ�"
        Exit Function
    End If

    ' ���� WAV �ļ�
    sCommand = "play myWAV wait"
    iReturn = mciSendString(sCommand, "", 0, 0)
    If iReturn <> 0 Then
        ' �ر��豸����Ϊ���˵�û�ܲ���
        mciSendString "close myWAV", "", 0, 0
        PlayWAV = False
        Debug.Print "�������� WAV �ļ�ʧ�ܣ�"
    Else
        PlayWAV = True
    End If
End Function


' ֹͣ���ر� WAV
Public Function StopWAV()
    Dim sCommand As String

    ' ֹͣ����
    sCommand = "stop myWAV"
    Call mciSendString(sCommand, "", 0, 0)

    ' �ر��豸
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
    
    ' ����ÿ��������������ÿ��Ƶ�ʳ���100����
    If freq1 > 0 And freq2 > 0 Then
        Beep freq1, 100
        Beep freq2, 100
    End If
End Sub

Sub PlayPhoneNumber(ByVal phoneNumber As String)
    Dim i As Integer
    For i = 1 To Len(phoneNumber)
        PlayDTMFTone (Mid$(phoneNumber, i, 1))
        ' ÿ������֮����ͣһС��ʱ�䣬����100����
        Beep 0, 100
    Next i
End Sub
