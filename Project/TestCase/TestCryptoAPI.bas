Attribute VB_Name = "TestCryptoAPI"
Option Explicit

Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" ( _
    ByRef phProv As Long, _
    ByVal pszContainer As String, _
    ByVal pszProvider As String, _
    ByVal dwProvType As Long, _
    ByVal dwFlags As Long) As Long

Private Declare Function CryptCreateHash Lib "advapi32.dll" ( _
    ByVal hProv As Long, _
    ByVal Algid As Long, _
    ByVal hKey As Long, _
    ByVal dwFlags As Long, _
    ByRef phHash As Long) As Long

Private Declare Function CryptGenKey Lib "advapi32.dll" ( _
    ByVal hProv As Long, _
    ByVal Algid As Long, _
    ByVal dwFlags As Long, _
    ByRef phKey As Long) As Long


Private Declare Function CryptDeriveKey Lib "advapi32.dll" ( _
    ByVal hProv As Long, _
    ByVal Algid As Long, _
    ByVal hBaseData As Long, _
    ByVal dwFlags As Long, _
    ByRef phKey As Long) As Long

Private Declare Function CryptEncrypt Lib "advapi32.dll" ( _
    ByVal hKey As Long, _
    ByVal hHash As Long, _
    ByVal Final As Long, _
    ByVal dwFlags As Long, _
    ByVal pbData As Byte, _
    ByRef pdwDataLen As Long, _
    ByVal dwBufLen As Long) As Long

Private Declare Function CryptHashData Lib "advapi32.dll" ( _
    ByVal hHash As Long, _
    ByRef pbData As Byte, _
    ByVal dwDataLen As Long, _
    ByVal dwFlags As Long) As Long

Private Declare Function CryptGetHashParam Lib "advapi32.dll" ( _
    ByVal hHash As Long, _
    ByVal dwParam As Long, _
    ByRef pbData As Byte, _
    ByRef pdwDataLen As Long, _
    ByVal dwFlags As Long) As Long

Private Declare Function CryptDestroyKey Lib "advapi32.dll" ( _
    ByVal hKey As Long) As Long

Private Declare Function CryptDestroyHash Lib "advapi32.dll" ( _
    ByVal hHash As Long) As Long

Private Declare Function CryptReleaseContext Lib "advapi32.dll" ( _
    ByVal hProv As Long, _
    ByVal dwFlags As Long) As Long

Private Const PROV_RSA_FULL As Long = 1
Private Const PROV_RSA_AES As Long = 24
Private Const HP_HASHVAL As Long = 2
Private Const ALG_CLASS_DATA_ENCRYPT As Long = (3 * 2 ^ 13)
Private Const ALG_CLASS_HASH As Long = (4 * 2 ^ 13)
Private Const ALG_TYPE_STREAM As Long = (4 * 2 ^ 9)
Private Const ALG_TYPE_ANY As Long = 0
Private Const ALG_SID_MD5 As Long = 3
Private Const ALG_MD5 As Long = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD5)
Private Const ALG_SID_SHA1 As Long = 4
Private Const ALG_SHA1 As Long = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA1)
Private Const ALG_SID_SHA_256 As Long = 12
Private Const ALG_SHA_256 As Long = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA_256)
Private Const ALG_SID_RC4 As Long = 1
Private Const ALG_RC4 As Long = (ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_STREAM Or ALG_SID_RC4)
Private Const CRYPT_VERIFYCONTEXT As Long = &HF0000000


Function MD5Hash(ByVal sText As String) As String
    Dim hProv As Long
    Dim hHash As Long
    Dim bData() As Byte
    Dim bHash(15) As Byte
    Dim sHash As String
    Dim i As Integer

    ' Acquire a cryptographic provider context
    If CryptAcquireContext(hProv, vbNullString, vbNullString, PROV_RSA_FULL, 0) Then
        ' Create an empty hash object
        If CryptCreateHash(hProv, ALG_MD5, 0, 0, hHash) Then
            ' Hash the data
            bData = StrConv(sText, vbFromUnicode)
            If CryptHashData(hHash, bData(0), UBound(bData) + 1, 0) Then
                Dim dwHashLen As Long
                dwHashLen = 16
                ' Get the hash value
                If CryptGetHashParam(hHash, HP_HASHVAL, bHash(0), dwHashLen, 0) Then
                    sHash = ""
                    For i = 0 To 15
                        sHash = sHash & LCase(Right("0" & Hex(bHash(i)), 2))
                    Next i
                    MD5Hash = sHash
                End If
            End If
            Call CryptDestroyHash(hHash)
        End If
        Call CryptReleaseContext(hProv, 0)
    End If
End Function


Function SHA1Hash(ByVal sText As String) As String
    Dim hProv As Long
    Dim hHash As Long
    Dim bData() As Byte
    Dim bHash(19) As Byte ' SHA-1 produces a 160-bit hash, which is 20 bytes long
    Dim sHash As String
    Dim i As Integer

    ' Acquire a cryptographic provider context
    If CryptAcquireContext(hProv, vbNullString, vbNullString, PROV_RSA_FULL, 0) Then
        ' Create an empty hash object
        If CryptCreateHash(hProv, ALG_SHA1, 0, 0, hHash) Then
            ' Hash the data
            bData = StrConv(sText, vbFromUnicode)
            If CryptHashData(hHash, bData(0), UBound(bData) + 1, 0) Then
                Dim dwHashLen As Long
                dwHashLen = 20
                ' Get the hash value
                If CryptGetHashParam(hHash, HP_HASHVAL, bHash(0), dwHashLen, 0) Then
                    sHash = ""
                    For i = 0 To 19
                        sHash = sHash & LCase(Right("0" & Hex(bHash(i)), 2))
                    Next i
                    SHA1Hash = sHash
                End If
            End If
            Call CryptDestroyHash(hHash)
        End If
        Call CryptReleaseContext(hProv, 0)
    End If
End Function


Function SHA256Hash(ByVal sText As String) As String
    Dim hProv As Long
    Dim hHash As Long
    Dim bData() As Byte
    Dim bHash(31) As Byte ' SHA-256 produces a 256-bit hash, which is 32 bytes long
    Dim sHash As String
    Dim i As Integer

    ' Acquire a cryptographic provider context
    If CryptAcquireContext(hProv, vbNullString, vbNullString, PROV_RSA_AES, 0) Then
        ' Create a hash object
        If CryptCreateHash(hProv, ALG_SHA_256, 0, 0, hHash) Then
            ' Hash the data
            bData = StrConv(sText, vbFromUnicode)
            If CryptHashData(hHash, bData(0), UBound(bData) + 1, 0) Then
                Dim dwHashLen As Long
                dwHashLen = 32
                ' Get the hash value
                If CryptGetHashParam(hHash, HP_HASHVAL, bHash(0), dwHashLen, 0) Then
                    sHash = ""
                    For i = 0 To 31
                        sHash = sHash & LCase(Right("0" & Hex(bHash(i)), 2))
                    Next i
                    SHA256Hash = sHash
                End If
            End If
            Call CryptDestroyHash(hHash)
        End If
        Call CryptReleaseContext(hProv, 0)
    End If
End Function

