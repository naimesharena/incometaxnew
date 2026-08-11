Attribute VB_Name = "HS256"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
'=====
'HS256 Version 3.1
'=====
'
'A class that calls the Crypto API to create HMAC-SHA256 signatures.
'
'Several useful encoding and decoding operations are included, along
'with conversion to and from UTF8 character encoding.
'
'OS requirements
'---------------
'
'   Windows XP or later.  Some functions require Vista or later.
'
'PROPERTIES
'----------
'
'   None.
'
'METHODS
'-------
'
'   Public Function Decode(
'       ByVal Encoded As String,
'       Optional ByVal Format As EncDecFormat = edfHexRaw) As Byte()
'
'           Accepts a String of data in any of several supported
'           formats and returns a decoded binary Byte array result.
'
'   Public Function Encode(
'       ByRef Bytes() As Byte,
'       Optional ByVal Format As EncDecFormat = edfHexRaw,
'       Optional ByVal Folding As EncodeFolding = efCrLf) As String
'
'           Accepts a Byte array of binary data and returns a String
'           encoded in any of several formats with various line
'           folding options (efNoFolding requires Vista or later).
'           The efNoFolding option does not apply to the edfHexRaw
'           format.
'
'   Public Function FromUTF8(ByRef UTF8() As Byte) As String
'
'           Accepts a Byte array of UTF8 text and returns a String
'           in UTF16 ("Unicode") encoding.
'
'   Public Function HmacSha256(ByRef Data() As Byte) As Byte()
'
'           Creates an HMAC-SHA256 signature using the previously
'           imported key and the data provided as a Byte array,
'           and returns a Byte array result.
'
'           InitHmac() must be called first, and called again
'           whenever a different key is to be used.
'
'   Public Sub InitHmac(ByRef Key() As Byte)
'
'           Prepares to create HMACs by importing the key supplied
'           as a Byte array.
'
'   Public Function MD5(ByRef Data() As Byte) As Byte()
'
'           Calculates the MD5 hash of the provided data in the
'           form of a Byte array and returns it as a Byte array.
'
'   Public Function ToUTF8(ByVal Text As String) As Byte()
'
'           Accepts a String of UTF16 ("Unicode") text and returns
'           a Byte array in UTF8 encoding.
'
'ENUMS
'-----
'
'   Public Enum EncDecFormat
'       edfBase64         Standard Base64 encoding, with line-wrap,
'                         no headers.
'
'       edfHex            Hexadecimal encoding with line-wrap.
'
'       edfHexAscii       Hex plus ASCII encoding with line-wrap.
'
'       edfHexAddr        Hex plus an address/count at the left of
'                         each wrapped line (with line-wrap).
'
'       edfHexAsciiAddr   Hex plus ASCII plus an address/count,
'                         with line-wrap.
'
'       edfHexRaw         Hexadecimal encoding with NO line-wrap.
'   End Enum
'
'   Public Enum EncodeFolding
'       efCrLf            Standard CrLf line-wrap.
'
'       efLf              Lf-only line-wrap.
'
'       efNoFolding       Suppress folding (requires Vista or later).
'   End Enum
'


Private Const CP_UTF8 As Long = 65001

Private Const CALG_RC2                   As Long = &H6602&
Private Const CALG_MD5                   As Long = &H8003&
Private Const CALG_SHA_256               As Long = &H800C&
Private Const CALG_HMAC                  As Long = &H8009&

Private Const PROV_RSA_FULL              As Long = 1
Private Const PROV_RSA_AES               As Long = 24
Private Const CRYPT_VERIFYCONTEXT        As Long = &HF0000000
Private Const CRYPT_MACHINE_KEYSET       As Long = 32
Private Const MS_DEFAULT_PROVIDER        As String = _
    "Microsoft Base Cryptographic Provider v1.0"
Private Const MS_ENH_RSA_AES_PROV        As String = _
    "Microsoft Enhanced RSA and AES Cryptographic Provider"
Private Const MS_ENH_RSA_AES_PROV_XP     As String = _
    "Microsoft Enhanced RSA and AES Cryptographic Provider (Prototype)"

Private Const HP_HASHVAL                 As Long = 2
Private Const HP_HASHSIZE                As Long = 4
Private Const HP_HMAC_INFO               As Long = 5

Private Const CRYPT_STRING_BASE64        As Long = &H1&
Private Const CRYPT_STRING_HEX           As Long = &H4&
Private Const CRYPT_STRING_HEXASCII      As Long = &H5&
Private Const CRYPT_STRING_HEXADDR       As Long = &HA&
Private Const CRYPT_STRING_HEXASCIIADDR  As Long = &HB&
Private Const CRYPT_STRING_HEXRAW        As Long = &HC&       'Requires Vista or later, so we emulate.
Private Const CRYPT_STRING_NOCR          As Long = &H80000000
Private Const CRYPT_STRING_NOCRLF        As Long = &H40000000 'Requires Vista or later!

Private Const CRYPT_IPSEC_HMAC_KEY       As Long = &H100&

Private Const PLAINTEXTKEYBLOB           As Byte = &H8
Private Const CUR_BLOB_VERSION           As Byte = &H2

Private Type HMAC_INFO
    HashAlgId As Long
    pbInnerString As Long
    cbInnerString As Long
    pbOuterString As Long
    cbOuterString As Long
End Type

Private Type BLOBHEADER
    bType As Byte
    bVersion As Byte
    reserved As Integer
    aiKeyAlg As Long
End Type

Private Type KEYBLOB
    hdr As BLOBHEADER
    cbKeySize As Long
    'rgbKeyData() As Byte 'We'll actually append this when we build the Byte array copy.
End Type

#If VBA7 Then
Private Declare PtrSafe Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" ( _
    ByVal Destination As Any, _
    ByVal Source As Any, _
    ByVal Length As LongPtr)

Private Declare PtrSafe Function MultiByteToWideChar Lib "Kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpMultiByteStr As LongPtr, _
    ByVal cbMultiByte As Long, _
    ByVal lpWideCharStr As LongPtr, _
    ByVal cchWideChar As Long) As Long

Private Declare PtrSafe Function WideCharToMultiByte Lib "Kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpWideCharStr As LongPtr, _
    ByVal cchWideChar As Long, _
    ByVal lpMultiByteStr As LongPtr, _
    ByVal cchMultiByte As Long, _
    ByVal lpDefaultChar As LongPtr, _
    ByVal lpUsedDefaultChar As LongPtr) As Long

Private Declare PtrSafe Function CryptAcquireContext Lib "Advapi32" Alias "CryptAcquireContextW" ( _
    ByRef phProv As LongPtr, _
    ByVal pszContainer As LongPtr, _
    ByVal pszProvider As LongPtr, _
    ByVal dwProvType As Long, _
    ByVal dwFlags As Long) As Long

Private Declare PtrSafe Function CryptCreateHash Lib "Advapi32" ( _
    ByVal hProv As LongPtr, _
    ByVal AlgId As Long, _
    ByVal hKey As LongPtr, _
    ByVal dwFlags As Long, _
    ByRef phHash As LongPtr) As Long
    
Private Declare PtrSafe Function CryptDestroyHash Lib "Advapi32" ( _
    ByVal hHash As LongPtr) As Long
    
Private Declare PtrSafe Function CryptDestroyKey Lib "Advapi32" ( _
    ByVal hKey As LongPtr) As Long

Private Declare PtrSafe Function CryptGetHashParam Lib "Advapi32" ( _
    ByVal hHash As LongPtr, _
    ByVal dwParam As Long, _
    ByRef pbData As Any, _
    ByRef pdwDataLen As Long, _
    ByVal dwFlags As Long) As Long

Private Declare PtrSafe Function CryptHashData Lib "Advapi32" ( _
    ByVal hHash As LongPtr, _
    ByRef pbData As Any, _
    ByVal dwDataLen As Long, _
    ByVal dwFlags As Long) As Long

Private Declare PtrSafe Function CryptImportKey Lib "Advapi32" ( _
    ByVal hProv As LongPtr, _
    ByVal pbData As Any, _
    ByVal dwDataLen As Long, _
    ByVal hPubKey As LongPtr, _
    ByVal dwFlags As Long, _
    ByRef phKey As LongPtr) As Long

Private Declare PtrSafe Function CryptReleaseContext Lib "Advapi32" ( _
    ByVal hProv As LongPtr, _
    ByVal dwFlags As Long) As Long

Private Declare PtrSafe Function CryptSetHashParam Lib "Advapi32" ( _
    ByVal hHash As LongPtr, _
    ByVal dwParam As Long, _
    ByRef pbData As HMAC_INFO, _
    ByVal dwFlags As Long) As Long

Private hBaseProvider As LongPtr
Private hAdvProvider As LongPtr
Private hKey As LongPtr
Private hHmacHash As LongPtr
Private TypeNameOfMe As String

#Else
Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" ( _
    ByVal Destination As Any, _
    ByVal Source As Any, _
    ByVal Length As Long)

Private Declare Function MultiByteToWideChar Lib "Kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpMultiByteStr As Long, _
    ByVal cbMultiByte As Long, _
    ByVal lpWideCharStr As Long, _
    ByVal cchWideChar As Long) As Long

Private Declare Function WideCharToMultiByte Lib "Kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpWideCharStr As Long, _
    ByVal cchWideChar As Long, _
    ByVal lpMultiByteStr As Long, _
    ByVal cchMultiByte As Long, _
    ByVal lpDefaultChar As Long, _
    ByVal lpUsedDefaultChar As Long) As Long

Private Declare Function CryptAcquireContext Lib "Advapi32" Alias "CryptAcquireContextW" ( _
    ByRef phProv As Long, _
    ByVal pszContainer As Long, _
    ByVal pszProvider As Long, _
    ByVal dwProvType As Long, _
    ByVal dwFlags As Long) As Long

Private Declare Function CryptCreateHash Lib "Advapi32" ( _
    ByVal hProv As Long, _
    ByVal AlgId As Long, _
    ByVal hKey As Long, _
    ByVal dwFlags As Long, _
    ByRef phHash As Long) As Long
    
Private Declare Function CryptDestroyHash Lib "Advapi32" ( _
    ByVal hHash As Long) As Long
    
Private Declare Function CryptDestroyKey Lib "Advapi32" ( _
    ByVal hKey As Long) As Long

Private Declare Function CryptGetHashParam Lib "Advapi32" ( _
    ByVal hHash As Long, _
    ByVal dwParam As Long, _
    ByRef pbData As Any, _
    ByRef pdwDataLen As Long, _
    ByVal dwFlags As Long) As Long

Private Declare Function CryptHashData Lib "Advapi32" ( _
    ByVal hHash As Long, _
    ByRef pbData As Any, _
    ByVal dwDataLen As Long, _
    ByVal dwFlags As Long) As Long

Private Declare Function CryptImportKey Lib "Advapi32" ( _
    ByVal hProv As Long, _
    ByVal pbData As Any, _
    ByVal dwDataLen As Long, _
    ByVal hPubKey As Long, _
    ByVal dwFlags As Long, _
    ByRef phKey As Long) As Long

Private Declare Function CryptReleaseContext Lib "Advapi32" ( _
    ByVal hProv As Long, _
    ByVal dwFlags As Long) As Long

Private Declare Function CryptSetHashParam Lib "Advapi32" ( _
    ByVal hHash As Long, _
    ByVal dwParam As Long, _
    ByRef pbData As HMAC_INFO, _
    ByVal dwFlags As Long) As Long

Private hBaseProvider As Long
Private hAdvProvider As Long
Private hKey As Long
Private hHmacHash As Long
Private TypeNameOfMe As String
#End If

Public Function FromUTF8(ByRef UTF8() As Byte) As String
    Dim lngOutLen As Long
    Dim strWide As String
    
    lngOutLen = MultiByteToWideChar(CP_UTF8, _
                                    0, _
                                    VarPtr(UTF8(LBound(UTF8))), _
                                    UBound(UTF8) - LBound(UTF8) + 1, _
                                    0, _
                                    0)
    If lngOutLen = 0 Then
        Err.Raise vbObjectError Or &HC312&, _
                  TypeNameOfMe, _
                  "Failed to decode string, system error " _
                & CStr(Err.LastDllError)
    Else
        strWide = String$(lngOutLen, 0)
        lngOutLen = MultiByteToWideChar(CP_UTF8, _
                                        0, _
                                        VarPtr(UTF8(LBound(UTF8))), _
                                        UBound(UTF8) - LBound(UTF8) + 1, _
                                        StrPtr(strWide), _
                                        lngOutLen)
        If lngOutLen = 0 Then
            Err.Raise vbObjectError Or &HC312&, _
                      TypeNameOfMe, _
                      "Failed to decode string, system error " _
                    & CStr(Err.LastDllError)
        Else
            FromUTF8 = strWide
        End If
    End If
End Function

Private Sub Class_Initialize()
    Dim strProvider As String
    
    TypeNameOfMe = TypeName(Me)
    
    'NOTE: Version probe hacks below.  These should defeat any use of
    '      version-lie appcompat shims by naive who try to use this class.
    '
    '      We need these because (a.) Windows 5.1 (XP and Server 2003) do
    '      not have support for CRYPT_STRING_NOCRLF or CRYPT_STRING_HEXRAW,
    '      and (b.) Windows XP does not support MS_ENH_RSA_AES_PROV and so
    '      we must request MS_ENH_RSA_AES_PROV_XP instead.
    
    On Error GoTo 0
    If CryptAcquireContext(hBaseProvider, _
                           0&, _
                           StrPtr(MS_DEFAULT_PROVIDER), _
                           PROV_RSA_FULL, _
                           CRYPT_VERIFYCONTEXT Or CRYPT_MACHINE_KEYSET) = 0 Then
        Err.Raise vbObjectError Or &HC366&, _
                  TypeNameOfMe, _
                  "Failed to obtain CryptoAPI Base context, system error " _
                & CStr(Err.LastDllError)
    ElseIf CryptAcquireContext(hAdvProvider, _
                               0&, _
                               StrPtr(MS_ENH_RSA_AES_PROV), _
                               PROV_RSA_AES, _
                               CRYPT_VERIFYCONTEXT Or CRYPT_MACHINE_KEYSET) = 0 Then
        If CryptAcquireContext(hAdvProvider, _
                                   0&, _
                                   StrPtr(MS_ENH_RSA_AES_PROV_XP), _
                                   PROV_RSA_AES, _
                                   CRYPT_VERIFYCONTEXT Or CRYPT_MACHINE_KEYSET) = 0 Then
            Err.Raise vbObjectError Or &HC368&, _
                      TypeNameOfMe, _
                      "Failed to obtain CryptoAPI RSA AES context, system error " _
                    & CStr(Err.LastDllError)
        End If
    End If
End Sub

Public Function HMACSHA256(ByRef data() As Byte) As Byte()
    Dim lngErr As Long
    Dim HmacInfo As HMAC_INFO
    Dim lngDataLen As Long
    Dim lngHashSize As Long
    Dim bytHashValue() As Byte
    
    If hKey = 0 Then
        Err.Raise vbObjectError Or &HC322&, _
                  TypeNameOfMe, _
                  "No key set, call InitHmac first"
    ElseIf CryptCreateHash(hAdvProvider, CALG_HMAC, hKey, 0, hHmacHash) = 0 Then
        lngErr = Err.LastDllError
        DestroyHandles
        Err.Raise vbObjectError Or &HC32A&, _
                  TypeNameOfMe, _
                  "Failed to create HMAC hash object, system error " _
                & CStr(lngErr)
    Else
        HmacInfo.HashAlgId = CALG_SHA_256
        If CryptSetHashParam(hHmacHash, HP_HMAC_INFO, HmacInfo, 0) = 0 Then
            lngErr = Err.LastDllError
            DestroyHandles
            Err.Raise vbObjectError Or &HC32C&, _
                      TypeNameOfMe, _
                      "Failed to set HMAC_INFO hash param, system error " _
                    & CStr(lngErr)
        ElseIf CryptHashData(hHmacHash, _
                             data(LBound(data)), _
                             UBound(data) - LBound(data) + 1, _
                             0&) = 0 Then
            lngErr = Err.LastDllError
            DestroyHandles
            Err.Raise vbObjectError Or &HC32E&, _
                      TypeNameOfMe, _
                      "Failed to hash data, system error " _
                    & CStr(lngErr)
        Else
            lngDataLen = 4 '4 bytes for Long length.
            If CryptGetHashParam(hHmacHash, HP_HASHSIZE, lngHashSize, lngDataLen, 0&) = 0 Then
                lngErr = Err.LastDllError
                DestroyHandles
                Err.Raise vbObjectError Or &HC332&, _
                          TypeNameOfMe, _
                          "Failed to obtain hash value length, system error " _
                        & CStr(lngErr)
            Else
                lngDataLen = lngHashSize
                ReDim bytHashValue(lngDataLen - 1)
                If CryptGetHashParam(hHmacHash, HP_HASHVAL, bytHashValue(0), lngDataLen, 0&) = 0 Then
                    lngErr = Err.LastDllError
                    DestroyHandles
                    Err.Raise vbObjectError Or &HC334&, _
                              TypeNameOfMe, _
                              "Failed to obtain hash value, system error " _
                            & CStr(lngErr)
                Else
                    DestroyHandles
                    HMACSHA256 = bytHashValue
                End If
            End If
        End If
    End If
End Function

Public Sub InitHmac(ByRef Key() As Byte)
    Dim kbKey As KEYBLOB
    Dim bytKbKey() As Byte
    Dim lngErr As Long

    DestroyHandles
    If hAdvProvider = 0 Then
        Err.Raise vbObjectError Or &HC342&, _
                  TypeNameOfMe, _
                  "No cryptographic RSA AES provider context"
    Else
        With kbKey
            With .hdr
                .bType = PLAINTEXTKEYBLOB
                .bVersion = CUR_BLOB_VERSION
                .aiKeyAlg = CALG_RC2
            End With
            .cbKeySize = UBound(Key) - LBound(Key) + 1
            ReDim bytKbKey(LenB(kbKey) + .cbKeySize - 1)
            CopyMemory VarPtr(bytKbKey(0)), VarPtr(kbKey), LenB(kbKey)
            CopyMemory VarPtr(bytKbKey(LenB(kbKey))), VarPtr(Key(LBound(Key))), .cbKeySize
        End With
        If CryptImportKey(hAdvProvider, _
                          VarPtr(bytKbKey(0)), _
                          UBound(bytKbKey) + 1, _
                          0, _
                          CRYPT_IPSEC_HMAC_KEY, _
                          hKey) = 0 Then
            lngErr = Err.LastDllError
            DestroyHandles
            Err.Raise vbObjectError Or &HC344&, _
                      TypeNameOfMe, _
                      "Failed to import key, system error " _
                    & CStr(lngErr)
        End If
    End If
End Sub

Public Function ToUTF8(ByVal text As String) As Byte()
    Dim lngOutLen As Long
    Dim UTF8() As Byte
    
    lngOutLen = WideCharToMultiByte(CP_UTF8, 0, StrPtr(text), Len(text), _
                                    0, 0, 0, 0)
    ReDim UTF8(lngOutLen - 1)
    WideCharToMultiByte CP_UTF8, 0, StrPtr(text), Len(text), _
                        VarPtr(UTF8(0)), lngOutLen, 0, 0
    ToUTF8 = UTF8
End Function

Private Sub DestroyHandles(Optional ByVal Release As Boolean = False)
    On Error Resume Next 'Handle all exceptions here!
    If hHmacHash <> 0 Then CryptDestroyHash hHmacHash: hHmacHash = 0
    If hKey <> 0 Then CryptDestroyKey hKey: hKey = 0
    If Release And (hBaseProvider <> 0) Then CryptReleaseContext hBaseProvider, 0&: hBaseProvider = 0
    If Release And (hAdvProvider <> 0) Then CryptReleaseContext hAdvProvider, 0&: hAdvProvider = 0
    Err.Clear
End Sub

Private Sub Class_Terminate()
    DestroyHandles Release:=True
End Sub



