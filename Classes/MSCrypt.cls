VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MSCrypt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
#If VBA7 Then
    'https://learn.microsoft.com/de-de/windows/win32/api/wincrypt/nf-wincrypt-cryptacquirecontextw
    Private Declare PtrSafe Function CryptAcquireContextW Lib "advapi32" (ByRef hProv_out As LongPtr, ByVal pszContainer As LongPtr, ByVal pszProvider As LongPtr, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
    'https://learn.microsoft.com/de-de/windows/win32/api/wincrypt/nf-wincrypt-cryptreleasecontext
    Private Declare PtrSafe Function CryptReleaseContext Lib "advapi32" (ByVal hProv As LongPtr, ByVal dwFlags As Long) As Long
    
    'https://learn.microsoft.com/de-de/windows/win32/api/wincrypt/nf-wincrypt-cryptcreatehash
    Private Declare PtrSafe Function CryptCreateHash Lib "advapi32" (ByVal hProv As LongPtr, ByVal Algid As Long, ByVal hSessionKey As LongPtr, ByVal dwFlags As Long, Byref hHash_out As LongPtr) As Long
    'https://learn.microsoft.com/de-de/windows/win32/api/wincrypt/nf-wincrypt-cryptdestroyhash
    Private Declare PtrSafe Function CryptDestroyHash Lib "advapi32" (ByVal hHash As LongPtr) As Long
    
    'https://learn.microsoft.com/de-de/windows/win32/api/wincrypt/nf-wincrypt-crypthashdata
    Private Declare PtrSafe Function CryptHashData Lib "advapi32" (ByVal hHash As LongPtr, ByVal pbData As LongPtr, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
    'https://learn.microsoft.com/de-de/windows/win32/api/wincrypt/nf-wincrypt-cryptgethashparam
    Private Declare PtrSafe Function CryptGetHashParam Lib "advapi32" (ByVal hHash As LongPtr, ByVal dwParam As Long, ByVal pbData As LongPtr, ByRef dwDataLen_out As Long, ByVal dwFlags As Long) As Long
    
    'https://learn.microsoft.com/de-de/windows/win32/api/wincrypt/nf-wincrypt-cryptderivekey
    Private Declare PtrSafe Function CryptDeriveKey Lib "advapi32" (ByVal hProv As LongPtr, ByVal Algid As Long, ByVal hBaseData As LongPtr, ByVal dwFlags As Long, ByVal hSessionKey As LongPtr) As Long
    'https://learn.microsoft.com/de-de/windows/win32/api/wincrypt/nf-wincrypt-cryptdestroykey
    Private Declare PtrSafe Function CryptDestroyKey Lib "advapi32" (ByVal hSessionKey As LongPtr) As Long
#Else
    'https://learn.microsoft.com/de-de/windows/win32/api/wincrypt/nf-wincrypt-cryptacquirecontextw
    Private Declare Function CryptAcquireContextW Lib "advapi32" (ByRef hProv_out As LongPtr, ByVal pszContainer As LongPtr, ByVal pszProvider As LongPtr, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
    'https://learn.microsoft.com/de-de/windows/win32/api/wincrypt/nf-wincrypt-cryptreleasecontext
    Private Declare Function CryptReleaseContext Lib "advapi32" (ByVal hProv As LongPtr, ByVal dwFlags As Long) As Long
    
    'https://learn.microsoft.com/de-de/windows/win32/api/wincrypt/nf-wincrypt-cryptcreatehash
    Private Declare Function CryptCreateHash Lib "advapi32" (ByVal hProv As LongPtr, ByVal Algid As Long, ByVal hSessionKey As LongPtr, ByVal dwFlags As Long, ByRef hHash_out As LongPtr) As Long
    'https://learn.microsoft.com/de-de/windows/win32/api/wincrypt/nf-wincrypt-cryptdestroyhash
    Private Declare Function CryptDestroyHash Lib "advapi32" (ByVal hHash As LongPtr) As Long
    
    'https://learn.microsoft.com/de-de/windows/win32/api/wincrypt/nf-wincrypt-crypthashdata
    Private Declare Function CryptHashData Lib "advapi32" (ByVal hHash As LongPtr, ByVal pbData As LongPtr, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
    'https://learn.microsoft.com/de-de/windows/win32/api/wincrypt/nf-wincrypt-cryptgethashparam
    Private Declare Function CryptGetHashParam Lib "advapi32" (ByVal hHash As LongPtr, ByVal dwParam As Long, ByVal pbData As LongPtr, ByRef dwDataLen_out As Long, ByVal dwFlags As Long) As Long
    
    'https://learn.microsoft.com/de-de/windows/win32/api/wincrypt/nf-wincrypt-cryptderivekey
    Private Declare Function CryptDeriveKey Lib "advapi32" (ByVal hProv As LongPtr, ByVal Algid As Long, ByVal hBaseData As LongPtr, ByVal dwFlags As Long, ByVal hSessionKey As LongPtr) As Long
    'https://learn.microsoft.com/de-de/windows/win32/api/wincrypt/nf-wincrypt-cryptdestroykey
    Private Declare Function CryptDestroyKey Lib "advapi32" (ByVal hSessionKey As LongPtr) As Long
#End If

Implements IHasher

Private Const SERVICE_PROVIDER As String = "Microsoft Base Cryptographic Provider v1.0" & vbNullChar

'Private Const ERROR_SUCCESS          As Long = 0
Private Const KEY_CONTAINER          As String = "MyHashKey" & vbNullChar
Private Const PROV_RSA_FULL          As Long = 1

'CryptCreateHash-dwFlags:
Private Const CRYPT_SECRETDIGEST     As Long = &H1&
'CryptAcquireContext-dwFlags:
Private Const CRYPT_NEWKEYSET        As Long = &H8&
Private Const CRYPT_UPDATE_KEY       As Long = &H8&
Private Const CRYPT_DELETEKEYSET     As Long = &H10&
Private Const CRYPT_MACHINE_KEYSET   As Long = &H20&
Private Const CRYPT_SILENT           As Long = &H40&
Private Const CRYPT_VERIFYCONTEXT    As Long = &HF0000000
Private Const CRYPT_DEFAULT_CONTAINER_OPTIONAL As Long = 0 '???
Private Const CRYPT_SERVER           As Long = &H400

Private Const CRYPT_EXPORTABLE       As Long = &H1
Private Const CRYPT_CREATE_SALT      As Long = &H4
Private Const CRYPT_NO_SALT          As Long = &H10

Private Const KP_SALT                As Long = 2

Private Const ALG_CLASS_DATA_ENCRYPT As Long = &H6000& '24576
Private Const ALG_CLASS_HASH         As Long = &H8000& '32768
Private Const ALG_TYPE_ANY           As Long = 0
Private Const ALG_TYPE_STREAM        As Long = &H800& '2048
Private Const ALG_SID_RC4            As Long = 1
Private Const ALG_SID_MD5            As Long = 3
Private Const ALG_SID_SHA            As Long = 4
'Private Const CALG_RC4               As Long = ((ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_STREAM) Or ALG_SID_RC4)
'Private Const CALG_RC4               As Long = ((ALG_CLASS_HASH Or ALG_TYPE_ANY) Or ALG_SID_RC4)
Private Const CALG_RC4               As Long = (ALG_CLASS_HASH Or ALG_SID_RC4)
'Private Const CALG_MD5               As Long = ((ALG_CLASS_HASH Or ALG_TYPE_ANY) Or ALG_SID_MD5)
Private Const CALG_MD5               As Long = (ALG_CLASS_HASH Or ALG_SID_MD5)
'Private Const CALG_SHA               As Long = ((ALG_CLASS_HASH Or ALG_TYPE_ANY) Or ALG_SID_SHA)
Private Const CALG_SHA               As Long = (ALG_CLASS_HASH Or ALG_SID_SHA)

Private Const AT_KEYEXCHANGE         As Long = 1
Private Const AT_SIGNATURE           As Long = 2

'CryptGetHashParam: dwParam:
Private Const HP_ALGID               As Long = &H1&
Private Const HP_HASHVAL             As Long = &H2&
Private Const HP_HASHSIZE            As Long = &H4&

Private Const SIMPLEBLOB             As Long = 1
Private Const PUBLICKEYBLOB          As Long = 6
Private Const PRIVATEKEYBLOB         As Long = 7

'returnvalues: GetLastError
Private Const ERROR_FILE_NOT_FOUND    As Long = &H2&
Private Const ERROR_INVALID_HANDLE    As Long = &H6&
Private Const ERROR_NOT_ENOUGH_MEMORY As Long = &H8&
Private Const ERROR_INVALID_PARAMETER As Long = &H57& ' =87
'Private Const ERROR_BUSY              As Long = &H68& '=104
Private Const ERROR_BUSY              As Long = &HAA& '=170&

Private Const NTE_BAD_UID             As Long = &H80090001
Private Const NTE_BAD_HASH            As Long = &H80090002
Private Const NTE_BAD_KEY             As Long = &H80090003
Private Const NTE_BAD_LEN             As Long = &H80090004
Private Const NTE_BAD_SIGNATURE       As Long = &H80090006 '-2146893818
Private Const NTE_BAD_ALGID           As Long = &H80090008
Private Const NTE_BAD_FLAGS           As Long = &H80090009
Private Const NTE_BAD_KEY_STATE       As Long = &H8009000B
Private Const NTE_BAD_HASH_STATE      As Long = &H8009000C
Private Const NTE_NO_MEMORY           As Long = &H8009000E
Private Const NTE_EXISTS              As Long = &H8009000F
Private Const NTE_BAD_PROV_TYPE       As Long = &H80090014
Private Const NTE_BAD_KEYSET          As Long = &H80090016
Private Const NTE_PROV_TYPE_NOT_DEF   As Long = &H80090017
Private Const NTE_PROV_TYPE_ENTRY_BAD As Long = &H80090018
Private Const NTE_KEYSET_NOT_DEF      As Long = &H80090019
Private Const NTE_KEYSET_ENTRY_BAD    As Long = &H8009001A
Private Const NTE_PROV_TYPE_NO_MATCH  As Long = &H8009001B
Private Const NTE_SIGNATURE_FILE_BAD  As Long = &H8009001C
Private Const NTE_PROVIDER_DLL_FAIL   As Long = &H8009001D
Private Const NTE_PROV_DLL_NOT_FOUND  As Long = &H8009001E
Private Const NTE_BAD_KEYSET_PARAM    As Long = &H8009001F
Private Const NTE_FAIL                As Long = &H80090020


Public Enum ECryptHashAlgo
    ha_RC4 = 1
    ha_MD5 = 3
    ha_SHA = 4
End Enum

'Private m_sSALT              As String
'Private m_sHash              As String
'Private m_hSessionKey        As LongPtr
'Private hKeyPair             As Long
'Private strPublicPrivateBlob As String
'Private strPublicBlob        As String
'Private strSessionBlob       As String
'Private lngType              As Long
'Private strSig               As String

Private m_hCryptoProv As LongPtr
Private m_HashAlgo    As ECryptHashAlgo
Private m_HashAlgoID  As Long
Private m_hHash       As LongPtr
Private m_bHash()     As Byte
Private m_hr          As Long

Private Sub Class_Initialize()
    ReDim m_bHash(0)
    m_hr = CryptAcquireContextW(m_hCryptoProv, StrPtr(KEY_CONTAINER), StrPtr(SERVICE_PROVIDER), PROV_RSA_FULL, CRYPT_NEWKEYSET)
    If m_hr = 0 Then
        ' Versuchen ein bereits erstelltes Handle zu nutzen, wenn nicht dann...
        m_hr = CryptAcquireContextW(m_hCryptoProv, StrPtr(KEY_CONTAINER), StrPtr(SERVICE_PROVIDER), PROV_RSA_FULL, 0)
        If m_hr = 0 Then
            Err.Raise Err.LastDllError, , "Error in acquiring crypto-provider context"
        End If
    End If
End Sub

Private Sub Class_Terminate()
    'If m_hSessionKey Then CryptDestroyKey m_hSessionKey   ' Key der Session zerstören
    If m_hCryptoProv Then CryptReleaseContext m_hCryptoProv, 0
End Sub

Friend Sub New_(ByVal aHashAlgo As ECryptHashAlgo)
    m_HashAlgo = aHashAlgo
    Select Case m_HashAlgo
    Case ECryptHashAlgo.ha_RC4: m_HashAlgoID = CALG_RC4
    Case ECryptHashAlgo.ha_MD5: m_HashAlgoID = CALG_MD5
    Case ECryptHashAlgo.ha_SHA: m_HashAlgoID = CALG_SHA
    End Select
End Sub

Public Property Get HashAlgo() As ECryptHashAlgo
    HashAlgo = m_HashAlgo
End Property

Public Property Get LastHash() As Byte()
    LastHash = m_bHash
End Property

Public Function GetHash(Value() As Byte) As Byte()
    
    m_hr = CryptCreateHash(m_hCryptoProv, m_HashAlgoID, 0, 0, m_hHash)
    If m_hr = 0 Then Err.Raise Err.LastDllError, , "Error creating hash algorithm"
    Dim lb As Long: lb = LBound(Value)
    Dim ub As Long: ub = UBound(Value)
    Dim ValueLen As Long: ValueLen = ub - lb + 1
    m_hr = CryptHashData(m_hHash, VarPtr(Value(lb)), ValueLen, 0)
    If m_hr = 0 Then Err.Raise Err.LastDllError, , "Error writing hash data"
    Dim hashlen As Long
    m_hr = CryptGetHashParam(m_hHash, HP_HASHVAL, vbNull, hashlen, 0)
    ReDim m_bHash(0 To hashlen - 1)
    m_hr = CryptGetHashParam(m_hHash, HP_HASHVAL, VarPtr(m_bHash(0)), hashlen, 0)
    If m_hr = 0 Then Err.Raise Err.LastDllError, , "Error invalid hash"
    GetHash = m_bHash
    If m_hHash <> 0 Then
        CryptDestroyHash m_hHash
        m_hHash = 0
    End If
    
End Function

' v ############################## v '    Implements IHasher    ' v ############################## v '
Private Function IHasher_GetHash(Value() As Byte) As Byte()
    IHasher_GetHash = Me.GetHash(Value)
End Function
