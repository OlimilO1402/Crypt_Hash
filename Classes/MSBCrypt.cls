VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MSBCrypt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IHasher

Public Enum EBCryptHashAlgo
    ha_None
    ha_RSA             'BCRYPT_RSA_ALGORITHM              'nope
    'ha_RSA_SIGN        'BCRYPT_RSA_SIGN_ALGORITHM 'This algorithm is not currently supported.
    ha_DH              'BCRYPT_DH_ALGORITHM               'nope
    ha_DSA             'BCRYPT_DSA_ALGORITHM              'nope
    ha_RC2             'BCRYPT_RC2_ALGORITHM              'nope
    ha_RC4             'BCRYPT_RC4_ALGORITHM              'nope
    ha_AES             'BCRYPT_AES_ALGORITHM              'nope
    ha_DES             'BCRYPT_DES_ALGORITHM              'nope
    ha_DESX            'BCRYPT_DESX_ALGORITHM             'nope
    ha_3DES            'BCRYPT_3DES_ALGORITHM             'nope
    ha_3DES_112        'BCRYPT_3DES_112_ALGORITHM         'nope
    ha_MD2             'BCRYPT_MD2_ALGORITHM                  'OK
    ha_MD4             'BCRYPT_MD4_ALGORITHM                  'OK
    ha_MD5             'BCRYPT_MD5_ALGORITHM                  'OK
    ha_SHA1            'BCRYPT_SHA1_ALGORITHM                 'OK
    ha_SHA256          'BCRYPT_SHA256_ALGORITHM               'OK
    ha_SHA384          'BCRYPT_SHA384_ALGORITHM               'OK
    ha_SHA512          'BCRYPT_SHA512_ALGORITHM               'OK
    ha_AES_GMAC        'BCRYPT_AES_GMAC_ALGORITHM         'nope
    ha_AES_CMAC        'BCRYPT_AES_CMAC_ALGORITHM         'nope
    ha_ECDSA_P256      'BCRYPT_ECDSA_P256_ALGORITHM       'nope
    ha_ECDSA_P384      'BCRYPT_ECDSA_P384_ALGORITHM       'nope
    ha_ECDSA_P521      'BCRYPT_ECDSA_P521_ALGORITHM       'nope
    ha_ECDH_P256       'BCRYPT_ECDH_P256_ALGORITHM        'nope
    ha_ECDH_P384       'BCRYPT_ECDH_P384_ALGORITHM        'nope
    ha_ECDH_P521       'BCRYPT_ECDH_P521_ALGORITHM        'nope
    ha_RNG             'BCRYPT_RNG_ALGORITHM              'nope
    ha_RNG_FIPS186_DSA 'BCRYPT_RNG_FIPS186_DSA_ALGORITHM  'nope
    ha_RNG_DUAL_EC     'BCRYPT_RNG_DUAL_EC_ALGORITHM      'nope
    ha_SHA3_256        'BCRYPT_SHA3_256_ALGORITHM             'OK
    ha_SHA3_384        'BCRYPT_SHA3_384_ALGORITHM             'OK
    ha_SHA3_512        'BCRYPT_SHA3_512_ALGORITHM             'OK
    ha_CSHAKE128       'BCRYPT_CSHAKE128_ALGORITHM            'OK
    ha_CSHAKE256       'BCRYPT_CSHAKE256_ALGORITHM            'OK
    ha_KMAC128         'BCRYPT_KMAC128_ALGORITHM              'OK
    ha_KMAC256         'BCRYPT_KMAC256_ALGORITHM              'OK
    ha_MAX
End Enum

'https://www.thesslstore.com/blog/difference-sha-1-sha-2-sha-256-hash-algorithms/
'https://learn.microsoft.com/de-de/windows/win32/seccng/creating-a-hash-with-cng

'bcrypt.h
Private Enum ENTSTATUS
'#define BCRYPT_SUCCESS(Status) (((NTSTATUS)(Status)) >= 0)
    STATUS_SUCCESS = 0                    'The function was successful.
    STATUS_UNSUCCESSFUL = &HC0000001      ' -1073741823 '
    STATUS_INVALID_HANDLE = &HC0000008    ' -1073741816 ' The algorithm handle in the hAlgorithm parameter is not valid.
    STATUS_INVALID_PARAMETER = &HC000000D ' -1073741811 ' One or more parameters are not valid.
    STATUS_BUFFER_TOO_SMALL = &HC0000023  ' -1073741789 ' The size of the hash object specified by the cbHashObject parameter is not large enough to hold the hash object.
    STATUS_NOT_SUPPORTED = &HC00000BB     ' -1073741637 ' The algorithm provider specified by the hAlgorithm parameter does not support the hash or MAC interface.
End Enum
#If VBA7 Then
    Private Declare PtrSafe Function BCryptOpenAlgorithmProvider Lib "BCrypt" (ByRef phAlgorithm_out As LongPtr, ByVal pszAlgId As LongPtr, ByVal pszImplementation As LongPtr, ByVal dwFlags As Long) As ENTSTATUS
    Private Declare PtrSafe Function BCryptCloseAlgorithmProvider Lib "BCrypt" (ByVal hAlgorithm As LongPtr, ByVal dwFlags As Long) As ENTSTATUS
    Private Declare PtrSafe Function BCryptGetProperty Lib "BCrypt" (ByVal hObject As LongPtr, ByVal pszProperty As LongPtr, ByVal pbOutput_out As LongPtr, ByVal cbOutput As Long, ByRef pcbResult_out As Long, ByVal dwFlags As Long) As ENTSTATUS
    Private Declare PtrSafe Function BCryptCreateHash Lib "BCrypt" (ByVal hAlgorithm_inout As LongPtr, ByVal phHash_out As LongPtr, ByRef pbHashObject_out As LongPtr, ByVal cbHashObject As Long, ByVal pbSecret As Long, ByVal cbSecret As Long, ByVal dwFlags As Long) As ENTSTATUS
    Private Declare PtrSafe Function BCryptHashData Lib "BCrypt" (ByVal hHash_inout As LongPtr, ByVal pbInput As LongPtr, ByVal cbInput As Long, ByVal dwFlags As Long) As ENTSTATUS
    Private Declare PtrSafe Function BCryptFinishHash Lib "BCrypt" (ByVal hHash_inout As LongPtr, ByVal pbOutput As LongPtr, ByVal cbOutput As Long, ByVal dwFlags As Long) As ENTSTATUS
    Private Declare PtrSafe Function BCryptDestroyHash Lib "BCrypt" (ByVal hHash As LongPtr) As ENTSTATUS
    Private Declare PtrSafe Function HeapAlloc Lib "kernel32" (ByVal hHeap As LongPtr, ByVal dwFlags As Long, ByVal dwBytes As Long) As LongPtr
    Private Declare PtrSafe Function GetProcessHeap Lib "kernel32" () As LongPtr
    Private Declare PtrSafe Function HeapFree Lib "kernel32" (ByVal hHeap As LongPtr, ByVal dwFlags As Long, ByVal lpMem As LongPtr) As Long
#Else
    'https://learn.microsoft.com/en-us/windows/win32/api/bcrypt/nf-bcrypt-bcryptopenalgorithmprovider
    'NTSTATUS BCryptOpenAlgorithmProvider( [out] BCRYPT_ALG_HANDLE *phAlgorithm, [in]  LPCWSTR           pszAlgId, [in]  LPCWSTR           pszImplementation, [in]  ULONG dwFlags);
    Private Declare Function BCryptOpenAlgorithmProvider Lib "BCrypt" (ByRef phAlgorithm_out As LongPtr, ByVal pszAlgId As LongPtr, ByVal pszImplementation As LongPtr, ByVal dwFlags As Long) As ENTSTATUS
    
    'https://learn.microsoft.com/en-us/windows/win32/api/bcrypt/nf-bcrypt-bcryptclosealgorithmprovider
    'NTSTATUS BCryptCloseAlgorithmProvider( [in, out] BCRYPT_ALG_HANDLE hAlgorithm, [in] ULONG dwFlags);
    Private Declare Function BCryptCloseAlgorithmProvider Lib "BCrypt" (ByVal hAlgorithm As LongPtr, ByVal dwFlags As Long) As ENTSTATUS
    
    'https://learn.microsoft.com/en-us/windows/win32/api/bcrypt/nf-bcrypt-bcryptgetproperty
    'NTSTATUS BCryptGetProperty( [in] BCRYPT_HANDLE hObject, [in] LPCWSTR pszProperty, [out] PUCHAR pbOutput, [in] ULONG cbOutput, [out] ULONG *pcbResult, [in] ULONG dwFlags);
    Private Declare Function BCryptGetProperty Lib "BCrypt" (ByVal hObject As LongPtr, ByVal pszProperty As LongPtr, ByVal pbOutput_out As LongPtr, ByVal cbOutput As Long, ByRef pcbResult_out As Long, ByVal dwFlags As Long) As ENTSTATUS
    
    'https://learn.microsoft.com/en-us/windows/win32/api/bcrypt/nf-bcrypt-bcryptcreatehash
    'NTSTATUS BCryptCreateHash([in, out] BCRYPT_ALG_HANDLE hAlgorithm, [out] BCRYPT_HASH_HANDLE *phHash, [out] PUCHAR pbHashObject, [in, optional] ULONG cbHashObject, [in, optional] PUCHAR pbSecret, [in] ULONG cbSecret, [in] ULONG dwFlags);
    Private Declare Function BCryptCreateHash Lib "BCrypt" (ByVal hAlgorithm_inout As LongPtr, ByVal phHash_out As LongPtr, ByRef pbHashObject_out As LongPtr, ByVal cbHashObject As Long, ByVal pbSecret As Long, ByVal cbSecret As Long, ByVal dwFlags As Long) As ENTSTATUS
    
    'https://learn.microsoft.com/en-us/windows/win32/api/bcrypt/nf-bcrypt-bcrypthashdata
    'NTSTATUS BCryptHashData([in, out] BCRYPT_HASH_HANDLE hHash, [in] PUCHAR pbInput, [in] ULONG cbInput, [in] ULONG dwFlags);
    Private Declare Function BCryptHashData Lib "BCrypt" (ByVal hHash_inout As LongPtr, ByVal pbInput As LongPtr, ByVal cbInput As Long, ByVal dwFlags As Long) As ENTSTATUS
    
    'https://learn.microsoft.com/de-de/windows/win32/api/bcrypt/nf-bcrypt-bcryptfinishhash
    'NTSTATUS BCryptFinishHash([in, out] BCRYPT_HASH_HANDLE hHash, [out] PUCHAR pbOutput, [in] ULONG cbOutput, [in] ULONG dwFlags);
    Private Declare Function BCryptFinishHash Lib "BCrypt" (ByVal hHash_inout As LongPtr, ByVal pbOutput As LongPtr, ByVal cbOutput As Long, ByVal dwFlags As Long) As ENTSTATUS
    
    'https://learn.microsoft.com/en-us/windows/win32/api/bcrypt/nf-bcrypt-bcryptdestroyhash
    'NTSTATUS BCryptDestroyHash( [in, out] BCRYPT_HASH_HANDLE hHash );
    Private Declare Function BCryptDestroyHash Lib "BCrypt" (ByVal hHash As LongPtr) As ENTSTATUS
    
    'https://learn.microsoft.com/en-us/windows/win32/api/heapapi/nf-heapapi-heapalloc
    'DECLSPEC_ALLOCATOR LPVOID HeapAlloc( [in] HANDLE hHeap, [in] DWORD dwFlags, [in] SIZE_T dwBytes);
    Private Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As LongPtr, ByVal dwFlags As Long, ByVal dwBytes As Long) As LongPtr
    'Private Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As LongPtr, ByVal dwFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
    
    'https://learn.microsoft.com/en-us/windows/win32/api/heapapi/nf-heapapi-getprocessheap
    'HANDLE GetProcessHeap();
    Private Declare Function GetProcessHeap Lib "kernel32" () As LongPtr
    
    'https://learn.microsoft.com/en-us/windows/win32/api/heapapi/nf-heapapi-heapfree
    'BOOL HeapFree( [in] HANDLE hHeap, [in] DWORD dwFlags, [in] _Frees_ptr_opt_ LPVOID lpMem);
    Private Declare Function HeapFree Lib "kernel32" (ByVal hHeap As LongPtr, ByVal dwFlags As Long, ByVal lpMem As LongPtr) As Long
#End If

Private Const BCRYPT_KDF_HASH                    As String = "HASH"
Private Const BCRYPT_KDF_HMAC                    As String = "HMAC"
Private Const BCRYPT_KDF_TLS_PRF                 As String = "TLS_PRF"
Private Const BCRYPT_KDF_SP80056A_CONCAT         As String = "SP800_56A_CONCAT"
Private Const BCRYPT_KDF_RAW_SECRET              As String = "TRUNCATE"
Private Const BCRYPT_KDF_HKDF                    As String = "HKDF"

Private Const MS_PRIMITIVE_PROVIDER              As String = "Microsoft Primitive Provider"
Private Const MS_PLATFORM_CRYPTO_PROVIDER        As String = "Microsoft Platform Crypto Provider"

'//
'// Common algorithm identifiers.
'//
Private Const BCRYPT_RSA_ALGORITHM               As String = "RSA"
'Private Const BCRYPT_RSA_SIGN_ALGORITHM          As String = "RSA_SIGN" 'This algorithm is not currently supported.
Private Const BCRYPT_DH_ALGORITHM                As String = "DH"
Private Const BCRYPT_DSA_ALGORITHM               As String = "DSA"
Private Const BCRYPT_RC2_ALGORITHM               As String = "RC2"
Private Const BCRYPT_RC4_ALGORITHM               As String = "RC4"
Private Const BCRYPT_AES_ALGORITHM               As String = "AES"
Private Const BCRYPT_DES_ALGORITHM               As String = "DES"
Private Const BCRYPT_DESX_ALGORITHM              As String = "DESX"
Private Const BCRYPT_3DES_ALGORITHM              As String = "3DES"
Private Const BCRYPT_3DES_112_ALGORITHM          As String = "3DES_112"
Private Const BCRYPT_MD2_ALGORITHM               As String = "MD2"
Private Const BCRYPT_MD4_ALGORITHM               As String = "MD4"
Private Const BCRYPT_MD5_ALGORITHM               As String = "MD5"
Private Const BCRYPT_SHA1_ALGORITHM              As String = "SHA1"
Private Const BCRYPT_SHA256_ALGORITHM            As String = "SHA256"
Private Const BCRYPT_SHA384_ALGORITHM            As String = "SHA384"
Private Const BCRYPT_SHA512_ALGORITHM            As String = "SHA512"
Private Const BCRYPT_AES_GMAC_ALGORITHM          As String = "AES-GMAC"
Private Const BCRYPT_AES_CMAC_ALGORITHM          As String = "AES-CMAC"
Private Const BCRYPT_ECDSA_P256_ALGORITHM        As String = "ECDSA_P256"
Private Const BCRYPT_ECDSA_P384_ALGORITHM        As String = "ECDSA_P384"
Private Const BCRYPT_ECDSA_P521_ALGORITHM        As String = "ECDSA_P521" '<<<<<-----------????????
Private Const BCRYPT_ECDH_P256_ALGORITHM         As String = "ECDH_P256"
Private Const BCRYPT_ECDH_P384_ALGORITHM         As String = "ECDH_P384"
Private Const BCRYPT_ECDH_P521_ALGORITHM         As String = "ECDH_P521" '<<<<<-----------????????
Private Const BCRYPT_RNG_ALGORITHM               As String = "RNG"
Private Const BCRYPT_RNG_FIPS186_DSA_ALGORITHM   As String = "FIPS186DSARNG"
Private Const BCRYPT_RNG_DUAL_EC_ALGORITHM       As String = "DUALECRNG"

'#if (NTDDI_VERSION >= NTDDI_WIN8)
Private Const BCRYPT_SP800108_CTR_HMAC_ALGORITHM As String = "SP800_108_CTR_HMAC"
Private Const BCRYPT_SP80056A_CONCAT_ALGORITHM   As String = "SP800_56A_CONCAT"
Private Const BCRYPT_PBKDF2_ALGORITHM            As String = "PBKDF2"
Private Const BCRYPT_CAPI_KDF_ALGORITHM          As String = "CAPI_KDF"
Private Const BCRYPT_TLS1_1_KDF_ALGORITHM        As String = "TLS1_1_KDF"
Private Const BCRYPT_TLS1_2_KDF_ALGORITHM        As String = "TLS1_2_KDF"
'#End If

'#if (NTDDI_VERSION >= NTDDI_WINTHRESHOLD)
Private Const BCRYPT_ECDSA_ALGORITHM             As String = "ECDSA"
Private Const BCRYPT_ECDH_ALGORITHM              As String = "ECDH"
Private Const BCRYPT_XTS_AES_ALGORITHM           As String = "XTS-AES"
'#End If

'#if (NTDDI_VERSION >= NTDDI_WIN10_RS4)
Private Const BCRYPT_HKDF_ALGORITHM              As String = "HKDF"
'#End If

'#if (NTDDI_VERSION >= NTDDI_WIN10_FE)
Private Const BCRYPT_CHACHA20_POLY1305_ALGORITHM  As String = "CHACHA20_POLY1305"
'#End If

'#if (NTDDI_VERSION >= NTDDI_WIN11_ZN)
Private Const BCRYPT_SHA3_256_ALGORITHM           As String = "SHA3-256"
Private Const BCRYPT_SHA3_384_ALGORITHM           As String = "SHA3-384"
Private Const BCRYPT_SHA3_512_ALGORITHM           As String = "SHA3-512"
Private Const BCRYPT_CSHAKE128_ALGORITHM          As String = "CSHAKE128"
Private Const BCRYPT_CSHAKE256_ALGORITHM          As String = "CSHAKE256"
Private Const BCRYPT_KMAC128_ALGORITHM            As String = "KMAC128"
Private Const BCRYPT_KMAC256_ALGORITHM            As String = "KMAC256"
'#End If

'//
'// Interfaces
'//
Private Const BCRYPT_CIPHER_INTERFACE     As Long = &H1
Private Const BCRYPT_HASH_INTERFACE       As Long = &H2
Private Const BCRYPT_ASYMMETRIC_ENCRYPTION_INTERFACE As Long = &H3
Private Const BCRYPT_SECRET_AGREEMENT_INTERFACE      As Long = &H4
Private Const BCRYPT_SIGNATURE_INTERFACE  As Long = &H5
Private Const BCRYPT_RNG_INTERFACE        As Long = &H6

'#if (NTDDI_VERSION >= NTDDI_WIN8)
Private Const BCRYPT_KEY_DERIVATION_INTERFACE As Long = &H7
'#End If

Private Const KDF_HASH_ALGORITHM          As Long = &H0
Private Const KDF_SECRET_PREPEND          As Long = &H1
Private Const KDF_SECRET_APPEND           As Long = &H2
Private Const KDF_HMAC_KEY                As Long = &H3
Private Const KDF_TLS_PRF_LABEL           As Long = &H4
Private Const KDF_TLS_PRF_SEED            As Long = &H5
Private Const KDF_SECRET_HANDLE           As Long = &H6

'#if (NTDDI_VERSION >= NTDDI_WIN7)
Private Const KDF_TLS_PRF_PROTOCOL        As Long = &H7
Private Const KDF_ALGORITHMID             As Long = &H8
Private Const KDF_PARTYUINFO              As Long = &H9
Private Const KDF_PARTYVINFO              As Long = &HA
Private Const KDF_SUPPPUBINFO             As Long = &HB
Private Const KDF_SUPPPRIVINFO            As Long = &HC

Private Const KDF_LABEL                   As Long = &HD
Private Const KDF_CONTEXT                 As Long = &HE
Private Const KDF_SALT                    As Long = &HF

Private Const KDF_ITERATION_COUNT         As Long = &H10

Private Const KDF_KEYBITLENGTH            As Long = &H12
Private Const KDF_HKDF_SALT               As Long = &H13 ' This is used only for testing purposes
Private Const KDF_HKDF_INFO               As Long = &H14

'// BCryptGetProperty strings
Private Const BCRYPT_OBJECT_LENGTH        As String = "ObjectLength"        ' As Long
Private Const BCRYPT_ALGORITHM_NAME       As String = "AlgorithmName"       ' As String
Private Const BCRYPT_PROVIDER_HANDLE      As String = "ProviderHandle"      ' As LongPtr
Private Const BCRYPT_CHAINING_MODE        As String = "ChainingMode"        ' As Long
Private Const BCRYPT_BLOCK_LENGTH         As String = "BlockLength"         ' As Long
Private Const BCRYPT_KEY_LENGTH           As String = "KeyLength"           ' As Long
Private Const BCRYPT_KEY_OBJECT_LENGTH    As String = "KeyObjectLength"     ' As Long
Private Const BCRYPT_KEY_STRENGTH         As String = "KeyStrength"         ' As Long
Private Const BCRYPT_KEY_LENGTHS          As String = "KeyLengths"          ' As Long
Private Const BCRYPT_BLOCK_SIZE_LIST      As String = "BlockSizeList"       ' As Long
Private Const BCRYPT_EFFECTIVE_KEY_LENGTH As String = "EffectiveKeyLength"  ' As Long
Private Const BCRYPT_HASH_LENGTH          As String = "HashDigestLength"    ' As Long
Private Const BCRYPT_HASH_OID_LIST        As String = "HashOIDList"         ' As Long
Private Const BCRYPT_PADDING_SCHEMES      As String = "PaddingSchemes"      ' As Long
Private Const BCRYPT_SIGNATURE_LENGTH     As String = "SignatureLength"     ' As Long
Private Const BCRYPT_HASH_BLOCK_LENGTH    As String = "HashBlockLength"     ' As Long
Private Const BCRYPT_AUTH_TAG_LENGTH      As String = "AuthTagLength"       ' As Long

'#if (NTDDI_VERSION >= NTDDI_WIN11_ZN)
Private Const BCRYPT_FUNCTION_NAME_STRING As String = "FunctionNameString"  ' As String
Private Const BCRYPT_CUSTOMIZATION_STRING As String = "CustomizationString" ' As String
'#End If

'#if (NTDDI_VERSION >= NTDDI_WIN7)
Private Const BCRYPT_PRIMITIVE_TYPE       As String = "PrimitiveType"       ' As Long
Private Const BCRYPT_IS_KEYED_HASH        As String = "IsKeyedHash"         ' As Boolean
'#End If

'#if (NTDDI_VERSION >= NTDDI_WIN8)
Private Const BCRYPT_IS_REUSABLE_HASH     As String = "IsReusableHash"      ' As Boolean
Private Const BCRYPT_MESSAGE_BLOCK_LENGTH As String = "MessageBlockLength"  ' As Long
'#End If

'#if (NTDDI_VERSION >= NTDDI_WIN8)
Private Const BCRYPT_PUBLIC_KEY_LENGTH    As String = "PublicKeyLength"     ' As Long
'#End If

'// Additional BCryptGetProperty strings for the RNG Platform Crypto Provider
Private Const BCRYPT_PCP_PLATFORM_TYPE_PROPERTY    As String = "PCP_PLATFORM_TYPE"
Private Const BCRYPT_PCP_PROVIDER_VERSION_PROPERTY As String = "PCP_PROVIDER_VERSION"

'#if (NTDDI_VERSION > NTDDI_WINBLUE || (NTDDI_VERSION == NTDDI_WINBLUE && defined(WINBLUE_KBSPRING14)))
Private Const BCRYPT_MULTI_OBJECT_LENGTH  As String = "MultiObjectLength"
'#End If

'#if (NTDDI_VERSION >= NTDDI_WIN10_RS4)
Private Const BCRYPT_IS_IFX_TPM_WEAK_KEY  As String = "IsIfxTpmWeakKey"

'Private Const BCRYPT_SHA256_ALGORITHM As String = "SHA256"
'Private Const BCRYPT_OBJECT_LENGTH    As Long = 0
Private Const sizeof_DWORD            As Long = 4

Private m_HashAlgo     As EBCryptHashAlgo
Private m_status       As ENTSTATUS 'NTSTATUS           status       = STATUS_UNSUCCESSFUL;
Private m_hAlgo        As LongPtr   'BCRYPT_ALG_HANDLE  hAlg         = NULL;
Private m_hHash        As LongPtr   'BCRYPT_HASH_HANDLE hHash        = NULL;
Private m_cbHashObject As Long      'DWORD              cbHashObject = 0;
Private m_pbHashObject As LongPtr   'PBYTE              pbHashObject = NULL;
Private m_cbHash       As Long      'DWORD              cbHash       = 0;
Private m_pbHash       As LongPtr   'PBYTE              pbHash       = NULL;
Private m_cbData       As Long      'DWORD              cbData       = 0;

Private Sub Class_Initialize()
    'm_HashAlgo = EBCryptHashAlgo.ha_SHA256
End Sub

Private Sub Class_Terminate()
    CleanUp
End Sub

Friend Sub New_(ByVal aHashAlgo As EBCryptHashAlgo)
    m_HashAlgo = aHashAlgo
    'open a handle to an algorithm provider
    If Not TryGetAlgorithmProvider(m_hAlgo, m_HashAlgo) Then
        'MsgBox "**** Error " & m_status & " returned by BCryptOpenAlgorithmProvider"
        ShowMsgIfErr m_status, "BCryptOpenAlgorithmProvider"
        Call CleanUp: Exit Sub
    End If
    'Debug.Print m_hAlgo & " = 0x" & Hex(m_hAlgo)
    'calculate the size of the buffer to hold the hash object
    m_cbHashObject = Me.ObjectLength
    If m_cbHashObject = 0 Then Exit Sub
    'If Not AlgoProv_TryGetObjectLength(m_cbHashObject) Then
    '    MsgBox "**** Error " & m_status & " returned by BCryptGetProperty"
    '    Call CleanUp: Exit Function
    'End If
    'Debug.Print m_cbHashObject & " = 0x" & Hex(m_cbHashObject)
    'allocate the hash object on the heap
    m_pbHashObject = HeapAlloc(GetProcessHeap, 0, m_cbHashObject)
    If m_pbHashObject = 0 Then
        'MsgBox "**** memory allocation for HashObject failed"
        ShowMsgIfErr m_status, "HeapAlloc memory allocation for HashObject failed"
        Call CleanUp: Exit Sub
    End If
    'Debug.Print m_pbHashObject & " = 0x" & Hex(m_pbHashObject)
    'm_cbHashObject = Me.ObjectLength
    m_cbHash = Me.HashLength
    If m_cbHash = 0 Then Exit Sub
    'If Not AlgoProv_TryGetHashLength(m_cbHash) Then
    '    MsgBox "**** Error " & m_status & " returned by BCryptGetProperty"
    '    Call CleanUp: Exit Function
    'End If
    
    'allocate the hash buffer on the heap
    m_pbHash = HeapAlloc(GetProcessHeap, 0, m_cbHash)
    If m_pbHash = 0 Then
        'MsgBox "**** memory allocation for Hash failed"
        ShowMsgIfErr True, "HeapAlloc memory allocation for Hash failed"
        Call CleanUp: Exit Sub
    End If
    'Debug.Print m_pbHash & " = 0x" & Hex(m_pbHash)
    
    'create a hash
    'm_status = BCryptCreateHash(m_hAlgo, VarPtr(m_hHash), ByVal m_pbHashObject, m_cbHashObject, 0, 0, 0)
    'If Not (m_status <> STATUS_SUCCESS) Then
    If Not TryCreateHash_Handle(m_hHash) Then
        'MsgBox "**** Error " & Hex(m_status) & " returned by BCryptCreateHash" & vbCrLf & MErr.WinApiError_ToStr(m_status)
        ShowMsgIfErr m_status, "BCryptCreateHash"
        Call CleanUp: Exit Sub
    End If
    'Debug.Print m_hHash & " = 0x" & Hex(m_hHash)
    
End Sub

Public Property Get WasSuccessful() As Boolean
    WasSuccessful = m_status = ENTSTATUS.STATUS_SUCCESS
End Property

Public Sub HashData(Buffer() As Byte)
    If Not TryHashData(Buffer) Then
        'MsgBox "**** Error " & Hex(m_status) & " returned by BCryptHashData" & vbCrLf & MErr.WinApiError_ToStr(m_status)
        ShowMsgIfErr m_status, "BCryptHashData"
    End If
End Sub

Public Function Finish() As Byte()
    '//close the hash
    Dim HashBytes() As Byte
    If Not TryFinishHash(HashBytes) Then
        'MsgBox "**** Error " & Hex(m_status) & " returned by BCryptFinishHash" & vbCrLf & MErr.WinApiError_ToStr(m_status)
        ShowMsgIfErr m_status, "BCryptFinishHash"
        Call CleanUp: Exit Function
    End If
    Finish = HashBytes
    'Call CleanUp    'Maybe not
End Function

Public Property Get ObjectLength() As Long
    If Not AlgoProv_TryGetObjectLength(ObjectLength) Then ShowMsgIfErr m_status, "BCryptGetProperty ObjectLength"
End Property

Public Property Get HashLength() As Long
    'If m_cbHash = 0 Then
    If Not AlgoProv_TryGetHashLength(HashLength) Then ShowMsgIfErr m_status, "BCryptGetProperty HashDigestLength"
    'End If
    'HashLength = m_cbHash
    'Debug.Print m_cbHash & " = 0x" & Hex(m_cbHash)
End Property

Private Sub ShowMsgIfErr(status As ENTSTATUS, ByVal FncNam As String)
    If status Then MsgBox "**** Error " & status & " returned by " & FncNam & vbCrLf & ENTSTATUS_ToStr(status)
End Sub

'Public Property Get ObjectLength() As Long:         AlgoProv_TryGetPropLng BCRYPT_OBJECT_LENGTH, ObjectLength:                         End Property
Public Property Get AlgorithmName() As String:       AlgoProv_TryGetPropStr BCRYPT_ALGORITHM_NAME, AlgorithmName:                       End Property
Public Property Get ProviderHandle() As Long:        AlgoProv_TryGetPropLng BCRYPT_PROVIDER_HANDLE, ProviderHandle:                     End Property
Public Property Get ChainingMode() As Long:          AlgoProv_TryGetPropLng BCRYPT_CHAINING_MODE, ChainingMode:                         End Property
Public Property Get BlockLength() As Long:           AlgoProv_TryGetPropLng BCRYPT_BLOCK_LENGTH, BlockLength:                           End Property
Public Property Get KeyLength() As Long:             AlgoProv_TryGetPropLng BCRYPT_KEY_LENGTH, KeyLength:                               End Property
Public Property Get KeyObjectLength() As Long:       AlgoProv_TryGetPropLng BCRYPT_KEY_OBJECT_LENGTH, KeyObjectLength:                  End Property
Public Property Get KeyStrength() As Long:           AlgoProv_TryGetPropLng BCRYPT_KEY_STRENGTH, KeyStrength:                           End Property
Public Property Get KeyLengths() As Long:            AlgoProv_TryGetPropLng BCRYPT_KEY_LENGTHS, KeyLengths:                             End Property
Public Property Get BlockSizeList() As Long:         AlgoProv_TryGetPropLng BCRYPT_BLOCK_SIZE_LIST, BlockSizeList:                      End Property
Public Property Get EffectiveKeyLength() As Long:    AlgoProv_TryGetPropLng BCRYPT_EFFECTIVE_KEY_LENGTH, EffectiveKeyLength:            End Property
Public Property Get HashDigestLength() As Long:      AlgoProv_TryGetPropLng BCRYPT_HASH_LENGTH, HashDigestLength:                       End Property
Public Property Get HashOIDList() As Long:           AlgoProv_TryGetPropLng BCRYPT_HASH_OID_LIST, HashOIDList:                          End Property
Public Property Get PaddingSchemes() As Long:        AlgoProv_TryGetPropLng BCRYPT_PADDING_SCHEMES, PaddingSchemes:                     End Property
Public Property Get SignatureLength() As Long:       AlgoProv_TryGetPropLng BCRYPT_SIGNATURE_LENGTH, SignatureLength:                   End Property
Public Property Get HashBlockLength() As Long:       AlgoProv_TryGetPropLng BCRYPT_HASH_BLOCK_LENGTH, HashBlockLength:                  End Property
Public Property Get AuthTagLength() As Long:         AlgoProv_TryGetPropLng BCRYPT_AUTH_TAG_LENGTH, AuthTagLength:                      End Property

'#if (NTDDI_VERSION >= NTDDI_WIN11_ZN)
Public Property Get FunctionNameString() As String:  AlgoProv_TryGetPropStr BCRYPT_FUNCTION_NAME_STRING, FunctionNameString:            End Property
Public Property Get CustomizationString() As String: AlgoProv_TryGetPropStr BCRYPT_CUSTOMIZATION_STRING, CustomizationString:           End Property

'#if (NTDDI_VERSION  >=  NTDDI_WIN7)
Public Property Get PrimitiveType() As Long:         AlgoProv_TryGetPropLng BCRYPT_PRIMITIVE_TYPE, PrimitiveType:                       End Property
Public Property Get IsKeyedHash() As Long:           AlgoProv_TryGetPropLng BCRYPT_IS_KEYED_HASH, IsKeyedHash:                          End Property

'#if (NTDDI_VERSION  >=  NTDDI_WIN8)
Public Property Get IsReusableHash() As Long:        AlgoProv_TryGetPropLng BCRYPT_IS_REUSABLE_HASH, IsReusableHash:                    End Property
Public Property Get MessageBlockLength() As Long:    AlgoProv_TryGetPropLng BCRYPT_MESSAGE_BLOCK_LENGTH, MessageBlockLength:            End Property

'#if (NTDDI_VERSION  >=  NTDDI_WIN8)
Public Property Get PublicKeyLength() As Long:       AlgoProv_TryGetPropLng BCRYPT_PUBLIC_KEY_LENGTH, PublicKeyLength:                  End Property

'// Additional BCryptGetProperty strings for the RNG Platform Crypto Provider
Public Property Get PCP_PLATFORM_TYPE() As Long:     AlgoProv_TryGetPropLng BCRYPT_PCP_PLATFORM_TYPE_PROPERTY, PCP_PLATFORM_TYPE:       End Property
Public Property Get PCP_PROVIDER_VERSION() As Long:  AlgoProv_TryGetPropLng BCRYPT_PCP_PROVIDER_VERSION_PROPERTY, PCP_PROVIDER_VERSION: End Property

'#if (NTDDI_VERSION > NTDDI_WINBLUE || (NTDDI_VERSION == NTDDI_WINBLUE && defined(WINBLUE_KBSPRING14)))
Public Property Get MultiObjectLength() As Long:     AlgoProv_TryGetPropLng BCRYPT_MULTI_OBJECT_LENGTH, MultiObjectLength:              End Property

'#if (NTDDI_VERSION >= NTDDI_WIN10_RS4)
Public Property Get IsIfxTpmWeakKey() As Long:       AlgoProv_TryGetPropLng BCRYPT_IS_IFX_TPM_WEAK_KEY, IsIfxTpmWeakKey:                End Property

Public Function ToStr() As String
    Dim s As String
    s = s & "AlgorithmName:      " & Me.AlgorithmName & vbCrLf
    s = s & "ProviderHandle:     " & Me.ProviderHandle & vbCrLf
    s = s & "ChainingMode:       " & Me.ChainingMode & vbCrLf
    s = s & "BlockLength:        " & Me.BlockLength & vbCrLf
    s = s & "KeyLength:          " & Me.KeyLength & vbCrLf
    s = s & "KeyObjectLength:    " & Me.KeyObjectLength & vbCrLf
    s = s & "KeyStrength:        " & Me.KeyStrength & vbCrLf
    s = s & "KeyLengths:         " & Me.KeyLengths & vbCrLf
    s = s & "BlockSizeList:      " & Me.BlockSizeList & vbCrLf
    s = s & "EffectiveKeyLength: " & Me.EffectiveKeyLength & vbCrLf
    s = s & "HashDigestLength:   " & Me.HashDigestLength & vbCrLf
    s = s & "HashOIDList:        " & Me.HashOIDList & vbCrLf
    s = s & "PaddingSchemes:     " & Me.PaddingSchemes & vbCrLf
    s = s & "SignatureLength:    " & Me.SignatureLength & vbCrLf
    s = s & "HashBlockLength:    " & Me.HashBlockLength & vbCrLf
    s = s & "AuthTagLength:      " & Me.AuthTagLength & vbCrLf
    s = s & "FunctionNameString: " & Me.FunctionNameString & vbCrLf
    s = s & "CustomizationString:" & Me.CustomizationString & vbCrLf
    s = s & "PrimitiveType:      " & Me.PrimitiveType & vbCrLf
    s = s & "IsKeyedHash:        " & Me.IsKeyedHash & vbCrLf
    s = s & "IsReusableHash:     " & Me.IsReusableHash & vbCrLf
    s = s & "MessageBlockLength: " & Me.MessageBlockLength & vbCrLf
    s = s & "PublicKeyLength:    " & Me.PublicKeyLength & vbCrLf
    s = s & "MultiObjectLength:  " & Me.MultiObjectLength & vbCrLf
    s = s & "IsIfxTpmWeakKey:    " & Me.IsIfxTpmWeakKey & vbCrLf
    ToStr = s
End Function

Public Function EBCryptHashAlgo_ToStr(e As EBCryptHashAlgo) As String
    Dim s As String
    Select Case e
    Case EBCryptHashAlgo.ha_RSA:             s = BCRYPT_RSA_ALGORITHM
    'Case EBCryptHashAlgo.ha_RSA_SIGN:        s = BCRYPT_RSA_SIGN_ALGORITHM 'This algorithm is not currently supported.
    Case EBCryptHashAlgo.ha_DH:              s = BCRYPT_DH_ALGORITHM
    Case EBCryptHashAlgo.ha_DSA:             s = BCRYPT_DSA_ALGORITHM
    Case EBCryptHashAlgo.ha_RC2:             s = BCRYPT_RC2_ALGORITHM
    Case EBCryptHashAlgo.ha_RC4:             s = BCRYPT_RC4_ALGORITHM
    Case EBCryptHashAlgo.ha_AES:             s = BCRYPT_AES_ALGORITHM
    Case EBCryptHashAlgo.ha_DES:             s = BCRYPT_DES_ALGORITHM
    Case EBCryptHashAlgo.ha_DESX:            s = BCRYPT_DESX_ALGORITHM
    Case EBCryptHashAlgo.ha_3DES:            s = BCRYPT_3DES_ALGORITHM
    Case EBCryptHashAlgo.ha_3DES_112:        s = BCRYPT_3DES_112_ALGORITHM
    Case EBCryptHashAlgo.ha_MD2:             s = BCRYPT_MD2_ALGORITHM
    Case EBCryptHashAlgo.ha_MD4:             s = BCRYPT_MD4_ALGORITHM
    Case EBCryptHashAlgo.ha_MD5:             s = BCRYPT_MD5_ALGORITHM
    Case EBCryptHashAlgo.ha_SHA1:            s = BCRYPT_SHA1_ALGORITHM
    Case EBCryptHashAlgo.ha_SHA256:          s = BCRYPT_SHA256_ALGORITHM
    Case EBCryptHashAlgo.ha_SHA384:          s = BCRYPT_SHA384_ALGORITHM
    Case EBCryptHashAlgo.ha_SHA512:          s = BCRYPT_SHA512_ALGORITHM
    Case EBCryptHashAlgo.ha_AES_GMAC:        s = BCRYPT_AES_GMAC_ALGORITHM
    Case EBCryptHashAlgo.ha_AES_CMAC:        s = BCRYPT_AES_CMAC_ALGORITHM
    Case EBCryptHashAlgo.ha_ECDSA_P256:      s = BCRYPT_ECDSA_P256_ALGORITHM
    Case EBCryptHashAlgo.ha_ECDSA_P384:      s = BCRYPT_ECDSA_P384_ALGORITHM
    Case EBCryptHashAlgo.ha_ECDSA_P521:      s = BCRYPT_ECDSA_P521_ALGORITHM        'As String = "ECDSA_P521" '521???????????<<<<<<<<-------------
    Case EBCryptHashAlgo.ha_ECDH_P256:       s = BCRYPT_ECDH_P256_ALGORITHM
    Case EBCryptHashAlgo.ha_ECDH_P384:       s = BCRYPT_ECDH_P384_ALGORITHM
    Case EBCryptHashAlgo.ha_ECDH_P521:       s = BCRYPT_ECDH_P521_ALGORITHM         'As String = "ECDH_P521" '521???????????<<<<<<<<-------------
    Case EBCryptHashAlgo.ha_RNG:             s = BCRYPT_RNG_ALGORITHM
    Case EBCryptHashAlgo.ha_RNG_FIPS186_DSA: s = BCRYPT_RNG_FIPS186_DSA_ALGORITHM
    Case EBCryptHashAlgo.ha_RNG_DUAL_EC:     s = BCRYPT_RNG_DUAL_EC_ALGORITHM
    Case EBCryptHashAlgo.ha_SHA3_256:        s = BCRYPT_SHA3_256_ALGORITHM
    Case EBCryptHashAlgo.ha_SHA3_384:        s = BCRYPT_SHA3_384_ALGORITHM
    Case EBCryptHashAlgo.ha_SHA3_512:        s = BCRYPT_SHA3_512_ALGORITHM
    Case EBCryptHashAlgo.ha_CSHAKE128:       s = BCRYPT_CSHAKE128_ALGORITHM
    Case EBCryptHashAlgo.ha_CSHAKE256:       s = BCRYPT_CSHAKE256_ALGORITHM
    Case EBCryptHashAlgo.ha_KMAC128:         s = BCRYPT_KMAC128_ALGORITHM
    Case EBCryptHashAlgo.ha_KMAC256:         s = BCRYPT_KMAC256_ALGORITHM
    End Select
    EBCryptHashAlgo_ToStr = s
End Function

Public Sub HashAlgos_ToCBLB(CBLB)
    Dim i As Long, s As String
    With CBLB
        .Clear
        For i = EBCryptHashAlgo.ha_None To EBCryptHashAlgo.ha_MAX
            s = EBCryptHashAlgo_ToStr(i)
            If Len(s) Then .AddItem s
        Next
    End With
End Sub

Private Function TryGetAlgorithmProvider(ByRef HandleAlgo_out As LongPtr, ByVal ha As EBCryptHashAlgo) As Boolean
Try: On Error GoTo Catch
    ' open an algorithm handle
    Dim sha As String: sha = EBCryptHashAlgo_ToStr(ha)
    m_status = BCryptOpenAlgorithmProvider(HandleAlgo_out, StrPtr(sha), 0, 0)
    TryGetAlgorithmProvider = m_status >= STATUS_SUCCESS
Catch:
End Function

Private Function AlgoProv_TryGetObjectLength(ByRef Value_out As Long) As Boolean
    AlgoProv_TryGetObjectLength = AlgoProv_TryGetPropLng(BCRYPT_OBJECT_LENGTH, Value_out)
End Function

Private Function AlgoProv_TryGetHashLength(ByRef Value_out As Long) As Boolean
    AlgoProv_TryGetHashLength = AlgoProv_TryGetPropLng(BCRYPT_HASH_LENGTH, Value_out)
End Function

Private Function AlgoProv_TryGetPropLng(ByVal PropName As String, Value_out As Long) As Boolean
Try: On Error GoTo Catch
    m_status = BCryptGetProperty(m_hAlgo, StrPtr(PropName), VarPtr(Value_out), sizeof_DWORD, m_cbData, 0)
    AlgoProv_TryGetPropLng = m_status >= STATUS_SUCCESS
    'If m_status Then MsgBox ENTSTATUS_ToStr(m_status)
Catch:
End Function

Private Function AlgoProv_TryGetPropStr(ByVal PropName As String, Value_out As String) As Boolean
Try: On Error GoTo Catch
    'obtain the required size
    Dim cbSize1 As Long, cbSize2 As Long
    m_status = BCryptGetProperty(m_hAlgo, StrPtr(PropName), 0, 0, cbSize1, 0)
    If cbSize1 <= 0 Then Exit Function
    Value_out = Space(cbSize1 / 2)
    m_status = BCryptGetProperty(m_hAlgo, StrPtr(PropName), StrPtr(Value_out), cbSize1, cbSize2, 0)
    Value_out = Left(Value_out, cbSize2 \ 2 - 1)
    AlgoProv_TryGetPropStr = m_status >= STATUS_SUCCESS
    If m_status Then MsgBox ENTSTATUS_ToStr(m_status)
Catch:
End Function

Private Function TryCreateHash_Handle(ByRef Value_out As LongPtr) As Boolean
Try: On Error GoTo Catch
    m_status = BCryptCreateHash(m_hAlgo, VarPtr(Value_out), ByVal m_pbHashObject, m_cbHashObject, 0, 0, 0)
    TryCreateHash_Handle = m_status >= STATUS_SUCCESS
Catch:
End Function

Private Function TryHashData(ByteBuffer() As Byte) As Boolean
Try: On Error GoTo Catch
    Dim lpInput As LongPtr: lpInput = VarPtr(ByteBuffer(0))
    Dim cbInput As Long:    cbInput = UBound(ByteBuffer) + 1
    m_status = BCryptHashData(m_hHash, lpInput, cbInput, 0)
    TryHashData = m_status >= STATUS_SUCCESS
Catch:
End Function

Private Function TryFinishHash(HashResult() As Byte) As Boolean
Try: On Error GoTo Catch
    m_status = BCryptFinishHash(m_hHash, m_pbHash, m_cbHash, 0)
    TryFinishHash = m_status >= STATUS_SUCCESS
    ReDim HashResult(0 To m_cbHash - 1)
    RtlMoveMemory HashResult(0), ByVal m_pbHash, m_cbHash
Catch:
End Function

Private Sub CleanUp()
    If m_status Then
        MsgBox ENTSTATUS_ToStr(m_status)
        m_status = 0
    End If
    If m_hHash Then
        BCryptDestroyHash m_hHash
        m_hHash = 0
    End If
    If m_hAlgo Then
        BCryptCloseAlgorithmProvider m_hAlgo, 0
        m_hAlgo = 0
    End If
    If m_pbHashObject Then
        HeapFree GetProcessHeap, 0, m_pbHashObject
        m_pbHashObject = 0
    End If
    If m_pbHash Then
        HeapFree GetProcessHeap, 0, m_pbHash
        m_pbHash = 0
    End If
End Sub

Private Function ENTSTATUS_ToStr(e As ENTSTATUS) As String
    Dim s As String
    Select Case e
    Case ENTSTATUS.STATUS_SUCCESS:           s = "STATUS_SUCCESS: The function was successful"
    Case ENTSTATUS.STATUS_INVALID_HANDLE:    s = "STATUS_INVALID_HANDLE: The algorithm handle in the hAlgorithm parameter is not valid."
    Case ENTSTATUS.STATUS_UNSUCCESSFUL:      s = "STATUS_UNSUCCESSFUL: No Success "
    Case ENTSTATUS.STATUS_BUFFER_TOO_SMALL:  s = "STATUS_BUFFER_TOO_SMALL: The size of the hash object specified by the cbHashObject parameter is not large enough to hold the hash object."
    Case ENTSTATUS.STATUS_NOT_SUPPORTED:     s = "STATUS_NOT_SUPPORTED: The algorithm provider specified by the hAlgorithm parameter does not support the hash or MAC interface."
    Case ENTSTATUS.STATUS_INVALID_PARAMETER: s = "STATUS_INVALID_PARAMETER: One or more parameters are not valid."
    End Select
    ENTSTATUS_ToStr = s
End Function

Public Function GetHash(Value() As Byte) As Byte()
    If m_HashAlgo = EBCryptHashAlgo.ha_None Then Me.New_ EBCryptHashAlgo.ha_SHA256
    Me.HashData Value: GetHash = Me.Finish
End Function

'das gewöhnliche Pattern von Hash-Algorithmen:
'* Init
'* einmal oder mehrfach Daten rein kippen,
'* Hash finalisieren lassen und abholen.

' v ############################## v '    Implements IHasher    ' v ############################## v '
Private Function IHasher_GetHash(Value() As Byte) As Byte()
    IHasher_GetHash = Me.GetHash(Value)
End Function
