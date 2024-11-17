Attribute VB_Name = "MBCrypt"
Option Explicit

'https://www.thesslstore.com/blog/difference-sha-1-sha-2-sha-256-hash-algorithms/
'https://learn.microsoft.com/de-de/windows/win32/seccng/creating-a-hash-with-cng


'bcrypt.h

Private Enum ENTSTATUS
'#define BCRYPT_SUCCESS(Status) (((NTSTATUS)(Status)) >= 0)
    STATUS_SUCCESS = 0
    STATUS_INVALID_HANDLE = -1
    STATUS_UNSUCCESSFUL = &HC0000001
    
End Enum
'https://learn.microsoft.com/en-us/windows/win32/api/bcrypt/nf-bcrypt-bcryptopenalgorithmprovider
'NTSTATUS BCryptOpenAlgorithmProvider( [out] BCRYPT_ALG_HANDLE *phAlgorithm, [in]  LPCWSTR           pszAlgId, [in]  LPCWSTR           pszImplementation, [in]  ULONG dwFlags);
Private Declare Function BCryptOpenAlgorithmProvider Lib "BCrypt" (ByRef phAlgorithm_out As LongPtr, ByVal pszAlgId As LongPtr, ByVal pszImplementation As LongPtr, ByVal dwFlags As Long) As ENTSTATUS

'https://learn.microsoft.com/en-us/windows/win32/api/bcrypt/nf-bcrypt-bcryptclosealgorithmprovider
'NTSTATUS BCryptCloseAlgorithmProvider( [in, out] BCRYPT_ALG_HANDLE hAlgorithm, [in] ULONG dwFlags);
Private Declare Function BCryptCloseAlgorithmProvider Lib "BCrypt" (ByRef hAlgorithm As LongPtr, ByVal dwFlags As Long) As ENTSTATUS

'https://learn.microsoft.com/en-us/windows/win32/api/bcrypt/nf-bcrypt-bcryptgetproperty
'NTSTATUS BCryptGetProperty( [in] BCRYPT_HANDLE hObject, [in] LPCWSTR pszProperty, [out] PUCHAR pbOutput, [in] ULONG cbOutput, [out] ULONG *pcbResult, [in] ULONG dwFlags);
Private Declare Function BCryptGetProperty Lib "BCrypt" (ByVal hObject As LongPtr, ByVal pszProperty As LongPtr, ByRef pbOutput_out As LongPtr, ByVal cbOutput As Long, ByRef pcbResult_out As Long, ByVal dwFlags As Long) As ENTSTATUS

'https://learn.microsoft.com/en-us/windows/win32/api/bcrypt/nf-bcrypt-bcryptcreatehash
'NTSTATUS BCryptCreateHash([in, out] BCRYPT_ALG_HANDLE hAlgorithm, [out] BCRYPT_HASH_HANDLE *phHash, [out] PUCHAR pbHashObject, [in, optional] ULONG cbHashObject, [in, optional] PUCHAR pbSecret, [in] ULONG cbSecret, [in] ULONG dwFlags);
Private Declare Function BCryptCreateHash Lib "BCrypt" (ByRef hAlgorithm_inout As LongPtr, ByRef phHash_out As LongPtr, ByRef pbHashObject_out As LongPtr, ByVal cbHashObject As Long, ByVal pbSecret As Long, ByVal cbSecret As Long, ByVal dwFlags As Long) As ENTSTATUS

'https://learn.microsoft.com/de-de/windows/win32/api/bcrypt/nf-bcrypt-bcrypthashdata
'NTSTATUS BCryptHashData([in, out] BCRYPT_HASH_HANDLE hHash, [in] PUCHAR pbInput, [in] ULONG cbInput, [in] ULONG dwFlags);
Private Declare Function BCryptHashData Lib "BCrypt" (ByRef hHash_inout As LongPtr, ByVal pbInput As LongPtr, ByVal cbInput As Long, ByVal dwFlags As Long) As ENTSTATUS

'https://learn.microsoft.com/de-de/windows/win32/api/bcrypt/nf-bcrypt-bcryptfinishhash
'NTSTATUS BCryptFinishHash([in, out] BCRYPT_HASH_HANDLE hHash, [out] PUCHAR pbOutput, [in] ULONG cbOutput, [in] ULONG dwFlags);
Private Declare Function BCryptFinishHash Lib "BCrypt" (ByRef hHash_inout As LongPtr, ByRef pbOutput As LongPtr, ByVal cbOutput As Long, ByVal dwFlags As Long) As ENTSTATUS

'https://learn.microsoft.com/en-us/windows/win32/api/bcrypt/nf-bcrypt-bcryptdestroyhash
'NTSTATUS BCryptDestroyHash( [in, out] BCRYPT_HASH_HANDLE hHash );
Private Declare Function BCryptDestroyHash Lib "BCrypt" (ByRef hHash As LongPtr) As ENTSTATUS

'https://learn.microsoft.com/en-us/windows/win32/api/heapapi/nf-heapapi-heapalloc
'DECLSPEC_ALLOCATOR LPVOID HeapAlloc( [in] HANDLE hHeap, [in] DWORD dwFlags, [in] SIZE_T dwBytes);
Private Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As LongPtr, ByVal dwFlags As Long, ByVal dwBytes As Long) As LongPtr

'https://learn.microsoft.com/en-us/windows/win32/api/heapapi/nf-heapapi-getprocessheap
'HANDLE GetProcessHeap();
Private Declare Function GetProcessHeap Lib "kernel32" () As LongPtr

'https://learn.microsoft.com/en-us/windows/win32/api/heapapi/nf-heapapi-heapfree
'BOOL HeapFree( [in] HANDLE hHeap, [in] DWORD dwFlags, [in] _Frees_ptr_opt_ LPVOID lpMem);
Private Declare Function HeapFree Lib "kernel32" (ByVal hHeap As LongPtr, ByVal dwFlags As Long, ByVal lpMem As LongPtr) As Long

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
Private Const BCRYPT_RSA_SIGN_ALGORITHM          As String = "RSA_SIGN"
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
Private Const BCRYPT_ECDSA_P521_ALGORITHM        As String = "ECDSA_P521"
Private Const BCRYPT_ECDH_P256_ALGORITHM         As String = "ECDH_P256"
Private Const BCRYPT_ECDH_P384_ALGORITHM         As String = "ECDH_P384"
Private Const BCRYPT_ECDH_P521_ALGORITHM         As String = "ECDH_P521"
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
Private Const BCRYPT_CIPHER_INTERFACE                As Long = &H1
Private Const BCRYPT_HASH_INTERFACE                  As Long = &H2
Private Const BCRYPT_ASYMMETRIC_ENCRYPTION_INTERFACE As Long = &H3
Private Const BCRYPT_SECRET_AGREEMENT_INTERFACE      As Long = &H4
Private Const BCRYPT_SIGNATURE_INTERFACE             As Long = &H5
Private Const BCRYPT_RNG_INTERFACE                   As Long = &H6

'#if (NTDDI_VERSION >= NTDDI_WIN8)
Private Const BCRYPT_KEY_DERIVATION_INTERFACE        As Long = &H7
'#End If

Private Const KDF_HASH_ALGORITHM                     As Long = &H0
Private Const KDF_SECRET_PREPEND                     As Long = &H1
Private Const KDF_SECRET_APPEND                      As Long = &H2
Private Const KDF_HMAC_KEY                           As Long = &H3
Private Const KDF_TLS_PRF_LABEL                      As Long = &H4
Private Const KDF_TLS_PRF_SEED                       As Long = &H5
Private Const KDF_SECRET_HANDLE                      As Long = &H6

'#if (NTDDI_VERSION >= NTDDI_WIN7)
Private Const KDF_TLS_PRF_PROTOCOL                   As Long = &H7
Private Const KDF_ALGORITHMID                        As Long = &H8
Private Const KDF_PARTYUINFO                         As Long = &H9
Private Const KDF_PARTYVINFO                         As Long = &HA
Private Const KDF_SUPPPUBINFO                        As Long = &HB
Private Const KDF_SUPPPRIVINFO                       As Long = &HC

Private Const KDF_LABEL                              As Long = &HD
Private Const KDF_CONTEXT                            As Long = &HE
Private Const KDF_SALT                               As Long = &HF

Private Const KDF_ITERATION_COUNT                    As Long = &H10

Private Const KDF_KEYBITLENGTH                       As Long = &H12
Private Const KDF_HKDF_SALT                          As Long = &H13 ' This is used only for testing purposes
Private Const KDF_HKDF_INFO                          As Long = &H14

'// BCryptGetProperty strings
Private Const BCRYPT_OBJECT_LENGTH        As String = "ObjectLength"
Private Const BCRYPT_ALGORITHM_NAME       As String = "AlgorithmName"
Private Const BCRYPT_PROVIDER_HANDLE      As String = "ProviderHandle"
Private Const BCRYPT_CHAINING_MODE        As String = "ChainingMode"
Private Const BCRYPT_BLOCK_LENGTH         As String = "BlockLength"
Private Const BCRYPT_KEY_LENGTH           As String = "KeyLength"
Private Const BCRYPT_KEY_OBJECT_LENGTH    As String = "KeyObjectLength"
Private Const BCRYPT_KEY_STRENGTH         As String = "KeyStrength"
Private Const BCRYPT_KEY_LENGTHS          As String = "KeyLengths"
Private Const BCRYPT_BLOCK_SIZE_LIST      As String = "BlockSizeList"
Private Const BCRYPT_EFFECTIVE_KEY_LENGTH As String = "EffectiveKeyLength"
Private Const BCRYPT_HASH_LENGTH          As String = "HashDigestLength"
Private Const BCRYPT_HASH_OID_LIST        As String = "HashOIDList"
Private Const BCRYPT_PADDING_SCHEMES      As String = "PaddingSchemes"
Private Const BCRYPT_SIGNATURE_LENGTH     As String = "SignatureLength"
Private Const BCRYPT_HASH_BLOCK_LENGTH    As String = "HashBlockLength"
Private Const BCRYPT_AUTH_TAG_LENGTH      As String = "AuthTagLength"

'#if (NTDDI_VERSION >= NTDDI_WIN11_ZN)
Private Const BCRYPT_FUNCTION_NAME_STRING As String = "FunctionNameString"
Private Const BCRYPT_CUSTOMIZATION_STRING As String = "CustomizationString"
'#End If

'#if (NTDDI_VERSION >= NTDDI_WIN7)
Private Const BCRYPT_PRIMITIVE_TYPE       As String = "PrimitiveType"
Private Const BCRYPT_IS_KEYED_HASH        As String = "IsKeyedHash"
'#End If

'#if (NTDDI_VERSION >= NTDDI_WIN8)
Private Const BCRYPT_IS_REUSABLE_HASH     As String = "IsReusableHash"
Private Const BCRYPT_MESSAGE_BLOCK_LENGTH As String = "MessageBlockLength"
'#End If

'#if (NTDDI_VERSION >= NTDDI_WIN8)
Private Const BCRYPT_PUBLIC_KEY_LENGTH    As String = "PublicKeyLength"
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




Private m_status       As Long
Private m_hAlgo        As LongPtr
Private m_hHash        As LongPtr
Private m_cbHashObject As Long
Private m_pbHashObject As LongPtr
Private m_cbData       As Long
Private m_pbHash       As LongPtr

Public Function TryGetHash(Bytes() As Byte) As Byte()
    
    'open an algorithm handle
    If Not TryGetAlgorithmProvider(m_hAlgo) Then
        MsgBox "**** Error " & m_status & " returned by BCryptOpenAlgorithmProvider"
        Call CleanUp: Exit Function
    End If
    
    'calculate the size of the buffer to hold the hash object
    If Not TryGetProperty_ObjectLength Then
        MsgBox "**** Error " & m_status & " returned by BCryptGetProperty"
        Call CleanUp: Exit Function
    End If
    
    'allocate the hash object on the heap
    m_pbHashObject = HeapAlloc(GetProcessHeap, 0, m_cbHashObject)
    If m_pbHashObject = 0 Then
        MsgBox "**** memory allocation for HashObject failed"
        Call CleanUp: Exit Function
    End If
    
    'calculate the length of the hash
    If Not TryGetProperty_HashLength Then
        MsgBox "**** Error " & m_status & " returned by BCryptGetProperty"
        Call CleanUp: Exit Function
    End If
    
    'allocate the hash buffer on the heap
    m_pbHash = HeapAlloc(GetProcessHeap, 0, m_cbHashObject)
    If m_pbHash = 0 Then
        MsgBox "**** memory allocation for Hash failed"
        Call CleanUp: Exit Function
    End If
    
    'create a hash
    m_status = BCryptCreateHash(m_hAlgo, m_pbHash, ByVal m_pbHashObject, m_cbHashObject, 0, 0, 0)
    If Not (m_status >= STATUS_SUCCESS) Then
        MsgBox "**** Error " & Hex(m_status) & " returned by BCryptCreateHash" & vbCrLf & MErr.WinApiError_ToStr(m_status)
        Call CleanUp: Exit Function
    End If
    
    'hash some data
    m_status = BCryptHashData(m_hHash, ByVal m_pbHashObject, m_cbHashObject, 0)
    If Not (m_status >= STATUS_SUCCESS) Then
        
    End If
    'if(!NT_SUCCESS(status = BCryptHashData(
    '                                    hHash,
    '                                    (PBYTE)rgbMsg,
    '                                    sizeof(rgbMsg),
    '                                    0)))
    '{
    '    wprintf(L"**** Error 0x%x returned by BCryptHashData\n", status);
    '    goto Cleanup;
    '}
    
    Call CleanUp

'    //create a hash
'    if(!NT_SUCCESS(status = BCryptCreateHash(
'                                        hAlg,
'                                        &hHash,
'                                        pbHashObject,
'                                        cbHashObject,
'                                        NULL,
'                                        0,
'                                        0)))
'    {
'        wprintf(L"**** Error 0x%x returned by BCryptCreateHash\n", status);
'        goto Cleanup;
'    }
'
'
'    //hash some data
'    if(!NT_SUCCESS(status = BCryptHashData(
'                                        hHash,
'                                        (PBYTE)rgbMsg,
'                                        sizeof(rgbMsg),
'                                        0)))
End Function

Private Function TryGetAlgorithmProvider(ByRef HandleAlgo_out As LongPtr) As Boolean
Try: On Error GoTo Catch
    m_status = BCryptOpenAlgorithmProvider(HandleAlgo_out, StrPtr(BCRYPT_SHA256_ALGORITHM), 0, 0)
    TryGetAlgorithmProvider = m_status >= STATUS_SUCCESS
Catch:
End Function

Private Function TryGetProperty_ObjectLength() As Boolean
Try: On Error GoTo Catch
    m_status = BCryptGetProperty(m_hAlgo, StrPtr(BCRYPT_OBJECT_LENGTH), m_cbHashObject, sizeof_DWORD, m_cbData, 0)
    TryGetProperty_ObjectLength = m_status >= STATUS_SUCCESS
Catch:
End Function

Private Function TryGetProperty_HashLength() As Boolean
Try: On Error GoTo Catch
    m_status = BCryptGetProperty(m_hAlgo, StrPtr(BCRYPT_HASH_LENGTH), m_pbHash, sizeof_DWORD, m_cbData, 0)
    TryGetProperty_HashLength = m_status >= STATUS_SUCCESS
Catch:
End Function

Private Function TryHeapAlloc() As Boolean
Try: On Error GoTo Catch
    m_pbHash = HeapAlloc(GetProcessHeap, 0, m_pbHash)
Catch:
End Function

Private Sub CleanUp()
    If m_hAlgo Then
        BCryptCloseAlgorithmProvider m_hAlgo, 0
        m_hAlgo = 0
    End If
    If m_hHash Then
        BCryptDestroyHash m_hHash
        m_hHash = 0
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
'// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF
'// ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO
'// THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
'// PARTICULAR PURPOSE.
'//
'// Copyright (C) Microsoft. All rights reserved.
'/*++
'
'Abstract:
'
'    Sample program for SHA 256 hashing using CNG
'
'--*/
'
'
'#include <windows.h>
'#include <stdio.h>
'#include <bcrypt.h>
'
'
'
'#define NT_SUCCESS(Status)          (((NTSTATUS)(Status)) >= 0)
'
'#define STATUS_UNSUCCESSFUL         ((NTSTATUS)0xC0000001L)
'
'
'static const BYTE rgbMsg[] =
'{
'    0x61, 0x62, 0x63
'};
'
'
'void __cdecl wmain(
'                   int                      argc,
'                   __in_ecount(argc) LPWSTR *wargv)
'{
'
'    BCRYPT_ALG_HANDLE       hAlg            = NULL;
'    BCRYPT_HASH_HANDLE      hHash           = NULL;
'    NTSTATUS                status          = STATUS_UNSUCCESSFUL;
'    DWORD                   cbData          = 0,
'                            cbHash          = 0,
'                            cbHashObject    = 0;
'    PBYTE                   pbHashObject    = NULL;
'    PBYTE                   pbHash          = NULL;
'
'    UNREFERENCED_PARAMETER(argc);
'    UNREFERENCED_PARAMETER(wargv);
'
'    //open an algorithm handle
'    if(!NT_SUCCESS(status = BCryptOpenAlgorithmProvider(
'                                                &hAlg,
'                                                BCRYPT_SHA256_ALGORITHM,
'                                                NULL,
'                                                0)))
'    {
'        wprintf(L"**** Error 0x%x returned by BCryptOpenAlgorithmProvider\n", status);
'        goto Cleanup;
'    }
'
'    //calculate the size of the buffer to hold the hash object
'    if(!NT_SUCCESS(status = BCryptGetProperty(
'                                        hAlg,
'                                        BCRYPT_OBJECT_LENGTH,
'                                        (PBYTE)&cbHashObject,
'                                        sizeof(DWORD),
'                                        &cbData,
'                                        0)))
'    {
'        wprintf(L"**** Error 0x%x returned by BCryptGetProperty\n", status);
'        goto Cleanup;
'    }
'
'    //allocate the hash object on the heap
'    pbHashObject = (PBYTE)HeapAlloc (GetProcessHeap (), 0, cbHashObject);
'    if(NULL == pbHashObject)
'    {
'        wprintf(L"**** memory allocation failed\n");
'        goto Cleanup;
'    }
'
'   //calculate the length of the hash
'    if(!NT_SUCCESS(status = BCryptGetProperty(
'                                        hAlg,
'                                        BCRYPT_HASH_LENGTH,
'                                        (PBYTE)&cbHash,
'                                        sizeof(DWORD),
'                                        &cbData,
'                                        0)))
'    {
'        wprintf(L"**** Error 0x%x returned by BCryptGetProperty\n", status);
'        goto Cleanup;
'    }
'
'    //allocate the hash buffer on the heap
'    pbHash = (PBYTE)HeapAlloc (GetProcessHeap (), 0, cbHash);
'    if(NULL == pbHash)
'    {
'        wprintf(L"**** memory allocation failed\n");
'        goto Cleanup;
'    }
'
'    //create a hash
'    if(!NT_SUCCESS(status = BCryptCreateHash(
'                                        hAlg,
'                                        &hHash,
'                                        pbHashObject,
'                                        cbHashObject,
'                                        NULL,
'                                        0,
'                                        0)))
'    {
'        wprintf(L"**** Error 0x%x returned by BCryptCreateHash\n", status);
'        goto Cleanup;
'    }
'
'
'    //hash some data
'    if(!NT_SUCCESS(status = BCryptHashData(
'                                        hHash,
'                                        (PBYTE)rgbMsg,
'                                        sizeof(rgbMsg),
'                                        0)))
'    {
'        wprintf(L"**** Error 0x%x returned by BCryptHashData\n", status);
'        goto Cleanup;
'    }
'
'    //close the hash
'    if(!NT_SUCCESS(status = BCryptFinishHash(
'                                        hHash,
'                                        pbHash,
'                                        cbHash,
'                                        0)))
'    {
'        wprintf(L"**** Error 0x%x returned by BCryptFinishHash\n", status);
'        goto Cleanup;
'    }
'
'    wprintf(L"Success!\n");
'
'Cleanup:
'
'    if(hAlg)
'    {
'        BCryptCloseAlgorithmProvider(hAlg,0);
'    }
'
'    if (hHash)
'    {
'        BCryptDestroyHash(hHash);
'    }
'
'    if(pbHashObject)
'    {
'        HeapFree(GetProcessHeap(), 0, pbHashObject);
'    }
'
'    if(pbHash)
'    {
'        HeapFree(GetProcessHeap(), 0, pbHash);
'    }
'
'}

