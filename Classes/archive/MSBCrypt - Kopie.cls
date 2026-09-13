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
    ha_SHA256    'Private Const BCRYPT_SHA3_256_ALGORITHM           As String = "SHA3-256"
    ha_SHA384    'Private Const BCRYPT_SHA3_384_ALGORITHM           As String = "SHA3-384"
    ha_SHA512    'Private Const BCRYPT_SHA3_512_ALGORITHM           As String = "SHA3-512"
    ha_CSHAKE128 'Private Const BCRYPT_CSHAKE128_ALGORITHM          As String = "CSHAKE128"
    ha_CSHAKE256 'Private Const BCRYPT_CSHAKE256_ALGORITHM          As String = "CSHAKE256"
    ha_KMAC128   'Private Const BCRYPT_KMAC128_ALGORITHM            As String = "KMAC128"
    ha_KMAC256   'Private Const BCRYPT_KMAC256_ALGORITHM            As String = "KMAC256"
End Enum



Private m_HashAlgo As EBCryptHashAlgo

Friend Sub New_(ByVal aHashAlgo As EBCryptHashAlgo)
    m_HashAlgo = aHashAlgo
End Sub

Public Function GetHash(Value() As Byte) As Byte()
    GetHash = MBCrypt.TryGetHash(Value())
End Function

'das gewöhnliche Pattern von Hash-Algorithmen:
'* Init
'* einmal oder mehrfach Daten rein kippen,
'* Hash finalisieren lassen und abholen.

' v ############################## v '    Implements IHasher    ' v ############################## v '
Private Function IHasher_GetHash(Value() As Byte) As Byte()
    IHasher_GetHash = Me.GetHash(Value)
End Function
