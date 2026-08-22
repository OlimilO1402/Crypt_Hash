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
    ha_SHA256 = 1
    ha_SHA384 = 2
    ha_SHA512 = 3
End Enum

Private m_HashAlgo As EBCryptHashAlgo

Friend Sub New_(ByVal aHashAlgo As EBCryptHashAlgo)
    m_HashAlgo = aHashAlgo
End Sub

Public Function GetHash(Value() As Byte) As Byte()
    GetHash = MBCrypt.TryGetHash(Value())
End Function

' v ############################## v '    Implements IHasher    ' v ############################## v '
Private Function IHasher_GetHash(Value() As Byte) As Byte()
    IHasher_GetHash = Me.GetHash(Value)
End Function
