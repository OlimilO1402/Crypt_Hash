Attribute VB_Name = "MNew"
Option Explicit

Public Function CRC32(aHashAlgo As ECRC32Algo) As CRC32
    Set CRC32 = New CRC32: CRC32.New_ aHashAlgo
End Function

Public Function CSHA256(aHashAlgo As ECRC32Algo) As CSHA256
    Set CSHA256 = New CSHA256 ': CSHA.New_ aHashAlgo
End Function

Public Function MSCrypt(ByVal alg As ECryptHashAlgo) As MSCrypt
    Set MSCrypt = New MSCrypt: MSCrypt.New_ alg
End Function

Public Function BCHash(ByVal alg As EBCryptHashAlgo) As BCHash
    Set BCHash = New BCHash: BCHash.New_ alg
End Function

