VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSHA256"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************
' MODULE:       CSHA256
' FILENAME:     CSHA256.Cls
' AUTHOR:       Phil Fresle
' CREATED:      10-Apr-2001
' COPYRIGHT:    Copyright 2001 Phil Fresle. All Rights Reserved.
'
' This is 'free' software with the following restrictions:
'
' You may not redistribute this code as a 'sample' or 'demo'.
' However, you are free to use the source code in your own code,
' but you may not claim that you created the sample code.
' It is expressly forbidden to sell or profit from this source
' code other than by the knowledge gained or the enhanced value
' added by your own code.
'
' Use of this software is also done so at your own risk. The code
' is supplied as is without warranty or guarantee of any kind.
'
' Should you wish to commission some derivative work based on this
' code provided here, or any consultancy work, please do not
' hesitate to contact us.
'
' Web Site: http://www.frez.co.uk
' e-mail:   sales@frez.co.uk
'****************************************************

Option Explicit

Private m_lOnBits(30)   As Long
Private m_l2Power(30)   As Long
Private K(63)           As Long

Private Const BITS_TO_A_BYTE  As Long = 8
Private Const BYTES_TO_A_WORD As Long = 4
Private Const BITS_TO_A_WORD  As Long = BYTES_TO_A_WORD * BITS_TO_A_BYTE
'----------------------------------------------------

Private Sub Class_Initialize()
    ' Could have done this with a loop calculating each value, but simply
    ' assigning the values is quicker - BITS SET FROM RIGHT
    Dim i As Long
    m_lOnBits(i) = 1:          m_l2Power(i) = 1:          i = i + 1 ' 00000000000000000000000000000001' 00000000000000000000000000000001
    m_lOnBits(i) = 3:          m_l2Power(i) = 2:          i = i + 1 ' 00000000000000000000000000000011' 00000000000000000000000000000010
    m_lOnBits(i) = 7:          m_l2Power(i) = 4:          i = i + 1 ' 00000000000000000000000000000111' 00000000000000000000000000000100
    m_lOnBits(i) = 15:         m_l2Power(i) = 8:          i = i + 1 ' 00000000000000000000000000001111' 00000000000000000000000000001000
    m_lOnBits(i) = 31:         m_l2Power(i) = 16:         i = i + 1 ' 00000000000000000000000000011111' 00000000000000000000000000010000
    m_lOnBits(i) = 63:         m_l2Power(i) = 32:         i = i + 1 ' 00000000000000000000000000111111' 00000000000000000000000000100000
    m_lOnBits(i) = 127:        m_l2Power(i) = 64:         i = i + 1 ' 00000000000000000000000001111111' 00000000000000000000000001000000
    m_lOnBits(i) = 255:        m_l2Power(i) = 128:        i = i + 1 ' 00000000000000000000000011111111' 00000000000000000000000010000000
    m_lOnBits(i) = 511:        m_l2Power(i) = 256:        i = i + 1 ' 00000000000000000000000111111111' 00000000000000000000000100000000
    m_lOnBits(i) = 1023:       m_l2Power(i) = 512:        i = i + 1 ' 00000000000000000000001111111111' 00000000000000000000001000000000
    m_lOnBits(i) = 2047:       m_l2Power(i) = 1024:       i = i + 1 ' 00000000000000000000011111111111' 00000000000000000000010000000000
    m_lOnBits(i) = 4095:       m_l2Power(i) = 2048:       i = i + 1 ' 00000000000000000000111111111111' 00000000000000000000100000000000
    m_lOnBits(i) = 8191:       m_l2Power(i) = 4096:       i = i + 1 ' 00000000000000000001111111111111' 00000000000000000001000000000000
    m_lOnBits(i) = 16383:      m_l2Power(i) = 8192:       i = i + 1 ' 00000000000000000011111111111111' 00000000000000000010000000000000
    m_lOnBits(i) = 32767:      m_l2Power(i) = 16384:      i = i + 1 ' 00000000000000000111111111111111' 00000000000000000100000000000000
    m_lOnBits(i) = 65535:      m_l2Power(i) = 32768:      i = i + 1 ' 00000000000000001111111111111111' 00000000000000001000000000000000
    m_lOnBits(i) = 131071:     m_l2Power(i) = 65536:      i = i + 1 ' 00000000000000011111111111111111' 00000000000000010000000000000000
    m_lOnBits(i) = 262143:     m_l2Power(i) = 131072:     i = i + 1 ' 00000000000000111111111111111111' 00000000000000100000000000000000
    m_lOnBits(i) = 524287:     m_l2Power(i) = 262144:     i = i + 1 ' 00000000000001111111111111111111' 00000000000001000000000000000000
    m_lOnBits(i) = 1048575:    m_l2Power(i) = 524288:     i = i + 1 ' 00000000000011111111111111111111' 00000000000010000000000000000000
    m_lOnBits(i) = 2097151:    m_l2Power(i) = 1048576:    i = i + 1 ' 00000000000111111111111111111111' 00000000000100000000000000000000
    m_lOnBits(i) = 4194303:    m_l2Power(i) = 2097152:    i = i + 1 ' 00000000001111111111111111111111' 00000000001000000000000000000000
    m_lOnBits(i) = 8388607:    m_l2Power(i) = 4194304:    i = i + 1 ' 00000000011111111111111111111111' 00000000010000000000000000000000
    m_lOnBits(i) = 16777215:   m_l2Power(i) = 8388608:    i = i + 1 ' 00000000111111111111111111111111' 00000000100000000000000000000000
    m_lOnBits(i) = 33554431:   m_l2Power(i) = 16777216:   i = i + 1 ' 00000001111111111111111111111111' 00000001000000000000000000000000
    m_lOnBits(i) = 67108863:   m_l2Power(i) = 33554432:   i = i + 1 ' 00000011111111111111111111111111' 00000010000000000000000000000000
    m_lOnBits(i) = 134217727:  m_l2Power(i) = 67108864:   i = i + 1 ' 00000111111111111111111111111111' 00000100000000000000000000000000
    m_lOnBits(i) = 268435455:  m_l2Power(i) = 134217728:  i = i + 1 ' 00001111111111111111111111111111' 00001000000000000000000000000000
    m_lOnBits(i) = 536870911:  m_l2Power(i) = 268435456:  i = i + 1 ' 00011111111111111111111111111111' 00010000000000000000000000000000
    m_lOnBits(i) = 1073741823: m_l2Power(i) = 536870912:  i = i + 1 ' 00111111111111111111111111111111' 00100000000000000000000000000000
    m_lOnBits(i) = 2147483647: m_l2Power(i) = 1073741824: i = i + 1 ' 01111111111111111111111111111111' 01000000000000000000000000000000
    ' Could have done this with a loop calculating each value, but simply
    ' assigning the values is quicker - POWERS OF 2
    
    ' Just put together the K array once
    K(0) = &H428A2F98:  K(1) = &H71374491:  K(2) = &HB5C0FBCF:  K(3) = &HE9B5DBA5:  K(4) = &H3956C25B:  K(5) = &H59F111F1:  K(6) = &H923F82A4:  K(7) = &HAB1C5ED5
    K(8) = &HD807AA98:  K(9) = &H12835B01:  K(10) = &H243185BE: K(11) = &H550C7DC3: K(12) = &H72BE5D74: K(13) = &H80DEB1FE: K(14) = &H9BDC06A7: K(15) = &HC19BF174
    K(16) = &HE49B69C1: K(17) = &HEFBE4786: K(18) = &HFC19DC6:  K(19) = &H240CA1CC: K(20) = &H2DE92C6F: K(21) = &H4A7484AA: K(22) = &H5CB0A9DC: K(23) = &H76F988DA
    K(24) = &H983E5152: K(25) = &HA831C66D: K(26) = &HB00327C8: K(27) = &HBF597FC7: K(28) = &HC6E00BF3: K(29) = &HD5A79147: K(30) = &H6CA6351:  K(31) = &H14292967
    K(32) = &H27B70A85: K(33) = &H2E1B2138: K(34) = &H4D2C6DFC: K(35) = &H53380D13: K(36) = &H650A7354: K(37) = &H766A0ABB: K(38) = &H81C2C92E: K(39) = &H92722C85
    K(40) = &HA2BFE8A1: K(41) = &HA81A664B: K(42) = &HC24B8B70: K(43) = &HC76C51A3: K(44) = &HD192E819: K(45) = &HD6990624: K(46) = &HF40E3585: K(47) = &H106AA070
    K(48) = &H19A4C116: K(49) = &H1E376C08: K(50) = &H2748774C: K(51) = &H34B0BCB5: K(52) = &H391C0CB3: K(53) = &H4ED8AA4A: K(54) = &H5B9CCA4F: K(55) = &H682E6FF3
    K(56) = &H748F82EE: K(57) = &H78A5636F: K(58) = &H84C87814: K(59) = &H8CC70208: K(60) = &H90BEFFFA: K(61) = &HA4506CEB: K(62) = &HBEF9A3F7: K(63) = &HC67178F2
End Sub

'----------------------------------------------------
' LShift As Long
' A left shift takes all the set binary bits and moves them left, in-filling
' with zeros in the vacated bits on the right. This function is equivalent to
' the << operator in Java and C++
' - IN lValue As Long: the shifting value
' - IN iShiftBits As Integer: number of bytess to shift
' returns the shifted number
'----------------------------------------------------
Private Function LShift(ByVal lValue As Long, ByVal iShiftBits As Integer) As Long
    ' NOTE: If you can guarantee that the Shift parameter will be in the
    ' range 1 to 30 you can safely strip of this first nested if structure for
    ' speed.
    
    If iShiftBits = 0 Then
        ' A shift of zero is no shift at all.
        LShift = lValue
        
        Exit Function
    ElseIf iShiftBits = 31 Then
        ' A shift of 31 will result in the right most bit becoming the left most
        ' bit and all other bits being cleared
        If lValue And 1 Then LShift = &H80000000 Else LShift = 0
        Exit Function
    ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
        ' A shift of less than zero or more than 31 is undefined
        Err.Raise 6
    End If
    
    ' If the left most bit that remains will end up in the negative bit
    ' position (&H80000000) we would end up with an overflow if we took the
    ' standard route. We need to strip the left most bit and add it back
    ' afterwards.
    If (lValue And m_l2Power(31 - iShiftBits)) Then
        ' (Value And OnBits(31 - (Shift + 1))) chops off the left most bits that
        ' we are shifting into, but also the left most bit we still want as this
        ' is going to end up in the negative bit marker position (&H80000000).
        ' After the multiplication/shift we Or the result with &H80000000 to
        ' turn the negative bit on.
        LShift = ((lValue And m_lOnBits(31 - (iShiftBits + 1))) * m_l2Power(iShiftBits)) Or &H80000000
    Else
        ' (Value And OnBits(31-Shift)) chops off the left most bits that we are
        ' shifting into so we do not get an overflow error when we do the
        ' multiplication/shift
        LShift = ((lValue And m_lOnBits(31 - iShiftBits)) * m_l2Power(iShiftBits))
    End If
End Function

'----------------------------------------------------
' RShift As Long
' The right shift of an unsigned long integer involves shifting all the set bits
' to the right and in-filling on the left with zeros. This function is
' equivalent to the >> operator in C++ when used on an unsigned long.
' - IN lValue As Long: the shifting value
' - IN iShiftBits As Integer: number of bytes to shift
' returns the shifted number
'----------------------------------------------------
Private Function RShift(ByVal lValue As Long, ByVal iShiftBits As Integer) As Long
    ' NOTE: If you can guarantee that the Shift parameter will be in the
    ' range 1 to 30 you can safely strip of this first nested if structure for
    ' speed.
    
    If iShiftBits = 0 Then
        ' A shift of zero is no shift at all
        RShift = lValue
        Exit Function
    ElseIf iShiftBits = 31 Then
        ' A shift of 31 will clear all bits and move the left most bit to the right
        ' most bit position
        If lValue And &H80000000 Then RShift = 1 Else RShift = 0
        Exit Function
    ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
        ' A shift of less than zero or more than 31 is undefined
        Err.Raise 6
    End If
    
    ' We do not care about the top most bit or the final bit, the top most bit
    ' will be taken into account in the next stage, the final bit (whether it
    ' is an odd number or not) is being shifted into, so we do not give a jot
    ' about it
    RShift = (lValue And &H7FFFFFFE) \ m_l2Power(iShiftBits)
    
    ' If the top most bit (&H80000000) was set we need to do things differently
    ' as in a normal VB signed long integer the top most bit is used to indicate
    ' the sign of the number, when it is set it is a negative number, so just
    ' deviding by a factor of 2 as above would not work.
    ' NOTE: (lValue And  &H80000000) is equivalent to (lValue < 0), you could
    ' get a very marginal speed improvement by changing the test to (lValue < 0)
    If (lValue And &H80000000) Then
        ' We take the value computed so far, and then add the left most negative
        ' bit after it has been shifted to the right the appropriate number of
        ' places
        RShift = (RShift Or (&H40000000 \ m_l2Power(iShiftBits - 1)))
    End If
End Function

'----------------------------------------------------
' AddUnsigned As Long
' Adds two potentially large unsigned numbers without overflowing
' - IN lX As Long: first value
' - IN lY As Long: second value
' returns added number
'----------------------------------------------------
Private Function AddUnsigned(ByVal lX As Long, ByVal lY As Long) As Long
    Dim lX4 As Long: lX4 = lX And &H40000000
    Dim lY4 As Long: lY4 = lY And &H40000000
    Dim lX8 As Long: lX8 = lX And &H80000000
    Dim lY8 As Long: lY8 = lY And &H80000000
    Dim lResult As Long: lResult = (lX And &H3FFFFFFF) + (lY And &H3FFFFFFF)
    If lX4 And lY4 Then
        lResult = lResult Xor &H80000000 Xor lX8 Xor lY8
    ElseIf lX4 Or lY4 Then
        If lResult And &H40000000 Then lResult = lResult Xor &HC0000000 Xor lX8 Xor lY8 Else lResult = lResult Xor &H40000000 Xor lX8 Xor lY8
    Else
        lResult = lResult Xor lX8 Xor lY8
    End If
    AddUnsigned = lResult
End Function

'----------------------------------------------------
' Ch As Long
'----------------------------------------------------
Private Function Ch(ByVal x As Long, ByVal y As Long, ByVal z As Long) As Long
    Ch = ((x And y) Xor ((Not x) And z))
End Function

'----------------------------------------------------
' Maj As Long
'OM: was genau macht diese Funktion?
'----------------------------------------------------
Private Function Maj(ByVal x As Long, ByVal y As Long, ByVal z As Long) As Long
    Maj = ((x And y) Xor (x And z) Xor (y And z))
End Function

'----------------------------------------------------
' s As Long
' right rotation
'----------------------------------------------------
Private Function s(ByVal x As Long, ByVal n As Long) As Long
    s = (RShift(x, (n And m_lOnBits(4))) Or LShift(x, (32 - (n And m_lOnBits(4)))))
End Function

'----------------------------------------------------
' R As Long
' right shift
'----------------------------------------------------
Private Function R(ByVal x As Long, ByVal n As Long) As Long
    R = RShift(x, CInt(n And m_lOnBits(4)))
End Function

'----------------------------------------------------
' Sigma0 As Long
'----------------------------------------------------
Private Function Sigma0(ByVal x As Long) As Long
    Sigma0 = (s(x, 2) Xor s(x, 13) Xor s(x, 22))
End Function

'----------------------------------------------------
' Sigma1 As Long
'----------------------------------------------------
Private Function Sigma1(ByVal x As Long) As Long
    Sigma1 = (s(x, 6) Xor s(x, 11) Xor s(x, 25))
End Function

'----------------------------------------------------
' Gamma0 As Long
'----------------------------------------------------
Private Function Gamma0(ByVal x As Long) As Long
    Gamma0 = (s(x, 7) Xor s(x, 18) Xor R(x, 3))
End Function

'----------------------------------------------------
' Gamma1 As Long
'----------------------------------------------------
Private Function Gamma1(ByVal x As Long) As Long
    Gamma1 = (s(x, 17) Xor s(x, 19) Xor R(x, 10))
End Function

'----------------------------------------------------
' ConvertToWordArray As Long()
' Takes the string message and puts it in a long array with padding according to
' the SHA-256 rules (similar to MD5 routine).
' - IN Message As String: message to convert
' returns the converted message as a Long array
'----------------------------------------------------
Private Function ConvertToWordArray(sMessage As String) As Long()
    
    Dim lWordCount       As Long
    Dim lByte            As Long
    
    Const MODULUS_BITS   As Long = 512
    Const CONGRUENT_BITS As Long = 448
    
    Dim lMessageLength As Long: lMessageLength = Len(sMessage)
    
    ' Get padded number of words. Message needs to be congruent to 448 bits,
    ' modulo 512 bits. If it is exactly congruent to 448 bits, modulo 512 bits
    ' it must still have another 512 bits added. 512 bits = 64 bytes
    ' (or 16 * 4 byte words), 448 bits = 56 bytes. This means lNumberOfWords must
    ' be a multiple of 1lMessageLength6 (i.e. 16 * 4 (bytes) * 8 (bits))
    Dim lNumberOfWords As Long: lNumberOfWords = (((((MODULUS_BITS - CONGRUENT_BITS) \ BITS_TO_A_BYTE)) \ (MODULUS_BITS \ BITS_TO_A_BYTE)) + 1) * (MODULUS_BITS \ BITS_TO_A_WORD)
    'Dim lWordArray()
    ReDim lWordArray(lNumberOfWords - 1) As Long
    
    ' Combine each block of 4 bytes (ascii code of character) into one long
    ' value and store in the message. The high-order (most significant) bit of
    ' each byte is listed first. However, unlike MD5 we put the high-order
    ' (most significant) byte first in each word.
    Dim lBytePosition As Long: lBytePosition = 0
    Dim lByteCount    As Long: lByteCount = 0
    Do Until lByteCount >= lMessageLength
        ' Each word is 4 bytes
        lWordCount = lByteCount \ BYTES_TO_A_WORD
        
        lBytePosition = (3 - (lByteCount Mod BYTES_TO_A_WORD)) * BITS_TO_A_BYTE
        
        ' NOTE: This is where we are using just the first byte of each unicode
        ' character, you may want to make the change here, or to the SHA256 method
        ' so it accepts a byte array.
        lByte = AscB(Mid$(sMessage, lByteCount + 1, 1))
        
        lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(lByte, lBytePosition)
        lByteCount = lByteCount + 1
    Loop
    
    ' Terminate according to SHA-256 rules with a 1 bit, zeros and the length in
    ' bits stored in the last two words
    lWordCount = lByteCount \ BYTES_TO_A_WORD
    lBytePosition = (3 - (lByteCount Mod BYTES_TO_A_WORD)) * BITS_TO_A_BYTE

    ' Add a terminating 1 bit, all the rest of the bits to the end of the
    ' word array will default to zero
    lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(&H80, lBytePosition)
    
    ' We put the length of the message in bits into the last two words, to get
    ' the length in bits we need to multiply by 8 (or left shift 3). This left
    ' shifted value is put in the last word. Any bits shifted off the left edge
    ' need to be put in the penultimate word, we can work out which bits by shifting
    ' right the length by 29 bits.
    lWordArray(lNumberOfWords - 1) = LShift(lMessageLength, 3)
    lWordArray(lNumberOfWords - 2) = RShift(lMessageLength, 29)
    
    ConvertToWordArray = lWordArray
End Function

'----------------------------------------------------
' SHA256 As Long()
' calculates the SHA hash of Message
' - IN Message As String: message to hash, assumes ANSI charset
' returns a Long array containing the hash
'----------------------------------------------------
Public Function SHA256(Message As String) As Long()
    Dim T1 As Long, T2 As Long
    Dim Hash(0 To 7) As Long
    ' Initial hash values
    Hash(0) = &H6A09E667:    Hash(1) = &HBB67AE85:    Hash(2) = &H3C6EF372:    Hash(3) = &HA54FF53A
    Hash(4) = &H510E527F:    Hash(5) = &H9B05688C:    Hash(6) = &H1F83D9AB:    Hash(7) = &H5BE0CD19
    
    ' Preprocessing. Append padding bits and length and convert to words
    Dim W(0 To 63) As Long
    Dim M() As Long: M = ConvertToWordArray(Message)
    
    Dim a As Long, b As Long, c As Long, d As Long, e As Long, f As Long, g As Long, h As Long
    Dim i As Long, j As Long
    For i = 0 To UBound(M) Step 16
        a = Hash(0): b = Hash(1): c = Hash(2): d = Hash(3)
        e = Hash(4): f = Hash(5): g = Hash(6): h = Hash(7)
        
        For j = 0 To 63
            If j < 16 Then
                W(j) = M(j + i)
            Else
                W(j) = AddUnsigned(AddUnsigned(AddUnsigned(Gamma1(W(j - 2)), W(j - 7)), Gamma0(W(j - 15))), W(j - 16))
            End If
            T1 = AddUnsigned(AddUnsigned(AddUnsigned(AddUnsigned(h, Sigma1(e)), Ch(e, f, g)), K(j)), W(j))
            T2 = AddUnsigned(Sigma0(a), Maj(a, b, c))
            h = g: g = f: f = e
            e = AddUnsigned(d, T1)
            d = c: c = b: b = a
            a = AddUnsigned(T1, T2)
        Next
        Hash(0) = AddUnsigned(a, Hash(0)): Hash(1) = AddUnsigned(b, Hash(1)): Hash(2) = AddUnsigned(c, Hash(2)): Hash(3) = AddUnsigned(d, Hash(3))
        Hash(4) = AddUnsigned(e, Hash(4)): Hash(5) = AddUnsigned(f, Hash(5)): Hash(6) = AddUnsigned(g, Hash(6)): Hash(7) = AddUnsigned(h, Hash(7))
    Next
    SHA256 = Hash
End Function
