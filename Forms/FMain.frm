VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "CryptHash"
   ClientHeight    =   5865
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10455
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FMain.frx":0000
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manuell
   ScaleHeight     =   5865
   ScaleWidth      =   10455
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox TxtBCryptSHA512 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   3720
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   25
      Text            =   "FMain.frx":1782
      Top             =   5040
      Width           =   6735
   End
   Begin VB.TextBox TxtBCryptSHA384 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3720
      MultiLine       =   -1  'True
      TabIndex        =   22
      Text            =   "FMain.frx":1786
      Top             =   4560
      Width           =   6735
   End
   Begin VB.TextBox TxtBCryptSHA256 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3720
      MultiLine       =   -1  'True
      TabIndex        =   19
      Text            =   "FMain.frx":178A
      Top             =   4080
      Width           =   6735
   End
   Begin VB.TextBox TxtMSCryptSHA 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3720
      MultiLine       =   -1  'True
      TabIndex        =   16
      Text            =   "FMain.frx":178E
      Top             =   3600
      Width           =   6735
   End
   Begin VB.TextBox TxtMSCryptMD5 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3720
      MultiLine       =   -1  'True
      TabIndex        =   13
      Text            =   "FMain.frx":1792
      Top             =   3120
      Width           =   6735
   End
   Begin VB.TextBox TxtMSCryptRC4 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3720
      MultiLine       =   -1  'True
      TabIndex        =   10
      Text            =   "FMain.frx":1796
      Top             =   2640
      Width           =   6735
   End
   Begin VB.TextBox TxtCRC32MEF 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3720
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "FMain.frx":179A
      Top             =   2160
      Width           =   6735
   End
   Begin VB.TextBox TxtCRC32JAM 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3720
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "FMain.frx":179E
      Top             =   1680
      Width           =   6735
   End
   Begin VB.CommandButton BtnBCryptSHA512 
      Caption         =   "BCrypt.SHA512"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   23
      Top             =   5040
      Width           =   1935
   End
   Begin VB.CommandButton BtnBCryptSHA384 
      Caption         =   "BCrypt.SHA384"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   20
      Top             =   4560
      Width           =   1935
   End
   Begin VB.CommandButton BtnCRC32MEF 
      Caption         =   "CRC-32/MEF"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CommandButton BtnMSCryptSHA 
      Caption         =   "MSCrypt.SHA"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   14
      Top             =   3600
      Width           =   1935
   End
   Begin VB.CommandButton BtnMSCryptRC4 
      Caption         =   "MSCrypt.RC4"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CommandButton BtnBCryptSHA256 
      Caption         =   "BCrypt.SHA256"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   17
      Top             =   4080
      Width           =   1935
   End
   Begin VB.CommandButton BtnCRC32JAM 
      Caption         =   "CRC-32/JAM"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   1680
      Width           =   1935
   End
   Begin VB.TextBox TxtUserText 
      Height          =   975
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   10455
   End
   Begin VB.CommandButton BtnMSCryptMD5 
      Caption         =   "MSCrypt.MD5"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label LblBCrypt3 
      AutoSize        =   -1  'True
      Caption         =   "BCrypt.SHA512:"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2160
      TabIndex        =   24
      Top             =   5040
      Width           =   1470
   End
   Begin VB.Label LblBCrypt2 
      AutoSize        =   -1  'True
      Caption         =   "BCrypt.SHA384:"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2160
      TabIndex        =   21
      Top             =   4560
      Width           =   1470
   End
   Begin VB.Label LblCRC322 
      AutoSize        =   -1  'True
      Caption         =   "CRC-32/MEF:"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2160
      TabIndex        =   6
      Top             =   2160
      Width           =   1155
   End
   Begin VB.Label LblMSCrypt3 
      AutoSize        =   -1  'True
      Caption         =   "MSCrypt.SHA:"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2160
      TabIndex        =   15
      Top             =   3600
      Width           =   1260
   End
   Begin VB.Label LblMSCrypt1 
      AutoSize        =   -1  'True
      Caption         =   "MSCrypt.RC4:"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2160
      TabIndex        =   9
      Top             =   2640
      Width           =   1260
   End
   Begin VB.Label LblBCrypt1 
      AutoSize        =   -1  'True
      Caption         =   "BCrypt.SHA256:"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2160
      TabIndex        =   18
      Top             =   4080
      Width           =   1470
   End
   Begin VB.Label LBLDragDropFileCRC32 
      Alignment       =   2  'Zentriert
      Caption         =   "Drag'n'drop file here, checksum-hash will be calculated!"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      OLEDropMode     =   1  'Manuell
      TabIndex        =   0
      Top             =   0
      Width           =   10455
   End
   Begin VB.Label LblCRC321 
      AutoSize        =   -1  'True
      Caption         =   "CRC-32/JAM:"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2160
      TabIndex        =   3
      Top             =   1680
      Width           =   1155
   End
   Begin VB.Label LblMSCrypt2 
      AutoSize        =   -1  'True
      Caption         =   "MSCrypt.MD5:"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2160
      TabIndex        =   12
      Top             =   3120
      Width           =   1260
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'CRC = Cyclic Redundancy Check
'https://de.wikipedia.org/wiki/Zyklische_Redundanzpr%C3%BCfung
'https://en.wikipedia.org/wiki/Cyclic_redundancy_check

'MD5 = Message Digest Algorithm 5
'https://en.wikipedia.org/wiki/MD5
'https://de.wikipedia.org/wiki/Message-Digest_Algorithm_5

'RC4 = Rivest Cipher 4
'https://en.wikipedia.org/wiki/RC4
'https://de.wikipedia.org/wiki/RC4

'SHA = Secure Hash Algorithm

'"The Quick Brown Fox Jumps Over The Lazy Dog"
'CRC32/JAMCRC  = &HF89F9449
'CRC32/MEF     = &HA26BF73F
'MSCrypt.RC4   = &H3E9B65EFC2EADA94C501D9DF83719322
'MSCrypt.MD5   = &H58826469C2606F4791B9F75880DFBE2A
'MSCrypt.SHA   = &H645218467886DD414EA66A09B6CCEEA806127FB5
'BCrypt.SHA256 = &HC6E68384699D2E81C02D4C3EEC53CEDE3EA420C1AE8A227DAC495AA00666FD13
'BCrypt.SHA384 = &H29713F65A24E97E66DA57499723359374326DD1498C9A26FDA84396A7A7D0A24C56A50343F5E0228778EA7BD53F9A179
'BCrypt.SHA512 = &H12a98085e307959d5d6e6d0ed361845b604a33f9b66d025f30cc0414d2fa374ea129e6e80a838dffc07e2334e9936119d5bb18443d3ecde58a2f1ec4306e6fb2


Private Sub Form_Load()
    TxtUserText.text = "The Quick Brown Fox Jumps Over The Lazy Dog"
    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Form_Resize()
    Dim l As Single, t As Single, W As Single, H As Single
    t = LBLDragDropFileCRC32.Top
    W = Me.ScaleWidth
    H = LBLDragDropFileCRC32.Height
    If W > 0 And H > 0 Then LBLDragDropFileCRC32.Move l, t, W, H
    t = TxtUserText.Top
    H = TxtUserText.Height
    If W > 0 And H > 0 Then TxtUserText.Move l, t, W, H
    l = TxtCRC32JAM.Left
    W = Me.ScaleWidth - l
    H = TxtCRC32JAM.Height
    t = TxtCRC32JAM.Top
    If W > 0 And H > 0 Then TxtCRC32JAM.Move l, t, W, H
    t = TxtCRC32MEF.Top
    If W > 0 And H > 0 Then TxtCRC32MEF.Move l, t, W, H
    t = TxtMSCryptRC4.Top
    If W > 0 And H > 0 Then TxtMSCryptRC4.Move l, t, W, H
    t = TxtMSCryptMD5.Top
    If W > 0 And H > 0 Then TxtMSCryptMD5.Move l, t, W, H
    t = TxtMSCryptSHA.Top
    If W > 0 And H > 0 Then TxtMSCryptSHA.Move l, t, W, H
    t = TxtBCryptSHA256.Top
    If W > 0 And H > 0 Then TxtBCryptSHA256.Move l, t, W, H
    t = TxtBCryptSHA384.Top
    If W > 0 And H > 0 Then TxtBCryptSHA384.Move l, t, W, H
    t = TxtBCryptSHA512.Top
    H = Me.ScaleHeight - t
    If W > 0 And H > 0 Then TxtBCryptSHA512.Move l, t, W, H
End Sub

Private Sub LBLDragDropFileCRC32_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    OnOLEDragDrop Data, Effect, Button, Shift, x, y
End Sub
Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    OnOLEDragDrop Data, Effect, Button, Shift, x, y
End Sub

Private Sub OnOLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not Data.GetFormat(ClipBoardConstants.vbCFFiles) Then Exit Sub
    Dim PFN As String: PFN = Data.Files(1)
    Dim Value() As Byte: Value = GetValueFromFile(PFN)
    TxtCRC32JAM.text = GetHashString(MNew.CRC32(ECRC32Algo.CRC32_JAMCRC), Value)
    TxtCRC32MEF.text = GetHashString(MNew.CRC32(ECRC32Algo.CRC32_MEF), Value)
    TxtMSCryptRC4.text = GetHashString(MNew.MSCrypt(ECryptHashAlgo.ha_RC4), Value)
    TxtMSCryptMD5.text = GetHashString(MNew.MSCrypt(ECryptHashAlgo.ha_MD5), Value)
    TxtMSCryptSHA.text = GetHashString(MNew.MSCrypt(ECryptHashAlgo.ha_SHA), Value)
    TxtBCryptSHA256.text = GetHashString(MNew.MSBCrypt(EBCryptHashAlgo.ha_SHA256), Value)
    TxtBCryptSHA384.text = GetHashString(MNew.MSBCrypt(EBCryptHashAlgo.ha_SHA384), Value)
    TxtBCryptSHA512.text = GetHashString(MNew.MSBCrypt(EBCryptHashAlgo.ha_SHA512), Value)
End Sub

Function GetValueFromFile(PFN As String) As Byte()
Try: On Error GoTo Catch
    Dim FNr As Integer: FNr = FreeFile
    Open PFN For Binary Access Read As FNr
    ReDim FileContent(0 To LOF(FNr) - 1) As Byte
    Get FNr, , FileContent
    GetValueFromFile = FileContent
    GoTo Finally
Catch:
    MsgBox "Error in FMain.GetFileContent"
Finally: Close FNr
End Function

Function GetValueFromUserInput() As Byte()
    Dim s As String: s = TxtUserText.text
    If Len(s) = 0 Then
        MsgBox "Please give a valid string in edittextbox"
        Exit Function
    End If
    GetValueFromUserInput = StrConv(s, vbFromUnicode)
End Function

Private Function GetHashString(hasher As IHasher, Value() As Byte) As String
Try: On Error GoTo Catch
    If MPtr.Array_Count(Value) = 0 Then
        MsgBox "Empty value!"
        Exit Function
    End If
    Dim hash() As Byte: hash = hasher.GetHash(Value)
    GetHashString = Hex_ToStr(hash)
    Exit Function
Catch:
    MsgBox "Error in FMain.GetHash maybe text is empty"
End Function

Private Sub BtnCRC32JAM_Click()
    TxtCRC32JAM.text = GetHashString(MNew.CRC32(ECRC32Algo.CRC32_JAMCRC), GetValueFromUserInput)
End Sub

Private Sub BtnCRC32MEF_Click()
    TxtCRC32MEF.text = GetHashString(MNew.CRC32(ECRC32Algo.CRC32_MEF), GetValueFromUserInput)
End Sub

Private Sub BtnMSCryptRC4_Click()
    TxtMSCryptRC4.text = GetHashString(MNew.MSCrypt(ECryptHashAlgo.ha_RC4), GetValueFromUserInput)
End Sub

Private Sub BtnMSCryptMD5_Click()
    TxtMSCryptMD5.text = GetHashString(MNew.MSCrypt(ECryptHashAlgo.ha_MD5), GetValueFromUserInput)
End Sub

Private Sub BtnMSCryptSHA_Click()
    TxtMSCryptSHA.text = GetHashString(MNew.MSCrypt(ECryptHashAlgo.ha_SHA), GetValueFromUserInput)
End Sub

Private Sub BtnBCryptSHA256_Click()
    TxtBCryptSHA256.text = GetHashString(MNew.MSBCrypt(EBCryptHashAlgo.ha_SHA256), GetValueFromUserInput)
End Sub

Private Sub BtnBCryptSHA384_Click()
    TxtBCryptSHA256.text = GetHashString(MNew.MSBCrypt(EBCryptHashAlgo.ha_SHA384), GetValueFromUserInput)
End Sub

Private Sub BtnBCryptSHA512_Click()
    TxtBCryptSHA256.text = GetHashString(MNew.MSBCrypt(EBCryptHashAlgo.ha_SHA512), GetValueFromUserInput)
End Sub
