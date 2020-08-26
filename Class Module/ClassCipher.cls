VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCipher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'clsCipher.CLS
Private mstrKey As String
Private mstrText As String

'~~~.KeyString
'A string (key) used in encryption and decryption
Public Property Let KeyString(strKey As String)
    mstrKey = strKey
    Initialize
End Property

'~~~.Text
'Write text to be encrypted or decrypted
Public Property Let Text(strText As String)
    mstrText = strText
End Property

'Read text that was encrypted or decrypted
Public Property Get Text() As String
    Text = mstrText
End Property

'~~~.DoXor
'Exclusive-or method to encrypt or decrypt
Public Sub DoXor()
    Dim lngC As Long
    Dim intB As Long
    Dim lngN As Long
    For lngN = 1 To Len(mstrText)
        lngC = Asc(Mid(mstrText, lngN, 1))
        intB = Int(Rnd * 256)
        Mid(mstrText, lngN, 1) = Chr(lngC Xor intB)
    Next lngN
End Sub

'~~~.Stretch
'Convert any string to a printable, displayable string
Public Sub Stretch()
    Dim lngC As Long
    Dim lngN As Long
    Dim lngJ As Long
    Dim lngK As Long
    Dim lngA As Long
    Dim strB As String
    lngA = Len(mstrText)
    strB = Space(lngA + (lngA + 2) \ 3)
    For lngN = 1 To lngA
        lngC = Asc(Mid(mstrText, lngN, 1))
        lngJ = lngJ + 1
        Mid(strB, lngJ, 1) = Chr((lngC And 63) + 59)
        Select Case lngN Mod 3
        Case 1
            lngK = lngK Or ((lngC \ 64) * 16)
        Case 2
            lngK = lngK Or ((lngC \ 64) * 4)
        Case 0
            lngK = lngK Or (lngC \ 64)
            lngJ = lngJ + 1
            Mid(strB, lngJ, 1) = Chr(lngK + 59)
            lngK = 0
        End Select
    Next lngN
    If lngA Mod 3 Then
        lngJ = lngJ + 1
        Mid(strB, lngJ, 1) = Chr(lngK + 59)
    End If
    mstrText = strB
End Sub

'~~~.Shrink
'Inverse of the Stretch method;
'result can contain any of the 256-byte values
Public Sub Shrink()
    Dim lngC As Long
    Dim lngD As Long
    Dim lngE As Long
    Dim lngA As Long
    Dim lngB As Long
    Dim lngN As Long
    Dim lngJ As Long
    Dim lngK As Long
    Dim strB As String
    lngA = Len(mstrText)
    lngB = lngA - 1 - (lngA - 1) \ 4
    strB = Space(lngB)
    For lngN = 1 To lngB
        lngJ = lngJ + 1
        lngC = Asc(Mid(mstrText, lngJ, 1)) - 59
        Select Case lngN Mod 3
        Case 1
            lngK = lngK + 4
            If lngK > lngA Then lngK = lngA
            lngE = Asc(Mid(mstrText, lngK, 1)) - 59
            lngD = ((lngE \ 16) And 3) * 64
        Case 2
            lngD = ((lngE \ 4) And 3) * 64
        Case 0
            lngD = (lngE And 3) * 64
            lngJ = lngJ + 1
        End Select
        Mid(strB, lngN, 1) = Chr(lngC Or lngD)
    Next lngN
    mstrText = strB
End Sub

'Initializes random numbers using the key string
Private Sub Initialize()
    Dim lngN As Long
    Randomize Rnd(-1)
    For lngN = 1 To Len(mstrKey)
        Randomize Rnd(-Rnd * Asc(Mid(mstrKey, lngN, 1)))
    Next lngN
End Sub
