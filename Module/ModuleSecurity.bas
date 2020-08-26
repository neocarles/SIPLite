Attribute VB_Name = "ModuleSecurity"
'Encryption function used in this program.
Public Function Encrypt(ByVal StrPword) As String
On Error Resume Next
Dim i, ct As Integer
Dim letter, enc, strHexvalappend, strHexval As String
enc = ""
For i = 1 To Len(StrPword)
    letter = Mid(StrPword, i, 1)
    enc = enc & Chr(Asc(letter) + i + 3)
Next

For ct = 1 To Len(enc)
    strHexvalappend = Hex(Asc(Mid(enc, ct, 1)))
    strHexval = strHexval & strHexvalappend
Next
Encrypt = StrReverse(strHexval)
End Function
'Decryption function used in this program.
Public Function Decrypt(ByVal strDecoded_Pword As String) As String
On Error Resume Next
Dim i, ct As Integer
Dim letter, dec, StrValappend, strVal As String
dec = ""
strDecoded_Pword = StrReverse(strDecoded_Pword)

For ct = 1 To Len(strDecoded_Pword) Step 2
    StrValappend = Chr(Val("&H" & (Mid(strDecoded_Pword, ct, 2))))
    strVal = strVal & StrValappend
Next
strDecoded_Pword = strVal

For i = 1 To Len(strDecoded_Pword)
    letter = Mid(strDecoded_Pword, i, 1)
    dec = dec & Chr(Asc(letter) - i - 3)
Next
Decrypt = dec
End Function

