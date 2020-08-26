Attribute VB_Name = "ModuleAplikasi"
Option Explicit

Private Const OF_EXIST         As Long = &H4000
Private Const OFS_MAXPATHNAME  As Long = 128
Private Const HFILE_ERROR      As Long = -1

Private Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type

Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, _
                        lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long

Public blnTambah            As Boolean
Public Sukses               As Boolean

'Fungsi Untuk Check File Exist Dengan Berbagai Atribut
Public Function FileExists(ByVal Fname As String) As Boolean
 
    Dim lRetVal As Long
    Dim OfSt As OFSTRUCT
    
    lRetVal = OpenFile(Fname, OfSt, OF_EXIST)
    If lRetVal <> HFILE_ERROR Then
        FileExists = True
    Else
        FileExists = False
    End If
    
End Function

'***********************************************
'fungsi untuk mengganti karakter ' => `
'***********************************************
Public Function GantiPetik(KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 96
    End If
End Function

Public Function GantiTF(Value As String) As String
    If Value = "True" Then
        GantiTF = "1"
    ElseIf Value = "False" Then
        GantiTF = "0"
    End If
End Function

Public Function Ganti01(Value As String) As String
    If Value = "1" Then
        Ganti01 = "True"
    ElseIf Value = "0" Then
        Ganti01 = "False"
    End If
End Function

