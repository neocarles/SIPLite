VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassBiaya"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim strSQL As String

Public Function AddBiaya(Kode As String, Nama As String, _
                         Conn As ADODB.Connection) As Boolean

    On Error GoTo errHandle

    strSQL = "INSERT INTO biaya (kode, nama) " & _
             "VALUES ('" & Kode & "','" & Nama & "')"

    Conn.Execute strSQL

    AddBiaya = True
    Exit Function
errHandle:
    MsgBox err.Description & vbCrLf & vbCrLf & strSQL
    AddBiaya = False
End Function

Public Function DeleteBiaya(KodeBiaya As String, Conn As ADODB.Connection) As Boolean
    On Error GoTo errHandle

    strSQL = "DELETE FROM biaya WHERE kode = '" & KodeBiaya & "'"

    Conn.Execute strSQL

    DeleteBiaya = True
    Exit Function
errHandle:
    MsgBox err.Description & vbCrLf & vbCrLf & strSQL
    DeleteBiaya = False
End Function

Public Function UpdateBiaya(KodeBaru As String, KodeLama As String, Nama As String, _
                            Conn As ADODB.Connection) As Boolean

    On Error GoTo errHandle

    strSQL = "UPDATE biaya SET kode='" & KodeBaru & "', nama = '" & Nama & _
             "' WHERE kode = '" & CStr(KodeLama) & "'"

    Conn.Execute strSQL

    UpdateBiaya = True
    Exit Function
errHandle:
    MsgBox err.Description & vbCrLf & strSQL
    UpdateBiaya = False
End Function

Public Function CekDuplikatKode(Kode As String, Conn As ADODB.Connection) As String
    On Error GoTo errHandle
    Dim Rs As ADODB.Recordset

    strSQL = "SELECT kode FROM biaya WHERE kode = '" & Kode & "'"
    Set Rs = Conn.Execute(strSQL)
    If Not Rs.EOF Then
        CekDuplikatKode = Rs("kode")
    End If
    Exit Function
errHandle:
    MsgBox err.Description & vbCrLf & vbCrLf & strSQL
End Function
