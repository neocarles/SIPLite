VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassDepartemen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim strSQL As String

Public Function AddDepartemen(AutoKode As String, Keterangan As String, _
                         Conn As ADODB.Connection) As Boolean

    On Error GoTo errHandle

    strSQL = "INSERT INTO departemen (keterangan, autokode_karyawan) " & _
             "VALUES ('" & AutoKode & "','" & Keterangan & "')"

    Conn.Execute strSQL

    AddDepartemen = True
    Exit Function
errHandle:
    MsgBox err.Description & vbCrLf & vbCrLf & strSQL
    AddDepartemen = False
End Function

Public Function DeleteDepartemen(AutoKode As String, Conn As ADODB.Connection) As Boolean
    On Error GoTo errHandle

    strSQL = "DELETE FROM departemen WHERE AutoKode = '" & AutoKode & "'"

    Conn.Execute strSQL

    DeleteDepartemen = True
    Exit Function
errHandle:
    MsgBox err.Description & vbCrLf & vbCrLf & strSQL
    DeleteDepartemen = False
End Function

Public Function UpdateDepartemen(KeteranganBaru As String, KeteranganLama As String, AutoKode As String, _
                            Conn As ADODB.Connection) As Boolean

    On Error GoTo errHandle

    strSQL = "UPDATE departemen SET keterangan='" & KeteranganBaru & "', autokode_karyawan = '" & AutoKode & _
             "' WHERE keterangan = '" & CStr(KeteranganLama) & "'"

    Conn.Execute strSQL

    UpdateDepartemen = True
    Exit Function
errHandle:
    MsgBox err.Description & vbCrLf & strSQL
    UpdateDepartemen = False
End Function

Public Function CekDuplikatKodeDepartemen(AutoKode As String, Conn As ADODB.Connection) As String
    On Error GoTo errHandle
    Dim Rs As ADODB.Recordset

    strSQL = "SELECT kode FROM departemen WHERE autokode_karyawan = '" & AutoKode & "'"
    Set Rs = Conn.Execute(strSQL)
    If Not Rs.EOF Then
        CekDuplikatKodeDepartemen = Rs("autokode_karyawan")
    End If
    Exit Function
errHandle:
    MsgBox err.Description & vbCrLf & vbCrLf & strSQL
End Function


