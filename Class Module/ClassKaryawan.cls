VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassKaryawan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim strSQL As String
Dim strSQL1 As String
Dim strSQL2 As String

Public Function AddKaryawan(IdKar As String, IdDep As String, Nama As String, Alamat As String, Kontak As String, Status As String, BasicHk As String, Gaji As String, Conn As ADODB.Connection) As Boolean
    On Error GoTo errHandle

    'Insert ke tabel karyawan
    strSQL = "INSERT INTO karyawan (id,nama,alamat,kontak,status,basic_hk,basic_gaji) " & _
             "VALUES ('" & IdKar & "','" & Nama & "','" & Alamat & "','" & Kontak & "','" & Status & "','" & BasicHk & "','" & Gaji & "')"
    Conn.Execute strSQL

    'insert ke tabel departemen_karyawan
    strSQL = "INSERT INTO departemen_karyawan (id,departemen_id) " & _
             "VALUES ('" & IdKar & "','" & IdDep & "')"
    Conn.Execute strSQL

    AddKaryawan = True

    Exit Function
errHandle:
    MsgBox err.Description & vbCrLf & vbCrLf & strSQL
    AddKaryawan = False
End Function

Public Function UpdateKaryawan(IdKar As String, IdDep As String, Nama As String, Alamat As String, Kontak As String, Status As String, BasicHk As String, Gaji As String, _
                               Conn As ADODB.Connection) As Boolean

    On Error GoTo errHandle

    strSQL = "UPDATE karyawan SET nama='" & Nama & "', alamat = '" & Alamat & "', kontak = '" & Kontak & "', status = '" & Status & "', basic_hk = '" & BasicHk & "', basic_gaji = '" & Gaji & _
             "' WHERE id = '" & CStr(IdKar) & "'"

    Conn.Execute strSQL

    UpdateKaryawan = True
    Exit Function
errHandle:
    MsgBox err.Description & vbCrLf & strSQL
    UpdateKaryawan = False
End Function

Public Function DeleteKaryawan(IdKar As String, Conn As ADODB.Connection) As Boolean
    On Error GoTo errHandle

    'delete data departemen_karyawan
    strSQL = "DELETE FROM departemen_karyawan WHERE id =  " & IdKar & ""
    Conn.Execute strSQL

    'delete data karyawan
    strSQL = "DELETE FROM karyawan WHERE id = " & IdKar & ""
    Conn.Execute strSQL

    DeleteKaryawan = True

    Exit Function
errHandle:
    MsgBox err.Description & vbCrLf & vbCrLf & strSQL
    DeleteKaryawan = False
End Function
