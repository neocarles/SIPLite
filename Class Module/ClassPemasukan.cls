VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPemasukan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim strSQL As String

Public Function GetPemasukan(Conn As ADODB.Connection) As ADODB.Recordset
    On Error GoTo errHandle

    strSQL = "SELECT * FROM Pemasukan ORDER BY tanggal asc"

    Set GetPemasukan = Conn.Execute(strSQL)

    Exit Function
errHandle:
    MsgBox err.Description & vbCrLf & vbCrLf & strSQL
    Set GetPemasukan = Nothing
End Function

Public Function AddPemasukan(Kode As String, Tanggal As String, Keterangan As String, _
                             Kode_Kas As String, Jumlah As String, Operator As String, _
                             Conn As ADODB.Connection) As Boolean

    On Error GoTo errHandle

    strSQL = "INSERT INTO Pemasukan (kode, tanggal, keterangan, kode_kas, " & _
             "jumlah, operator) " & _
             "VALUES ('" & Kode & "','" & Format(Tanggal, "yyyy-mm-dd") & "', '" & Keterangan & "', '" & _
             Kode_Kas & "', '" & Format(Jumlah, "###0") & "', '" & Operator & "')"

    Conn.Execute strSQL

    AddPemasukan = True
    Exit Function
errHandle:
    MsgBox err.Description & vbCrLf & vbCrLf & strSQL
    AddPemasukan = False
End Function

Public Function DeletePemasukan(KodePemasukan As String, Conn As ADODB.Connection) As Boolean
    On Error GoTo errHandle

    strSQL = "DELETE FROM Pemasukan WHERE kode = '" & KodePemasukan & "'"

    Conn.Execute strSQL

    DeletePemasukan = True
    Exit Function
errHandle:
    MsgBox err.Description & vbCrLf & vbCrLf & strSQL
    DeletePemasukan = False
End Function

Public Function UpdatePemasukan(Kode As String, Tanggal As String, Keterangan As String, _
                                Kode_Kas As String, Jumlah As String, Operator As String, _
                                Conn As ADODB.Connection) As Boolean

    On Error GoTo errHandle

    strSQL = "UPDATE Pemasukan SET tanggal='" & Format(Tanggal, "yyyy-mm-dd") & "', keterangan = '" & Keterangan & _
             "', kode_kas='" & Kode_Kas & "', jumlah='" & Format(Jumlah, "###0") & "', operator='" & _
             Operator & "' Where SUBSTR(kode,3,10) = " & Mid(Kode, 3, 10) & ""

    Conn.Execute strSQL

    UpdatePemasukan = True
    Exit Function
errHandle:
    MsgBox err.Description & vbCrLf & strSQL
    UpdatePemasukan = False
End Function



