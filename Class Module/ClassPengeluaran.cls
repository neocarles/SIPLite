VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPengeluaran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim strSQL As String

Public Function GetPengeluaran(Conn As ADODB.Connection) As ADODB.Recordset
    On Error GoTo errHandle

    strSQL = "SELECT * FROM pengeluaran ORDER BY tanggal asc"

    Set GetPengeluaran = Conn.Execute(strSQL)

    Exit Function
errHandle:
    MsgBox err.Description & vbCrLf & vbCrLf & strSQL
    Set GetPengeluaran = Nothing
End Function

Public Function AddPengeluaran(Kode As String, Tanggal As String, Kd_Biaya As String, Keterangan As String, _
                               Kode_Kas As String, Jumlah As String, Operator As String, _
                               Conn As ADODB.Connection) As Boolean

    On Error GoTo errHandle

    strSQL = "INSERT INTO pengeluaran (id, tanggal, kd_biaya, keterangan, kode_kas, " & _
             "jumlah, operator) " & _
             "VALUES ('" & Kode & "','" & Format(Tanggal, "yyyy-mm-dd") & "', '" & Kd_Biaya & "', '" & Keterangan & "', '" & _
             Kode_Kas & "', '" & Format(Jumlah, "###0") & "', '" & Operator & "')"

    Conn.Execute strSQL

    AddPengeluaran = True
    Exit Function
errHandle:
    MsgBox err.Description & vbCrLf & vbCrLf & strSQL
    AddPengeluaran = False
End Function

Public Function DeletePengeluaran(KodePengeluaran As String, Conn As ADODB.Connection) As Boolean
    On Error GoTo errHandle

    strSQL = "DELETE FROM pengeluaran WHERE id = '" & KodePengeluaran & "'"

    Conn.Execute strSQL

    DeletePengeluaran = True
    Exit Function
errHandle:
    MsgBox err.Description & vbCrLf & vbCrLf & strSQL
    DeletePengeluaran = False
End Function

Public Function UpdatePengeluaran(Kode As String, Tanggal As String, Kd_Biaya As String, Keterangan As String, _
                                  Kode_Kas As String, Jumlah As String, Operator As String, _
                                  Conn As ADODB.Connection) As Boolean

    On Error GoTo errHandle

    strSQL = "UPDATE pengeluaran SET tanggal='" & Format(Tanggal, "yyyy-mm-dd") & "', kd_biaya = '" & Kd_Biaya & "', keterangan = '" & Keterangan & _
             "', kode_kas='" & Kode_Kas & "', jumlah='" & Format(Jumlah, "###0") & "', operator='" & _
             Operator & "' Where SUBSTR(id,3,10) = " & Mid(Kode, 3, 10) & ""

    Conn.Execute strSQL

    UpdatePengeluaran = True
    Exit Function
errHandle:
    MsgBox err.Description & vbCrLf & strSQL
    UpdatePengeluaran = False
End Function

