VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassInventaris"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim strSQL As String

Public Function GetInventaris(Conn As ADODB.Connection) As ADODB.Recordset
    On Error GoTo errHandle

    strSQL = "SELECT * FROM Inventaris ORDER BY tanggal asc"

    Set GetInventaris = Conn.Execute(strSQL)

    Exit Function
errHandle:
    MsgBox err.Description & vbCrLf & vbCrLf & strSQL
    Set GetInventaris = Nothing
End Function

Public Function AddInventaris(Id As String, Tanggal As String, Nama As String, Qty As String, _
                              Perolehan As String, Kategori As String, Operator As String, _
                              Conn As ADODB.Connection) As Boolean

    On Error GoTo errHandle

    strSQL = "INSERT INTO inventaris (id, tanggal, nama, qty, perolehan, " & _
             "kategori, operator) " & _
             "VALUES ('" & Id & "','" & Format(Tanggal, "yyyy-mm-dd") & "', '" & Nama & "', '" & Qty & "', '" & _
             Perolehan & "', '" & Kategori & "', '" & Operator & "')"

    Conn.Execute strSQL

    AddInventaris = True
    Exit Function
errHandle:
    MsgBox err.Description & vbCrLf & vbCrLf & strSQL
    AddInventaris = False
End Function

Public Function DeleteInventaris(KodeInventaris As String, Conn As ADODB.Connection) As Boolean
    On Error GoTo errHandle

    strSQL = "DELETE FROM Inventaris WHERE id = '" & KodeInventaris & "'"

    Conn.Execute strSQL

    DeleteInventaris = True
    Exit Function
errHandle:
    MsgBox err.Description & vbCrLf & vbCrLf & strSQL
    DeleteInventaris = False
End Function

Public Function UpdateInventaris(Kode As String, Tanggal As String, Nama As String, Qty As String, _
                                 Perolehan As String, Kategori As String, Operator As String, _
                                 Conn As ADODB.Connection) As Boolean

    On Error GoTo errHandle

    strSQL = "UPDATE Inventaris SET tanggal='" & Format(Tanggal, "yyyy-mm-dd") & "', nama = '" & Nama & "', qty = '" & Qty & _
             "', perolehan='" & Perolehan & "', kategori='" & Kategori & "', operator='" & _
             Operator & "' Where id = '" & Kode & "'"

    Conn.Execute strSQL

    UpdateInventaris = True
    Exit Function
errHandle:
    MsgBox err.Description & vbCrLf & strSQL
    UpdateInventaris = False
End Function



