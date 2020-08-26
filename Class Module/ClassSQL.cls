VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
Dim Baris As Integer

Dim strSQL As String

Public Function FillGrid(objectToIgnore As VSFlexGrid, Rs As ADODB.Recordset, Table As String, field As String, where As String, contents As String, Conn As ADODB.Connection)
On Error GoTo ErrFillGrid
    Dim RsAS As ADODB.Recordset
    Set RsAS = New ADODB.Recordset
    'RsAS.Close
    SQL = "select " & field & " from " & Table & " " & where
    'Set Rs = Conn.Execute(strSQL)
    Set RsAS = Conn.Execute(SQL) '.Open SQL, Conn, adOpenDynamic, adLockPessimistic
        With RsAS
            .MoveFirst
            Do Until .EOF
                Baris = Baris + 1
                objectToIgnore.Rows = Baris
                objectToIgnore.AddItem Baris & vbTab & RsAS!Keterangan & vbTab & RsAS!autokode_karyawan
                .MoveNext
            Loop
            'MsgBox FillGrid
's            Call BarisLunasWarna
        End With
    RsAS.Close
Exit Function
ErrFillGrid:
    MsgBox "Maaf data tidak bisa ditampilkan di flexgrid." + err.Description, vbCritical
End Function
