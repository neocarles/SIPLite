VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassKas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim strSQL As String

Public Function AddKas(Kode As String, Nama As String, Saldo As String, _
                       Conn As ADODB.Connection) As Boolean

    On Error GoTo errHandle

    strSQL = "INSERT INTO kas (kode, nama, saldo) " & _
             "VALUES ('" & Kode & "','" & Nama & "', '" & Format(Saldo, "###0") & "')"

    Conn.Execute strSQL

    AddKas = True
    Exit Function
errHandle:
    MsgBox err.Description & vbCrLf & vbCrLf & strSQL
    AddKas = False
End Function

Public Function DeleteKas(KodeKas As String, Conn As ADODB.Connection) As Boolean
    On Error GoTo errHandle

    strSQL = "DELETE FROM kas WHERE kode = '" & KodeKas & "'"

    Conn.Execute strSQL

    DeleteKas = True
    Exit Function
errHandle:
    MsgBox err.Description & vbCrLf & vbCrLf & strSQL
    DeleteKas = False
End Function

Public Function UpdateKas(KodeBaru As String, KodeLama As String, Nama As String, Saldo As String, _
                          Conn As ADODB.Connection) As Boolean

    On Error GoTo errHandle

    strSQL = "UPDATE kas SET kode='" & KodeBaru & "', nama = '" & Nama & "', saldo = '" & Format(Saldo, "###0") & _
             "' WHERE kode = '" & CStr(KodeLama) & "'"

    Conn.Execute strSQL

    UpdateKas = True
    Exit Function
errHandle:
    MsgBox err.Description & vbCrLf & strSQL
    UpdateKas = False
End Function

Public Function UpdateJumlahKas(strKode As String, strJumlahBaru As String, Conn As ADODB.Connection, Optional Nama) As Boolean
    On Error GoTo errHandle

    If Not IsMissing(Nama) Then
        If Trim$(strJumlahBaru) = "" Then
            strSQL = "UPDATE kas SET saldo ='" & CStr(strJumlahBaru) & "' WHERE kode = '" & strKode & "'"
        Else
            strSQL = "UPDATE kas SET saldo ='" & CStr(strJumlahBaru) & "', nama ='" & CStr(Nama) & "' WHERE kode = '" & strKode & "'"
        End If
    Else
        strSQL = "UPDATE kas SET saldo = '" & strJumlahBaru & "' WHERE kode = '" & strKode & "'"
    End If

    Conn.Execute strSQL

    UpdateJumlahKas = True
    Exit Function
errHandle:
    MsgBox err.Description & vbCrLf & strSQL
    UpdateJumlahKas = False
End Function

Public Function UpdateJumlahKasHapus(strKode As String, strJumlahBaru As String, Conn As ADODB.Connection, Optional Nama) As Boolean
    On Error GoTo errHandle

    If Not IsMissing(Nama) Then
        If Trim$(strJumlahBaru) = "" Then
            strSQL = "UPDATE kas SET saldo=saldo + '" & Format(strJumlahBaru, "###0") & "' WHERE kode = '" & strKode & "'"
        Else
            strSQL = "UPDATE kas SET saldo=saldo + '" & Format(strJumlahBaru, "###0") & "', nama ='" & CStr(Nama) & "' WHERE kode = '" & strKode & "'"
        End If
    Else
        strSQL = "UPDATE kas SET saldo=saldo + '" & Format(strJumlahBaru, "###0") & "' WHERE kode = '" & strKode & "'"
    End If

    Conn.Execute strSQL

    UpdateJumlahKasHapus = True
    Exit Function
errHandle:
    MsgBox err.Description & vbCrLf & strSQL
    UpdateJumlahKasHapus = False
End Function

Public Function UpdateKasPemasukan(strKode As String, strJumlahBaru As String, Conn As ADODB.Connection, Optional Nama) As Boolean
    On Error GoTo errHandle

    If Not IsMissing(Nama) Then
        If Trim$(strJumlahBaru) = "" Then
            strSQL = "UPDATE kas SET saldo=saldo + '" & Format(strJumlahBaru, "###0") & "' WHERE kode = '" & strKode & "'"
        Else
            strSQL = "UPDATE kas SET saldo=saldo + '" & Format(strJumlahBaru, "###0") & "', nama ='" & CStr(Nama) & "' WHERE kode = '" & strKode & "'"
        End If
    Else
        strSQL = "UPDATE kas SET saldo=saldo + '" & Format(strJumlahBaru, "###0") & "' WHERE kode = '" & strKode & "'"
    End If

    Conn.Execute strSQL

    UpdateKasPemasukan = True
    Exit Function
errHandle:
    MsgBox err.Description & vbCrLf & strSQL
    UpdateKasPemasukan = False
End Function

Public Function GetKodeKasNama(Nama As String, Conn As ADODB.Connection) As String
    On Error GoTo errHandle
    Dim Rs As ADODB.Recordset

    strSQL = "SELECT kode FROM kas WHERE nama = '" & Nama & "'"
    Set Rs = Conn.Execute(strSQL)
    If Not Rs.EOF Then
        GetKodeKasNama = Rs("kode")
    End If
    Exit Function
errHandle:
    MsgBox err.Description & vbCrLf & vbCrLf & strSQL
End Function
