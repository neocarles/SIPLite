VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim strSQL As String

Dim RSHakAkses As New ADODB.Recordset

Private Function AutoNumber() As String
    Dim RSAuto As New ADODB.Recordset
    Set RSAuto = New ADODB.Recordset
    RSAuto.Open "SELECT id FROM pos_akses WHERE id in(select max(id) from pos_akses)order by id desc", Conn
    RSAuto.Requery
    Dim Urut As String * 5
    Dim Hitung As Long
    With RSAuto
        If .EOF Then
            Urut = "00001"
            AutoNumber = Urut
        Else
            Hitung = Right(!Id, 5) + 1
            Urut = Right("00000" & Hitung, 5)
        End If
        AutoNumber = Urut
    End With
End Function

Public Function GetTeksHasilCipher(strPlainTeks) As String
    On Error GoTo errHandle
    Dim ClassCipher As ClassCipher

    Set ClassCipher = New ClassCipher
    With ClassCipher
        .KeyString = "SIPLITE2017"
        .Text = strPlainTeks
        .DoXor
        .Stretch
        GetTeksHasilCipher = .Text
    End With
    Set ClassCipher = Nothing

    Exit Function
errHandle:
    MsgBox err.Description & vbCrLf & vbCrLf & strSQL
    GetTeksHasilCipher = ""
End Function

Public Function OldPassword(strUser As String, strPassLama As String, Conn As ADODB.Connection) As Boolean
    Dim Rs As ADODB.Recordset

    strSQL = "SELECT pword FROM pos_akses WHERE uname='" & strUser & "' AND " & _
             "pword = '" & strPassLama & "'"
    Set Rs = Conn.Execute(strSQL)
    If Rs.EOF Then
        OldPassword = False
    Else
        OldPassword = True
    End If
    Rs.Close
    Set Rs = Nothing
End Function

Public Function UpdatePassword(strUser As String, strPassBaru As String, Conn As ADODB.Connection) As Boolean
    On Error GoTo errHandle

    strSQL = "UPDATE pos_akses SET pword = '" & strPassBaru & "' WHERE uname = '" & strUser & "'"

    Conn.Execute strSQL

    UpdatePassword = True
    Exit Function
errHandle:
    MsgBox err.Description & vbCrLf & strSQL
    UpdatePassword = False
End Function

Public Function AddUser(strUser As String, strNama As String, strJabatan As String, strLevel As String, strGender As String, Conn As ADODB.Connection) As Boolean
    On Error GoTo errHandle

    strSQL = "INSERT INTO pos_akses (id,uname,pword,namauser,jabatan,kode,registered,gender) VALUES ('" & AutoNumber & "','" & _
             strUser & "', '', '" & strNama & "', '" & strJabatan & "', '" & strLevel & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & strGender & "')"
    Conn.Execute strSQL

    AddUser = True
    Exit Function
errHandle:
    MsgBox err.Description & vbCrLf & vbCrLf & strSQL
    AddUser = False
End Function

Public Function UpdateUser(strUser As String, strNama As String, strJabatan As String, strLevel As String, strGender As String, Conn As ADODB.Connection) As Boolean
    On Error GoTo errHandle

    strSQL = "UPDATE pos_akses SET namauser = '" & strNama & "', jabatan = '" & strJabatan & "', kode = '" & _
             strLevel & "', gender = '" & strGender & "' WHERE uname = '" & strUser & "'"

    Conn.Execute strSQL

    UpdateUser = True
    Exit Function
errHandle:
    MsgBox err.Description & vbCrLf & strSQL
    UpdateUser = False
End Function

Public Function DeleteUser(strUser As String, Conn As ADODB.Connection) As Boolean
    On Error GoTo errHandle

    strSQL = "DELETE FROM pos_akses WHERE uname = '" & strUser & "'"
    Conn.Execute strSQL

    DeleteUser = True
    Exit Function
errHandle:
    MsgBox err.Description & vbCrLf & vbCrLf & strSQL
    DeleteUser = False
End Function

Public Function HakAkses(KodeUser As String) As Boolean
    Set RSHakAkses = New ADODB.Recordset
    RSHakAkses.Open "SELECT * FROM pos_level WHERE kode='" & KodeUser & "'", Conn
    If Not RSHakAkses.EOF Then
        With FormMain
            .MenuNavigasi(0).Enabled = RSHakAkses!karyawan_view
            .MenuNavigasi(1).Enabled = RSHakAkses!potongan_view
            .MenuNavigasi(2).Enabled = RSHakAkses!tunjangan_view
            .MenuNavigasi(3).Enabled = RSHakAkses!karyawan_view
            .MenuNavigasi(4).Enabled = RSHakAkses!pengeluaran_view
            .MenuNavigasi(5).Enabled = RSHakAkses!pemasukan_view
            .MenuNavigasi(6).Enabled = RSHakAkses!kas_view
            .MenuNavigasi(7).Enabled = RSHakAkses!biaya_view
            .MenuNavigasi(8).Enabled = RSHakAkses!inventaris_view
            .MenuNavigasi(9).Enabled = RSHakAkses!laporan_view
        End With
        With FormKaryawan
            .mKaryawan(0).Enabled = RSHakAkses!karyawan_create
            .mKaryawan(1).Enabled = RSHakAkses!karyawan_update
            .mKaryawan(2).Enabled = RSHakAkses!karyawan_delete
        End With
        With FormPotongan
            .mPembelian(0).Enabled = RSHakAkses!potongan_create
        End With
        With FormTunjangan
            .mPembelian(0).Enabled = RSHakAkses!tunjangan_create
        End With
        With FormPenggajian
            .mPenggajian(0).Enabled = RSHakAkses!gaji_create
            .mPenggajian(1).Enabled = RSHakAkses!gaji_update
            .mPenggajian(2).Enabled = RSHakAkses!gaji_delete
            .mPenggajian(3).Enabled = RSHakAkses!gaji_generate
        End With
        With FormPengeluaran
            .mPengeluaran(0).Enabled = RSHakAkses!pengeluaran_create
            .mPengeluaran(1).Enabled = RSHakAkses!pengeluaran_update
            .mPengeluaran(2).Enabled = RSHakAkses!pengeluaran_delete
        End With
        With FormPemasukan
            .mPengeluaran(0).Enabled = RSHakAkses!pemasukan_create
            .mPengeluaran(1).Enabled = RSHakAkses!pemasukan_update
            .mPengeluaran(2).Enabled = RSHakAkses!pemasukan_delete
        End With
        With FormKas
            .mKas(0).Enabled = RSHakAkses!kas_create
            .mKas(1).Enabled = RSHakAkses!kas_update
            .mKas(2).Enabled = RSHakAkses!kas_delete
        End With
        With FormBiaya
            .mBiaya(0).Enabled = RSHakAkses!biaya_create
            .mBiaya(1).Enabled = RSHakAkses!biaya_update
            .mBiaya(2).Enabled = RSHakAkses!biaya_delete
        End With
        With FormInventaris
            .mPengeluaran(0).Enabled = RSHakAkses!inventaris_create
            .mPengeluaran(1).Enabled = RSHakAkses!inventaris_update
            .mPengeluaran(2).Enabled = RSHakAkses!inventaris_delete
        End With
        FormMain.LabelUser.Enabled = RSHakAkses!user_view
        With FormUser
            .mUser(0).Enabled = RSHakAkses!user_create
            .mUser(1).Enabled = RSHakAkses!user_update
            .mUser(2).Enabled = RSHakAkses!user_delete
        End With
        FormMain.LabelChangePW.Enabled = RSHakAkses!user_change
        FormMain.LabelLevel.Enabled = RSHakAkses!level_view
        With FormLevel
            .mLevel(0).Enabled = RSHakAkses!level_create
            .mLevel(1).Enabled = RSHakAkses!level_update
            .mLevel(2).Enabled = RSHakAkses!level_delete
        End With
    Else
        With FormMain
            .MenuNavigasi(0).Enabled = False
            .MenuNavigasi(1).Enabled = False
            .MenuNavigasi(2).Enabled = False
            .MenuNavigasi(3).Enabled = False
            .MenuNavigasi(4).Enabled = False
            .MenuNavigasi(5).Enabled = False
            .MenuNavigasi(6).Enabled = False
            .MenuNavigasi(7).Enabled = False
            .MenuNavigasi(8).Enabled = False
            .MenuNavigasi(9).Enabled = False
        End With
        With FormKaryawan
            .mKaryawan(0).Enabled = False
            .mKaryawan(1).Enabled = False
            .mKaryawan(2).Enabled = False
        End With
        With FormPotongan
            .mPembelian(0).Enabled = False
        End With
        With FormTunjangan
            .mPembelian(0).Enabled = False
        End With
        With FormPenggajian
            .mPenggajian(0).Enabled = False
            .mPenggajian(1).Enabled = False
            .mPenggajian(2).Enabled = False
            .mPenggajian(3).Enabled = False
        End With
        With FormPengeluaran
            .mPengeluaran(0).Enabled = False
            .mPengeluaran(1).Enabled = False
            .mPengeluaran(2).Enabled = False
        End With
        With FormPemasukan
            .mPengeluaran(0).Enabled = False
            .mPengeluaran(1).Enabled = False
            .mPengeluaran(2).Enabled = False
        End With
        With FormKas
            .mKas(0).Enabled = False
            .mKas(1).Enabled = False
            .mKas(2).Enabled = False
        End With
        With FormBiaya
            .mBiaya(0).Enabled = False
            .mBiaya(1).Enabled = False
            .mBiaya(2).Enabled = False
        End With
        With FormInventaris
            .mPengeluaran(0).Enabled = False
            .mPengeluaran(1).Enabled = False
            .mPengeluaran(2).Enabled = False
        End With
        FormMain.LabelUser.Enabled = False
        With FormUser
            .mUser(0).Enabled = False
            .mUser(1).Enabled = False
            .mUser(2).Enabled = False
        End With
        FormMain.LabelChangePW.Enabled = True
        FormMain.LabelLevel.Enabled = False
        With FormLevel
            .mLevel(0).Enabled = False
            .mLevel(1).Enabled = False
            .mLevel(2).Enabled = False
        End With
    End If
End Function

