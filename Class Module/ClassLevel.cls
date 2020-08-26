VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassLevel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim strSQL As String

Private Function AutoNumber() As String
    Dim RSAuto As New ADODB.Recordset
    Set RSAuto = New ADODB.Recordset
    RSAuto.Open "SELECT id FROM pos_level WHERE id in(select max(id) from pos_level)order by id desc", Conn
    RSAuto.Requery
    Dim Urut As String * 3
    Dim Hitung As Long
    With RSAuto
        If .EOF Then
            Urut = "001"
            AutoNumber = Urut
        Else
            Hitung = Right(!Id, 3) + 1
            Urut = Right("000" & Hitung, 3)
        End If
        AutoNumber = Urut
    End With
End Function

Public Function AddLevel(strKode As String, strNama As String, _
                         Conn As ADODB.Connection) As Boolean

    On Error GoTo errHandle

    strSQL = "INSERT INTO pos_level (id,kode,nama,karyawan_view,karyawan_create,karyawan_update,karyawan_delete,potongan_view,potongan_create,potongan_update,potongan_delete,tunjangan_view,tunjangan_create,tunjangan_update,tunjangan_delete,gaji_view,gaji_create,gaji_update,gaji_delete,gaji_generate " & _
             ", pengeluaran_view,pengeluaran_create,pengeluaran_update,pengeluaran_delete,pemasukan_view,pemasukan_create,pemasukan_update,pemasukan_delete,kas_view,kas_create,kas_update,kas_delete,biaya_view,biaya_create,biaya_update,biaya_delete " & _
             ", inventaris_view,inventaris_create,inventaris_update,inventaris_delete,laporan_view,user_view,user_create,user_update,user_delete,user_change,level_view,level_create,level_update,level_delete) " & _
             "VALUES ('" & AutoNumber & "','" & strKode & "','" & strNama & "','True','True','False','False','False','False','False','False','False','False','False','False','False','False','False','False','False','True','True','False','False','False','False','False','False','False','False','False','False','False','False','False','False','True','True','False','False','False','False','False','False','False','True','False','False','False','False')"

    Conn.Execute strSQL

    AddLevel = True
    Exit Function
errHandle:
    MsgBox err.Description & vbCrLf & vbCrLf & strSQL
    AddLevel = False
End Function

Public Function DeleteLevel(KodeLevel As String, Conn As ADODB.Connection) As Boolean
    On Error GoTo errHandle

    strSQL = "DELETE FROM pos_level WHERE kode = '" & KodeLevel & "'"

    Conn.Execute strSQL

    DeleteLevel = True
    Exit Function
errHandle:
    MsgBox err.Description & vbCrLf & vbCrLf & strSQL
    DeleteLevel = False
End Function

Public Function UpdateLevel(strKode As String, strKodeBaru As String, strNama As String, KarView As String, KarC As String, KarU As String, KarD As String, PotView As String, PotC As String, PotU As String, PotD As String, _
                            TunView As String, TunC As String, TunU As String, TunD As String, GajiView As String, GajiC As String, GajiU As String, GajiD As String, GajiG As String, _
                            PengView As String, PengC As String, PengU As String, PengD As String, PemView As String, PemC As String, PemU As String, PemD As String, _
                            KasView As String, KasC As String, KasU As String, KasD As String, BiayaView As String, BiayaC As String, BiayaU As String, BiayaD As String, _
                            InvView As String, InvC As String, InvU As String, InvD As String, UserView As String, UserC As String, UserU As String, UserD As String, UserP As String, _
                            LevelView As String, LevelC As String, LevelU As String, LevelD As String, _
                            Conn As ADODB.Connection) As Boolean
    On Error GoTo errHandle

    strSQL = "UPDATE pos_level SET kode = '" & strKodeBaru & "', nama = '" & strNama & "',karyawan_view = '" & KarView & "', karyawan_create = '" & KarC & "', karyawan_update = '" & KarU & "', karyawan_delete = '" & KarD & "', potongan_view = '" & PotView & "', potongan_create = '" & PotC & "', potongan_update = '" & PotU & "', potongan_delete = '" & PotD & "', tunjangan_view = '" & _
             TunView & "', tunjangan_create = '" & TunC & "', tunjangan_update = '" & TunU & "', tunjangan_delete = '" & TunD & "', gaji_view = '" & GajiView & "', gaji_create = '" & GajiC & "', gaji_update = '" & GajiU & "', gaji_delete = '" & GajiD & "', gaji_generate = '" & _
             GajiG & "', pengeluaran_view = '" & PengView & "', pengeluaran_create = '" & PengC & "', pengeluaran_update = '" & PengU & "', pengeluaran_delete = '" & PengD & "', pemasukan_view = '" & PemView & "', pemasukan_create = '" & PemC & "', pemasukan_update = '" & PemU & "', pemasukan_delete = '" & _
             PemD & "', kas_view = '" & KasView & "', kas_create = '" & KasC & "', kas_update = '" & KasU & "', kas_delete = '" & KasD & "', biaya_view = '" & BiayaView & "', biaya_create = '" & BiayaC & "', biaya_update = '" & BiayaU & "', biaya_delete = '" & _
             BiayaD & "', inventaris_view = '" & InvView & "', inventaris_create = '" & InvC & "', inventaris_update = '" & InvU & "', inventaris_delete = '" & InvD & "', user_view = '" & UserView & "', user_create = '" & UserC & "', user_update = '" & UserU & "', user_delete = '" & _
             UserD & "', user_change = '" & UserP & "', level_view = '" & LevelView & "', level_create = '" & LevelC & "', level_update = '" & LevelU & "', level_delete = '" & LevelD & "' WHERE kode = '" & strKode & "'"

    Conn.Execute strSQL


    strSQL = "UPDATE pos_akses set kode = '" & strKodeBaru & "' WHERE kode = '" & strKode & "'"

    Conn.Execute strSQL

    UpdateLevel = True
    Exit Function
errHandle:
    MsgBox err.Description & vbCrLf & strSQL
    UpdateLevel = False
End Function
