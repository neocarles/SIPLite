VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPenggajian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim strSQL As String

Private Function AutoNumber() As String
    Dim RSPenggajian As New ADODB.Recordset
    Set RSPenggajian = New ADODB.Recordset
    RSPenggajian.Open "SELECT id FROM karyawan_gaji WHERE id in(select max(id) from karyawan_gaji)order by id desc", Conn
    RSPenggajian.Requery
    Dim Urut As String * 12
    Dim Hitung As Long
    With RSPenggajian
        If .EOF Then
            Urut = "GK" + Format(Date, "yymmdd") + "0001"
        Else
            If Mid(!Id, 3, 6) <> Format(Date, "yymmdd") Then
                Urut = "GK" + Format(Date, "yymmdd") + "0001"
            Else
                Hitung = Right$(!Id, 4) + 1
                Urut = "GK" + Format(Date, "yymmdd") + Right("0000" & Hitung, 4)
            End If
        End If
        AutoNumber = Urut
    End With
End Function

Public Function AddPenggajian(Id As String, Kode As String, Nama As String, Basic_Hk As String, Basic_Gaji As String, Hk_Miss As String, Hk_Potongan As String, Hk_TotPotong As String, _
                              Tunjangan As String, Potongan As String, Total_Gaji As String, Tanggal As String, Gaji_Bulan As String, Operator As String, _
                              Conn As ADODB.Connection) As Boolean

    On Error GoTo errHandle

    strSQL = "INSERT INTO karyawan_gaji (id, kode, nama, basic_hk, basic_gaji, hk_miss, hk_potongan " & _
             ", hk_totpotong, tunjangan, potongan, total_gaji," & _
             "tanggal, gaji_bulan, operator) " & _
             "VALUES ('" & Id & "','" & Kode & "','" & Nama & "','" & Basic_Hk & "','" & Format(Basic_Gaji, "###0") & "','" & Hk_Miss & "', '" & _
             Format(Hk_Potongan, "###0") & "', '" & Format(Hk_TotPotong, "###0") & "', '" & Format(Tunjangan, "###0") & "', '" & Format(Potongan, "###0") & "', '" & Format(Total_Gaji, "###0") & "', '" & _
             Format(Tanggal, "yyyy-mm-dd") & "', '" & Format(Tanggal, "mm") & "', '" & Operator & "')"

    Conn.Execute strSQL

    AddPenggajian = True
    Exit Function
errHandle:
    MsgBox err.Description & vbCrLf & vbCrLf & strSQL
    AddPenggajian = False
End Function

Public Function DeletePenggajian(KodePenggajian As String, Conn As ADODB.Connection) As Boolean
    On Error GoTo errHandle

    strSQL = "DELETE FROM karyawan_gaji WHERE id = '" & KodePenggajian & "'"

    Conn.Execute strSQL

    DeletePenggajian = True
    Exit Function
errHandle:
    MsgBox err.Description & vbCrLf & vbCrLf & strSQL
    DeletePenggajian = False
End Function

Public Function UpdatePenggajian(Id As String, Kode As String, Nama As String, Basic_Hk As String, Basic_Gaji As String, Hk_Miss As String, Hk_Potongan As String, Hk_TotPotong As String, _
                                 Tunjangan As String, Potongan As String, Total_Gaji As String, Tanggal As String, Gaji_Bulan As String, Operator As String, _
                                 Conn As ADODB.Connection) As Boolean

    On Error GoTo errHandle

    strSQL = "UPDATE karyawan_gaji SET kode='" & Kode & "', nama = '" & Nama & "', basic_hk = '" & Format(Basic_Hk, "###0") & "', basic_gaji = '" & Format(Basic_Gaji, "###0") & _
             "', hk_miss='" & Hk_Miss & "', hk_potongan='" & Format(Hk_Potongan, "###0") & "', tunjangan='" & _
             Format(Tunjangan, "###0") & "', potongan='" & Format(Potongan, "###0") & "', total_gaji='" & _
             Format(Total_Gaji, "###0") & "', tanggal='" & Format(Tanggal, "yyyy-mm-dd") & "', operator='" & _
             Operator & "' Where SUBSTR(id,3,10) = " & Mid(Id, 3, 10) & ""

    Conn.Execute strSQL

    UpdatePenggajian = True
    Exit Function
errHandle:
    MsgBox err.Description & vbCrLf & strSQL
    UpdatePenggajian = False
End Function

Public Function GeneratePenggajian(Hk_Miss As String, Hk_Potongan As String, Hk_TotPotong As String, _
                                   Tunjangan As String, Potongan As String, Total_Gaji As String, Tanggal As String, Gaji_Bulan As String, Operator As String, _
                                   Conn As ADODB.Connection) As Boolean
    On Error GoTo errHandle
    Dim Rs As ADODB.Recordset

    strSQL = "SELECT karyawan.nama as knama,departemen.keterangan as kdep,karyawan.id as kid,karyawan.status as kstatus,karyawan.Basic_Hk as kbasichk, karyawan.basic_gaji as kbasicgaji " & _
           " From karyawan INNER JOIN departemen_karyawan ON karyawan.id = departemen_karyawan.id" & _
           " INNER JOIN departemen ON departemen_karyawan.departemen_id = departemen.id"
    Set Rs = Conn.Execute(strSQL)
    If Not Rs.EOF Then
        Do While Not Rs.EOF
            strSQL = "INSERT INTO karyawan_gaji (id, kode, nama, basic_hk, basic_gaji, hk_miss, hk_potongan " & _
                     ", hk_totpotong, tunjangan, potongan, total_gaji," & _
                     "tanggal, gaji_bulan, operator) " & _
                     "VALUES ('" & AutoNumber & "','" & Rs("kid") & "','" & Rs("knama") & "','" & Rs("kbasichk") & "','" & Format(Rs("kbasicgaji"), "###0") & "','" & Hk_Miss & "', '" & _
                     Format(Hk_Potongan, "###0") & "', '" & Format(Hk_TotPotong, "###0") & "', '" & Format(Tunjangan, "###0") & "', '" & Format(Potongan, "###0") & "', '" & Format(Total_Gaji, "###0") & "', '" & _
                     Format(Tanggal, "yyyy-mm-dd") & "', '" & Format(Tanggal, "mm") & "', '" & Operator & "')"
            Conn.Execute strSQL
            Rs.MoveNext
        Loop
    End If

    Rs.Close
    Set Rs = Nothing

    GeneratePenggajian = True
    Exit Function
errHandle:
    MsgBox err.Description & vbCrLf & vbCrLf & strSQL
    GeneratePenggajian = False
End Function
