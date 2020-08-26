VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} CetakSlipGaji 
   Caption         =   "Cetak Slip Gaji"
   ClientHeight    =   7935
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12765
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   22516
   _ExtentY        =   13996
   SectionData     =   "CetakSlipGaji.dsx":0000
End
Attribute VB_Name = "CetakSlipGaji"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Detail_Format()
On Error Resume Next
'TextNomor.Text = Val(TextNomor.Text) + 1
'Dim RSSupplier As New ADODB.Recordset
'RSSupplier.Open "SELECT kode,nama_barang,satuan,qty,harga_beli,subtotal FROM itempembelian", Conn, adOpenDynamic, adLockOptimistic
    With DC1.Recordset
        If Not .EOF Then
            EmpID.Text = .Fields("gkode").Value
            Dep.Text = .Fields("kdep").Value
            HK.Text = .Fields("ghk").Value
            Names.Text = .Fields("gnama").Value
            PayID.Text = .Fields("gid").Value
            DD.Text = Format(.Fields("gtggl").Value, "YYYY-MM")
            'LabelDD.Caption = .Fields("gtggl").Value
            TotalGaji.Text = Format(.Fields("gtotgaji").Value, "#,##0")
            TTun.Text = Format(.Fields("gtun").Value, "#,##0")
            TPot.Text = Format(.Fields("gpot").Value, "#,##0")
            
            LabelDD.Caption = Mid(DD.Text, 1, 4) & " - " & Mid(DD.Text, 6, 7)
                        
            strSQL1 = "SELECT * FROM potongan WHERE id='" & EmpID.Text & "' and MONTH(tanggal) = " & Mid(DD.Text, 6, 7) & " AND YEAR(tanggal) = " & Mid(DD.Text, 1, 4)
            'strSQL1 = "SELECT * FROM potongan"
            'strSQL2 = "SELECT * FROM tunjangan"
            strSQL2 = "SELECT * FROM tunjangan WHERE id='" & EmpID.Text & "' and MONTH(tanggal) = " & Mid(DD.Text, 6, 7) & " AND YEAR(tanggal) = " & Mid(DD.Text, 1, 4)
            'Set RSSupplier = Conn.Ex ecute(StrSQL)
            'RSSupplier.Requery
            'Memanggil Sub Report
            Set Potongan.object = New SubSlipPotongan
            With Potongan.object.DC
                .ConnectionString = "DRIVER={MySQL ODBC 5.1 Driver};SERVER=" & Servers & ";DATABASE=" & DBase & ";UID=" & Username & ";PWD=" & Password & ";PORT=" & Ports & ";OPTION=2+3+8+32+2048+16384+4194304;"
                .Source = strSQL1
            End With
            
            Set Tunjangan.object = New SubSlipTunjangan
            With Tunjangan.object.DC
                .ConnectionString = "DRIVER={MySQL ODBC 5.1 Driver};SERVER=" & Servers & ";DATABASE=" & DBase & ";UID=" & Username & ";PWD=" & Password & ";PORT=" & Ports & ";OPTION=2+3+8+32+2048+16384+4194304;"
                .Source = strSQL2
            End With
                
        End If
    End With
End Sub
