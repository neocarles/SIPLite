VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} SubSlipPotongan 
   Caption         =   "Sub Slip Potongan"
   ClientHeight    =   5100
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12150
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   21431
   _ExtentY        =   8996
   SectionData     =   "SubSlipPotongan.dsx":0000
End
Attribute VB_Name = "SubSlipPotongan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Detail_Format()
    With DC.Recordset
        If Not .EOF Then
            'txtNo.Text = Val(txtNo.Text) + 1
            
            Keterangan.Text = .Fields("keterangan").Value '"*** " & .Fields("nama_barang").Value
            Nominal.Text = Format(.Fields("jumlah").Value, "#,##0")
            'Qty.Text = .Fields("qty").Value
            'HargaBeli.Text = Format(.Fields("harga_beli").Value, "#,##0")
            'Jumlah.Text = Format(.Fields("subtotal").Value, "#,##0") & " ***"
        End If
    End With
End Sub
