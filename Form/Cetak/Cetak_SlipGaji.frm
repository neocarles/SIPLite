VERSION 5.00
Begin VB.Form Cetak_SlipGaji 
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Slip Gaji"
   ClientHeight    =   8760
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13620
   LinkTopic       =   "Form1"
   ScaleHeight     =   8760
   ScaleWidth      =   13620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00D9FBDB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   13620
      TabIndex        =   11
      Top             =   0
      Width           =   13620
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "CETAK SLIP GAJI"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   180
         TabIndex        =   13
         Top             =   60
         Width           =   4815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Preview atau Cetak Data Gaji Yang Tersedia."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   195
         TabIndex        =   12
         Top             =   300
         Width           =   3240
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   20
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   5775
      TabIndex        =   10
      Top             =   615
      Width           =   5775
   End
   Begin VB.OptionButton OptBulan 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   1860
      TabIndex        =   9
      Top             =   1260
      Width           =   315
   End
   Begin VB.OptionButton OptDept 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Cetak Data Gaji Berdasar Departemen"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   1860
      TabIndex        =   8
      Top             =   1980
      Width           =   3075
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00D9FBDB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   13620
      TabIndex        =   5
      Top             =   8130
      Width           =   13620
      Begin Project1.isButton ButtonKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   11220
         TabIndex        =   6
         Top             =   60
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   767
         Style           =   8
         Caption         =   "&Keluar [ESC]"
         iNonThemeStyle  =   0
         Object.ToolTipText     =   ""
         ToolTipTitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         RoundedBordersByTheme=   0   'False
      End
      Begin Project1.isButton ButtonPreview 
         Height          =   435
         Left            =   180
         TabIndex        =   7
         Top             =   60
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Icon            =   "Cetak_SlipGaji.frx":0000
         Style           =   8
         Caption         =   "&Preview"
         CaptionAlign    =   2
         iNonThemeStyle  =   0
         Enabled         =   0   'False
         Object.ToolTipText     =   ""
         ToolTipTitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         RoundedBordersByTheme=   0   'False
      End
   End
   Begin VB.ComboBox ComboTahun 
      Enabled         =   0   'False
      Height          =   315
      Left            =   2220
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1260
      Width           =   1575
   End
   Begin VB.ComboBox ComboBulan 
      Enabled         =   0   'False
      Height          =   315
      Left            =   4080
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1260
      Width           =   1155
   End
   Begin VB.ComboBox ComboDepartemen 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1860
      TabIndex        =   2
      Text            =   "-- Pilih Departemen --"
      Top             =   2400
      Width           =   2655
   End
   Begin VB.ComboBox ComboBln 
      Enabled         =   0   'False
      Height          =   315
      Left            =   6720
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2400
      Width           =   1155
   End
   Begin VB.ComboBox ComboThn 
      Enabled         =   0   'False
      Height          =   315
      Left            =   4860
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   1350
      Left            =   300
      Picture         =   "Cetak_SlipGaji.frx":039A
      Stretch         =   -1  'True
      Top             =   960
      Width           =   1350
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cetak Berdasarkan Bulan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1860
      TabIndex        =   15
      Top             =   960
      Width           =   2130
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cetak Berdasarkan Departemen"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1860
      TabIndex        =   14
      Top             =   1680
      Width           =   2730
   End
End
Attribute VB_Name = "Cetak_SlipGaji"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MyRS As New ADODB.Recordset
Dim strKode As String

Private Sub ButtonKeluar_Click()
    Unload Me
End Sub

Private Sub ButtonPreview_Click()
    On Error Resume Next
    Dim StrJenis As String
    Dim strSQL As String
    Dim StrTitle As String
    Dim StrAlamat As String
    Set MyRS = New ADODB.Recordset
    Unload CetakSlipGaji

    If OptBulan.Value = True Then
        MyRS.Open "SELECT departemen.keterangan as kdep,karyawan_gaji.id as gid,karyawan_gaji.kode as gkode,karyawan_gaji.nama as gnama,karyawan_gaji.basic_hk as ghk,karyawan_gaji.basic_gaji as ggaji,karyawan_gaji.hk_miss as gmiss," & _
                  " karyawan_gaji.hk_potongan as ghkpot, karyawan_gaji.hk_totpotong as gtotpot, karyawan_gaji.tunjangan as gtun, karyawan_gaji.potongan as gpot, karyawan_gaji.total_gaji as gtotgaji, karyawan_gaji.tanggal AS gtggl" & _
                  " From karyawan_gaji INNER JOIN departemen_karyawan ON karyawan_gaji.kode = departemen_karyawan.id" & _
                  " INNER JOIN departemen ON departemen_karyawan.departemen_id = departemen.id WHERE YEAR(tanggal)='" & ComboTahun.Text & "' and MONTH(tanggal) ='" & ComboBulan.Text & "' ORDER BY gkode", Conn, adOpenForwardOnly, adLockReadOnly
                  
        strSQL = "SELECT departemen.keterangan as kdep,karyawan_gaji.id as gid,karyawan_gaji.kode as gkode,karyawan_gaji.nama as gnama,karyawan_gaji.basic_hk as ghk,karyawan_gaji.basic_gaji as ggaji,karyawan_gaji.hk_miss as gmiss," & _
                  " karyawan_gaji.hk_potongan as ghkpot, karyawan_gaji.hk_totpotong as gtotpot, karyawan_gaji.tunjangan as gtun, karyawan_gaji.potongan as gpot, karyawan_gaji.total_gaji as gtotgaji, karyawan_gaji.tanggal AS gtggl" & _
                  " From karyawan_gaji INNER JOIN departemen_karyawan ON karyawan_gaji.kode = departemen_karyawan.id" & _
                  " INNER JOIN departemen ON departemen_karyawan.departemen_id = departemen.id WHERE YEAR(tanggal)='" & ComboTahun.Text & "' and MONTH(tanggal) ='" & ComboBulan.Text & "' ORDER BY gkode"
        strData = MyRS.RecordCount
    ElseIf OptDept.Value = True Then
        MyRS.Open "SELECT departemen.keterangan as kdep,karyawan_gaji.id as gid,karyawan_gaji.kode as gkode,karyawan_gaji.nama as gnama,karyawan_gaji.basic_hk as ghk,karyawan_gaji.basic_gaji as ggaji,karyawan_gaji.hk_miss as gmiss," & _
                  " karyawan_gaji.hk_potongan as ghkpot, karyawan_gaji.hk_totpotong as gtotpot, karyawan_gaji.tunjangan as gtun, karyawan_gaji.potongan as gpot, karyawan_gaji.total_gaji as gtotgaji, karyawan_gaji.tanggal AS gtggl" & _
                  " From karyawan_gaji INNER JOIN departemen_karyawan ON karyawan_gaji.kode = departemen_karyawan.id" & _
                  " INNER JOIN departemen ON departemen_karyawan.departemen_id = departemen.id WHERE YEAR(tanggal)='" & ComboThn.Text & "' and MONTH(tanggal) ='" & ComboBln.Text & "' and departemen.keterangan='" & ComboDepartemen.Text & "' ORDER BY gkode", Conn, adOpenForwardOnly, adLockReadOnly
                  
        strSQL = "SELECT departemen.keterangan as kdep,karyawan_gaji.id as gid,karyawan_gaji.kode as gkode,karyawan_gaji.nama as gnama,karyawan_gaji.basic_hk as ghk,karyawan_gaji.basic_gaji as ggaji,karyawan_gaji.hk_miss as gmiss," & _
                  " karyawan_gaji.hk_potongan as ghkpot, karyawan_gaji.hk_totpotong as gtotpot, karyawan_gaji.tunjangan as gtun, karyawan_gaji.potongan as gpot, karyawan_gaji.total_gaji as gtotgaji, karyawan_gaji.tanggal AS gtggl" & _
                  " From karyawan_gaji INNER JOIN departemen_karyawan ON karyawan_gaji.kode = departemen_karyawan.id" & _
                  " INNER JOIN departemen ON departemen_karyawan.departemen_id = departemen.id WHERE YEAR(tanggal)='" & ComboThn.Text & "' and MONTH(tanggal) ='" & ComboBln.Text & "' and departemen.keterangan='" & ComboDepartemen.Text & "' ORDER BY gkode"
        strData = MyRS.RecordCount
    Else
        MyRS.Open "SELECT departemen.keterangan as kdep,karyawan_gaji.id as gid,karyawan_gaji.kode as gkode,karyawan_gaji.nama as gnama,karyawan_gaji.basic_hk as ghk,karyawan_gaji.basic_gaji as ggaji,karyawan_gaji.hk_miss as gmiss," & _
                  " karyawan_gaji.hk_potongan as ghkpot, karyawan_gaji.hk_totpotong as gtotpot, karyawan_gaji.tunjangan as gtun, karyawan_gaji.potongan as gpot, karyawan_gaji.total_gaji as gtotgaji, karyawan_gaji.tanggal AS gtggl" & _
                  " From karyawan_gaji INNER JOIN departemen_karyawan ON karyawan_gaji.kode = departemen_karyawan.id" & _
                  " INNER JOIN departemen ON departemen_karyawan.departemen_id = departemen.id WHERE YEAR(tanggal)='" & Format(Now, "yyyy") & "' and MONTH(tanggal) ='" & Format(Now, "mm") & "' ORDER BY gkode", Conn, adOpenForwardOnly, adLockReadOnly
                  
        strSQL = "SELECT departemen.keterangan as kdep,karyawan_gaji.id as gid,karyawan_gaji.kode as gkode,karyawan_gaji.nama as gnama,karyawan_gaji.basic_hk as ghk,karyawan_gaji.basic_gaji as ggaji,karyawan_gaji.hk_miss as gmiss," & _
                  " karyawan_gaji.hk_potongan as ghkpot, karyawan_gaji.hk_totpotong as gtotpot, karyawan_gaji.tunjangan as gtun, karyawan_gaji.potongan as gpot, karyawan_gaji.total_gaji as gtotgaji, karyawan_gaji.tanggal AS gtggl" & _
                  " From karyawan_gaji INNER JOIN departemen_karyawan ON karyawan_gaji.kode = departemen_karyawan.id" & _
                  " INNER JOIN departemen ON departemen_karyawan.departemen_id = departemen.id WHERE YEAR(tanggal)='" & Format(Now, "yyyy") & "' and MONTH(tanggal) ='" & Format(Now, "mm") & "' ORDER BY gkode"
        strData = MyRS.RecordCount
    End If
    Set MyRS = New ADODB.Recordset
    strKode = 1
    With MyRS
        .Open "SELECT * FROM Pengaturan where Kode like '" & strKode & "%'", Conn, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            StrTitle = .Fields("Nama").Value
            StrAlamat = .Fields("Alamat").Value
            .MoveNext
        Loop
    End With
    With CetakSlipGaji
        .Refresh
        .DC1.CursorLocation = ddADOUseClient
        .DC1.CursorType = ddADOOpenStatic
        .DC1.LockType = 1
        .DC1.ConnectionString = "DRIVER={MySQL ODBC 5.1 Driver};SERVER=" & Servers & ";DATABASE=" & DBase & ";UID=" & Username & ";PWD=" & Password & ";PORT=" & Ports & ";OPTION=2+3+8+32+2048+16384+4194304;"
        .DC1.Source = strSQL

        .WindowState = 2
        '.LabelTanggal.Caption = "Tanggal Cetak: " & Format(Date, "dd mmmm yyyy")
        If OptBulan.Value = True Then
        '    .LabelBerdasar.Caption = "Berdasarkan Tanggal: " & Format(Tanggal1.Value, "dd mmmm yyyy") & " ~ " & Format(Tanggal2.Value, "dd mmmm yyyy")
        ElseIf OptDep.Value = True Then
        '    .LabelBerdasar.Caption = "Berdasarkan Kategori: " & ComboKategori.Text
        End If
        .lblTitle.Caption = StrTitle
        .LblAlamat.Caption = StrAlamat
        .PayDate.Caption = Format(Now, "DD-MM-YYYY")
        '.LabelDD.Caption = Format(Now, "YYYY-MM")
        
        
        .Show
        .WindowState = 2
    End With
End Sub

Private Sub Form_Load()
    Call SetIcon(Me.hWnd, "FORMICON", False)
    With Me
        .Top = 0
        .Height = Screen.Height
        .Left = 0
        .Width = Screen.Width
    End With
    
    For A = 1 To 12
        ComboBulan.AddItem A
    Next
    
    For A2 = Val(TahunK) To Val(TahunT)
        ComboTahun.AddItem A2
    Next
    
    For A3 = Val(TahunK) To Val(TahunT)
        ComboThn.AddItem A3
    Next
    
    For A4 = 1 To 12
        ComboBln.AddItem A4
    Next
    
    Call initCombo
End Sub

Private Sub initCombo()
    Dim RSDepartemen As New ADODB.Recordset
    Set RSDepartemen = New ADODB.Recordset
    Dim Combo1 As String
'    Combo1 = ComboListSingle(ComboDepartemen, RSDepartemen, "departemen", "*", "where id>'1' order by keterangan asc", "keterangan")
    Combo1 = ComboListSingle(ComboDepartemen, RSDepartemen, "departemen", "*", "ORDER BY keterangan ASC", "keterangan")
End Sub

Private Sub Form_Resize()
'On Error Resume Next
    Call AutoResize
End Sub

Public Sub AutoResize()

    Picture3.Move 0, FormMain.PictureBeranda.ScaleHeight - Picture3.ScaleHeight, FormMain.PictureBeranda.ScaleWidth

    Picture1.Move 0, Picture3.ScaleHeight + 15, Picture3.ScaleWidth

    ButtonPreview.Move 180, 90

    ButtonKeluar.Move Picture3.ScaleWidth - ButtonKeluar.Width - 180, 90

End Sub

Private Sub optBulan_Click()
    If OptBulan.Value = True Then
        ComboDepartemen.Enabled = False
        ComboTahun.Enabled = True
        ComboBulan.Enabled = True
        ComboThn.Enabled = False
        ComboBln.Enabled = False
        ButtonPreview.Enabled = True
    ElseIf OptBulan.Value = False Then
        ComboDepartemen.Enabled = True
    End If
End Sub

Private Sub OptDept_Click()
    If OptDept.Value = True Then
        ComboDepartemen.Enabled = True
        ComboDepartemen.SetFocus
        ComboTahun.Enabled = False
        ComboBulan.Enabled = False
        ComboThn.Enabled = True
        ComboBln.Enabled = True
        ButtonPreview.Enabled = True
    ElseIf OptDept.Value = False Then
        ComboDepartemen.Enabled = False
    End If
End Sub
