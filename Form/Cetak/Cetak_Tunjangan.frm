VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Cetak_Tunjangan 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Cetak Data Tunjangan"
   ClientHeight    =   8880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12930
   Icon            =   "Cetak_Tunjangan.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8880
   ScaleWidth      =   12930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox ComboUrutkan 
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
      ItemData        =   "Cetak_Tunjangan.frx":000C
      Left            =   2880
      List            =   "Cetak_Tunjangan.frx":001C
      TabIndex        =   15
      Text            =   "-- Urutkan Berdasar --"
      Top             =   2700
      Width           =   3195
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00D9FBDB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   12765
      TabIndex        =   12
      Top             =   8100
      Width           =   12765
      Begin Project1.isButton ButtonKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   11220
         TabIndex        =   13
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
         TabIndex        =   14
         Top             =   60
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Icon            =   "Cetak_Tunjangan.frx":0054
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
   Begin VB.OptionButton OptAll 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Left            =   2580
      TabIndex        =   11
      Top             =   2700
      Width           =   255
   End
   Begin VB.ComboBox ComboKaryawan 
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
      Left            =   2880
      TabIndex        =   10
      Text            =   "-- Pilih Karyawan --"
      Top             =   1980
      Width           =   2655
   End
   Begin VB.OptionButton OptDate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2580
      TabIndex        =   9
      Top             =   1260
      Width           =   315
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
      TabIndex        =   8
      Top             =   615
      Width           =   5775
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00D9FBDB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   12930
      TabIndex        =   5
      Top             =   0
      Width           =   12930
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "CETAK DATA TUNJANGAN"
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
         TabIndex        =   7
         Top             =   60
         Width           =   4815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Preview atau Cetak Data Tunjangan Yang Tersedia."
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
         TabIndex        =   6
         Top             =   300
         Width           =   3735
      End
   End
   Begin VB.OptionButton OptKaryawan 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "OptKaryawan"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2580
      TabIndex        =   2
      Top             =   1980
      Width           =   255
   End
   Begin MSComCtl2.DTPicker Tanggal4 
      Height          =   315
      Left            =   7500
      TabIndex        =   0
      Top             =   1980
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   94961665
      CurrentDate     =   42993
   End
   Begin MSComCtl2.DTPicker Tanggal3 
      Height          =   315
      Left            =   5700
      TabIndex        =   1
      Top             =   1980
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   94961665
      CurrentDate     =   42993
   End
   Begin MSComCtl2.DTPicker Tanggal2 
      Height          =   315
      Left            =   4680
      TabIndex        =   3
      Top             =   1260
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   94961665
      CurrentDate     =   42993
   End
   Begin MSComCtl2.DTPicker Tanggal1 
      Height          =   315
      Left            =   2880
      TabIndex        =   4
      Top             =   1260
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   94961665
      CurrentDate     =   42993
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cetak Semua Data Tunjangan"
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
      Left            =   2580
      TabIndex        =   20
      Top             =   2400
      Width           =   2520
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cetak Berdasarkan Tanggal"
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
      Left            =   2580
      TabIndex        =   19
      Top             =   960
      Width           =   2340
   End
   Begin VB.Image Image1 
      Height          =   2130
      Left            =   300
      Picture         =   "Cetak_Tunjangan.frx":03EE
      Stretch         =   -1  'True
      Top             =   960
      Width           =   2130
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "~"
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
      Left            =   4440
      TabIndex        =   18
      Top             =   1320
      Width           =   135
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Cetak Berdasarkan Karyawan Pada Bulan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2580
      TabIndex        =   17
      Top             =   1680
      Width           =   3480
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "~"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7260
      TabIndex        =   16
      Top             =   2040
      Width           =   120
   End
End
Attribute VB_Name = "Cetak_Tunjangan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RSKaryawan As New ADODB.Recordset
Dim RSID As New ADODB.Recordset
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
    Unload DataKaryawan

    If OptDate.Value = True Then
        MyRS.Open "SELECT tunjangan.id,tunjangan.tanggal,tunjangan.keterangan,tunjangan.jumlah,tunjangan.kode,karyawan.nama AS knama" & _
                " FROM tunjangan INNER JOIN karyawan ON tunjangan.id = karyawan.id WHERE tanggal BETWEEN '" & Format(Tanggal1.Value, "yyyy-mm-dd") & "' and '" & Format(Tanggal2.Value, "yyyy-mm-dd") & "'", Conn, adOpenForwardOnly, adLockReadOnly
        strSQL = "SELECT tunjangan.id,tunjangan.tanggal,tunjangan.keterangan,tunjangan.jumlah,tunjangan.kode,karyawan.nama AS knama" & _
               " FROM tunjangan INNER JOIN karyawan ON tunjangan.id = karyawan.id WHERE tanggal BETWEEN '" & Format(Tanggal1.Value, "yyyy-mm-dd") & "' and '" & Format(Tanggal2.Value, "yyyy-mm-dd") & "'"
        strData = MyRS.RecordCount
    ElseIf OptAll.Value = True Then
        MyRS.Open "SELECT tunjangan.id,tunjangan.tanggal,tunjangan.keterangan,tunjangan.jumlah,tunjangan.kode,karyawan.nama AS knama" & _
                " FROM tunjangan INNER JOIN karyawan ON tunjangan.id = karyawan.id", Conn, adOpenForwardOnly, adLockReadOnly
        strSQL = "SELECT tunjangan.id,tunjangan.tanggal,tunjangan.keterangan,tunjangan.jumlah,tunjangan.kode,karyawan.nama AS knama" & _
               " FROM tunjangan INNER JOIN karyawan ON tunjangan.id = karyawan.id ORDER " + ComboUrutkan.Text
        strData = MyRS.RecordCount
    ElseIf OptKaryawan.Value = True Then
        MyRS.Open "SELECT tunjangan.id,tunjangan.tanggal,tunjangan.keterangan,tunjangan.jumlah,tunjangan.kode,karyawan.nama AS knama" & _
                " FROM tunjangan INNER JOIN karyawan ON tunjangan.id = karyawan.id WHERE (tanggal BETWEEN '" & Format(Tanggal3.Value, "yyyy-mm-dd") & "' and '" & Format(Tanggal4.Value, "yyyy-mm-dd") & "') and (karyawan.nama='" & ComboKaryawan.Text & "')", Conn, adOpenForwardOnly, adLockReadOnly
        strSQL = "SELECT tunjangan.id,tunjangan.tanggal,tunjangan.keterangan,tunjangan.jumlah,tunjangan.kode,karyawan.nama AS knama" & _
               " FROM tunjangan INNER JOIN karyawan ON tunjangan.id = karyawan.id WHERE (tanggal BETWEEN '" & Format(Tanggal3.Value, "yyyy-mm-dd") & "' and '" & Format(Tanggal4.Value, "yyyy-mm-dd") & "') and (karyawan.nama='" & ComboKaryawan.Text & "') ORDER BY tanggal ASC"
        strData = MyRS.RecordCount
    Else
        MyRS.Open "SELECT karyawan.nama as knama,departemen.keterangan as kdep,karyawan.id as kid" & _
                " From karyawan INNER JOIN departemen_karyawan ON karyawan.id = departemen_karyawan.id" & _
                " INNER JOIN departemen ON departemen_karyawan.departemen_id = departemen.id ORDER by nama ASC", Conn, adOpenForwardOnly, adLockReadOnly
        strSQL = "SELECT karyawan.nama as knama,departemen.keterangan as kdep,karyawan.id as kid" & _
               " From karyawan INNER JOIN departemen_karyawan ON karyawan.id = departemen_karyawan.id" & _
               " INNER JOIN departemen ON departemen_karyawan.departemen_id = departemen.id ORDER " + ComboUrutkan.Text
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
    With DataTunjangan
        .DC1.ConnectionString = "DRIVER={MySQL ODBC 5.1 Driver};SERVER=" & Servers & ";DATABASE=" & DBase & ";UID=" & Username & ";PWD=" & Password & ";PORT=" & Ports & ";OPTION=2+3+8+32+2048+16384+4194304;"
        .DC1.Source = strSQL

        .WindowState = 2
        .LabelTanggal.Caption = "Tanggal Cetak: " & Format(Date, "dd mmmm yyyy")
        If OptKaryawan.Value = True Then
            .lblJudul.Caption = "LAPORAN DATA TUNJANGAN"
            .LabelBerdasar.Caption = "Berdasarkan Karyawan: " & ComboKaryawan.Text
        ElseIf OptAll.Value = True Then
            .lblJudul.Caption = "LAPORAN DATA TUNJANGAN"
            .LabelBerdasar.Caption = "SEMUA DATA"
        ElseIf OptDate.Value = True Then
            .lblJudul.Caption = "LAPORAN DATA TUNJANGAN"
            .LabelBerdasar.Caption = "Berdasarkan Tanggal: " & Format(Tanggal1.Value, "dd mmmm yyyy") & " ~ " & Format(Tanggal2.Value, "dd mmmm yyyy")
        End If
        .lblTitle.Caption = StrTitle
        .LblAlamat.Caption = StrAlamat

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
    Call initCombo
End Sub

Private Sub initCombo()
    Set RSKaryawan = New ADODB.Recordset
    Dim Combo1 As String
    Combo1 = ComboListSingle(ComboKaryawan, RSKaryawan, "karyawan", "*", "order by nama asc", "nama")
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

Private Sub OptAll_Click()
    If OptAll.Value = True Then
        Tanggal1.Enabled = False
        Tanggal2.Enabled = False
        Tanggal3.Enabled = False
        Tanggal4.Enabled = False
        ComboKaryawan.Enabled = False
        ComboUrutkan.Enabled = True
        ComboUrutkan.SetFocus
        ButtonPreview.Enabled = True
    ElseIf OptAll.Value = False Then
    End If
End Sub

Private Sub OptDate_Click()
    If OptDate.Value = True Then
        Tanggal1.Enabled = True
        Tanggal2.Enabled = True
        Tanggal3.Enabled = False
        Tanggal4.Enabled = False
        Tanggal1.SetFocus
        ComboKaryawan.Enabled = False
        ComboUrutkan.Enabled = False
        ButtonPreview.Enabled = True
    ElseIf OptDate.Value = False Then
        Tanggal1.Enabled = False
        Tanggal2.Enabled = False
        Tanggal3.Enabled = False
        Tanggal4.Enabled = False
    End If
End Sub

Private Sub OptKaryawan_Click()
    If OptKaryawan.Value = True Then
        Tanggal1.Enabled = False
        Tanggal2.Enabled = False
        Tanggal3.Enabled = True
        Tanggal4.Enabled = True
        ComboKaryawan.Enabled = True
        ComboKaryawan.SetFocus
        ComboUrutkan.Enabled = False
        ButtonPreview.Enabled = True
    ElseIf OptKaryawan.Value = False Then
        Tanggal1.Enabled = False
        Tanggal2.Enabled = False
        Tanggal3.Enabled = False
        Tanggal4.Enabled = False
    End If
End Sub


