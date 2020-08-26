VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Cetak_Inventaris 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9150
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13950
   Icon            =   "Cetak_Inventaris.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9150
   ScaleWidth      =   13950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton OptKategori 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2520
      TabIndex        =   15
      Top             =   2040
      Width           =   315
   End
   Begin VB.ComboBox ComboKategori 
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
      Left            =   2820
      TabIndex        =   14
      Text            =   "-- Silahkan Pilih Kategori --"
      Top             =   2040
      Width           =   3195
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
      ScaleWidth      =   13950
      TabIndex        =   6
      Top             =   0
      Width           =   13950
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "CETAK DATA INVENTARIS"
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
         TabIndex        =   8
         Top             =   60
         Width           =   4815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Preview atau Cetak Data Inventaris Yang Tersedia."
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
         TabIndex        =   7
         Top             =   300
         Width           =   3705
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
      TabIndex        =   5
      Top             =   615
      Width           =   5775
   End
   Begin VB.OptionButton OptTanggal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2520
      TabIndex        =   4
      Top             =   1320
      Width           =   315
   End
   Begin VB.OptionButton OptAll 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Cetak Semua Data Inventaris"
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
      Left            =   2520
      TabIndex        =   3
      Top             =   2520
      Width           =   2775
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00D9FBDB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   13950
      TabIndex        =   0
      Top             =   8550
      Width           =   13950
      Begin Project1.isButton ButtonKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   11220
         TabIndex        =   1
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
         TabIndex        =   2
         Top             =   60
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Icon            =   "Cetak_Inventaris.frx":000C
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
   Begin MSComCtl2.DTPicker Tanggal2 
      Height          =   315
      Left            =   4620
      TabIndex        =   11
      Top             =   1320
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
      Left            =   2820
      TabIndex        =   12
      Top             =   1320
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   94961665
      CurrentDate     =   42993
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
      Left            =   4380
      TabIndex        =   13
      Top             =   1380
      Width           =   135
   End
   Begin VB.Image Image1 
      Height          =   2010
      Left            =   360
      Picture         =   "Cetak_Inventaris.frx":03A6
      Stretch         =   -1  'True
      Top             =   960
      Width           =   2010
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
      Left            =   2520
      TabIndex        =   10
      Top             =   1020
      Width           =   2340
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cetak Berdasarkan Kategori"
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
      Left            =   2520
      TabIndex        =   9
      Top             =   1740
      Width           =   2385
   End
End
Attribute VB_Name = "Cetak_Inventaris"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RSInventaris As New ADODB.Recordset
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
    Unload DataInventaris

    If OptTanggal.Value = True Then
        MyRS.Open "SELECT *" & _
                " FROM inventaris WHERE tanggal BETWEEN '" & Format(Tanggal1.Value, "yyyy-mm-dd") & "' and '" & Format(Tanggal2.Value, "yyyy-mm-dd") & "'", Conn, adOpenForwardOnly, adLockReadOnly
        strSQL = "SELECT *" & _
               " FROM inventaris WHERE tanggal BETWEEN '" & Format(Tanggal1.Value, "yyyy-mm-dd") & "' and '" & Format(Tanggal2.Value, "yyyy-mm-dd") & "'"
        strData = MyRS.RecordCount
    ElseIf OptAll.Value = True Then
        MyRS.Open "SELECT * FROM inventaris ORDER BY nama ASC", Conn, adOpenForwardOnly, adLockReadOnly
        strSQL = "SELECT * FROM inventaris ORDER BY nama ASC"
        strData = MyRS.RecordCount
    ElseIf OptKaryawan.Value = True Then
        MyRS.Open "SELECT *" & _
                " FROM inventaris WHERE (kategori='" & ComboKategori.Text & "') ORDER BY tanggal ASC", Conn, adOpenForwardOnly, adLockReadOnly
        strSQL = "SELECT *" & _
               " FROM inventaris WHERE (kategori='" & ComboKategori.Text & "') ORDER BY tanggal ASC"
        strData = MyRS.RecordCount
    Else
        MyRS.Open "SELECT * FROM inventaris ORDER BY nama ASC", Conn, adOpenForwardOnly, adLockReadOnly
        strSQL = "SELECT * FROM inventaris ORDER BY nama ASC"
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
    With DataInventaris
        .DC1.ConnectionString = "DRIVER={MySQL ODBC 5.1 Driver};SERVER=" & Servers & ";DATABASE=" & DBase & ";UID=" & Username & ";PWD=" & Password & ";PORT=" & Ports & ";OPTION=2+3+8+32+2048+16384+4194304;"
        .DC1.Source = strSQL

        .WindowState = 2
        .LabelTanggal.Caption = "Tanggal Cetak: " & Format(Date, "dd mmmm yyyy")
        If OptTanggal.Value = True Then
            .LabelBerdasar.Caption = "Berdasarkan Tanggal: " & Format(Tanggal1.Value, "dd mmmm yyyy") & " ~ " & Format(Tanggal2.Value, "dd mmmm yyyy")
        ElseIf OptKategori.Value = True Then
            .LabelBerdasar.Caption = "Berdasarkan Kategori: " & ComboKategori.Text
        ElseIf OptAll.Value = True Then
            .LabelBerdasar.Caption = "SEMUA DATA"
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
    Set RSInventaris = New ADODB.Recordset
    Dim Combo1 As String
    Combo1 = ComboListSingle(ComboKategori, RSInventaris, "jenis_inventaris", "*", "order by nama asc", "nama")
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
        ComboKategori.Enabled = False
        ButtonPreview.Enabled = True
    End If
End Sub

Private Sub OptKategori_Click()
    If OptKategori.Value = True Then
        Tanggal1.Enabled = False
        Tanggal2.Enabled = False
        ComboKategori.Enabled = True
        ComboKategori.SetFocus
        ButtonPreview.Enabled = True
    ElseIf OptKategori.Value = False Then
        Tanggal1.Enabled = False
        Tanggal2.Enabled = False
        ComboKategori.Enabled = False
    End If
End Sub

Private Sub OptTanggal_Click()
    If OptTanggal.Value = True Then
        Tanggal1.Enabled = True
        Tanggal2.Enabled = True
        Tanggal1.SetFocus
        ComboKategori.Enabled = False
        ButtonPreview.Enabled = True
    ElseIf OptTanggal.Value = False Then
        Tanggal1.Enabled = False
        Tanggal2.Enabled = False
    End If
End Sub
