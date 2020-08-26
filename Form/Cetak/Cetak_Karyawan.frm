VERSION 5.00
Begin VB.Form Cetak_Karyawan 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Cetak Karyawan"
   ClientHeight    =   8745
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14430
   Icon            =   "Cetak_Karyawan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8745
   ScaleWidth      =   14430
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
      ItemData        =   "Cetak_Karyawan.frx":000C
      Left            =   2160
      List            =   "Cetak_Karyawan.frx":001C
      TabIndex        =   12
      Text            =   "-- Urutkan Berdasar --"
      Top             =   1980
      Width           =   2655
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
      TabIndex        =   8
      Top             =   8100
      Width           =   12765
      Begin Project1.isButton ButtonKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   11220
         TabIndex        =   9
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
         TabIndex        =   10
         Top             =   60
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Icon            =   "Cetak_Karyawan.frx":0056
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
      Left            =   1860
      TabIndex        =   7
      Top             =   1980
      Width           =   255
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
      Left            =   2160
      TabIndex        =   6
      Text            =   "-- Pilih Departemen --"
      Top             =   1260
      Width           =   2655
   End
   Begin VB.OptionButton OptDept 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   1860
      TabIndex        =   5
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
      TabIndex        =   1
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
      ScaleWidth      =   14430
      TabIndex        =   0
      Top             =   0
      Width           =   14430
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "CETAK DATA KARYAWAN"
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
         TabIndex        =   3
         Top             =   60
         Width           =   4815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Preview atau Cetak Data Karyawan Yang Tersedia."
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
         TabIndex        =   2
         Top             =   300
         Width           =   3690
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cetak Semua Data Karyawan"
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
      TabIndex        =   11
      Top             =   1680
      Width           =   2460
   End
   Begin VB.Label Label2 
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
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1860
      TabIndex        =   4
      Top             =   960
      Width           =   2730
   End
   Begin VB.Image Image1 
      Height          =   1350
      Left            =   300
      Picture         =   "Cetak_Karyawan.frx":03F0
      Stretch         =   -1  'True
      Top             =   960
      Width           =   1350
   End
End
Attribute VB_Name = "Cetak_Karyawan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RSDepartemen As New ADODB.Recordset
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

    If OptDept.Value = True Then
        MyRS.Open "SELECT karyawan.nama as knama,departemen.keterangan as kdep,karyawan.id as kid" & _
                " From karyawan INNER JOIN departemen_karyawan ON karyawan.id = departemen_karyawan.id" & _
                " INNER JOIN departemen ON departemen_karyawan.departemen_id = departemen.id where departemen.keterangan='" & ComboDepartemen.Text & "'", Conn, adOpenForwardOnly, adLockReadOnly
        strSQL = "SELECT karyawan.nama as knama,departemen.keterangan as kdep,karyawan.id as kid" & _
               " From karyawan INNER JOIN departemen_karyawan ON karyawan.id = departemen_karyawan.id" & _
               " INNER JOIN departemen ON departemen_karyawan.departemen_id = departemen.id where departemen.keterangan='" & ComboDepartemen.Text & "'"
        strData = MyRS.RecordCount
    ElseIf OptAll.Value = True Then
        MyRS.Open "SELECT karyawan.nama as knama,departemen.keterangan as kdep,karyawan.id as kid" & _
                " From karyawan INNER JOIN departemen_karyawan ON karyawan.id = departemen_karyawan.id" & _
                " INNER JOIN departemen ON departemen_karyawan.departemen_id = departemen.id ORDER by nama ASC", Conn, adOpenForwardOnly, adLockReadOnly
        strSQL = "SELECT karyawan.nama as knama,departemen.keterangan as kdep,karyawan.id as kid" & _
               " From karyawan INNER JOIN departemen_karyawan ON karyawan.id = departemen_karyawan.id" & _
               " INNER JOIN departemen ON departemen_karyawan.departemen_id = departemen.id ORDER " + ComboUrutkan.Text
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
    With DataKaryawan
        .DC1.ConnectionString = "DRIVER={MySQL ODBC 5.1 Driver};SERVER=" & Servers & ";DATABASE=" & DBase & ";UID=" & Username & ";PWD=" & Password & ";PORT=" & Ports & ";OPTION=2+3+8+32+2048+16384+4194304;"
        .DC1.Source = strSQL

        .WindowState = 2
        .lblData.Caption = strData & " KARYAWAN"
        .lblJudul.Caption = "LAPORAN DATA KARYAWAN " & Format(Date, "dd mmmm yyyy")
        If OptDept.Value = True Then
            .Label10.Caption = "Berdasarkan Departemen : "
            .lblFilter.Caption = ComboDepartemen.Text
        ElseIf OptAll.Value = True Then
            .Label10.Caption = "Berdasarkan : "
            .lblFilter.Caption = "SEMUA"
        End If
        .lblTitle.Caption = StrTitle
        .LblAlamat.Caption = StrAlamat

        .Show
        .WindowState = 2
    End With
End Sub

Private Sub Form_Load()
    Call SetIcon(Me.hWnd, "FORMICON", False)
    strKode = 1
    With Me
        .Top = 0
        .Height = Screen.Height
        .Left = 0
        .Width = Screen.Width
    End With

    Call initCombo
End Sub

Private Sub initCombo()
    Set RSDepartemen = New ADODB.Recordset
    Dim Combo1 As String
    Combo1 = ComboListSingle(ComboDepartemen, RSDepartemen, "departemen", "*", "where id>'1' order by keterangan asc", "keterangan")
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

Private Sub OptDept_Click()
    If OptDept.Value = True Then
        ComboDepartemen.Enabled = True
        ComboDepartemen.SetFocus
        ComboUrutkan.Enabled = False
        ButtonPreview.Enabled = True
    ElseIf OptDept.Value = False Then
        ComboDepartemen.Enabled = False
    End If
End Sub

Private Sub OptAll_Click()
    If OptAll.Value = True Then
        ComboDepartemen.Enabled = False
        ComboUrutkan.Enabled = True
        ComboUrutkan.SetFocus
        ButtonPreview.Enabled = True
    ElseIf OptAll.Value = False Then
    End If
End Sub
