VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form FormUser 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Manajemen User"
   ClientHeight    =   7245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13170
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   13170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00D9FBDB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   60
      ScaleHeight     =   600
      ScaleWidth      =   12765
      TabIndex        =   1
      Top             =   6420
      Width           =   12765
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00D9FBDB&
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   6840
         ScaleHeight     =   345
         ScaleWidth      =   4155
         TabIndex        =   2
         Top             =   120
         Width           =   4150
         Begin VB.CommandButton btnFirst 
            Height          =   315
            Left            =   2820
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "First 250"
            Top             =   15
            Width           =   315
         End
         Begin VB.CommandButton btnPrev 
            Height          =   315
            Left            =   3135
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Previous 250"
            Top             =   15
            Width           =   315
         End
         Begin VB.CommandButton btnLast 
            Height          =   315
            Left            =   3765
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Last 250"
            Top             =   15
            Width           =   315
         End
         Begin VB.CommandButton btnNext 
            Height          =   315
            Left            =   3450
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Next 250"
            Top             =   15
            Width           =   315
         End
         Begin VB.Label lblPageInfo 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0 - 0 of 0"
            Height          =   255
            Left            =   180
            TabIndex        =   7
            Top             =   60
            Width           =   2535
         End
      End
      Begin Project1.isButton mUser 
         Height          =   435
         Index           =   1
         Left            =   2520
         TabIndex        =   8
         Top             =   0
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   767
         Icon            =   "FormUser.frx":000C
         Style           =   8
         Caption         =   "&Ubah"
         iNonThemeStyle  =   0
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
      Begin Project1.isButton mUser 
         Height          =   435
         Index           =   2
         Left            =   4680
         TabIndex        =   9
         Top             =   0
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   767
         Icon            =   "FormUser.frx":0A1E
         Style           =   8
         Caption         =   "&Hapus"
         iNonThemeStyle  =   0
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
      Begin Project1.isButton mUser 
         Height          =   435
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   0
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   767
         Icon            =   "FormUser.frx":1430
         Style           =   8
         Caption         =   "&Tambah"
         iNonThemeStyle  =   0
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
      Begin Project1.isButton ButtonKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   11220
         TabIndex        =   11
         Top             =   0
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
   End
   Begin VB.ComboBox ComboDepartemen 
      Height          =   315
      Left            =   6780
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   140
      Width           =   2355
   End
   Begin VSFlex8Ctl.VSFlexGrid GridUser 
      Height          =   4965
      Left            =   0
      TabIndex        =   12
      Top             =   600
      Width           =   10395
      _cx             =   18336
      _cy             =   8758
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   8421504
      ForeColorFixed  =   16777215
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   15531501
      BackColorAlternate=   16774388
      GridColor       =   -2147483633
      GridColorFixed  =   0
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FormUser.frx":1F42
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kelola Informasi Data User Login Yang Tersedia."
      Height          =   195
      Left            =   195
      TabIndex        =   15
      Top             =   300
      Width           =   3450
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "DATA USER LOGIN"
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
      TabIndex        =   14
      Top             =   60
      Width           =   4815
   End
   Begin VB.Label LabelFilter 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Filter Data:"
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
      Left            =   5400
      TabIndex        =   13
      Top             =   180
      Width           =   930
   End
   Begin VB.Shape shpBar 
      BackColor       =   &H00D9FBDB&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   600
      Left            =   0
      Top             =   0
      Width           =   9315
   End
End
Attribute VB_Name = "FormUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Baris As Integer
Dim Hal As Integer      ' variable untuk menampung halaman
Dim TotHal As Integer   ' variabel untuk menampugn total halaman
Dim A As String
Dim x As Double         'untk menampilkan record dimulai dr record ke x
Dim y As Double         'variabel untuk menampung jumlah data perhalaman
Dim i As Double

Public blnPilih As Boolean

Dim ClassUser As ClassUser
Dim RSUser As New ADODB.Recordset
Dim RSDepartemen As New ADODB.Recordset

'Menampilkan Data Ke GridUser
Sub TampilGrid()
'On Error Resume Next
    Baris = 0
    If RSUser.EOF Then
        Exit Sub
    Else
        With RSUser
            .MoveFirst
            Do Until .EOF
                Baris = Baris + 1
                GridUser.Rows = Baris
                GridUser.AddItem RSUser!Id & vbTab & RSUser!uname & vbTab & RSUser!NamaUser & vbTab & RSUser!jabatan & vbTab & RSUser!Kode & vbTab & RSUser!Registered & vbTab & RSUser!Gender
                .MoveNext
            Loop
            GridUser.Select 1, 1, 1, GridUser.Cols - 1
        End With
    End If
End Sub

Sub TampilanAwal()
    x = 0
    y = 75
    A = x & "," & y
    Hal = 1

    'Menampilkan Data ke Grid
    Set RSUser = New ADODB.Recordset

    RSUser.Open "SELECT *" & _
              " From pos_akses WHERE id>000 LIMIT " & A, _
                Conn, adOpenForwardOnly, adLockReadOnly
    Call TampilGrid

    'Menentukan Jml Halaman
    Set RSUser = New ADODB.Recordset

    RSUser.Open "SELECT *" & _
              " From pos_akses WHERE id>000", Conn, adOpenForwardOnly, adLockReadOnly
    i = RSUser.RecordCount / y

    TotHal = Int(i) + IIf(CDbl(CStr(i - Int(i))) > 0.1, 1, 0)
    If TotHal = 0 Then TotHal = 1
    lblPageInfo.Caption = "Page " & Hal & " of " & TotHal & " Pages"
End Sub

Private Sub btnFirst_Click()
    x = 0
    y = y
    A = x & "," & y
    Set RSUser = New ADODB.Recordset
    RSUser.Open "SELECT *" & _
              " From pos_akses WHERE id>000 LIMIT " & A, _
                Conn, adOpenForwardOnly, adLockReadOnly
    Call TampilGrid
    Hal = 1
    lblPageInfo.Caption = "Page " & Hal & " of " & TotHal & " Pages"
End Sub

Private Sub btnLast_Click()
    x = (TotHal - 1) * y
    y = y
    A = x & "," & y

    Set RSUser = New ADODB.Recordset
    RSUser.Open "SELECT *" & _
              " From pos_akses WHERE id>000 LIMIT " & A, _
                Conn, adOpenForwardOnly, adLockReadOnly
    Call TampilGrid
    Hal = TotHal
    lblPageInfo.Caption = "Page " & Hal & " of " & TotHal & " Pages"
End Sub

Private Sub btnNext_Click()
    If Hal = TotHal Then
        btnPrev.Enabled = True
        btnLast.Enabled = False
        btnNext.Enabled = False
        btnFirst.Enabled = True
        Exit Sub
    Else
        btnPrev.Enabled = True
        btnLast.Enabled = True
        btnNext.Enabled = True
        btnFirst.Enabled = True
        x = x + y
        y = y
        A = x & "," & y
        Set RSUser = New ADODB.Recordset
        RSUser.Open "SELECT *" & _
                  " From pos_akses WHERE id>000 LIMIT " & A, _
                    Conn, adOpenForwardOnly, adLockReadOnly
        Call TampilGrid
        Hal = Hal + 1
        lblPageInfo.Caption = "Page " & Hal & " of " & TotHal & " Pages"
    End If
End Sub

Private Sub btnPrev_Click()
    If Hal = 1 Then
        btnPrev.Enabled = False
        btnLast.Enabled = True
        btnNext.Enabled = True
        btnFirst.Enabled = False
        Exit Sub
    Else
        btnPrev.Enabled = True
        btnLast.Enabled = True
        btnNext.Enabled = True
        btnFirst.Enabled = True
        x = x - y
        y = y
        A = x & "," & y
        Set RSUser = New ADODB.Recordset
        RSUser.Open "SELECT *" & _
                  " From pos_akses WHERE id>000 LIMIT " & A, _
                    Conn, adOpenForwardOnly, adLockReadOnly
        Call TampilGrid
        Hal = Hal - 1
        lblPageInfo.Caption = "Page " & Hal & " of " & TotHal & " Pages"
    End If
End Sub

Private Sub ButtonKeluar_Click()
    blnPilih = False
    Unload Me
End Sub

Private Sub ComboDepartemen_Click()
    Set RSUser = New ADODB.Recordset

    RSUser.Open "SELECT *" & _
              " From pos_akses where kode='" & ComboDepartemen.Text & "' ORDER by namauser ASC LIMIT " & A, _
                Conn, adOpenForwardOnly, adLockReadOnly
    Call TampilGrid
End Sub

Private Sub Form_Load()
    Call SetIcon(Me.hWnd, "FORMICON", False)
    With Me
        .Top = 0
        .Height = Screen.Height    'FormMain.PictureMaster.Height'
        .Left = 0
        .Width = Screen.Width    'FormMain.PictureMaster.Width
    End With

    Call AutoResize

    With FormUI
        btnFirst.Picture = .iList16x16.ListImages(3).Picture
        btnPrev.Picture = .iList16x16.ListImages(4).Picture
        btnNext.Picture = .iList16x16.ListImages(5).Picture
        btnLast.Picture = .iList16x16.ListImages(6).Picture

        btnFirst.DisabledPicture = .iList16x16G.ListImages(3).Picture
        btnPrev.DisabledPicture = .iList16x16G.ListImages(4).Picture
        btnNext.DisabledPicture = .iList16x16G.ListImages(5).Picture
        btnLast.DisabledPicture = .iList16x16G.ListImages(6).Picture
    End With

    Call TampilanAwal
    'Call BarisLunasWarna
    Call initCombo
End Sub

Private Sub initCombo()
    Set RSDepartemen = New ADODB.Recordset
    Dim Combo1 As String
    Combo1 = ComboListSingle(ComboDepartemen, RSDepartemen, "pos_level", "kode", "WHERE id>000 ORDER BY kode ASC", "kode")
End Sub

Private Sub Form_Resize()
'On Error Resume Next
    Call AutoResize
End Sub

Public Sub AutoResize()
    shpBar.Width = ScaleWidth

    With GridUser
        .Top = shpBar.Height + 120
        .Left = 120
        .Width = FormMain.PictureBeranda.ScaleWidth - 240
        .Height = (FormMain.PictureBeranda.ScaleHeight - Picture1.ScaleHeight) - GridUser.Top - 120
    End With

    Picture1.Move 0, GridUser.Height + shpBar.Height + 240, GridUser.Width + 240

    ButtonKeluar.Move Picture1.ScaleWidth - ButtonKeluar.Width - 180, 90

    For i = 0 To mUser.Count - 1
        mUser(i).Move 180 + (i * 1875) + (i * 120), 90, 1875
    Next

    ComboDepartemen.Move FormMain.PictureBeranda.Width - ComboDepartemen.Width - 360, 140
    LabelFilter.Move FormMain.PictureBeranda.Width - ComboDepartemen.Width - 360 - LabelFilter.Width - 360, 180
End Sub

Private Sub BarisLunasWarna()

    For i = 1 To GridUser.Rows - 1

        If Left(GridUser.TextMatrix(i, 2), 1) = "W" Then
            GridUser.Cell(flexcpBackColor, i, 2) = &H80FFFF
        End If

        If Left(GridUser.TextMatrix(i, 2), 1) = "H" Then
            GridUser.Cell(flexcpBackColor, i, 2) = &H80FF80
        End If

        If Left(GridUser.TextMatrix(i, 2), 1) = "C" Then
            GridUser.Cell(flexcpBackColor, i, 2) = RGB(188, 63, 67)    '&H80FFFF
        End If

        If Left(GridUser.TextMatrix(i, 2), 1) = "K" Then
            GridUser.Cell(flexcpBackColor, i, 2) = RGB(111, 146, 200)    '&H80FF80
        End If

        If Left(GridUser.TextMatrix(i, 2), 1) = "A" Then
            GridUser.Cell(flexcpBackColor, i, 2) = RGB(253, 206, 5)    '&H80FF80
        End If

    Next i

End Sub

Private Sub GridUser_Click()
    blnPilih = True
End Sub

Private Sub GridUser_DblClick()
    Dim MyRS As New ADODB.Recordset
    Set MyRS = New ADODB.Recordset
    MyRS.Open "select * from pos_akses Where id Like '" & GridUser.TextMatrix(GridUser.Row, 0) & "%'", Conn, adOpenDynamic, adLockOptimistic
    If GridUser.Rows > 1 Then
        With Add_User
            blnTambah = False
            .Caption = "Edit Data User"
            .TextUser.Text = MyRS.Fields("uname")
            .TextNama.Text = MyRS.Fields("namauser")
            .TextJabatan.Text = MyRS.Fields("jabatan")
            .ComboLevel.Text = MyRS.Fields("kode")
            .cmbGender.Text = MyRS.Fields("gender")
            .TextUser.Enabled = False
            .Show vbModal
        End With
    End If
End Sub

Private Sub GridUser_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    GridUser.Redraw = False
End Sub

Private Sub GridUser_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    GridUser.RowSel = GridUser.Row
    GridUser.Redraw = True
End Sub

Private Sub mUser_Click(Index As Integer)
    Select Case Index
    Case 0
        With Add_User
            blnTambah = True
            '.TextUser.Enabled = False
            .Caption = "Input Data User"
            .Show vbModal
        End With
    Case 1
        GridUser_DblClick
    Case 2

        If blnPilih = False Then
            MsgBox "Data Kas belum dipilih", vbInformation, "Informasi"
            Exit Sub
        End If

        If GridUser.TextMatrix(GridUser.Row, 1) = FormMain.LabelUsername.Caption Then
            MsgBox "Tidak Bisa Menghapus User Yang Sedang Login, Silahkan Hapus Dari Akun Lain", vbExclamation, "Gagal Menghapus"
            Exit Sub
        End If

        Pesan_Peringatan "Question", "Apakah Data User dengan Nama " & Chr(34) & GridUser.TextMatrix(GridUser.Row, 2) & Chr(34) & " Ingin dihapus ?", "Konfirmasi"
        If Respon = "Iya" Then
            Set ClassUser = New ClassUser
            Sukses = ClassUser.DeleteUser(GridUser.TextMatrix(GridUser.Row, 1), Conn)
            If Sukses = True Then
                FormUser.TampilanAwal
            Else
                MsgBox "Data User Gagal Dihapus", vbExclamation, "Peringatan"
            End If
        End If

    End Select
End Sub

Private Sub Picture1_Resize()
    Picture2.Left = Picture1.ScaleWidth - Picture2.ScaleWidth - ButtonKeluar.Width - 360
End Sub
