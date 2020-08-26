VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form FormDepartemen 
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Data Departemen"
   ClientHeight    =   7125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13155
   LinkTopic       =   "Form1"
   ScaleHeight     =   7125
   ScaleWidth      =   13155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCari 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   10680
      TabIndex        =   14
      Top             =   360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00D9FBDB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   60
      ScaleHeight     =   600
      ScaleWidth      =   12765
      TabIndex        =   0
      Top             =   6420
      Width           =   12765
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00D9FBDB&
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   6840
         ScaleHeight     =   345
         ScaleWidth      =   4155
         TabIndex        =   1
         Top             =   120
         Width           =   4150
         Begin VB.CommandButton btnNext 
            Height          =   315
            Left            =   3450
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Next 250"
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
         Begin VB.CommandButton btnPrev 
            Height          =   315
            Left            =   3135
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Previous 250"
            Top             =   15
            Width           =   315
         End
         Begin VB.CommandButton btnFirst 
            Height          =   315
            Left            =   2820
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "First 250"
            Top             =   15
            Width           =   315
         End
         Begin VB.Label lblPageInfo 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0 - 0 of 0"
            Height          =   255
            Left            =   180
            TabIndex        =   6
            Top             =   60
            Width           =   2535
         End
      End
      Begin Project1.isButton mDepartemen 
         Height          =   435
         Index           =   1
         Left            =   2520
         TabIndex        =   7
         Top             =   0
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   767
         Icon            =   "FormDepartemen.frx":0000
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
      Begin Project1.isButton mDepartemen 
         Height          =   435
         Index           =   2
         Left            =   4680
         TabIndex        =   8
         Top             =   0
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   767
         Icon            =   "FormDepartemen.frx":0A12
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
      Begin Project1.isButton mDepartemen 
         Height          =   435
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   0
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   767
         Icon            =   "FormDepartemen.frx":1424
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
         TabIndex        =   10
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
   Begin VSFlex8Ctl.VSFlexGrid GridDepartemen 
      Height          =   4965
      Left            =   0
      TabIndex        =   11
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
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FormDepartemen.frx":1F36
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
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "DATA DEPARTEMEN"
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
      Caption         =   "Kelola Informasi Data Departemen Yang Tersedia."
      Height          =   195
      Left            =   195
      TabIndex        =   12
      Top             =   300
      Width           =   3540
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
Attribute VB_Name = "FormDepartemen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

Dim Baris As Integer
Dim Hal As Integer      ' variable untuk menampung halaman
Dim TotHal As Integer   ' variabel untuk menampugn total halaman
Dim A As String
Dim x As Double         'untk menampilkan record dimulai dr record ke x
Dim y As Double         'variabel untuk menampung jumlah data perhalaman
Dim i As Double

Public blnPilih As Boolean
Dim RSDepartemen As ADODB.Recordset

'Menampilkan Data Ke GridKaryawan
Sub TampilGrid()
On Error GoTo errHandle
    Baris = 0
    If RSDepartemen.EOF Then
        Exit Sub
    Else
        With RSDepartemen
            .MoveFirst
            Do Until .EOF
                Baris = Baris + 1
                GridDepartemen.Rows = Baris
                GridDepartemen.AddItem _
                        Baris & vbTab & _
                        RSDepartemen!Keterangan & vbTab & _
                                     RSDepartemen!autokode_karyawan
                .MoveNext
            Loop
            'Call BarisStokWarna
            GridDepartemen.Select 1, 1, 1, GridDepartemen.Cols - 1
        End With
    End If
Exit Sub
errHandle:
    MsgBox err.Description & vbCrLf & vbCrLf & strSQL
End Sub

Sub TampilanAwal()
    x = 0
    y = 75
    A = x & "," & y
    Hal = 1

    'Menampilkan Data ke Grid
    Set RSDepartemen = New ADODB.Recordset
    RSDepartemen.Open "SELECT autokode_karyawan, keterangan FROM departemen where keterangan like '" & Trim(txtCari) & "%' LIMIT " & A, _
                 Conn, adOpenForwardOnly, adLockReadOnly
    Call TampilGrid

    'Menentukan Jml Halaman
    Set RSDepartemen = New ADODB.Recordset
    RSDepartemen.Open "SELECT autokode_karyawan, keterangan FROM departemen where keterangan like '" & Trim(txtCari) & "%'", Conn, adOpenForwardOnly, adLockReadOnly
    i = RSDepartemen.RecordCount / y

    TotHal = Int(i) + IIf(CDbl(CStr(i - Int(i))) > 0.1, 1, 0)
    If TotHal = 0 Then TotHal = 1
    lblPageInfo.Caption = "Page " & Hal & " of " & TotHal & " Pages"
    '    HighlightRowG 1, GridDepartemen
End Sub

Private Sub btnFirst_Click()
    x = 0
    y = y
    A = x & "," & y
    Set RSDepartemen = New ADODB.Recordset
    RSDepartemen.Open "SSELECT autokode_karyawan, keterangan FROM departemen where keterangan like '" & Trim(txtCari) & "%'  LIMIT " & A, Conn, adOpenForwardOnly, adLockReadOnly
    Call TampilGrid
    Hal = 1
    lblPageInfo.Caption = "Page " & Hal & " of " & TotHal & " Pages"
End Sub

Private Sub btnLast_Click()
    x = (TotHal - 1) * y
    y = y
    A = x & "," & y

    Set RSDepartemen = New ADODB.Recordset
    RSDepartemen.Open "SELECT autokode_karyawan, keterangan FROM departemen where keterangan like '" & Trim(txtCari) & "%' LIMIT " & A, Conn, adOpenForwardOnly, adLockReadOnly
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
        'MsgBox "Sudah sampai pada halaman terakhir !", vbCritical, ".:: Informasi "
        Exit Sub
    Else
        btnPrev.Enabled = True
        btnLast.Enabled = True
        btnNext.Enabled = True
        btnFirst.Enabled = True
        x = x + y
        y = y
        A = x & "," & y
        Set RSDepartemen = New ADODB.Recordset
        RSDepartemen.Open "SELECT autokode_karyawan, keterangan FROM departemen where keterangan like '" & Trim(txtCari) & "%' LIMIT " & A, Conn, adOpenForwardOnly, adLockReadOnly
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
        'MsgBox "Ini sudah sampai pada halaman pertama..!", vbCritical, ".:: Informasi "
        Exit Sub
    Else
        btnPrev.Enabled = True
        btnLast.Enabled = True
        btnNext.Enabled = True
        btnFirst.Enabled = True
        x = x - y
        y = y
        A = x & "," & y
        Set RSDepartemen = New ADODB.Recordset
        RSDepartemen.Open "SELECT autokode_karyawan, keterangan FROM departemen where keterangan like '" & Trim(txtCari) & "%' LIMIT " & A, Conn, adOpenForwardOnly, adLockReadOnly
        Call TampilGrid
        Hal = Hal - 1
        lblPageInfo.Caption = "Page " & Hal & " of " & TotHal & " Pages"
    End If
End Sub

Private Sub ButtonKeluar_Click()
    blnPilih = False
    Unload Me
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
End Sub

Private Sub Form_Resize()
'On Error Resume Next
    Call AutoResize
End Sub

Public Sub AutoResize()
    shpBar.Width = ScaleWidth

    With GridDepartemen
        .Top = shpBar.Height + 120
        .Left = 120
        .Width = FormMain.PictureBeranda.ScaleWidth - 240
        .Height = (FormMain.PictureBeranda.ScaleHeight - Picture1.ScaleHeight) - GridDepartemen.Top - 120
    End With

    Picture1.Move 0, GridDepartemen.Height + shpBar.Height + 240, GridDepartemen.Width + 240

    ButtonKeluar.Move Picture1.ScaleWidth - ButtonKeluar.Width - 180, 90

    For i = 0 To mDepartemen.Count - 1
        mDepartemen(i).Move 180 + (i * 1875) + (i * 120), 90, 1875
    Next
End Sub

Private Sub GridDepartemen_Click()
    blnPilih = True
    If GridDepartemen.TextMatrix(GridDepartemen.Row, 2) = "00001" Then
        mDepartemen(1).Enabled = False
        mDepartemen(2).Enabled = False
    Else
        mDepartemen(1).Enabled = True
        mDepartemen(2).Enabled = True
    End If
End Sub

Private Sub GridDepartemen_DblClick()
    Dim MyRS As New ADODB.Recordset
    Set MyRS = New ADODB.Recordset
    
    If GridDepartemen.TextMatrix(GridDepartemen.Row, 2) = "00001" Then
        mDepartemen(1).Enabled = False
        mDepartemen(2).Enabled = False
    Else
        MyRS.Open "select * from departemen Where autokode_karyawan Like '" & GridDepartemen.TextMatrix(GridDepartemen.Row, 2) & "%'", Conn, adOpenDynamic, adLockOptimistic
        If GridDepartemen.Rows > 1 Then
            With Add_Departemen
                blnTambah = False
                .Caption = "Edit Data Departemen"
                .TextKode = MyRS.Fields("autokode_karyawan")
                .LabelDepartemen.Caption = MyRS.Fields("keterangan")
                .TextDepartemen = MyRS.Fields("keterangan")
                .Show vbModal
            End With
        End If
    End If
End Sub

Private Sub GridDepartemen_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    GridDepartemen.Redraw = False
End Sub

Private Sub GridDepartemen_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    GridDepartemen.RowSel = GridDepartemen.Row
    GridDepartemen.Redraw = True
End Sub

Private Sub mDepartemen_Click(Index As Integer)
    Select Case Index
    Case 0
        With Add_Departemen
            blnTambah = True
            .Caption = "Input Data Departemen"
            .Show vbModal
        End With
    Case 1
        GridDepartemen_DblClick
    Case 2

        If blnPilih = False Then
            MsgBox "Data Departemen belum dipilih", vbInformation, "Informasi"
            Exit Sub
        End If

        Pesan_Peringatan "Question", "Apakah Data Departemen dengan Nama " & Chr(34) & GridDepartemen.TextMatrix(GridDepartemen.Row, 2) & Chr(34) & " Ingin dihapus ?", "Konfirmasi"
        If Respon = "Iya" Then
            Set ClassDepartemen = New ClassDepartemen
            Sukses = ClassDepartemen.DeleteDepartemen(GridDepartemen.TextMatrix(GridDepartemen.Row, 2), Conn)
            If Sukses = True Then
                FormDepartemen.TampilanAwal
            Else
                MsgBox "Data Kas gagal dihapus", vbExclamation, "Peringatan"
            End If
        End If

    End Select
End Sub

Private Sub Picture1_Resize()
    Picture2.Left = Picture1.ScaleWidth - Picture2.ScaleWidth - ButtonKeluar.Width - 360
End Sub
