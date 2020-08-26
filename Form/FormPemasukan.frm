VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form FormPemasukan 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Data Pemasukan"
   ClientHeight    =   7140
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13065
   Icon            =   "FormPemasukan.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7140
   ScaleWidth      =   13065
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
         Left            =   6660
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
      Begin Project1.isButton ButtonKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   11220
         TabIndex        =   8
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
      Begin Project1.isButton mPengeluaran 
         Height          =   435
         Index           =   1
         Left            =   2040
         TabIndex        =   9
         Top             =   0
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   767
         Icon            =   "FormPemasukan.frx":000C
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
      Begin Project1.isButton mPengeluaran 
         Height          =   435
         Index           =   2
         Left            =   4140
         TabIndex        =   10
         Top             =   0
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   767
         Icon            =   "FormPemasukan.frx":0A1E
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
      Begin Project1.isButton mPengeluaran 
         Height          =   435
         Index           =   0
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   767
         Icon            =   "FormPemasukan.frx":1430
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
   End
   Begin VB.TextBox txtCari 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7200
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VSFlex8Ctl.VSFlexGrid GridPemasukan 
      Height          =   5205
      Left            =   0
      TabIndex        =   12
      Top             =   600
      Width           =   11295
      _cx             =   19923
      _cy             =   9181
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
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FormPemasukan.frx":1F42
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
      Caption         =   "Kelola Informasi Data Pemasukan Yang Tersedia."
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
      Left            =   195
      TabIndex        =   14
      Top             =   300
      Width           =   3510
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "DATA PEMASUKAN"
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
   Begin VB.Shape shpBar 
      BackColor       =   &H00D9FBDB&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   600
      Left            =   0
      Top             =   0
      Width           =   7155
   End
End
Attribute VB_Name = "FormPemasukan"
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

Dim RSPengeluaran As New ADODB.Recordset

'Menampilkan Data Ke GridPemasukan
Sub TampilGrid()
'On Error Resume Next
'Call AktifGrid
    Baris = 0
    If RSPengeluaran.EOF Then
        Exit Sub
    Else
        With RSPengeluaran
            .MoveFirst
            Do Until .EOF
                Baris = Baris + 1
                GridPemasukan.Rows = Baris
                GridPemasukan.AddItem _
                        Format(RSPengeluaran!Tanggal, "dd/mm/yyyy") & vbTab & _
                                                                    RSPengeluaran!Kode & vbTab & _
                                                                    RSPengeluaran!Keterangan & vbTab & _
                                                                    RSPengeluaran!Kode_Kas & vbTab & _
                                                                    Format(RSPengeluaran!Jumlah, "#,##0") & vbTab & _
                                                                    RSPengeluaran!Operator
                .MoveNext
            Loop
            'Call BarisStokWarna
            GridPemasukan.Select 1, 1, 1, GridPemasukan.Cols - 1
        End With
    End If
End Sub

Sub TampilanAwal()
    x = 0
    y = 75
    A = x & "," & y
    Hal = 1

    'Menampilkan Data ke Grid
    Set RSPengeluaran = New ADODB.Recordset
    RSPengeluaran.Open "SELECT kode,tanggal,keterangan,jumlah,operator,kode_kas FROM pemasukan ORDER BY tanggal ASC LIMIT " & A, _
                       Conn, adOpenForwardOnly, adLockReadOnly
    Call TampilGrid

    'Menentukan Jml Halaman
    Set RSPengeluaran = New ADODB.Recordset
    RSPengeluaran.Open "SELECT kode,tanggal,keterangan,jumlah,operator,kode_kas FROM pemasukan ORDER BY tanggal ASC", Conn, adOpenForwardOnly, adLockReadOnly
    i = RSPengeluaran.RecordCount / y

    TotHal = Int(i) + IIf(CDbl(CStr(i - Int(i))) > 0.1, 1, 0)
    If TotHal = 0 Then TotHal = 1
    lblPageInfo.Caption = "Page " & Hal & " of " & TotHal & " Pages"
    '    HighlightRowG 1, GridPemasukan
End Sub

Private Sub btnFirst_Click()
    x = 0
    y = y
    A = x & "," & y
    Set RSPengeluaran = New ADODB.Recordset
    RSPengeluaran.Open "SELECT kode,tanggal,keterangan,jumlah,operator,kode_kas FROM pemasukan ORDER BY tanggal ASC  LIMIT " & A, Conn, adOpenForwardOnly, adLockReadOnly
    Call TampilGrid
    Hal = 1
    lblPageInfo.Caption = "Page " & Hal & " of " & TotHal & " Pages"
End Sub

Private Sub btnLast_Click()
    x = (TotHal - 1) * y
    y = y
    A = x & "," & y

    Set RSPengeluaran = New ADODB.Recordset
    RSPengeluaran.Open "SELECT kode,tanggal,keterangan,jumlah,operator,kode_kas FROM pemasukan ORDER BY tanggal ASC LIMIT " & A, Conn, adOpenForwardOnly, adLockReadOnly
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
        Set RSPengeluaran = New ADODB.Recordset
        RSPengeluaran.Open "SELECT kode,tanggal,keterangan,jumlah,operator,kode_kas FROM pemasukan ORDER BY tanggal ASC LIMIT " & A, Conn, adOpenForwardOnly, adLockReadOnly
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
        Set RSPengeluaran = New ADODB.Recordset
        RSPengeluaran.Open "SELECT kode,tanggal,keterangan,jumlah,operator,kode_kas FROM pemasukan ORDER BY tanggal ASC LIMIT " & A, Conn, adOpenForwardOnly, adLockReadOnly
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

    With GridPemasukan
        .Top = shpBar.Height + 120
        .Left = 120
        .Width = FormMain.PictureBeranda.ScaleWidth - 240
        .Height = (FormMain.PictureBeranda.ScaleHeight - Picture1.ScaleHeight) - GridPemasukan.Top - 120
    End With

    Picture1.Move 0, GridPemasukan.Height + shpBar.Height + 240, GridPemasukan.Width + 240

    ButtonKeluar.Move Picture1.ScaleWidth - ButtonKeluar.Width - 180, 90

    For i = 0 To mPengeluaran.Count - 1
        mPengeluaran(i).Move 180 + (i * 1875) + (i * 120), 90, 1875
    Next
End Sub

Private Sub GridPemasukan_Click()
    blnPilih = True
End Sub

Private Sub GridPemasukan_DblClick()
    Dim MyRS As New ADODB.Recordset
    Dim MyRS2 As New ADODB.Recordset
    Set MyRS = New ADODB.Recordset
    Set MyRS2 = New ADODB.Recordset
    MyRS.Open "select * from pemasukan Where kode Like '" & GridPemasukan.TextMatrix(GridPemasukan.Row, 1) & "%'", Conn, adOpenDynamic, adLockOptimistic
    MyRS2.Open "select * from kas Where kode Like '" & GridPemasukan.TextMatrix(GridPemasukan.Row, 3) & "%'", Conn, adOpenDynamic, adLockOptimistic
    If GridPemasukan.Rows > 1 Then
        With Add_Pemasukan
            blnTambah = False
            .Caption = "Edit Data Pemasukan"
            .TextKode.Text = MyRS.Fields("kode")
            .TextTanggal.Text = MyRS.Fields("tanggal")
            .Tanggal.Value = MyRS.Fields("tanggal")
            .TextKeterangan = MyRS.Fields("keterangan")
            .ComboKas.Text = MyRS2.Fields("nama")
            .LabelKode.Caption = MyRS2.Fields("kode")
            .LabelKodeAwal.Caption = MyRS2.Fields("kode")
            .TextJumlah.Text = MyRS.Fields("jumlah")
            .TextJumlah.Alignment = 1
            .TextJumlah.Text = Format(.TextJumlah.Text, "#,##0")
            .LabelJumlahKas.Caption = Val(Format(MyRS2.Fields("saldo"), "###0")) - Val(Format(MyRS.Fields("jumlah"), "###0"))
            .LabelSaldoAwal.Caption = Val(Format(MyRS2.Fields("saldo"), "###0")) - Val(Format(MyRS.Fields("jumlah"), "###0"))
            .Show vbModal
        End With
    End If
End Sub

Private Sub GridPemasukan_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    GridPemasukan.Redraw = False
End Sub

Private Sub GridPemasukan_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    GridPemasukan.RowSel = GridPemasukan.Row
    GridPemasukan.Redraw = True
End Sub

Private Sub BarisStokWarna()

    For i = 1 To GridPemasukan.Rows - 1

        If Val(Left(GridPemasukan.TextMatrix(i, 4), 1)) <= "1" Then
            GridPemasukan.Cell(flexcpBackColor, i, 4) = &H80FFFF
        End If

        If Val(Left(GridPemasukan.TextMatrix(i, 4), 1)) >= "1" Then
            GridPemasukan.Cell(flexcpBackColor, i, 4) = &H80FF80
        End If

    Next i

End Sub

Private Sub mPengeluaran_Click(Index As Integer)
    Select Case Index
    Case 0
        With Add_Pemasukan
            blnTambah = True
            .AutoNumber
            .Tanggal.Value = Now
            .TextTanggal.Text = Format(.Tanggal.Value, "dd/mm/yyyy")
            .Caption = "Input Data Pemasukan"
            .Show vbModal
        End With
    Case 1
        GridPemasukan_DblClick
    Case 2

    End Select
End Sub

Private Sub Picture1_Resize()
    Picture2.Left = Picture1.ScaleWidth - Picture2.ScaleWidth - ButtonKeluar.Width - 360
End Sub


