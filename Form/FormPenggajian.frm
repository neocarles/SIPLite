VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form FormPenggajian 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Data Gaji Karyawan"
   ClientHeight    =   8850
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15285
   Icon            =   "FormPenggajian.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8850
   ScaleWidth      =   15285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox shpBar 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00D9FBDB&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   15255
      TabIndex        =   11
      Top             =   0
      Width           =   15285
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kelola Informasi Data Gaji Karyawan Yang Tersedia."
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
         TabIndex        =   13
         Top             =   300
         Width           =   3735
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "TRANSAKSI GAJI KARYAWAN"
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
         TabIndex        =   12
         Top             =   60
         Width           =   4815
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00D9FBDB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   12765
      TabIndex        =   0
      Top             =   8220
      Width           =   12765
      Begin Project1.isButton mPenggajian 
         Height          =   435
         Index           =   3
         Left            =   6120
         TabIndex        =   14
         Top             =   0
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   767
         Icon            =   "FormPenggajian.frx":000C
         Style           =   8
         Caption         =   "&Generate"
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
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00D9FBDB&
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   6660
         ScaleHeight     =   345
         ScaleWidth      =   4155
         TabIndex        =   1
         Top             =   120
         Width           =   4150
         Begin VB.CommandButton btnFirst 
            Height          =   315
            Left            =   2820
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "First 250"
            Top             =   15
            Width           =   315
         End
         Begin VB.CommandButton btnPrev 
            Height          =   315
            Left            =   3135
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Previous 250"
            Top             =   15
            Width           =   315
         End
         Begin VB.CommandButton btnLast 
            Height          =   315
            Left            =   3765
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Last 250"
            Top             =   15
            Width           =   315
         End
         Begin VB.CommandButton btnNext 
            Height          =   315
            Left            =   3450
            Style           =   1  'Graphical
            TabIndex        =   2
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
            TabIndex        =   6
            Top             =   60
            Width           =   2535
         End
      End
      Begin Project1.isButton ButtonKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   11220
         TabIndex        =   7
         Top             =   90
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
      Begin Project1.isButton mPenggajian 
         Height          =   435
         Index           =   1
         Left            =   2040
         TabIndex        =   8
         Top             =   0
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   767
         Icon            =   "FormPenggajian.frx":03A6
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
      Begin Project1.isButton mPenggajian 
         Height          =   435
         Index           =   2
         Left            =   4140
         TabIndex        =   9
         Top             =   0
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   767
         Icon            =   "FormPenggajian.frx":0DB8
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
      Begin Project1.isButton mPenggajian 
         Height          =   435
         Index           =   0
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   767
         Icon            =   "FormPenggajian.frx":17CA
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
   Begin VSFlex8Ctl.VSFlexGrid GridGaji 
      Height          =   7365
      Left            =   120
      TabIndex        =   15
      Top             =   720
      Width           =   15015
      _cx             =   26485
      _cy             =   12991
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
      Cols            =   13
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FormPenggajian.frx":22DC
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
End
Attribute VB_Name = "FormPenggajian"
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

Dim RSPenggajian As New ADODB.Recordset
Dim ClassPenggajian As ClassPenggajian

'Menampilkan Data Ke GridGaji
Sub TampilGrid()
'On Error Resume Next
'Call AktifGrid
    Baris = 0
    If RSPenggajian.EOF Then
        Exit Sub
    Else
        With RSPenggajian
            .MoveFirst
            Do Until .EOF
                Baris = Baris + 1
                GridGaji.Rows = Baris
                GridGaji.AddItem _
                        RSPenggajian!Id & vbTab & _
                                        Format(RSPenggajian!Tanggal, "dd/mm/yyyy") & vbTab & _
                                        RSPenggajian!Kode & vbTab & _
                                        RSPenggajian!Nama & vbTab & _
                                        RSPenggajian!Basic_Hk & vbTab & _
                                        Format(RSPenggajian!Basic_Gaji, "#,##0") & vbTab & _
                                        RSPenggajian!Hk_Miss & vbTab & _
                                        Format(RSPenggajian!Hk_Potongan, "#,##0") & vbTab & _
                                        Format(RSPenggajian!Hk_TotPotong, "#,##0") & vbTab & _
                                        Format(RSPenggajian!Tunjangan, "#,##0") & vbTab & _
                                        Format(RSPenggajian!Potongan, "#,##0") & vbTab & _
                                        Format(RSPenggajian!Total_Gaji, "#,##0") & vbTab & _
                                        Format(RSPenggajian!Tanggal, "mmmm") & vbTab & _
                                        RSPenggajian!Operator
                .MoveNext
            Loop
            'Call BarisStokWarna
            GridGaji.Select 1, 1, 1, GridGaji.Cols - 1
        End With
    End If
End Sub

Sub TampilanAwal()
    x = 0
    y = 75
    A = x & "," & y
    Hal = 1

    'Menampilkan Data ke Grid
    Set RSPenggajian = New ADODB.Recordset
    RSPenggajian.Open "SELECT * FROM karyawan_gaji ORDER BY kode ASC LIMIT " & A, _
                      Conn, adOpenForwardOnly, adLockReadOnly
    Call TampilGrid

    'Menentukan Jml Halaman
    Set RSPenggajian = New ADODB.Recordset
    RSPenggajian.Open "SELECT * FROM karyawan_gaji ORDER BY kode ASC", Conn, adOpenForwardOnly, adLockReadOnly
    i = RSPenggajian.RecordCount / y

    TotHal = Int(i) + IIf(CDbl(CStr(i - Int(i))) > 0.1, 1, 0)
    If TotHal = 0 Then TotHal = 1
    lblPageInfo.Caption = "Page " & Hal & " of " & TotHal & " Pages"
    '    HighlightRowG 1, GridGaji
End Sub

Private Sub btnFirst_Click()
    x = 0
    y = y
    A = x & "," & y
    Set RSPenggajian = New ADODB.Recordset
    RSPenggajian.Open "SELECT * FROM karyawan_gaji ORDER BY kode ASC  LIMIT " & A, Conn, adOpenForwardOnly, adLockReadOnly
    Call TampilGrid
    Hal = 1
    lblPageInfo.Caption = "Page " & Hal & " of " & TotHal & " Pages"
End Sub

Private Sub btnLast_Click()
    x = (TotHal - 1) * y
    y = y
    A = x & "," & y

    Set RSPenggajian = New ADODB.Recordset
    RSPenggajian.Open "SELECT * FROM karyawan_gaji ORDER BY kode ASC LIMIT " & A, Conn, adOpenForwardOnly, adLockReadOnly
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
        Set RSPenggajian = New ADODB.Recordset
        RSPenggajian.Open "SELECT * FROM karyawan_gaji ORDER BY kode ASC LIMIT " & A, Conn, adOpenForwardOnly, adLockReadOnly
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
        Set RSPenggajian = New ADODB.Recordset
        RSPenggajian.Open "SELECT * FROM karyawan_gaji ORDER BY kode ASC LIMIT " & A, Conn, adOpenForwardOnly, adLockReadOnly
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
    With GridGaji
        .Top = shpBar.Height + 120
        .Left = 120
        .Width = FormMain.PictureBeranda.ScaleWidth - 240
        .Height = (FormMain.PictureBeranda.ScaleHeight - Picture1.ScaleHeight) - GridGaji.Top - 120
    End With

    'Picture1.Move 0, GridGaji.Height + shpBar.Height + 240, GridGaji.Width + 240
    Picture1.Move 0, FormMain.PictureBeranda.ScaleHeight - Picture1.ScaleHeight, FormMain.PictureBeranda.ScaleWidth

    ButtonKeluar.Move Picture1.ScaleWidth - ButtonKeluar.Width - 180, 90

    For i = 0 To mPenggajian.Count - 1
        mPenggajian(i).Move 180 + (i * 1875) + (i * 120), 90, 1875
    Next
End Sub

Private Sub GridGaji_Click()
    blnPilih = True
End Sub

Private Sub GridGaji_DblClick()
    On Error Resume Next
    Dim MyRS As New ADODB.Recordset
    Dim MyRS2 As New ADODB.Recordset
    Set MyRS = New ADODB.Recordset
    Set MyRS2 = New ADODB.Recordset
    MyRS.Open "select * from karyawan_gaji Where id Like '" & GridGaji.TextMatrix(GridGaji.Row, 0) & "%'", Conn, adOpenDynamic, adLockOptimistic
    MyRS2.Open "SELECT karyawan.nama as knama,departemen.keterangan as kdep,karyawan.id as kid,karyawan.status as kstatus,karyawan.Basic_Hk as kbasichk, karyawan.basic_gaji as kbasicgaji" & _
             " From karyawan INNER JOIN departemen_karyawan ON karyawan.id = departemen_karyawan.id" & _
             " INNER JOIN departemen ON departemen_karyawan.departemen_id = departemen.id WHERE karyawan.id='" & GridGaji.TextMatrix(GridGaji.Row, 2) & "'", _
               Conn, adOpenForwardOnly, adLockReadOnly
    If GridGaji.Rows > 1 Then
        With Add_Penggajian
            blnTambah = False
            .Caption = "Edit Data Penggajian"
            .CmbID.Enabled = False
            .TextKode.Text = MyRS.Fields("id")
            .CmbID.Text = MyRS2.Fields("kid") & " - " & MyRS2.Fields("knama")
            .LblID.Caption = MyRS2.Fields("kid")
            .TxtNama.Text = MyRS2.Fields("knama")
            .TextStatus.Text = MyRS2.Fields("kstatus")
            .TextTanggal.Text = Format(MyRS.Fields("tanggal"), "dd/mm/yyyy")
            .Tanggal.Value = Format(MyRS.Fields("tanggal"), "dd/mm/yyyy")
            .TextDepartemen.Text = MyRS2.Fields("kdep")
            .TextHK.Text = MyRS2.Fields("kbasichk")
            .TextGaji.Text = Format(MyRS2.Fields("kbasicgaji"), "Rp ###,##,0")
            .TextHKMiss.Text = MyRS.Fields("hk_miss")
            .TextPotong.Text = MyRS.Fields("hk_potongan")
            .TextAkhirHK.Text = MyRS.Fields("hk_totpotong")
            .TextGPokok.Text = Format(MyRS2.Fields("kbasicgaji"), "Rp ###,##,0")
            .LabelPotong.Caption = Val(Format(.TextGaji.Text, "###0")) / Val(.TextHK.Text)
            .QueryDataPotongan
            .QueryDataTunjangan
            Dim Hasil As String: Hasil = Val(.TextHKMiss.Text) * Val(Format(.TextPotong.Text, "###0"))
            .TextAkhirHK.Text = Format(Hasil, "Rp ###,##,0")
            Dim TGaji As String: TGaji = Val(Format(.TextGPokok.Text, "###0")) + Val(Format(.TextTotTunjangan.Text, "###0"))
            .TextTGaji.Text = Format(TGaji, "Rp ###,##,0")
            Dim TGajiPot As String: TGajiPot = Val(Format(.TextTGaji.Text, "###0")) - Val(Format(.TextTotPotongan.Text, "###0")) - Val(Format(.TextAkhirHK.Text, "###0"))
            .TextGajiPotong.Text = Format(TGajiPot, "Rp ###,##,0")
            .TextGajiAkhir.Text = Format(TGajiPot, "Rp ###,##,0")
            .TextGajiAkhir.Alignment = 1
            .Show vbModal
        End With
    End If
End Sub

Private Sub mPenggajian_Click(Index As Integer)
    Select Case Index
    Case 0
        With Add_Penggajian
            blnTambah = True
            .AutoNumber
            .CmbID.Enabled = True
            .Caption = "Tambah Data Penggajian"
            .Tanggal.Value = Now
            .TextTanggal.Text = Format(Now, "dd/mm/yyyy/")
            .Show vbModal
        End With
    Case 1
        GridGaji_DblClick
    Case 2
        If blnPilih = False Then
            MsgBox "Data Gaji belum dipilih", vbInformation, "Informasi"
            Exit Sub
        End If

        Pesan_Peringatan "Question", "Apakah Data Gaji dengan Kode " & Chr(34) & GridGaji.TextMatrix(GridGaji.Row, 0) & Chr(34) & " Ingin dihapus ?", "Konfirmasi"
        If Respon = "Iya" Then
            Set ClassPenggajian = New ClassPenggajian
            Sukses = ClassPenggajian.DeletePenggajian(GridGaji.TextMatrix(GridGaji.Row, 0), Conn)
            If Sukses = True Then
                FormPenggajian.TampilanAwal
            Else
                MsgBox "Data Gaji Gagal Dihapus", vbExclamation, "Peringatan"
            End If
        End If
    Case 3
        With Add_GenerateGaji
            .AutoNumber
            .Tanggal.Value = Now
            .TextTanggal.Text = Format(Now, "dd/mm/yyyy/")
            .Show vbModal
        End With
    End Select
End Sub

Private Sub Picture1_Resize()
    Picture2.Left = Picture1.ScaleWidth - Picture2.ScaleWidth - ButtonKeluar.Width - 360
End Sub
