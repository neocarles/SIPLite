VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form FormKaryawan 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Data Karyawan"
   ClientHeight    =   7170
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13050
   Icon            =   "FormKaryawan.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7170
   ScaleWidth      =   13050
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox ComboDepartemen 
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
      Left            =   6780
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   140
      Width           =   2355
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
      TabIndex        =   1
      Top             =   6420
      Width           =   12765
      Begin Project1.isButton mKaryawan 
         Height          =   435
         Index           =   3
         Left            =   6780
         TabIndex        =   16
         Top             =   0
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   767
         Icon            =   "FormKaryawan.frx":000C
         Style           =   8
         Caption         =   "&Departemen"
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
         Left            =   6840
         ScaleHeight     =   345
         ScaleWidth      =   4155
         TabIndex        =   2
         Top             =   120
         Width           =   4150
         Begin VB.CommandButton btnNext 
            Height          =   315
            Left            =   3450
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Next 250"
            Top             =   15
            Width           =   315
         End
         Begin VB.CommandButton btnLast 
            Height          =   315
            Left            =   3765
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Last 250"
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
         Begin VB.CommandButton btnFirst 
            Height          =   315
            Left            =   2820
            Style           =   1  'Graphical
            TabIndex        =   3
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
            TabIndex        =   7
            Top             =   60
            Width           =   2535
         End
      End
      Begin Project1.isButton mKaryawan 
         Height          =   435
         Index           =   1
         Left            =   2520
         TabIndex        =   10
         Top             =   0
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   767
         Icon            =   "FormKaryawan.frx":109E
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
      Begin Project1.isButton mKaryawan 
         Height          =   435
         Index           =   2
         Left            =   4680
         TabIndex        =   11
         Top             =   0
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   767
         Icon            =   "FormKaryawan.frx":1AB0
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
      Begin Project1.isButton mKaryawan 
         Height          =   435
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   0
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   767
         Icon            =   "FormKaryawan.frx":24C2
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
         TabIndex        =   13
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
   Begin VSFlex8Ctl.VSFlexGrid GridKaryawan 
      Height          =   4965
      Left            =   0
      TabIndex        =   0
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
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FormKaryawan.frx":2FD4
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
      TabIndex        =   15
      Top             =   180
      Width           =   930
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "DATA KARYAWAN"
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
      TabIndex        =   9
      Top             =   60
      Width           =   4815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kelola Informasi Data Karyawan Yang Tersedia."
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
      TabIndex        =   8
      Top             =   300
      Width           =   3420
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
Attribute VB_Name = "FormKaryawan"
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

Dim RSKaryawan As New ADODB.Recordset
Dim RSDepartemen As New ADODB.Recordset
Dim ClassKaryawan As ClassKaryawan

'Menampilkan Data Ke GridKaryawan
Sub TampilGrid()
'On Error Resume Next
    Baris = 0
    If RSKaryawan.EOF Then
        Exit Sub
    Else
        With RSKaryawan
            .MoveFirst
            Do Until .EOF
                Baris = Baris + 1
                GridKaryawan.Rows = Baris
                GridKaryawan.AddItem RSKaryawan!kid & vbTab & RSKaryawan!knama & vbTab & RSKaryawan!kdep & vbTab & RSKaryawan!kkontak & vbTab & RSKaryawan!kalamat
                .MoveNext
            Loop
            Call BarisLunasWarna
            GridKaryawan.Select 1, 1, 1, GridKaryawan.Cols - 1
        End With
    End If
End Sub

Sub TampilanAwal()
    x = 0
    y = 75
    A = x & "," & y
    Hal = 1

    'Menampilkan Data ke Grid
    Set RSKaryawan = New ADODB.Recordset

    RSKaryawan.Open "SELECT karyawan.nama as knama,departemen.keterangan as kdep,karyawan.id as kid,karyawan.alamat as kalamat,karyawan.kontak as kkontak" & _
                  " From karyawan INNER JOIN departemen_karyawan ON karyawan.id = departemen_karyawan.id" & _
                  " INNER JOIN departemen ON departemen_karyawan.departemen_id = departemen.id ORDER by kid ASC LIMIT " & A, _
                    Conn, adOpenForwardOnly, adLockReadOnly
    Call TampilGrid

    'Menentukan Jml Halaman
    Set RSKaryawan = New ADODB.Recordset

    RSKaryawan.Open "SELECT karyawan.nama as knama,departemen.keterangan as kdep,karyawan.id as kid,karyawan.alamat as kalamat,karyawan.kontak as kkontak" & _
                  " From karyawan INNER JOIN departemen_karyawan ON karyawan.id = departemen_karyawan.id" & _
                  " INNER JOIN departemen ON departemen_karyawan.departemen_id = departemen.id", Conn, adOpenForwardOnly, adLockReadOnly
    i = RSKaryawan.RecordCount / y

    TotHal = Int(i) + IIf(CDbl(CStr(i - Int(i))) > 0.1, 1, 0)
    If TotHal = 0 Then TotHal = 1
    lblPageInfo.Caption = "Page " & Hal & " of " & TotHal & " Pages"
End Sub

Private Sub btnFirst_Click()
    x = 0
    y = y
    A = x & "," & y
    Set RSKaryawan = New ADODB.Recordset
    RSKaryawan.Open "SELECT karyawan.nama as knama,departemen.keterangan as kdep,karyawan.id as kid,karyawan.alamat as kalamat,karyawan.kontak as kkontak" & _
                  " From karyawan INNER JOIN departemen_karyawan ON karyawan.id = departemen_karyawan.id" & _
                  " INNER JOIN departemen ON departemen_karyawan.departemen_id = departemen.id ORDER by kid ASC LIMIT " & A, _
                    Conn, adOpenForwardOnly, adLockReadOnly
    Call TampilGrid
    Hal = 1
    lblPageInfo.Caption = "Page " & Hal & " of " & TotHal & " Pages"
End Sub

Private Sub btnLast_Click()
    x = (TotHal - 1) * y
    y = y
    A = x & "," & y

    Set RSKaryawan = New ADODB.Recordset
    RSKaryawan.Open "SELECT karyawan.nama as knama,departemen.keterangan as kdep,karyawan.id as kid,karyawan.alamat as kalamat,karyawan.kontak as kkontak" & _
                  " From karyawan INNER JOIN departemen_karyawan ON karyawan.id = departemen_karyawan.id" & _
                  " INNER JOIN departemen ON departemen_karyawan.departemen_id = departemen.id ORDER by kid ASC LIMIT " & A, _
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
        Set RSKaryawan = New ADODB.Recordset
        RSKaryawan.Open "SELECT karyawan.nama as knama,departemen.keterangan as kdep,karyawan.id as kid,karyawan.alamat as kalamat,karyawan.kontak as kkontak" & _
                      " From karyawan INNER JOIN departemen_karyawan ON karyawan.id = departemen_karyawan.id" & _
                      " INNER JOIN departemen ON departemen_karyawan.departemen_id = departemen.id ORDER by kid ASC LIMIT " & A, _
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
        Set RSKaryawan = New ADODB.Recordset
        RSKaryawan.Open "SELECT karyawan.nama as knama,departemen.keterangan as kdep,karyawan.id as kid,karyawan.alamat as kalamat,karyawan.kontak as kkontak" & _
                      " From karyawan INNER JOIN departemen_karyawan ON karyawan.id = departemen_karyawan.id" & _
                      " INNER JOIN departemen ON departemen_karyawan.departemen_id = departemen.id ORDER by kid ASC LIMIT " & A, _
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
    If ComboDepartemen.Text = "ALL DEPARTEMEN" Then
        Set RSKaryawan = New ADODB.Recordset

        RSKaryawan.Open "SELECT karyawan.nama as knama,departemen.keterangan as kdep,karyawan.id as kid,karyawan.alamat as kalamat,karyawan.kontak as kkontak" & _
                      " From karyawan INNER JOIN departemen_karyawan ON karyawan.id = departemen_karyawan.id" & _
                      " INNER JOIN departemen ON departemen_karyawan.departemen_id = departemen.id ORDER by kid ASC LIMIT " & A, _
                        Conn, adOpenForwardOnly, adLockReadOnly
        Call TampilGrid
    Else
        Set RSKaryawan = New ADODB.Recordset

        RSKaryawan.Open "SELECT karyawan.nama as knama,departemen.keterangan as kdep,karyawan.id as kid,karyawan.alamat as kalamat,karyawan.kontak as kkontak" & _
                      " From karyawan INNER JOIN departemen_karyawan ON karyawan.id = departemen_karyawan.id" & _
                      " INNER JOIN departemen ON departemen_karyawan.departemen_id = departemen.id where departemen.keterangan='" & ComboDepartemen.Text & "' ORDER by kid ASC LIMIT " & A, _
                        Conn, adOpenForwardOnly, adLockReadOnly
        Call TampilGrid
    End If
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
    Call BarisLunasWarna
    Call initCombo
End Sub

Private Sub initCombo()
    Set RSDepartemen = New ADODB.Recordset
    Dim Combo1 As String
    Combo1 = ComboListSingle(ComboDepartemen, RSDepartemen, "departemen", "*", "order by keterangan asc", "keterangan")
End Sub

Private Sub Form_Resize()
'On Error Resume Next
    Call AutoResize
End Sub

Public Sub AutoResize()
    shpBar.Width = ScaleWidth

    With GridKaryawan
        .Top = shpBar.Height + 120
        .Left = 120
        .Width = FormMain.PictureBeranda.ScaleWidth - 240
        .Height = (FormMain.PictureBeranda.ScaleHeight - Picture1.ScaleHeight) - GridKaryawan.Top - 120
    End With

    Picture1.Move 0, GridKaryawan.Height + shpBar.Height + 240, GridKaryawan.Width + 240

    ButtonKeluar.Move Picture1.ScaleWidth - ButtonKeluar.Width - 180, 90

    For i = 0 To mKaryawan.Count - 1
        mKaryawan(i).Move 180 + (i * 1875) + (i * 120), 90, 1875
    Next

    ComboDepartemen.Move FormMain.PictureBeranda.Width - ComboDepartemen.Width - 360, 140
    LabelFilter.Move FormMain.PictureBeranda.Width - ComboDepartemen.Width - 360 - LabelFilter.Width - 360, 180
End Sub

Private Sub BarisLunasWarna()

    For i = 1 To GridKaryawan.Rows - 1

        If Left(GridKaryawan.TextMatrix(i, 2), 1) = "W" Then
            GridKaryawan.Cell(flexcpBackColor, i, 2) = &H80FFFF
        End If

        If Left(GridKaryawan.TextMatrix(i, 2), 1) = "H" Then
            GridKaryawan.Cell(flexcpBackColor, i, 2) = &H80FF80
        End If

        If Left(GridKaryawan.TextMatrix(i, 2), 1) = "C" Then
            GridKaryawan.Cell(flexcpBackColor, i, 2) = RGB(188, 63, 67)    '&H80FFFF
        End If

        If Left(GridKaryawan.TextMatrix(i, 2), 1) = "K" Then
            GridKaryawan.Cell(flexcpBackColor, i, 2) = RGB(111, 146, 200)    '&H80FF80
        End If

        If Left(GridKaryawan.TextMatrix(i, 2), 1) = "A" Then
            GridKaryawan.Cell(flexcpBackColor, i, 2) = RGB(253, 206, 5)    '&H80FF80
        End If

    Next i

End Sub

Private Sub GridKaryawan_Click()
    blnPilih = True
End Sub

Private Sub GridKaryawan_DblClick()
    Dim MyRS As New ADODB.Recordset
    Dim MyRS2 As New ADODB.Recordset
    Dim MyRS3 As New ADODB.Recordset
    Set MyRS = New ADODB.Recordset
    Set MyRS2 = New ADODB.Recordset
    Set MyRS3 = New ADODB.Recordset
    MyRS.Open "SELECT * FROM karyawan WHERE id LIKE '" & GridKaryawan.TextMatrix(GridKaryawan.Row, 0) & "%'", Conn, adOpenDynamic, adLockOptimistic
    MyRS2.Open "SELECT * FROM departemen_karyawan WHERE id LIKE '" & GridKaryawan.TextMatrix(GridKaryawan.Row, 0) & "%'", Conn, adOpenDynamic, adLockOptimistic
    MyRS3.Open "SELECT * FROM departemen WHERE id LIKE '" & MyRS2.Fields("departemen_id") & "%'", Conn, adOpenDynamic, adLockOptimistic
    If GridKaryawan.Rows > 1 Then
        With Add_Karyawan
            blnTambah = False
            .Caption = "Edit Data Karyawan"
            .TextKodeLama.Text = MyRS.Fields("id")
            .TextNama.Text = MyRS.Fields("nama")
            .TextAlamat.Text = MyRS.Fields("alamat")
            .TextKontak.Text = MyRS.Fields("kontak")
            .TextStatus.Text = MyRS.Fields("status")
            .TextHK.Text = MyRS.Fields("basic_hk")
            .TextGaji.Text = MyRS.Fields("basic_gaji")
            .TextGaji.Alignment = 1
            .TextGaji.Text = Format(.TextGaji.Text, "#,##0")
            .ComboDepartemen.Text = MyRS3.Fields("id") & " - " & MyRS3.Fields("keterangan")
            .ComboDepartemen.Enabled = False
            .Show vbModal
        End With
    End If
End Sub

Private Sub GridKaryawan_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    GridKaryawan.Redraw = False
End Sub

Private Sub GridKaryawan_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    GridKaryawan.RowSel = GridKaryawan.Row
    GridKaryawan.Redraw = True
End Sub

Private Sub mKaryawan_Click(Index As Integer)
On Error GoTo errHandle
    Select Case Index
    Case 0
        With Add_Karyawan
            blnTambah = True
            .Caption = "Input Data Karyawan"
            .Show vbModal
        End With
    Case 1
        GridKaryawan_DblClick
    Case 2

        If blnPilih = False Then
            MsgBox "Data Karyawan Belum Dipilih", vbInformation, "Informasi"
            Exit Sub
        End If

        Pesan_Peringatan "Question", "Apakah Data Karyawan dengan Nama " & Chr(34) & GridKaryawan.TextMatrix(GridKaryawan.Row, 1) & Chr(34) & " Ingin dihapus ?", "Konfirmasi"
        If Respon = "Iya" Then
            Set ClassKaryawan = New ClassKaryawan
            Sukses = ClassKaryawan.DeleteKaryawan(GridKaryawan.TextMatrix(GridKaryawan.Row, 0), Conn)
            If Sukses = True Then
                FormKaryawan.TampilanAwal
            Else
                MsgBox "Data Karyawan Gagal Dihapus", vbExclamation, "Peringatan"
            End If
        End If

    Case 3
        Call Form_Diatas(FormDepartemen, FormMain.PictureBeranda)
    End Select
    
Exit Sub
errHandle:
    MsgBox err.Description & vbCrLf & vbCrLf & strSQL
End Sub

Private Sub Picture1_Resize()
    Picture2.Left = Picture1.ScaleWidth - Picture2.ScaleWidth - ButtonKeluar.Width - 360
End Sub
