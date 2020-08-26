VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form FormPengeluaran 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Data Pengeluaran"
   ClientHeight    =   7095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12960
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormPengeluaran.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7095
   ScaleWidth      =   12960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox FrameSetting 
      Appearance      =   0  'Flat
      BackColor       =   &H004F4736&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   9000
      ScaleHeight     =   1695
      ScaleWidth      =   3660
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   3660
      Begin VB.OptionButton optSemua 
         Appearance      =   0  'Flat
         BackColor       =   &H004F4736&
         Caption         =   "&Semua"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   300
         TabIndex        =   23
         Top             =   1200
         Width           =   1035
      End
      Begin VB.OptionButton optBulan 
         Appearance      =   0  'Flat
         BackColor       =   &H004F4736&
         Caption         =   "Perbulan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   300
         TabIndex        =   22
         Top             =   900
         Width           =   1035
      End
      Begin VB.OptionButton optHari 
         Appearance      =   0  'Flat
         BackColor       =   &H004F4736&
         Caption         =   "&Hari Ini"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   300
         TabIndex        =   21
         Top             =   600
         Width           =   1035
      End
      Begin VB.ComboBox cmbBulan 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "FormPengeluaran.frx":000C
         Left            =   1380
         List            =   "FormPengeluaran.frx":0034
         TabIndex        =   20
         Top             =   840
         Width           =   1995
      End
      Begin Project1.N_Image ButtonClose 
         Height          =   600
         Left            =   3060
         TabIndex        =   18
         Top             =   0
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   1058
         Picture         =   "FormPengeluaran.frx":00D7
         PictureHover    =   "FormPengeluaran.frx":13E9
         PictureDown     =   "FormPengeluaran.frx":26FB
      End
      Begin VB.Label Label_Transaksi 
         AutoSize        =   -1  'True
         BackColor       =   &H002D261E&
         BackStyle       =   0  'Transparent
         Caption         =   "TAMPILKAN TRANSAKSI"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00969089&
         Height          =   240
         Left            =   180
         TabIndex        =   19
         Top             =   180
         Width           =   2325
      End
   End
   Begin Project1.N_Image ButtonSetting 
      Height          =   600
      Left            =   12300
      TabIndex        =   16
      Top             =   0
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   1058
      Picture         =   "FormPengeluaran.frx":3A0D
      PictureHover    =   "FormPengeluaran.frx":4D1F
      PictureDown     =   "FormPengeluaran.frx":6031
   End
   Begin VB.TextBox txtCari 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6600
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00D9FBDB&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6660
         ScaleHeight     =   345
         ScaleWidth      =   4155
         TabIndex        =   1
         Top             =   120
         Width           =   4150
         Begin VB.CommandButton btnNext 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3450
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Next 250"
            Top             =   15
            Width           =   315
         End
         Begin VB.CommandButton btnLast 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3765
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Last 250"
            Top             =   15
            Width           =   315
         End
         Begin VB.CommandButton btnPrev 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3135
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Previous 250"
            Top             =   15
            Width           =   315
         End
         Begin VB.CommandButton btnFirst 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
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
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
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
         TabIndex        =   8
         Top             =   0
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   767
         Icon            =   "FormPengeluaran.frx":7343
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
         TabIndex        =   9
         Top             =   0
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   767
         Icon            =   "FormPengeluaran.frx":7D55
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
         TabIndex        =   10
         Top             =   0
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   767
         Icon            =   "FormPengeluaran.frx":8767
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
   Begin VSFlex8Ctl.VSFlexGrid GridPengeluaran 
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
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FormPengeluaran.frx":9279
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
   Begin VB.Label LabelTotal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "%TotalPengeluaran%"
      Height          =   195
      Left            =   4800
      TabIndex        =   15
      Top             =   180
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "DATA PENGELUARAN"
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kelola Informasi Data Pengeluaran Yang Tersedia."
      Height          =   195
      Left            =   195
      TabIndex        =   13
      Top             =   300
      Width           =   3600
   End
   Begin VB.Shape shpBar 
      BackColor       =   &H00D9FBDB&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   600
      Left            =   0
      Top             =   0
      Width           =   12915
   End
End
Attribute VB_Name = "FormPengeluaran"
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
Dim ClassPengeluaran As ClassPengeluaran
Dim ClassKas As ClassKas

'Menampilkan Data Ke GridPengeluaran
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
                GridPengeluaran.Rows = Baris
                GridPengeluaran.AddItem _
                        Format(RSPengeluaran!Tanggal, "dd/mm/yyyy") & vbTab & _
                                                                    RSPengeluaran!Id & vbTab & _
                                                                    RSPengeluaran!Kd_Biaya & vbTab & _
                                                                    RSPengeluaran!Keterangan & vbTab & _
                                                                    RSPengeluaran!Kode_Kas & vbTab & _
                                                                    Format(RSPengeluaran!Jumlah, "#,##0") & vbTab & _
                                                                    RSPengeluaran!Operator
                .MoveNext
            Loop
            'Call BarisStokWarna
            GridPengeluaran.Select 1, 1, 1, GridPengeluaran.Cols - 1
            Hitung
        End With
    End If
End Sub

Private Sub Hitung()
    Dim jum(3), nil(3) As Long
    jum(0) = 0
    jum(1) = 0
    jum(2) = 0

    With GridPengeluaran
        For i = 0 To GridPengeluaran.Rows - 1
            If IsNumeric(.TextMatrix(i, 5)) Then
                nil(0) = .TextMatrix(i, 5)
            Else
                nil(0) = Val(.TextMatrix(i, 5))
            End If
            jum(0) = jum(0) + nil(0)

        Next
        LabelTotal.Caption = Format(jum(0), "Rp ###,##,0")
    End With
End Sub

Sub TampilanAwal()
    x = 0
    y = 75
    A = x & "," & y
    Hal = 1

    If optSemua.Value = True Then
        'Menampilkan Data ke Grid
        Set RSPengeluaran = New ADODB.Recordset
        RSPengeluaran.Open "SELECT id,tanggal,kd_biaya,keterangan,jumlah,operator,kode_kas FROM pengeluaran ORDER BY tanggal ASC LIMIT " & A, _
                           Conn, adOpenForwardOnly, adLockReadOnly
        Call TampilGrid

        'Menentukan Jml Halaman
        Set RSPengeluaran = New ADODB.Recordset
        RSPengeluaran.Open "SELECT id,tanggal,kd_biaya,keterangan,jumlah,operator,kode_kas FROM pengeluaran ORDER BY tanggal ASC", Conn, adOpenForwardOnly, adLockReadOnly
        i = RSPengeluaran.RecordCount / y
    ElseIf optHari.Value = True Then
        Set RSPengeluaran = New ADODB.Recordset
        RSPengeluaran.Open "SELECT id,tanggal,kd_biaya,keterangan,jumlah,operator,kode_kas FROM pengeluaran WHERE YEAR(tanggal)=" & Format(Now, "yyyy") & " and MONTH(tanggal) =" & Format(Now, "mm") & " and DAY(tanggal) = " & Format(Now, "dd") & " ORDER BY tanggal ASC LIMIT " & A, _
                           Conn, adOpenForwardOnly, adLockReadOnly
        Call TampilGrid

        'Menentukan Jml Halaman
        Set RSPengeluaran = New ADODB.Recordset
        RSPengeluaran.Open "SELECT id,tanggal,kd_biaya,keterangan,jumlah,operator,kode_kas FROM pengeluaran WHERE YEAR(tanggal)=" & Format(Now, "yyyy") & " and MONTH(tanggal) =" & Format(Now, "mm") & " and DAY(tanggal) = " & Format(Now, "dd") & " ORDER BY tanggal ASC", Conn, adOpenForwardOnly, adLockReadOnly
        i = RSPengeluaran.RecordCount / y
    ElseIf OptBulan.Value = True Then
        Set RSPengeluaran = New ADODB.Recordset
        RSPengeluaran.Open "SELECT id,tanggal,kd_biaya,keterangan,jumlah,operator,kode_kas FROM pengeluaran WHERE YEAR(tanggal)=" & Format(Now, "yyyy") & " and MONTH(tanggal) =" & Left(cmbBulan.Text, 2) & " ORDER BY tanggal ASC LIMIT " & A, _
                           Conn, adOpenForwardOnly, adLockReadOnly
        Call TampilGrid

        'Menentukan Jml Halaman
        Set RSPengeluaran = New ADODB.Recordset
        RSPengeluaran.Open "SELECT id,tanggal,kd_biaya,keterangan,jumlah,operator,kode_kas FROM pengeluaran WHERE YEAR(tanggal)=" & Format(Now, "yyyy") & " and MONTH(tanggal) =" & Format(Now, "mm") & " ORDER BY tanggal ASC", Conn, adOpenForwardOnly, adLockReadOnly
        i = RSPengeluaran.RecordCount / y
    End If

    TotHal = Int(i) + IIf(CDbl(CStr(i - Int(i))) > 0.1, 1, 0)
    If TotHal = 0 Then TotHal = 1
    lblPageInfo.Caption = "Page " & Hal & " of " & TotHal & " Pages"
    '    HighlightRowG 1, GridPengeluaran
End Sub

Private Sub btnFirst_Click()
    x = 0
    y = y
    A = x & "," & y

    If optSemua.Value = True Then
        Set RSPengeluaran = New ADODB.Recordset
        RSPengeluaran.Open "SELECT id,tanggal,kd_biaya,keterangan,jumlah,operator,kode_kas FROM pengeluaran ORDER BY tanggal ASC  LIMIT " & A, Conn, adOpenForwardOnly, adLockReadOnly
        Call TampilGrid
    ElseIf optHari.Value = True Then
        Set RSPengeluaran = New ADODB.Recordset
        RSPengeluaran.Open "SELECT id,tanggal,kd_biaya,keterangan,jumlah,operator,kode_kas FROM pengeluaran WHERE YEAR(tanggal)=" & Format(Now, "yyyy") & " and MONTH(tanggal) =" & Format(Now, "mm") & " and DAY(tanggal) = " & Format(Now, "dd") & " ORDER BY tanggal ASC  LIMIT " & A, Conn, adOpenForwardOnly, adLockReadOnly
        Call TampilGrid
    ElseIf OptBulan.Value = True Then
        Set RSPengeluaran = New ADODB.Recordset
        RSPengeluaran.Open "SELECT id,tanggal,kd_biaya,keterangan,jumlah,operator,kode_kas FROM pengeluaran WHERE YEAR(tanggal)=" & Format(Now, "yyyy") & " and MONTH(tanggal) =" & Format(Now, "mm") & " ORDER BY tanggal ASC  LIMIT " & A, Conn, adOpenForwardOnly, adLockReadOnly
        Call TampilGrid
    End If

    Hal = 1
    lblPageInfo.Caption = "Page " & Hal & " of " & TotHal & " Pages"
End Sub

Private Sub btnLast_Click()
    x = (TotHal - 1) * y
    y = y
    A = x & "," & y

    If optSemua.Value = True Then
        Set RSPengeluaran = New ADODB.Recordset
        RSPengeluaran.Open "SELECT id,tanggal,kd_biaya,keterangan,jumlah,operator,kode_kas FROM pengeluaran ORDER BY tanggal ASC  LIMIT " & A, Conn, adOpenForwardOnly, adLockReadOnly
        Call TampilGrid
    ElseIf optHari.Value = True Then
        Set RSPengeluaran = New ADODB.Recordset
        RSPengeluaran.Open "SELECT id,tanggal,kd_biaya,keterangan,jumlah,operator,kode_kas FROM pengeluaran WHERE YEAR(tanggal)=" & Format(Now, "yyyy") & " and MONTH(tanggal) =" & Format(Now, "mm") & " and DAY(tanggal) = " & Format(Now, "dd") & " ORDER BY tanggal ASC  LIMIT " & A, Conn, adOpenForwardOnly, adLockReadOnly
        Call TampilGrid
    ElseIf OptBulan.Value = True Then
        Set RSPengeluaran = New ADODB.Recordset
        RSPengeluaran.Open "SELECT id,tanggal,kd_biaya,keterangan,jumlah,operator,kode_kas FROM pengeluaran WHERE YEAR(tanggal)=" & Format(Now, "yyyy") & " and MONTH(tanggal) =" & Format(Now, "mm") & " ORDER BY tanggal ASC  LIMIT " & A, Conn, adOpenForwardOnly, adLockReadOnly
        Call TampilGrid
    End If

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

        If optSemua.Value = True Then
            Set RSPengeluaran = New ADODB.Recordset
            RSPengeluaran.Open "SELECT id,tanggal,kd_biaya,keterangan,jumlah,operator,kode_kas FROM pengeluaran ORDER BY tanggal ASC  LIMIT " & A, Conn, adOpenForwardOnly, adLockReadOnly
            Call TampilGrid
        ElseIf optHari.Value = True Then
            Set RSPengeluaran = New ADODB.Recordset
            RSPengeluaran.Open "SELECT id,tanggal,kd_biaya,keterangan,jumlah,operator,kode_kas FROM pengeluaran WHERE YEAR(tanggal)=" & Format(Now, "yyyy") & " and MONTH(tanggal) =" & Format(Now, "mm") & " and DAY(tanggal) = " & Format(Now, "dd") & " ORDER BY tanggal ASC  LIMIT " & A, Conn, adOpenForwardOnly, adLockReadOnly
            Call TampilGrid
        ElseIf OptBulan.Value = True Then
            Set RSPengeluaran = New ADODB.Recordset
            RSPengeluaran.Open "SELECT id,tanggal,kd_biaya,keterangan,jumlah,operator,kode_kas FROM pengeluaran WHERE YEAR(tanggal)=" & Format(Now, "yyyy") & " and MONTH(tanggal) =" & Format(Now, "mm") & " ORDER BY tanggal ASC  LIMIT " & A, Conn, adOpenForwardOnly, adLockReadOnly
            Call TampilGrid
        End If

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

        If optSemua.Value = True Then
            Set RSPengeluaran = New ADODB.Recordset
            RSPengeluaran.Open "SELECT id,tanggal,kd_biaya,keterangan,jumlah,operator,kode_kas FROM pengeluaran ORDER BY tanggal ASC  LIMIT " & A, Conn, adOpenForwardOnly, adLockReadOnly
            Call TampilGrid
        ElseIf optHari.Value = True Then
            Set RSPengeluaran = New ADODB.Recordset
            RSPengeluaran.Open "SELECT id,tanggal,kd_biaya,keterangan,jumlah,operator,kode_kas FROM pengeluaran WHERE YEAR(tanggal)=" & Format(Now, "yyyy") & " and MONTH(tanggal) =" & Format(Now, "mm") & " and DAY(tanggal) = " & Format(Now, "dd") & " ORDER BY tanggal ASC  LIMIT " & A, Conn, adOpenForwardOnly, adLockReadOnly
            Call TampilGrid
        ElseIf OptBulan.Value = True Then
            Set RSPengeluaran = New ADODB.Recordset
            RSPengeluaran.Open "SELECT id,tanggal,kd_biaya,keterangan,jumlah,operator,kode_kas FROM pengeluaran WHERE YEAR(tanggal)=" & Format(Now, "yyyy") & " and MONTH(tanggal) =" & Format(Now, "mm") & " ORDER BY tanggal ASC  LIMIT " & A, Conn, adOpenForwardOnly, adLockReadOnly
            Call TampilGrid
        End If

        Hal = Hal - 1
        lblPageInfo.Caption = "Page " & Hal & " of " & TotHal & " Pages"
    End If
End Sub

Private Sub ButtonClose_Click()
    FrameSetting.Visible = False
End Sub

Private Sub ButtonKeluar_Click()
    blnPilih = False
    Unload Me
End Sub

Private Sub ButtonSetting_Click()
    FrameSetting.Visible = True
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

    Call TampilkanPengeluaran

End Sub

Private Sub TampilkanPengeluaran()
    Dim Tampilkan As String: Tampilkan = ReadINI("Tampil", "Pengeluaran", Lokasi_File_Konfigurasi)
    If Tampilkan = "Hari" Then
        optHari_Click
        optHari.Value = True
    ElseIf Tampilkan = "Bulan" Then
        optBulan_Click
        OptBulan.Value = True
        cmbBulan.Text = ReadINI("Tampil", "Bulan", Lokasi_File_Konfigurasi)
        Set RSPengeluaran = New ADODB.Recordset
        RSPengeluaran.Open "SELECT id,tanggal,kd_biaya,keterangan,jumlah,operator,kode_kas FROM pengeluaran WHERE YEAR(tanggal)= " & Format(Now, "yyyy") & " and MONTH(tanggal) =" & Left(cmbBulan.Text, 2) & " ORDER BY tanggal ASC LIMIT " & A, Conn, adOpenForwardOnly, adLockReadOnly
        Call TampilGrid
    ElseIf Tampilkan = "Semua" Then
        optSemua_Click
        optSemua.Value = True
    Else
        optSemua_Click
        optSemua.Value = True
    End If
End Sub

Private Sub Form_Resize()
'On Error Resume Next
    Call AutoResize
End Sub

Public Sub AutoResize()
    shpBar.Width = ScaleWidth

    With ButtonSetting
        .Top = 0
        .Left = FormMain.PictureBeranda.ScaleWidth - .Width
    End With

    With FrameSetting
        .Top = 0
        .Left = FormMain.PictureBeranda.ScaleWidth - .Width
    End With

    With GridPengeluaran
        .Top = shpBar.Height + 120
        .Left = 120
        .Width = FormMain.PictureBeranda.ScaleWidth - 240
        .Height = (FormMain.PictureBeranda.ScaleHeight - Picture1.ScaleHeight) - GridPengeluaran.Top - 120
    End With

    Picture1.Move 0, GridPengeluaran.Height + shpBar.Height + 240, GridPengeluaran.Width + 240

    ButtonKeluar.Move Picture1.ScaleWidth - ButtonKeluar.Width - 180, 90

    For i = 0 To mPengeluaran.Count - 1
        mPengeluaran(i).Move 180 + (i * 1875) + (i * 120), 90, 1875
    Next
End Sub

Private Sub GridPengeluaran_Click()
    blnPilih = True
End Sub

Private Sub GridPengeluaran_DblClick()
    Dim MyRS As New ADODB.Recordset
    Dim MyRS2 As New ADODB.Recordset
    Dim MyRS3 As New ADODB.Recordset
    Set MyRS = New ADODB.Recordset
    Set MyRS2 = New ADODB.Recordset
    Set MyRS3 = New ADODB.Recordset
    MyRS.Open "select * from pengeluaran Where id Like '" & GridPengeluaran.TextMatrix(GridPengeluaran.Row, 1) & "%'", Conn, adOpenDynamic, adLockOptimistic
    MyRS2.Open "select * from biaya Where kode Like '" & GridPengeluaran.TextMatrix(GridPengeluaran.Row, 2) & "%'", Conn, adOpenDynamic, adLockOptimistic
    MyRS3.Open "select * from kas Where kode Like '" & GridPengeluaran.TextMatrix(GridPengeluaran.Row, 4) & "%'", Conn, adOpenDynamic, adLockOptimistic
    If GridPengeluaran.Rows > 1 Then
        With Add_Pengeluaran
            blnTambah = False
            .Caption = "Edit Data Pengeluaran"
            .txtKode.Text = MyRS.Fields("id")
            .TextTanggal.Text = Format(MyRS.Fields("tanggal"), "dd/mm/yyyy")
            .ComboBiaya.Text = MyRS2.Fields("nama")
            .LabelBiaya.Caption = MyRS.Fields("kd_biaya")
            .TextKeterangan.Text = MyRS.Fields("keterangan")
            .ComboKas.Text = MyRS3.Fields("nama")
            .LabelJumlah.Caption = Format(Val(MyRS3.Fields("saldo")), "Rp ###,##,0")
            .LabelKas.Caption = MyRS.Fields("kode_kas")
            .TextJumlahAwal = MyRS.Fields("jumlah")
            .TextJumlah.Text = Format(MyRS.Fields("jumlah"), "#,##0")
            .LabelJumlahKas.Caption = Val(Format(MyRS3.Fields("saldo"), "###0")) + Val(Format(MyRS.Fields("jumlah"), "###0"))
            .LabelJumlah.Visible = True
            .TextJumlah.Alignment = 1
            .TextJumlah.Text = Format(.TextJumlah.Text, "#,##0")
            .Show vbModal
        End With
    End If
End Sub

Private Sub GridPengeluaran_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    GridPengeluaran.Redraw = False
End Sub

Private Sub GridPengeluaran_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    GridPengeluaran.RowSel = GridPengeluaran.Row
    GridPengeluaran.Redraw = True
End Sub

Private Sub BarisStokWarna()

    For i = 1 To GridPengeluaran.Rows - 1

        If Val(Left(GridPengeluaran.TextMatrix(i, 4), 1)) <= "1" Then
            GridPengeluaran.Cell(flexcpBackColor, i, 4) = &H80FFFF
        End If

        If Val(Left(GridPengeluaran.TextMatrix(i, 4), 1)) >= "1" Then
            GridPengeluaran.Cell(flexcpBackColor, i, 4) = &H80FF80
        End If

    Next i

End Sub

Private Sub mPengeluaran_Click(Index As Integer)
    Select Case Index
    Case 0
        With Add_Pengeluaran
            blnTambah = True
            .Caption = "Input Data Pengeluaran"
            .Show vbModal
        End With
    Case 1
        GridPengeluaran_DblClick
    Case 2
        If blnPilih = False Then
            MsgBox "Data Kas belum dipilih", vbInformation, "Informasi"
            Exit Sub
        End If

        Pesan_Peringatan "Question", "Apakah Data Kas dengan Kode " & Chr(34) & GridPengeluaran.TextMatrix(GridPengeluaran.Row, 1) & Chr(34) & " Ingin dihapus ?", "Konfirmasi"
        If Respon = "Iya" Then
            Set ClassPengeluaran = New ClassPengeluaran
            Set ClassKas = New ClassKas
            Sukses = ClassPengeluaran.DeletePengeluaran(GridPengeluaran.TextMatrix(GridPengeluaran.Row, 1), Conn)
            Sukses = ClassKas.UpdateJumlahKasHapus(GridPengeluaran.TextMatrix(GridPengeluaran.Row, 4), GridPengeluaran.TextMatrix(GridPengeluaran.Row, 5), Conn)
            If Sukses = True Then
                FormPengeluaran.TampilanAwal
                FormKas.TampilanAwal
            Else
                MsgBox "Data Kas gagal dihapus", vbExclamation, "Peringatan"
            End If
        End If
    End Select
End Sub

Private Sub optBulan_Click()
    If OptBulan.Value = True Then
        cmbBulan.Enabled = True
    ElseIf OptBulan.Value = False Then
        cmbBulan.Enabled = False
    End If

    Call WriteINI("Tampil", "Pengeluaran", "Bulan", Lokasi_File_Konfigurasi)
End Sub

Private Sub cmbBulan_Click()
    Set RSPengeluaran = New ADODB.Recordset
    RSPengeluaran.Open "SELECT id,tanggal,kd_biaya,keterangan,jumlah,operator,kode_kas FROM pengeluaran WHERE YEAR(tanggal)=" & Format(Now, "yyyy") & " and MONTH(tanggal) =" & Left(cmbBulan.Text, 2) & " ORDER BY tanggal ASC LIMIT " & A, _
                       Conn, adOpenForwardOnly, adLockReadOnly
    Call TampilGrid

    Call WriteINI("Tampil", "Bulan", cmbBulan.Text, Lokasi_File_Konfigurasi)
End Sub

Private Sub optHari_Click()
    If optHari.Value = True Then
        cmbBulan.Enabled = False
    ElseIf optHari.Value = False Then
        cmbBulan.Enabled = False
    End If

    Set RSPengeluaran = New ADODB.Recordset
    RSPengeluaran.Open "SELECT id,tanggal,kd_biaya,keterangan,jumlah,operator,kode_kas FROM pengeluaran WHERE YEAR(tanggal)=" & Format(Now, "yyyy") & " and MONTH(tanggal) =" & Format(Now, "mm") & " and DAY(tanggal) = " & Format(Now, "dd") & " ORDER BY tanggal ASC LIMIT " & A, _
                       Conn, adOpenForwardOnly, adLockReadOnly
    Call TampilGrid

    Call WriteINI("Tampil", "Pengeluaran", "Hari", Lokasi_File_Konfigurasi)
End Sub

Private Sub optSemua_Click()
    If optSemua.Value = True Then
        cmbBulan.Enabled = False
    ElseIf optSemua.Value = False Then
        cmbBulan.Enabled = False
    End If

    Set RSPengeluaran = New ADODB.Recordset
    RSPengeluaran.Open "SELECT id,tanggal,kd_biaya,keterangan,jumlah,operator,kode_kas FROM pengeluaran ORDER BY tanggal ASC LIMIT " & A, _
                       Conn, adOpenForwardOnly, adLockReadOnly
    Call TampilGrid

    Call WriteINI("Tampil", "Pengeluaran", "Semua", Lokasi_File_Konfigurasi)
End Sub

Private Sub Picture1_Resize()
    Picture2.Left = Picture1.ScaleWidth - Picture2.ScaleWidth - ButtonKeluar.Width - 360
End Sub
