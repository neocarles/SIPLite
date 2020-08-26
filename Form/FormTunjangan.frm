VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FormTunjangan 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8370
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14040
   Icon            =   "FormTunjangan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   14040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00D9FBDB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   180
      ScaleHeight     =   600
      ScaleWidth      =   12765
      TabIndex        =   3
      Top             =   7620
      Width           =   12765
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00D9FBDB&
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   1620
         ScaleHeight     =   345
         ScaleWidth      =   4155
         TabIndex        =   5
         Top             =   120
         Width           =   4150
         Begin VB.CommandButton btnNext 
            Height          =   315
            Left            =   3450
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Next 250"
            Top             =   15
            Width           =   315
         End
         Begin VB.CommandButton btnLast 
            Height          =   315
            Left            =   3765
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Last 250"
            Top             =   15
            Width           =   315
         End
         Begin VB.CommandButton btnPrev 
            Height          =   315
            Left            =   3135
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Previous 250"
            Top             =   15
            Width           =   315
         End
         Begin VB.CommandButton btnFirst 
            Height          =   315
            Left            =   2820
            Style           =   1  'Graphical
            TabIndex        =   6
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
            TabIndex        =   10
            Top             =   60
            Width           =   2535
         End
      End
      Begin Project1.isButton ButtonKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   11280
         TabIndex        =   4
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
   End
   Begin VB.TextBox TextPotongan 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   3900
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox ComboDepartemen 
      Height          =   315
      Left            =   6780
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   2355
   End
   Begin VSFlex8Ctl.VSFlexGrid GridTunjangan 
      Height          =   2325
      Left            =   6060
      TabIndex        =   0
      Top             =   5160
      Width           =   7875
      _cx             =   13891
      _cy             =   4101
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
      BackColorBkg    =   16777215
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
      Rows            =   1
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FormTunjangan.frx":000C
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
   Begin Project1.XPFrame FrameData 
      Height          =   4515
      Left            =   6060
      TabIndex        =   11
      Top             =   600
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   7964
      HeaderLightColor=   8421504
      HeaderDarkColor =   8421504
      TextColor       =   16777215
      Caption         =   "TAMBAH DATA TUNJANGAN"
      Curvature       =   0
      Begin VB.ComboBox CmbID 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1620
         TabIndex        =   16
         Top             =   1380
         Width           =   2655
      End
      Begin VB.TextBox TxtNama 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1620
         TabIndex        =   15
         Top             =   1860
         Width           =   4335
      End
      Begin VB.ComboBox CmbKet 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1620
         TabIndex        =   14
         Top             =   2280
         Width           =   2655
      End
      Begin VB.TextBox TxtJumlah 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1620
         TabIndex        =   13
         Top             =   2760
         Width           =   4335
      End
      Begin VB.TextBox TxtKey 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   4380
         TabIndex        =   12
         Top             =   900
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker Tanggal 
         Height          =   315
         Left            =   1620
         TabIndex        =   17
         Top             =   900
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   149356545
         CurrentDate     =   42991
      End
      Begin Project1.isButton mPembelian 
         Height          =   435
         Index           =   1
         Left            =   1920
         TabIndex        =   18
         Top             =   3360
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   767
         Icon            =   "FormTunjangan.frx":00B6
         Style           =   8
         Caption         =   "&Ubah"
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
      Begin Project1.isButton mPembelian 
         Height          =   435
         Index           =   2
         Left            =   3300
         TabIndex        =   19
         Top             =   3360
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   767
         Icon            =   "FormTunjangan.frx":0830
         Style           =   8
         Caption         =   "&Hapus"
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
      Begin Project1.isButton mPembelian 
         Height          =   435
         Index           =   0
         Left            =   540
         TabIndex        =   20
         Top             =   3360
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   767
         Icon            =   "FormTunjangan.frx":1242
         Style           =   8
         Caption         =   "&Tambah"
         CaptionAlign    =   2
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
      Begin Project1.isButton mPembelian 
         Height          =   435
         Index           =   3
         Left            =   4680
         TabIndex        =   21
         Top             =   3360
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   767
         Icon            =   "FormTunjangan.frx":1D54
         Style           =   8
         Caption         =   "&Batal"
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Input Data Tunjangan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   180
         TabIndex        =   29
         Top             =   480
         Width           =   1860
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
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
         Left            =   240
         TabIndex        =   28
         Top             =   2340
         Width           =   1155
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
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
         Left            =   240
         TabIndex        =   27
         Top             =   1920
         Width           =   1155
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ID Karyawan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   26
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   960
         Width           =   1155
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah (Rp)"
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
         Left            =   240
         TabIndex        =   24
         Top             =   2820
         Width           =   1155
      End
      Begin VB.Label LblID 
         BackStyle       =   0  'Transparent
         Height          =   315
         Left            =   4320
         TabIndex        =   23
         Top             =   1380
         Width           =   1635
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Pinjam"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   22
         Top             =   960
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000010&
         Height          =   255
         Left            =   120
         Top             =   480
         Width           =   6195
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid GridKaryawan 
      Height          =   6945
      Left            =   0
      TabIndex        =   30
      Top             =   600
      Width           =   5955
      _cx             =   10504
      _cy             =   12250
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
      BackColorBkg    =   16777215
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
      FormatString    =   $"FormTunjangan.frx":2766
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
      Caption         =   "DATA TUNJANGAN"
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
      TabIndex        =   33
      Top             =   60
      Width           =   4815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kelola Informasi Data Tunjangan Yang Tersedia."
      Height          =   195
      Left            =   195
      TabIndex        =   32
      Top             =   300
      Width           =   3435
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
      TabIndex        =   31
      Top             =   165
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
Attribute VB_Name = "FormTunjangan"
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

Dim RSKaryawan As New ADODB.Recordset
Dim RSTunjangan As New ADODB.Recordset
Dim RSDepartemen As New ADODB.Recordset
Dim RSID As New ADODB.Recordset
Dim RSKet As New ADODB.Recordset

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
                GridKaryawan.AddItem RSKaryawan!kid & vbTab & RSKaryawan!knama & vbTab & RSKaryawan!kdep
                .MoveNext
            Loop
            Call BarisLunasWarna
        End With
    End If
End Sub

Sub TampilGrid2()
'On Error Resume Next
    Baris = 0
    If RSTunjangan.EOF Then
        Exit Sub
    Else
        With RSTunjangan
            .MoveFirst
            Do Until .EOF
                Baris = Baris + 1
                GridTunjangan.Rows = Baris
                GridTunjangan.AddItem RSTunjangan!Id & vbTab & RSTunjangan!Tanggal & vbTab & Format(RSTunjangan!Jumlah, "#,##0") & vbTab & RSTunjangan!Keterangan & vbTab & RSTunjangan!Kode
                .MoveNext
            Loop
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

    RSKaryawan.Open "SELECT karyawan.nama as knama,departemen.keterangan as kdep,karyawan.id as kid" & _
                  " From karyawan INNER JOIN departemen_karyawan ON karyawan.id = departemen_karyawan.id" & _
                  " INNER JOIN departemen ON departemen_karyawan.departemen_id = departemen.id ORDER by kid ASC LIMIT " & A, _
                    Conn, adOpenForwardOnly, adLockReadOnly
    Call TampilGrid

    'Menentukan Jml Halaman
    Set RSKaryawan = New ADODB.Recordset

    RSKaryawan.Open "SELECT karyawan.nama as knama,departemen.keterangan as kdep,karyawan.id as kid" & _
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
    RSKaryawan.Open "SELECT karyawan.nama as knama,departemen.keterangan as kdep,karyawan.id as kid" & _
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
    RSKaryawan.Open "SELECT karyawan.nama as knama,departemen.keterangan as kdep,karyawan.id as kid" & _
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
        RSKaryawan.Open "SELECT karyawan.nama as knama,departemen.keterangan as kdep,karyawan.id as kid" & _
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
        RSKaryawan.Open "SELECT karyawan.nama as knama,departemen.keterangan as kdep,karyawan.id as kid" & _
                      " From karyawan INNER JOIN departemen_karyawan ON karyawan.id = departemen_karyawan.id" & _
                      " INNER JOIN departemen ON departemen_karyawan.departemen_id = departemen.id ORDER by kid ASC LIMIT " & A, _
                        Conn, adOpenForwardOnly, adLockReadOnly
        Call TampilGrid
        Hal = Hal - 1
        lblPageInfo.Caption = "Page " & Hal & " of " & TotHal & " Pages"
    End If
End Sub

Private Sub ButtonKeluar_Click()
    Unload Me
End Sub

Private Sub CmbID_Click()
    Set RSID = New ADODB.Recordset
    RSID.Open "SELECT * FROM karyawan WHERE id='" & Left(CmbID.Text, 5) & "'", _
              Conn, adOpenDynamic, adLockOptimistic
    RSID.Requery
    With RSID
        If .EOF And .BOF Then
            MsgBox "ID TIDAK DITEMUKAN", _
                   vbOKOnly + vbCritical, "Error"
            Exit Sub
        Else
            LblID.Caption = !Id
            TxtNama.Text = !Nama
        End If
    End With
    RSID.Close
    'MsgBox Right(CmbID.Text, (Len(CmbID.Text)) - 8)
    'LblID.Caption = Left(CmbID.Text, 5)
    Call FormatTxt
    Call AutoKey
End Sub

Private Sub CmbKet_Click()
    Call FormatTxt
End Sub

Private Sub ComboDepartemen_Click()
    If ComboDepartemen.Text = "ALL DEPARTEMEN" Then
        Set RSKaryawan = New ADODB.Recordset

        RSKaryawan.Open "SELECT karyawan.nama as knama,departemen.keterangan as kdep,karyawan.id as kid" & _
                      " From karyawan INNER JOIN departemen_karyawan ON karyawan.id = departemen_karyawan.id" & _
                      " INNER JOIN departemen ON departemen_karyawan.departemen_id = departemen.id ORDER by kid ASC LIMIT " & A, _
                        Conn, adOpenForwardOnly, adLockReadOnly
        Call TampilGrid
    Else
        Set RSKaryawan = New ADODB.Recordset

        RSKaryawan.Open "SELECT karyawan.nama as knama,departemen.keterangan as kdep,karyawan.id as kid" & _
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

    Tanggal.Value = Now
End Sub

Private Sub initCombo()
    Set RSDepartemen = New ADODB.Recordset
    Set RSID = New ADODB.Recordset
    Set RSKet = New ADODB.Recordset
    Dim Combo1 As String: Combo1 = ComboListSingle(ComboDepartemen, RSDepartemen, "departemen", "*", "order by keterangan asc", "keterangan")
    Dim Combo2 As String: Combo2 = ComboList(CmbID, RSID, "karyawan", "*", "order by id asc", "id", "nama")
    Dim Combo3 As String: Combo3 = ComboListSingle(CmbKet, RSKet, "jenis_tunjangan", "*", "order by nama desc", "nama")
End Sub

Private Sub Form_Resize()
'On Error Resume Next
    Call AutoResize
End Sub

Public Sub AutoResize()
'On Error Resume Next
    shpBar.Width = ScaleWidth

    With GridKaryawan
        .Left = 120
        .Top = shpBar.Height + 120
        .Height = (FormMain.PictureBeranda.ScaleHeight) - GridKaryawan.Top - Picture1.Height - 120
        .Width = 5955
    End With

    With FrameData
        .Left = GridKaryawan.Width + 240
        .Top = shpBar.Height + 120
        .Width = FormMain.PictureBeranda.ScaleWidth - GridKaryawan.Width - 360
    End With

    With GridTunjangan
        .Left = GridKaryawan.Width + 240
        .Top = FrameData.Top + FrameData.Height + 120
        .Width = FrameData.Width
        .Height = FormMain.PictureBeranda.ScaleHeight - shpBar.Height - FrameData.Height - Picture1.ScaleHeight - 360
    End With

    Picture1.Move 0, GridKaryawan.Height + shpBar.Height + 240, shpBar.Width

    ButtonKeluar.Move FormMain.PictureBeranda.ScaleWidth - ButtonKeluar.Width - 180, 90

    Picture2.Move GridKaryawan.Width - Picture2.Width + 120

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
    TextPotongan.Text = GridKaryawan.TextMatrix(GridKaryawan.Row, 0)
    mPembelian_Click (3)
End Sub

Private Sub GridKaryawan_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    GridKaryawan.Redraw = False
End Sub

Private Sub GridKaryawan_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    GridKaryawan.RowSel = GridKaryawan.Row
    GridKaryawan.Redraw = True
End Sub

Private Sub GridTunjangan_Click()
    Dim MyRS As New ADODB.Recordset
    Set MyRS = New ADODB.Recordset
    Dim MyRS2 As New ADODB.Recordset
    Set MyRS2 = New ADODB.Recordset
    MyRS.Open "select * from tunjangan Where kode Like '" & GridTunjangan.TextMatrix(GridTunjangan.Row, 4) & "%'", Conn, adOpenDynamic, adLockOptimistic
    MyRS2.Open "select * from karyawan Where id Like '" & GridTunjangan.TextMatrix(GridTunjangan.Row, 0) & "%'", Conn, adOpenDynamic, adLockOptimistic
    If GridTunjangan.Rows > 1 Then
        CmbID.Text = MyRS2!Id & " - " & MyRS2!Nama
        LblID.Caption = MyRS!Id
        TxtKey.Text = MyRS!Kode
        TxtNama.Text = MyRS2!Nama
        TxtJumlah.Text = MyRS!Jumlah
        CmbKet.Text = MyRS!Keterangan

        Tanggal.Enabled = True
        CmbID.Enabled = True
        CmbKet.Enabled = True
        TxtJumlah.Enabled = True
        mPembelian(0).Enabled = False
        mPembelian(1).Enabled = True
        mPembelian(2).Enabled = True
        mPembelian(3).Enabled = True
        mPembelian(1).Caption = "&Ubah"

        TxtJumlah.Alignment = 1
        TxtJumlah.Text = Format(TxtJumlah.Text, "#,##0")
    End If
End Sub

Private Sub mPembelian_Click(Index As Integer)
    Dim SQLTambah As String
    Dim SQLEdit As String
    Dim SQLHapus As String
    Select Case Index
    Case 0
        Tanggal.Enabled = True
        CmbID.Enabled = True
        CmbKet.Enabled = True
        TxtJumlah.Enabled = True
        mPembelian(0).Enabled = False
        mPembelian(1).Enabled = True
        mPembelian(2).Enabled = False
        mPembelian(3).Enabled = True
        mPembelian(1).Caption = "&Simpan"
    Case 1
        If mPembelian(1).Caption = "&Simpan" Then
            If TxtNama.Text = "" Or CmbKet.Text = "" Or TxtJumlah.Text = "" Then
                MsgBox "Data Belum Terisi Lengkap Mohon dilengkapi", vbInformation + vbOKOnly, "Lengkapi Data"
            Else
                SQLTambah = "INSERT INTO tunjangan VALUES ('" & LblID.Caption & "','" & Format(Tanggal.Value, "yyyy-mm-dd") & "','" & CmbKet.Text & "','" & Format(TxtJumlah.Text, "###0") & "','" & TxtKey.Text & "')"
                Conn.Execute SQLTambah

                Tanggal.Enabled = False
                CmbID.Enabled = False
                CmbKet.Enabled = False
                TxtJumlah.Enabled = False
                mPembelian(0).Enabled = True
                mPembelian(1).Enabled = False
                mPembelian(2).Enabled = False
                mPembelian(3).Enabled = False
                CmbID.Text = "-- Pilih ID Karyawan --"
                CmbKet.Text = "-- Pilih Keterangan --"
                TxtJumlah.Text = ""
                TxtNama.Text = ""
                TxtKey.Text = ""
                LblID.Caption = ""

                MsgBox "Data Berhasil Disimpan", vbInformation + vbOKOnly, "Sukses"
            End If
        ElseIf mPembelian(1).Caption = "&Ubah" Then
            If TxtNama.Text = "" Or CmbKet.Text = "" Or TxtJumlah.Text = "" Then
                MsgBox "Data Belum Terisi Lengkap Mohon dilengkapi", vbInformation + vbOKOnly, "Lengkapi Data"
            Else

                SQLEdit = "Update tunjangan Set id= '" & LblID.Caption & "', tanggal='" & Tanggal.Value & "', keterangan='" & CmbKet.Text & "',jumlah='" & Format(TxtJumlah.Text, "###0") & "' where kode='" & TxtKey.Text & "'"
                Conn.Execute SQLEdit

                Tanggal.Enabled = False
                CmbID.Enabled = False
                CmbKet.Enabled = False
                TxtJumlah.Enabled = False
                mPembelian(0).Enabled = True
                mPembelian(1).Enabled = False
                mPembelian(2).Enabled = False
                mPembelian(3).Enabled = False
                CmbID.Text = "-- Pilih ID Karyawan --"
                CmbKet.Text = "-- Pilih Keterangan --"
                TxtJumlah.Text = ""
                TxtNama.Text = ""
                TxtKey.Text = ""
                LblID.Caption = ""

                QueryDataTunjangan

                MsgBox "Data Berhasil Diperbaharui", vbInformation + vbOKOnly, "Sukses"
            End If
        End If
    Case 2
        On Error Resume Next
        Dim RSHapus As New ADODB.Recordset
        Dim pesan As String
        If RSHapus.State = 1 Then Set RSHapus = Nothing
        RSHapus.Open "SELECT * from [karyawan] where [kode]='" & TxtKey.Text & "'", Conn, adOpenDynamic, adLockPessimistic
        With RSHapus
            If Not .EOF Then
                pesan = MsgBox("Apakah Anda Ingin Menghapus Potongan Atas Nama " & TxtNama.Text & " Dengan ID " & TxtKey.Text & " ?", vbCritical + vbYesNo, "Hapus?")
                If pesan = vbYes Then
                    SQLHapus = "Delete From tunjangan where kode= '" & TxtKey.Text & "'"
                    Conn.Execute SQLHapus
                    QueryDataTunjangan
                End If
            Else
                MsgBox TxtNama.Text & " Sudah Dihapus !", vbExclamation + vbOKOnly
            End If
        End With
    Case 3
        Tanggal.Enabled = False
        CmbID.Enabled = False
        CmbKet.Enabled = False
        TxtJumlah.Enabled = False
        mPembelian(0).Enabled = True
        mPembelian(1).Enabled = False
        mPembelian(2).Enabled = False
        mPembelian(3).Enabled = False
        CmbID.Text = "-- Pilih ID Karyawan --"
        CmbKet.Text = "-- Pilih Keterangan --"
        TxtJumlah.Text = ""
        TxtNama.Text = ""
        TxtKey.Text = ""
        LblID.Caption = ""
    End Select
End Sub

Private Sub TextPotongan_Change()
    QueryDataTunjangan
End Sub

Sub QueryDataTunjangan()
    Set RSTunjangan = New ADODB.Recordset

    LookGridTunjangan

    RSTunjangan.Open "SELECT * FROM tunjangan WHERE id='" & TextPotongan.Text & "'", Conn, adOpenForwardOnly, adLockReadOnly
    RSTunjangan.Requery

    Call TampilGrid2
End Sub

Sub LookGridTunjangan()
    GridTunjangan.Clear
    GridTunjangan.Rows = 50
    GridTunjangan.Cols = 5
    '---------------
    GridTunjangan.Col = 0
    GridTunjangan.Row = 0
    GridTunjangan.Text = "ID KARYAWAN"
    'GridTunjangan.CellFontBold = True
    '
    GridTunjangan.Col = 1
    GridTunjangan.Row = 0
    GridTunjangan.Text = "TANGGAL"
    'GridTunjangan.CellFontBold = True
    '
    GridTunjangan.Col = 2
    GridTunjangan.Row = 0
    GridTunjangan.Text = "TOTAL"
    'GridTunjangan.CellFontBold = True
    '
    GridTunjangan.Col = 3
    GridTunjangan.Row = 0
    GridTunjangan.Text = "KETERANGAN"
    'GridTunjangan.CellFontBold = True
    '
    GridTunjangan.Col = 4
    GridTunjangan.Row = 0
    GridTunjangan.Text = "ID TUNJANGAN"
    'GridTunjangan.CellFontBold = True
    '
End Sub

Private Sub TxtJumlah_Click()
    TxtJumlah.Alignment = 0
    TxtJumlah.Text = Format(TxtJumlah.Text, "###0")
End Sub

Private Sub TxtJumlah_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then KeyAscii = 0
    If KeyAscii = 13 Then
        TxtJumlah.Alignment = 1
        TxtJumlah.Text = Format(TxtJumlah.Text, "#,##0")
    End If
End Sub

Sub FormatTxt()
    TxtJumlah.Alignment = 1
    TxtJumlah.Text = Format(TxtJumlah.Text, "#,##0")
End Sub

Private Sub TxtNama_Click()
    Call FormatTxt
End Sub

Private Sub AutoKey()
    Dim RSAuto As New ADODB.Recordset
    RSAuto.Open ("SELECT * FROM tunjangan WHERE kode in(select max(kode) from tunjangan)order by kode desc"), Conn
    RSAuto.Requery
    Dim Urut As String * 8
    Dim Hitung As Long
    With RSAuto
        If .EOF Then
            Urut = "00000001"
            TxtKey = Urut
        Else
            Hitung = Right(!Kode, 8) + 1
            Urut = Right("00000000" & Hitung, 8)
        End If
        TxtKey = Urut
    End With
End Sub

