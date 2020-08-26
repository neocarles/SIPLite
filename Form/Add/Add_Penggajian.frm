VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Add_Penggajian 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Input Data Gaji"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12075
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Add_Penggajian.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   12075
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FrameGaji 
      Appearance      =   0  'Flat
      Caption         =   "INFORMASI PENGGAJIAN"
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
      Height          =   7095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11835
      Begin MSComCtl2.DTPicker Tanggal 
         Height          =   315
         Left            =   9960
         TabIndex        =   39
         Top             =   360
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         _Version        =   393216
         Format          =   95092737
         CurrentDate     =   43008
      End
      Begin VB.TextBox TextGajiAkhir 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8865
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   6600
         Width           =   2715
      End
      Begin VB.TextBox TextPotong 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8400
         TabIndex        =   34
         Top             =   5400
         Width           =   1035
      End
      Begin VB.TextBox TextHKMiss 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7020
         TabIndex        =   31
         Top             =   5400
         Width           =   615
      End
      Begin VB.TextBox TextGajiPotong 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8865
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   5940
         Width           =   2715
      End
      Begin VB.TextBox TextAkhirHK 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9765
         TabIndex        =   27
         Top             =   5400
         Width           =   1815
      End
      Begin VB.TextBox TextTGaji 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   25
         Top             =   5940
         Width           =   2715
      End
      Begin VB.TextBox TextGPokok 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   23
         Top             =   5400
         Width           =   2715
      End
      Begin VB.TextBox TextTotTunjangan 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   10
         Top             =   4800
         Width           =   2715
      End
      Begin VB.TextBox TextTotPotongan 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8880
         TabIndex        =   9
         Top             =   4800
         Width           =   2715
      End
      Begin VB.TextBox TextGaji 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7860
         TabIndex        =   8
         Top             =   1800
         Width           =   2595
      End
      Begin VB.TextBox TextHK 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7860
         TabIndex        =   7
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox TextDepartemen 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7860
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   840
         Width           =   2595
      End
      Begin VB.TextBox TextStatus 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1920
         TabIndex        =   5
         Top             =   1800
         Width           =   2595
      End
      Begin VB.TextBox TextTanggal 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7860
         TabIndex        =   4
         Top             =   360
         Width           =   2115
      End
      Begin VB.ComboBox CmbID 
         Height          =   315
         Left            =   1920
         TabIndex        =   3
         Text            =   "-- Silahkan Pilih ID Karyawan --"
         Top             =   840
         Width           =   2655
      End
      Begin VB.TextBox TxtNama 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1920
         TabIndex        =   2
         Top             =   1320
         Width           =   3795
      End
      Begin VB.TextBox TextKode 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1920
         TabIndex        =   1
         Top             =   360
         Width           =   2595
      End
      Begin VSFlex8Ctl.VSFlexGrid GridPotongan 
         Height          =   2505
         Left            =   6000
         TabIndex        =   11
         Top             =   2220
         Width           =   5595
         _cx             =   9869
         _cy             =   4419
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
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"Add_Penggajian.frx":000C
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
      Begin VSFlex8Ctl.VSFlexGrid GridTunjangan 
         Height          =   2505
         Left            =   240
         TabIndex        =   12
         Top             =   2220
         Width           =   5595
         _cx             =   9869
         _cy             =   4419
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
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"Add_Penggajian.frx":00A0
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
      Begin VB.Label LabelPotong 
         AutoSize        =   -1  'True
         Caption         =   "%BasicPotongan%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   9360
         TabIndex        =   40
         Top             =   1380
         Width           =   1590
      End
      Begin VB.Label LblID 
         BackStyle       =   0  'Transparent
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
         Left            =   4620
         TabIndex        =   38
         Top             =   840
         Width           =   1635
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Gaji Akhir"
         Height          =   195
         Left            =   7980
         TabIndex        =   37
         Top             =   6660
         Width           =   675
      End
      Begin VB.Line Line2 
         X1              =   300
         X2              =   11580
         Y1              =   6420
         Y2              =   6420
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "="
         Height          =   195
         Left            =   9540
         TabIndex        =   35
         Top             =   5460
         Width           =   120
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "X"
         Height          =   195
         Left            =   8100
         TabIndex        =   33
         Top             =   5460
         Width           =   90
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "HK"
         Height          =   195
         Left            =   7680
         TabIndex        =   32
         Top             =   5460
         Width           =   195
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total Gaji Setelah Potongan"
         Height          =   195
         Left            =   6660
         TabIndex        =   30
         Top             =   6000
         Width           =   1995
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Potongan "
         Height          =   195
         Left            =   6255
         TabIndex        =   28
         Top             =   5460
         Width           =   735
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total Gaji"
         Height          =   195
         Left            =   2235
         TabIndex        =   26
         Top             =   6000
         Width           =   675
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Gaji Pokok"
         Height          =   195
         Left            =   2175
         TabIndex        =   24
         Top             =   5460
         Width           =   735
      End
      Begin VB.Line Line1 
         X1              =   300
         X2              =   11580
         Y1              =   5280
         Y2              =   5280
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total Tunjangan"
         Height          =   195
         Left            =   1740
         TabIndex        =   22
         Top             =   4860
         Width           =   1170
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total Potongan"
         Height          =   195
         Left            =   7560
         TabIndex        =   21
         Top             =   4860
         Width           =   1095
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Basic Gaji"
         Height          =   255
         Left            =   5940
         TabIndex        =   20
         Top             =   1860
         Width           =   1695
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Basic HK"
         Height          =   255
         Left            =   5940
         TabIndex        =   19
         Top             =   1380
         Width           =   1695
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Departemen"
         Height          =   255
         Left            =   5940
         TabIndex        =   18
         Top             =   900
         Width           =   1695
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Status Karyawan"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1860
         Width           =   1515
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Tanggal Penggajian"
         Height          =   255
         Left            =   5940
         TabIndex        =   16
         Top             =   420
         Width           =   1695
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Karyawan"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   1380
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ID Karyawan"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   900
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Kode Penggajian"
         Height          =   255
         Left            =   180
         TabIndex        =   13
         Top             =   420
         Width           =   1575
      End
   End
   Begin Project1.isButton ButtonBatal 
      Cancel          =   -1  'True
      Height          =   435
      Left            =   7140
      TabIndex        =   41
      Top             =   7440
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   767
      Style           =   8
      Caption         =   "&Batal [Esc]"
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   0
      RoundedBordersByTheme=   0   'False
   End
   Begin Project1.isButton ButtonSimpan 
      Height          =   435
      Left            =   2820
      TabIndex        =   42
      Top             =   7440
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   767
      Style           =   8
      Caption         =   "&Simpan [F8]"
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   0
      RoundedBordersByTheme=   0   'False
   End
End
Attribute VB_Name = "Add_Penggajian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RSTunjangan As New ADODB.Recordset
Dim RSPotongan As New ADODB.Recordset
Dim RSPenggajian As New ADODB.Recordset
Dim RSID As New ADODB.Recordset

Dim ClassPenggajian As ClassPenggajian

Private Sub ButtonBatal_Click()
    Unload Me
End Sub

Private Sub ButtonSimpan_Click()
    If blnTambah = True Then    'data baru
        If Trim(TxtNama.Text) = "" Or Trim(TextAkhirHK.Text) = "" Or CmbID.Text = "-- Silahkan Pilih ID Karyawan --" Or Trim(TextGajiAkhir.Text) = "" Or Trim(TextHKMiss.Text) = "" Then
            Beep
            Exit Sub
        End If

        Dim Rs As ADODB.Recordset

        strSQL = "SELECT kode, nama FROM karyawan_gaji WHERE kode=" & LblID.Caption & " and gaji_bulan = " & Format(Tanggal, "mm") & " ORDER BY kode"
        Set Rs = Conn.Execute(strSQL)
        If Not Rs.EOF Then
            MsgBox "Data Gaji Karyawan dengan nama " & Chr(34) & Rs!Nama & Chr(34) & " Telah Ada di Database!", vbExclamation, "Peringatan"
            Exit Sub
        End If

        If TextTotTunjangan.Text = "" Then
            TextTotTunjangan.Text = "Rp 0"
        End If
        If TextTotPotongan.Text = "" Then
            TextTotPotongan.Text = "Rp 0"
        End If

        Set ClassPenggajian = New ClassPenggajian
        Sukses = ClassPenggajian.AddPenggajian(TextKode.Text, LblID.Caption, TxtNama.Text, TextHK, TextGaji, TextHKMiss, _
                                               TextPotong.Text, TextAkhirHK, TextTotTunjangan, TextTotPotongan, TextGajiAkhir, _
                                               Tanggal.Value, Tanggal.Value, FormMain.LabelNama.Caption, _
                                               Conn)

        If Sukses Then
            FormPenggajian.TampilanAwal
            Unload Me
        Else
            MsgBox "Data Gaji Karyawan Mungkin Telah Dibuat", vbExclamation, "Peringatan"
        End If

    Else
        If Trim(TxtNama.Text) = "" Or Trim(TextAkhirHK.Text) = "" Or CmbID.Text = "-- Silahkan Pilih ID Karyawan --" Or Trim(TextGajiAkhir.Text) = "" Or Trim(TextHKMiss.Text) = "" Then
            Beep
            Exit Sub
        End If

        If TextTotTunjangan.Text = "" Then
            TextTotTunjangan.Text = "Rp 0"
        End If
        If TextTotPotongan.Text = "" Then
            TextTotPotongan.Text = "Rp 0"
        End If

        Set ClassPenggajian = New ClassPenggajian
        Sukses = ClassPenggajian.UpdatePenggajian(TextKode.Text, LblID.Caption, TxtNama.Text, TextHK, TextGaji, TextHKMiss, _
                                                  TextPotong.Text, TextAkhirHK, TextTotTunjangan, TextTotPotongan, TextGajiAkhir, _
                                                  Tanggal.Value, Tanggal.Value, FormMain.LabelNama.Caption, _
                                                  Conn)

        If Sukses Then
            FormPenggajian.TampilanAwal
            Unload Me
        Else
            MsgBox "Data Gaji Karyawan Gagal Disimpan", vbExclamation, "Peringatan"
        End If

    End If

    Set ClassPenggajian = Nothing
End Sub

Private Sub CmbID_Click()
    Set RSID = New ADODB.Recordset
    'RSID.Open "SELECT * FROM karyawan WHERE id='" & Left(CmbID.Text, 5) & "'", _
     '                 Conn, adOpenDynamic, adLockOptimistic
    RSID.Open "SELECT karyawan.nama as knama,departemen.keterangan as kdep,karyawan.id as kid,karyawan.status as kstatus,karyawan.Basic_Hk as kbasichk, karyawan.basic_gaji as kbasicgaji" & _
            " From karyawan INNER JOIN departemen_karyawan ON karyawan.id = departemen_karyawan.id" & _
            " INNER JOIN departemen ON departemen_karyawan.departemen_id = departemen.id WHERE karyawan.id='" & Left(CmbID.Text, 5) & "'", _
              Conn, adOpenForwardOnly, adLockReadOnly
    RSID.Requery
    With RSID
        If .EOF And .BOF Then
            MsgBox "ID TIDAK DITEMUKAN", _
                   vbOKOnly + vbCritical, "Error"
            Exit Sub
        Else
            LblID.Caption = !kid
            TxtNama.Text = !knama
            TextStatus.Text = !kstatus
            TextDepartemen.Text = !kdep
            TextHK.Text = !kbasichk

            TextHKMiss.Text = "0"
            TextGaji.Text = Format(!kbasicgaji, "Rp ###,##,0")
            TextGPokok.Text = Format(!kbasicgaji, "Rp ###,##,0")
            LabelPotong.Caption = Val(Format(TextGaji.Text, "###0")) / Val(TextHK.Text)
            TextPotong.Text = Format(Bulatkan(Val(LabelPotong.Caption), 0), "Rp ###,##,0")

            Dim Hasil As String: Hasil = Val(TextHKMiss.Text) * Val(Format(TextPotong.Text, "###0"))
            TextAkhirHK.Text = Format(Hasil, "Rp ###,##,0")
        End If
    End With
    RSID.Close

    TextTotTunjangan.Text = ""
    TextTotPotongan.Text = ""

    QueryDataPotongan
    QueryDataTunjangan
    Dim TGaji As String: TGaji = Val(Format(TextGPokok.Text, "###0")) + Val(Format(TextTotTunjangan.Text, "###0"))
    TextTGaji.Text = Format(TGaji, "Rp ###,##,0")
    Dim TGajiPot As String: TGajiPot = Val(Format(TextTGaji.Text, "###0")) - Val(Format(TextTotPotongan.Text, "###0")) - Val(Format(TextAkhirHK.Text, "###0"))
    TextGajiPotong.Text = Format(TGajiPot, "Rp ###,##,0")
    TextGajiAkhir.Text = Format(TGajiPot, "Rp ###,##,0")
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case Is = vbKeyEscape
        Unload Me
    Case Is = vbKeyF8
        ButtonSimpan_Click
    End Select
End Sub

Private Sub Form_Load()
    CenterForm Me

    Call initCombo
End Sub

Private Sub initCombo()
    Set RSID = New ADODB.Recordset
    Dim Combo As String: Combo = ComboList(CmbID, RSID, "karyawan", "*", "order by id asc", "id", "nama")
End Sub

Public Sub AutoNumber()
    Set RSPenggajian = New ADODB.Recordset
    RSPenggajian.Open "SELECT id FROM karyawan_gaji WHERE id in(select max(id) from karyawan_gaji)order by id desc", Conn
    RSPenggajian.Requery
    Dim Urut As String * 12
    Dim Hitung As Long
    With RSPenggajian
        If .EOF Then
            Urut = "GK" + Format(Date, "yymmdd") + "0001"
        Else
            If Mid(!Id, 3, 6) <> Format(Date, "yymmdd") Then
                Urut = "GK" + Format(Date, "yymmdd") + "0001"
            Else
                Hitung = Right$(!Id, 4) + 1
                Urut = "GK" + Format(Date, "yymmdd") + Right("0000" & Hitung, 4)
            End If
        End If
        TextKode.Text = Urut
    End With
End Sub

Private Sub Tanggal_Change()
    TextTanggal.Text = Format(Tanggal.Value, "dd/mm/yyyy/")
End Sub

Private Sub TextHKMiss_Change()
    Dim Hasil As String: Hasil = Val(TextHKMiss.Text) * Val(Format(TextPotong.Text, "###0"))
    TextAkhirHK.Text = Format(Hasil, "Rp ###,##,0")

    Dim TGajiPot As String: TGajiPot = Val(Format(TextTGaji.Text, "###0")) - Val(Format(TextTotPotongan.Text, "###0")) - Val(Format(TextAkhirHK.Text, "###0"))
    TextGajiPotong.Text = Format(TGajiPot, "Rp ###,##,0")
    TextGajiAkhir.Text = Format(TGajiPot, "Rp ###,##,0")
End Sub

Public Sub QueryDataTunjangan()
    Set RSTunjangan = New ADODB.Recordset

    LookGridTunjangan

    RSTunjangan.Open "SELECT * FROM tunjangan WHERE id='" & LblID.Caption & "' and MONTH(tanggal) = " & Format(Tanggal.Value, "mm") & " AND YEAR(tanggal) = " & Format(Tanggal.Value, "yyyy"), Conn, adOpenForwardOnly, adLockReadOnly
    RSTunjangan.Requery

    Call TampilGrid
End Sub

Sub TampilGrid()
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
                GridTunjangan.AddItem RSTunjangan!Tanggal & vbTab & Format(RSTunjangan!Jumlah, "#,##0") & vbTab & RSTunjangan!Keterangan
                .MoveNext
            Loop
            HitungTunjangan
        End With
    End If
End Sub

Private Sub HitungTunjangan()
    Dim jum(3), nil(3) As Long
    jum(0) = 0
    jum(1) = 0
    jum(2) = 0

    With GridTunjangan
        For i = 0 To GridTunjangan.Rows - 1
            If IsNumeric(.TextMatrix(i, 1)) Then
                nil(0) = Format(.TextMatrix(i, 1), "###0")
            Else
                nil(0) = Format(Val(.TextMatrix(i, 1)), "###0")
            End If
            jum(0) = jum(0) + nil(0)

        Next
        TextTotTunjangan.Text = Format(jum(0), "Rp ###,##,0")
    End With
End Sub

Sub LookGridTunjangan()
    GridTunjangan.Clear
    GridTunjangan.Rows = 50
    GridTunjangan.Cols = 5
    '---------------
    GridTunjangan.Col = 0
    GridTunjangan.Row = 0
    GridTunjangan.Text = "TANGGAL"
    'GridTunjangan.CellFontBold = True
    '
    GridTunjangan.Col = 1
    GridTunjangan.Row = 0
    GridTunjangan.Text = "TOTAL TUNJANGAN"
    'GridTunjangan.CellFontBold = True
    '
    GridTunjangan.Col = 2
    GridTunjangan.Row = 0
    GridTunjangan.Text = "KETERANGAN"
    'GridTunjangan.CellFontBold = True
    '
End Sub

Public Sub QueryDataPotongan()
    Set RSPotongan = New ADODB.Recordset

    LookGridPotong

    RSPotongan.Open "SELECT * FROM potongan WHERE id='" & LblID.Caption & "' and MONTH(tanggal) = " & Format(Tanggal.Value, "mm") & " AND YEAR(tanggal) = " & Format(Tanggal.Value, "yyyy"), Conn, adOpenForwardOnly, adLockReadOnly
    RSPotongan.Requery

    Call TampilGrid2
End Sub

Sub TampilGrid2()
'On Error Resume Next
    Baris = 0
    If RSPotongan.EOF Then
        Exit Sub
    Else
        With RSPotongan
            .MoveFirst
            Do Until .EOF
                Baris = Baris + 1
                GridPotongan.Rows = Baris
                GridPotongan.AddItem RSPotongan!Tanggal & vbTab & Format(RSPotongan!Jumlah, "#,##0") & vbTab & RSPotongan!Keterangan
                .MoveNext
            Loop
            HitungPotongan
        End With
    End If
End Sub

Private Sub HitungPotongan()
    Dim jum(3), nil(3) As Long
    jum(0) = 0
    jum(1) = 0
    jum(2) = 0

    With GridPotongan
        For i = 0 To GridPotongan.Rows - 1
            If IsNumeric(.TextMatrix(i, 1)) Then
                nil(0) = Format(.TextMatrix(i, 1), "###0")
            Else
                nil(0) = Format(Val(.TextMatrix(i, 1)), "###0")
            End If
            jum(0) = jum(0) + nil(0)

        Next
        TextTotPotongan.Text = Format(jum(0), "Rp ###,##,0")
    End With
End Sub

Sub LookGridPotong()
    GridPotongan.Clear
    GridPotongan.Rows = 50
    GridPotongan.Cols = 5
    '---------------
    GridPotongan.Col = 0
    GridPotongan.Row = 0
    GridPotongan.Text = "TANGGAL"
    'GridPotongan.CellFontBold = True
    '
    GridPotongan.Col = 1
    GridPotongan.Row = 0
    GridPotongan.Text = "TOTAL POTONGAN"
    'GridPotongan.CellFontBold = True
    '
    GridPotongan.Col = 2
    GridPotongan.Row = 0
    GridPotongan.Text = "KETERANGAN"
    'GridPotongan.CellFontBold = True
    '
End Sub

Private Sub TextTanggal_Change()
    Call QueryDataTunjangan
    Call QueryDataPotongan

    Dim TGaji As String: TGaji = Val(Format(TextGPokok.Text, "###0")) + Val(Format(TextTotTunjangan.Text, "###0"))
    TextTGaji.Text = Format(TGaji, "Rp ###,##,0")
    Dim TGajiPot As String: TGajiPot = Val(Format(TextTGaji.Text, "###0")) - Val(Format(TextTotPotongan.Text, "###0")) - Val(Format(TextAkhirHK.Text, "###0"))
    TextGajiPotong.Text = Format(TGajiPot, "Rp ###,##,0")
    TextGajiAkhir.Text = Format(TGajiPot, "Rp ###,##,0")
End Sub

