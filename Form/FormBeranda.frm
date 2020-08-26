VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form FormBeranda 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   0  'None
   Caption         =   "Form Beranda"
   ClientHeight    =   6960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13950
   Icon            =   "FormBeranda.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6960
   ScaleWidth      =   13950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VSFlex8Ctl.VSFlexGrid GridDepartemen 
      Height          =   2625
      Left            =   5760
      TabIndex        =   7
      Top             =   3420
      Width           =   7875
      _cx             =   13891
      _cy             =   4630
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
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FormBeranda.frx":000C
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
   Begin VB.PictureBox Picture1 
      Height          =   2475
      Left            =   8220
      ScaleHeight     =   2415
      ScaleWidth      =   4995
      TabIndex        =   0
      Top             =   4440
      Visible         =   0   'False
      Width           =   5055
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Test"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1905
         TabIndex        =   1
         Top             =   1020
         Width           =   705
      End
   End
   Begin Project1.NEOBadge NEOBadge4 
      Height          =   1650
      Left            =   11220
      TabIndex        =   2
      Top             =   1140
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   2910
      Picture1        =   "FormBeranda.frx":0078
      Picture2        =   "FormBeranda.frx":19D9
      BeginProperty FontCap1 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontCap2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption1        =   "190"
      Caption2        =   "Data Karyawan"
      Color1          =   15712256
      Color2          =   14200064
      Themes          =   3
   End
   Begin Project1.NEOBadge NEOBadge1 
      Height          =   1650
      Left            =   60
      TabIndex        =   3
      Top             =   1200
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   2910
      Picture1        =   "FormBeranda.frx":1A6F
      Picture2        =   "FormBeranda.frx":21B2
      BeginProperty FontCap1 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontCap2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption1        =   "190"
      Caption2        =   "Data Departemen"
      Color1          =   3754973
      Color2          =   3359943
      Themes          =   1
   End
   Begin Project1.NEOBadge NEOBadge2 
      Height          =   1650
      Left            =   3780
      TabIndex        =   4
      Top             =   1140
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   2910
      Picture1        =   "FormBeranda.frx":2248
      Picture2        =   "FormBeranda.frx":1822C
      BeginProperty FontCap1 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontCap2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption1        =   "190"
      Caption2        =   "Data Pengeluaran/Bln"
      Color1          =   5940736
      Color2          =   5346816
      Themes          =   2
   End
   Begin Project1.NEOBadge NEOBadge3 
      Height          =   1650
      Left            =   7500
      TabIndex        =   5
      Top             =   1140
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   2910
      Picture1        =   "FormBeranda.frx":182C2
      Picture2        =   "FormBeranda.frx":1A3C2
      BeginProperty FontCap1 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontCap2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption1        =   "190"
      Caption2        =   "Data Inventaris"
      Color1          =   1219827
      Color2          =   1084891
      Themes          =   4
   End
   Begin VSFlex8Ctl.VSFlexGrid GridPotongan 
      Height          =   2625
      Left            =   60
      TabIndex        =   6
      Top             =   3060
      Width           =   7875
      _cx             =   13891
      _cy             =   4630
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
      FormatString    =   $"FormBeranda.frx":1A458
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
   Begin Project1.isButton ButtonKeluar 
      Cancel          =   -1  'True
      Height          =   435
      Left            =   780
      TabIndex        =   8
      Top             =   240
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
Attribute VB_Name = "FormBeranda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RSPotongan As New ADODB.Recordset
Dim RSDepartemen As New ADODB.Recordset

Sub TampilanAwal()
    Set RSPotongan = New ADODB.Recordset

    'RSPotongan.Open "SELECT * FROM potongan WHERE jumlah in(select max(jumlah) from potongan)order by jumlah desc LIMIT 10", Conn, adOpenForwardOnly, adLockReadOnly
    RSPotongan.Open "SELECT * FROM potongan order by jumlah desc LIMIT 10", Conn, adOpenForwardOnly, adLockReadOnly
    RSPotongan.Requery

    Call TampilGrid
End Sub

Sub TampilGrid()
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
                GridPotongan.AddItem RSPotongan!Id & vbTab & RSPotongan!Tanggal & vbTab & Format(RSPotongan!Jumlah, "#,##0") & vbTab & RSPotongan!Keterangan & vbTab & RSPotongan!Kode
                .MoveNext
            Loop
        End With
    End If
End Sub

Sub TampilanAwal2()
    Set RSDepartemen = New ADODB.Recordset

    'RSPotongan.Open "SELECT * FROM potongan WHERE jumlah in(select max(jumlah) from potongan)order by jumlah desc LIMIT 10", Conn, adOpenForwardOnly, adLockReadOnly
    RSDepartemen.Open "SELECT * FROM departemen where id>'1'", Conn, adOpenForwardOnly, adLockReadOnly
    RSDepartemen.Requery

    Call TampilGrid2
End Sub

Sub TampilGrid2()
'On Error Resume Next
    Baris = 0
    If RSDepartemen.EOF Then
        Exit Sub
    Else
        With RSDepartemen
            .MoveFirst
            Do Until .EOF
                Baris = Baris + 1
                GridDepartemen.Rows = Baris
                GridDepartemen.AddItem RSDepartemen!Id & vbTab & RSDepartemen!Keterangan
                .MoveNext
            Loop
        End With
    End If
End Sub

Private Sub ButtonKeluar_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    Me.Hide
    'Unload Me
End Sub

Private Sub Form_Load()
    Call SetIcon(Me.hWnd, "FORMICON", False)
    With Me
        .Top = 0
        .Height = Screen.Height
        .Left = 0
        .Width = Screen.Width
    End With
    Call AutoResize

    Call HitungData

    Call TampilanAwal
    Call TampilanAwal2

    Call BarisLunasWarna
End Sub

Function HitungData()
    NEOBadge1.Caption1 = getRecordCount("departemen", "WHERE id>'1'")
    NEOBadge2.Caption1 = getRecordCount("pengeluaran", "WHERE MONTH(tanggal) =" & Format(Now, "mm") & " AND YEAR(tanggal) = " & Format(Now, "yyyy"))
    NEOBadge3.Caption1 = getRecordCount("inventaris", "")
    NEOBadge4.Caption1 = getRecordCount("karyawan", "")
End Function

Private Sub BarisLunasWarna()

    For i = 1 To GridDepartemen.Rows - 1

        If Left(GridDepartemen.TextMatrix(i, 1), 1) = "W" Then
            GridDepartemen.Cell(flexcpBackColor, i, 1) = &H80FFFF
        End If

        If Left(GridDepartemen.TextMatrix(i, 1), 1) = "H" Then
            GridDepartemen.Cell(flexcpBackColor, i, 1) = &H80FF80
        End If

        If Left(GridDepartemen.TextMatrix(i, 1), 1) = "C" Then
            GridDepartemen.Cell(flexcpBackColor, i, 1) = RGB(188, 63, 67)    '&H80FFFF
        End If

        If Left(GridDepartemen.TextMatrix(i, 1), 1) = "K" Then
            GridDepartemen.Cell(flexcpBackColor, i, 1) = RGB(111, 146, 200)    '&H80FF80
        End If

        If Left(GridDepartemen.TextMatrix(i, 1), 1) = "A" Then
            GridDepartemen.Cell(flexcpBackColor, i, 1) = RGB(253, 206, 5)    '&H80FF80
        End If

    Next i

End Sub

Private Sub Form_Resize()
    Call AutoResize
End Sub

Public Sub AutoResize()
    If Me.WindowState = 1 Then Exit Sub
    'With Picture1
    '    .Move (FormMain.PictureBeranda.ScaleWidth * 0.5) - (.Width * 0.5), (FormMain.PictureBeranda.ScaleHeight * 0.5) - (.Height * 0.5)
    'End With
    Dim BadgeTrue As String: BadgeTrue = ReadINI("Settings", "Badge", Lokasi_File_Konfigurasi)
    If BadgeTrue = "True" Then
        'Neo Badge AutoSize
        With NEOBadge1
            .Left = 180
            .Top = 180
            .Width = (FormMain.PictureBeranda.Width / 4) - 180
        End With

        With NEOBadge2
            .Left = NEOBadge1.Left + NEOBadge1.Width + 135
            .Top = 180
            .Width = (FormMain.PictureBeranda.Width / 4) - 180
        End With

        With NEOBadge3
            .Left = NEOBadge2.Left + NEOBadge2.Width + 135
            .Top = 180
            .Width = (FormMain.PictureBeranda.Width / 4) - 180
        End With

        With NEOBadge4
            .Left = NEOBadge3.Left + NEOBadge3.Width + 135
            .Top = 180
            .Width = (FormMain.PictureBeranda.Width / 4) - 180
        End With
        'Neo Badge AutoSize

        With GridPotongan
            .Left = 180
            .Top = NEOBadge1.Top + NEOBadge1.Height + 180
            .Width = NEOBadge1.Width + NEOBadge2.Width + 140
        End With

        With GridDepartemen
            .Left = NEOBadge3.Left
            .Top = NEOBadge3.Top + NEOBadge3.Height + 180
            .Width = NEOBadge3.Width + NEOBadge4.Width + 120
        End With
    Else
        With NEOBadge1
            .Left = 180
            .Top = 180
            .Width = (FormMain.PictureBeranda.Width / 4) - 180
            .Visible = False
        End With

        With NEOBadge2
            .Left = NEOBadge1.Left + NEOBadge1.Width + 135
            .Top = 180
            .Width = (FormMain.PictureBeranda.Width / 4) - 180
            .Visible = False
        End With

        With NEOBadge3
            .Left = NEOBadge2.Left + NEOBadge2.Width + 135
            .Top = 180
            .Width = (FormMain.PictureBeranda.Width / 4) - 180
            .Visible = False
        End With

        With NEOBadge4
            .Left = NEOBadge3.Left + NEOBadge3.Width + 135
            .Top = 180
            .Width = (FormMain.PictureBeranda.Width / 4) - 180
            .Visible = False
        End With

        With GridPotongan
            .Left = 180
            '.Top = NEOBadge1.Top + NEOBadge1.Height + 180
            .Top = NEOBadge1.Top
            .Width = NEOBadge1.Width + NEOBadge2.Width + 140
        End With

        With GridDepartemen
            .Left = NEOBadge3.Left
            '.Top = NEOBadge3.Top + NEOBadge3.Height + 180
            .Top = NEOBadge3.Top
            .Width = NEOBadge3.Width + NEOBadge4.Width + 120
        End With
    End If

    With ButtonKeluar
        .Move FormMain.PictureBeranda.ScaleWidth - ButtonKeluar.Width - 180, FormMain.PictureBeranda.ScaleHeight - ButtonKeluar.Height - 180
    End With
End Sub
