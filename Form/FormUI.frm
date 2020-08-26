VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{62E1564C-1E05-44A7-A921-FD8347F324A5}#1.0#0"; "HookMenu.ocx"
Begin VB.Form FormUI 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   Caption         =   "Form User Interface"
   ClientHeight    =   4725
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   9165
   FillColor       =   &H00404040&
   ForeColor       =   &H00404040&
   Icon            =   "FormUI.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   9165
   StartUpPosition =   3  'Windows Default
   Begin HookMenu.ctxHookMenu ctxHookMenu1 
      Left            =   4140
      Top             =   540
      _ExtentX        =   900
      _ExtentY        =   900
      BitmapSize      =   24
      BmpCount        =   9
      Bmp:1           =   "FormUI.frx":000C
      Key:1           =   "#mLap1"
      Bmp:2           =   "FormUI.frx":0814
      Key:2           =   "#mLap2"
      Bmp:3           =   "FormUI.frx":101C
      Key:3           =   "#mLap3"
      Bmp:4           =   "FormUI.frx":1824
      Key:4           =   "#mLap4"
      Bmp:5           =   "FormUI.frx":202C
      Key:5           =   "#mLap5"
      Bmp:6           =   "FormUI.frx":2834
      Key:6           =   "#mLap6"
      Bmp:7           =   "FormUI.frx":303C
      Key:7           =   "#mLap7"
      Bmp:8           =   "FormUI.frx":3844
      Key:8           =   "#mLap8"
      Bmp:9           =   "FormUI.frx":404C
      Key:9           =   "#mLap9"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H00BC8D3C&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   0
      Picture         =   "FormUI.frx":4854
      ScaleHeight     =   765
      ScaleWidth      =   660
      TabIndex        =   30
      Top             =   0
      Width           =   660
      Begin VB.Shape Shape1 
         BorderColor     =   &H00BC8D3C&
         FillColor       =   &H00BC8D3C&
         FillStyle       =   0  'Solid
         Height          =   795
         Left            =   60
         Top             =   0
         Width           =   675
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   230
         X2              =   420
         Y1              =   300
         Y2              =   300
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   230
         X2              =   430
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   230
         X2              =   430
         Y1              =   420
         Y2              =   420
      End
      Begin VB.Image Image1 
         Height          =   795
         Left            =   0
         Top             =   0
         Width           =   675
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00E6C8DD&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   5160
      ScaleHeight     =   585
      ScaleWidth      =   1305
      TabIndex        =   29
      Top             =   360
      Width           =   1335
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   3300
      ScaleHeight     =   1095
      ScaleWidth      =   1515
      TabIndex        =   25
      Top             =   3120
      Width           =   1515
      Begin VB.PictureBox S 
         Appearance      =   0  'Flat
         BackColor       =   &H00312D22&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   5
         Left            =   0
         ScaleHeight     =   855
         ScaleWidth      =   495
         TabIndex        =   28
         Top             =   260
         Width           =   495
      End
      Begin VB.PictureBox H1 
         Appearance      =   0  'Flat
         BackColor       =   &H001955D1&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   495
         TabIndex        =   27
         Top             =   0
         Width           =   495
      End
      Begin VB.PictureBox H2 
         Appearance      =   0  'Flat
         BackColor       =   &H00317FE4&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   480
         ScaleHeight     =   255
         ScaleWidth      =   1095
         TabIndex        =   26
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   1560
      ScaleHeight     =   1095
      ScaleWidth      =   1515
      TabIndex        =   21
      Top             =   3120
      Width           =   1515
      Begin VB.PictureBox H2 
         Appearance      =   0  'Flat
         BackColor       =   &H00B55C9B&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   480
         ScaleHeight     =   255
         ScaleWidth      =   1095
         TabIndex        =   24
         Top             =   0
         Width           =   1095
      End
      Begin VB.PictureBox H1 
         Appearance      =   0  'Flat
         BackColor       =   &H00AB488E&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   495
         TabIndex        =   23
         Top             =   0
         Width           =   495
      End
      Begin VB.PictureBox S 
         Appearance      =   0  'Flat
         BackColor       =   &H00312D22&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   4
         Left            =   0
         ScaleHeight     =   855
         ScaleWidth      =   495
         TabIndex        =   22
         Top             =   260
         Width           =   495
      End
   End
   Begin VB.TextBox Text_Layar_Penuh 
      Appearance      =   0  'Flat
      Height          =   465
      Left            =   360
      TabIndex        =   16
      Top             =   1020
      Width           =   735
   End
   Begin VB.PictureBox PictureB 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   1560
      ScaleHeight     =   1095
      ScaleWidth      =   1515
      TabIndex        =   12
      Top             =   1800
      Width           =   1515
      Begin VB.PictureBox H2 
         Appearance      =   0  'Flat
         BackColor       =   &H00BC8D3C&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   480
         ScaleHeight     =   255
         ScaleWidth      =   1095
         TabIndex        =   15
         Top             =   0
         Width           =   1095
      End
      Begin VB.PictureBox H1 
         Appearance      =   0  'Flat
         BackColor       =   &H00A87F36&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   495
         TabIndex        =   14
         Top             =   0
         Width           =   495
      End
      Begin VB.PictureBox S 
         Appearance      =   0  'Flat
         BackColor       =   &H00312D22&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   2
         Left            =   0
         ScaleHeight     =   855
         ScaleWidth      =   495
         TabIndex        =   13
         Top             =   260
         Width           =   495
      End
   End
   Begin VB.PictureBox PictureR 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   3300
      ScaleHeight     =   1095
      ScaleWidth      =   1515
      TabIndex        =   8
      Top             =   1800
      Width           =   1515
      Begin VB.PictureBox S 
         Appearance      =   0  'Flat
         BackColor       =   &H00312D22&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   0
         Left            =   0
         ScaleHeight     =   855
         ScaleWidth      =   495
         TabIndex        =   11
         Top             =   260
         Width           =   495
      End
      Begin VB.PictureBox H1 
         Appearance      =   0  'Flat
         BackColor       =   &H002539D7&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   495
         TabIndex        =   10
         Top             =   0
         Width           =   495
      End
      Begin VB.PictureBox H2 
         Appearance      =   0  'Flat
         BackColor       =   &H00394BDD&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   480
         ScaleHeight     =   255
         ScaleWidth      =   1095
         TabIndex        =   9
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.PictureBox PictureG 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   5040
      ScaleHeight     =   1095
      ScaleWidth      =   1515
      TabIndex        =   4
      Top             =   1800
      Width           =   1515
      Begin VB.PictureBox H2 
         Appearance      =   0  'Flat
         BackColor       =   &H005AA600&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   480
         ScaleHeight     =   255
         ScaleWidth      =   1095
         TabIndex        =   7
         Top             =   0
         Width           =   1095
      End
      Begin VB.PictureBox H1 
         Appearance      =   0  'Flat
         BackColor       =   &H00498700&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   495
         TabIndex        =   6
         Top             =   0
         Width           =   495
      End
      Begin VB.PictureBox S 
         Appearance      =   0  'Flat
         BackColor       =   &H00312D22&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   1
         Left            =   0
         ScaleHeight     =   855
         ScaleWidth      =   495
         TabIndex        =   5
         Top             =   260
         Width           =   495
      End
   End
   Begin VB.PictureBox PictureY 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   6780
      ScaleHeight     =   1095
      ScaleWidth      =   1515
      TabIndex        =   0
      Top             =   1800
      Width           =   1515
      Begin VB.PictureBox S 
         Appearance      =   0  'Flat
         BackColor       =   &H00312D22&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   3
         Left            =   0
         ScaleHeight     =   855
         ScaleWidth      =   495
         TabIndex        =   3
         Top             =   260
         Width           =   495
      End
      Begin VB.PictureBox H1 
         Appearance      =   0  'Flat
         BackColor       =   &H000B8BDB&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   495
         TabIndex        =   2
         Top             =   0
         Width           =   495
      End
      Begin VB.PictureBox H2 
         Appearance      =   0  'Flat
         BackColor       =   &H00129CF3&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   480
         ScaleHeight     =   255
         ScaleWidth      =   1095
         TabIndex        =   1
         Top             =   0
         Width           =   1095
      End
   End
   Begin MSComctlLib.ImageList iList16x16 
      Left            =   1980
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormUI.frx":62E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormUI.frx":6CF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormUI.frx":7706
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormUI.frx":7CA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormUI.frx":823A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormUI.frx":87D4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList iList16x16G 
      Left            =   3300
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormUI.frx":8D6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormUI.frx":9308
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormUI.frx":98A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormUI.frx":9E3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormUI.frx":A3D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormUI.frx":A970
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image Icon_Error 
      Enabled         =   0   'False
      Height          =   720
      Left            =   6240
      Picture         =   "FormUI.frx":AF0A
      Top             =   3360
      Width           =   720
   End
   Begin VB.Image Icon_Quest 
      Enabled         =   0   'False
      Height          =   720
      Left            =   6960
      Picture         =   "FormUI.frx":CA4C
      Top             =   3360
      Width           =   720
   End
   Begin VB.Image Icon_Info 
      Enabled         =   0   'False
      Height          =   720
      Left            =   7680
      Picture         =   "FormUI.frx":E58E
      Top             =   3360
      Width           =   720
   End
   Begin VB.Label Label_Left 
      Height          =   375
      Left            =   720
      TabIndex        =   20
      Top             =   420
      Width           =   495
   End
   Begin VB.Label Label_Top 
      Height          =   315
      Left            =   720
      TabIndex        =   19
      Top             =   0
      Width           =   555
   End
   Begin VB.Label Label_Width 
      Height          =   375
      Left            =   0
      TabIndex        =   18
      Top             =   420
      Width           =   555
   End
   Begin VB.Label Label_Height 
      Height          =   315
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   555
   End
   Begin VB.Menu mLaporan 
      Caption         =   "Laporan"
      Begin VB.Menu mLap1 
         Caption         =   "Data Karyawan"
      End
      Begin VB.Menu mLap2 
         Caption         =   "Data Tunjangan"
      End
      Begin VB.Menu mLap3 
         Caption         =   "Data Potongan"
      End
      Begin VB.Menu mLap4 
         Caption         =   "Data Gaji"
      End
      Begin VB.Menu mSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mLap5 
         Caption         =   "Laporan Pemasukan"
      End
      Begin VB.Menu mLap6 
         Caption         =   "Laporan Pengeluaran"
      End
      Begin VB.Menu mLap7 
         Caption         =   "Laporan Kas"
      End
      Begin VB.Menu mLap8 
         Caption         =   "Laporan Inventaris"
      End
      Begin VB.Menu mSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mLap9 
         Caption         =   "Cetak Slip Gaji"
      End
   End
End
Attribute VB_Name = "FormUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Call SetIcon(Me.hWnd, "FORMICON", False)
End Sub

Private Sub mLap1_Click()
    Call Form_Diatas(Cetak_Karyawan, FormMain.PictureBeranda)
End Sub

Private Sub mLap2_Click()
    Call Form_Diatas(Cetak_Tunjangan, FormMain.PictureBeranda)
End Sub

Private Sub mLap3_Click()
    Call Form_Diatas(Cetak_Potongan, FormMain.PictureBeranda)
End Sub

Private Sub mLap4_Click()
    Call Form_Diatas(Cetak_DataGaji, FormMain.PictureBeranda)
End Sub

Private Sub mLap5_Click()
    Call Form_Diatas(Cetak_Pemasukan, FormMain.PictureBeranda)
End Sub

Private Sub mLap6_Click()
    Call Form_Diatas(Cetak_Pengeluaran, FormMain.PictureBeranda)
End Sub

Private Sub mLap8_Click()
    Call Form_Diatas(Cetak_Inventaris, FormMain.PictureBeranda)
End Sub

Private Sub mLap9_Click()
    Call Form_Diatas(Cetak_SlipGaji, FormMain.PictureBeranda)
End Sub
