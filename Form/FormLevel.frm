VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FormLevel 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Manajemen Level"
   ClientHeight    =   8295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14715
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormLevel.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8295
   ScaleWidth      =   14715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Project1.XPFrame FrameLevel 
      Height          =   975
      Left            =   120
      TabIndex        =   68
      Top             =   6300
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   1720
      HeaderLightColor=   8421504
      HeaderDarkColor =   8421504
      TextColor       =   16777215
      Caption         =   "TAMBAH LEVEL"
      Curvature       =   0
      Begin VB.TextBox TextNama 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1440
         TabIndex        =   70
         Top             =   480
         Width           =   2955
      End
      Begin VB.TextBox TextKode 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   120
         TabIndex        =   69
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label LabelKode 
         Height          =   255
         Left            =   3360
         TabIndex        =   71
         Top             =   60
         Visible         =   0   'False
         Width           =   1035
      End
   End
   Begin Project1.XPFrame FrameAkses 
      Height          =   6735
      Left            =   4740
      TabIndex        =   8
      Top             =   720
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   11880
      HeaderLightColor=   8421504
      HeaderDarkColor =   8421504
      TextColor       =   16777215
      Caption         =   "HAK AKSES"
      CaptionAlignment=   2
      Curvature       =   0
      Begin TabDlg.SSTab TabAkses 
         Height          =   6135
         Left            =   5
         TabIndex        =   9
         Top             =   480
         Width           =   9675
         _ExtentX        =   17066
         _ExtentY        =   10821
         _Version        =   393216
         Style           =   1
         TabsPerRow      =   5
         TabHeight       =   564
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Master Data"
         TabPicture(0)   =   "FormLevel.frx":000C
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "PictureFrame"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Picture3"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Laporan"
         TabPicture(1)   =   "FormLevel.frx":0028
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Picture4"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Lainnya"
         TabPicture(2)   =   "FormLevel.frx":0044
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Picture2"
         Tab(2).ControlCount=   1
         Begin VB.PictureBox Picture4 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   4035
            Left            =   -74940
            ScaleHeight     =   4035
            ScaleWidth      =   8595
            TabIndex        =   63
            Top             =   360
            Width           =   8595
            Begin VB.Line Line3 
               BorderColor     =   &H00404040&
               X1              =   3420
               X2              =   3420
               Y1              =   3840
               Y2              =   420
            End
            Begin VB.Line Line4 
               X1              =   120
               X2              =   7320
               Y1              =   420
               Y2              =   420
            End
            Begin VB.Shape Shape2 
               BorderColor     =   &H00404040&
               Height          =   3735
               Left            =   120
               Top             =   120
               Width           =   8355
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BackColor       =   &H00404040&
               Caption         =   "FITUR"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   3420
               TabIndex        =   65
               Top             =   120
               Width           =   5055
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               BackColor       =   &H00404040&
               Caption         =   "NAMA FORM"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   120
               TabIndex        =   64
               Top             =   120
               Width           =   3315
            End
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   3975
            Left            =   60
            ScaleHeight     =   3975
            ScaleWidth      =   8595
            TabIndex        =   23
            Top             =   360
            Width           =   8595
            Begin VB.CheckBox chkGajiGenerate 
               Appearance      =   0  'Flat
               Caption         =   "&Generate"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   7140
               TabIndex        =   24
               Top             =   1620
               Width           =   1155
            End
            Begin VB.CheckBox chkKaryView 
               Appearance      =   0  'Flat
               Caption         =   "&Data Karyawan"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   450
               TabIndex        =   60
               Top             =   540
               Width           =   2500
            End
            Begin VB.CheckBox chkKaryAdd 
               Appearance      =   0  'Flat
               Caption         =   "&Tambah"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   3660
               TabIndex        =   59
               Top             =   540
               Width           =   1155
            End
            Begin VB.CheckBox chkKaryEdit 
               Appearance      =   0  'Flat
               Caption         =   "&Ubah"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   4980
               TabIndex        =   58
               Top             =   540
               Width           =   975
            End
            Begin VB.CheckBox chkKaryDelete 
               Appearance      =   0  'Flat
               Caption         =   "&Hapus"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   6000
               TabIndex        =   57
               Top             =   540
               Width           =   1155
            End
            Begin VB.CheckBox chkPotView 
               Appearance      =   0  'Flat
               Caption         =   "&Data Potongan"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   450
               TabIndex        =   56
               Top             =   900
               Width           =   2500
            End
            Begin VB.CheckBox chkPotAdd 
               Appearance      =   0  'Flat
               Caption         =   "&Tambah"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   3660
               TabIndex        =   55
               Top             =   900
               Width           =   1155
            End
            Begin VB.CheckBox chkPotEdit 
               Appearance      =   0  'Flat
               Caption         =   "&Ubah"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   4980
               TabIndex        =   54
               Top             =   900
               Width           =   975
            End
            Begin VB.CheckBox chkPotDelete 
               Appearance      =   0  'Flat
               Caption         =   "&Hapus"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   6000
               TabIndex        =   53
               Top             =   900
               Width           =   1155
            End
            Begin VB.CheckBox chkTunView 
               Appearance      =   0  'Flat
               Caption         =   "&Data Tunjangan"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   450
               TabIndex        =   52
               Top             =   1260
               Width           =   2500
            End
            Begin VB.CheckBox chkTunAdd 
               Appearance      =   0  'Flat
               Caption         =   "&Tambah"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   3660
               TabIndex        =   51
               Top             =   1260
               Width           =   1155
            End
            Begin VB.CheckBox chkTunEdit 
               Appearance      =   0  'Flat
               Caption         =   "&Ubah"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   4980
               TabIndex        =   50
               Top             =   1260
               Width           =   975
            End
            Begin VB.CheckBox chkTunDelete 
               Appearance      =   0  'Flat
               Caption         =   "&Hapus"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   6000
               TabIndex        =   49
               Top             =   1260
               Width           =   1155
            End
            Begin VB.CheckBox chkGajView 
               Appearance      =   0  'Flat
               Caption         =   "&Data Gaji"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   450
               TabIndex        =   48
               Top             =   1620
               Width           =   2500
            End
            Begin VB.CheckBox chkGajAdd 
               Appearance      =   0  'Flat
               Caption         =   "&Tambah"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   3660
               TabIndex        =   47
               Top             =   1620
               Width           =   1155
            End
            Begin VB.CheckBox chkGajEdit 
               Appearance      =   0  'Flat
               Caption         =   "&Ubah"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   4980
               TabIndex        =   46
               Top             =   1620
               Width           =   975
            End
            Begin VB.CheckBox chkGajDelete 
               Appearance      =   0  'Flat
               Caption         =   "&Hapus"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   6000
               TabIndex        =   45
               Top             =   1620
               Width           =   1155
            End
            Begin VB.CheckBox chkPengView 
               Appearance      =   0  'Flat
               Caption         =   "&Pengeluaran"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   450
               TabIndex        =   44
               Top             =   1980
               Width           =   2500
            End
            Begin VB.CheckBox chkPengAdd 
               Appearance      =   0  'Flat
               Caption         =   "&Tambah"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   3660
               TabIndex        =   43
               Top             =   1980
               Width           =   1155
            End
            Begin VB.CheckBox chkPengEdit 
               Appearance      =   0  'Flat
               Caption         =   "&Ubah"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   4980
               TabIndex        =   42
               Top             =   1980
               Width           =   975
            End
            Begin VB.CheckBox chkPengDelete 
               Appearance      =   0  'Flat
               Caption         =   "&Hapus"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   6000
               TabIndex        =   41
               Top             =   1980
               Width           =   1155
            End
            Begin VB.CheckBox chkPemView 
               Appearance      =   0  'Flat
               Caption         =   "&Pemasukan"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   450
               TabIndex        =   40
               Top             =   2340
               Width           =   2500
            End
            Begin VB.CheckBox chkPemAdd 
               Appearance      =   0  'Flat
               Caption         =   "&Tambah"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   3660
               TabIndex        =   39
               Top             =   2340
               Width           =   1155
            End
            Begin VB.CheckBox chkPemEdit 
               Appearance      =   0  'Flat
               Caption         =   "&Ubah"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   4980
               TabIndex        =   38
               Top             =   2340
               Width           =   975
            End
            Begin VB.CheckBox chkPemDelete 
               Appearance      =   0  'Flat
               Caption         =   "&Hapus"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   6000
               TabIndex        =   37
               Top             =   2340
               Width           =   1155
            End
            Begin VB.CheckBox chkKasView 
               Appearance      =   0  'Flat
               Caption         =   "&Kas"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   450
               TabIndex        =   36
               Top             =   2700
               Width           =   2500
            End
            Begin VB.CheckBox chkKasAdd 
               Appearance      =   0  'Flat
               Caption         =   "&Tambah"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   3660
               TabIndex        =   35
               Top             =   2700
               Width           =   1155
            End
            Begin VB.CheckBox chkKasEdit 
               Appearance      =   0  'Flat
               Caption         =   "&Ubah"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   4980
               TabIndex        =   34
               Top             =   2700
               Width           =   975
            End
            Begin VB.CheckBox chkKasDelete 
               Appearance      =   0  'Flat
               Caption         =   "&Hapus"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   6000
               TabIndex        =   33
               Top             =   2700
               Width           =   1155
            End
            Begin VB.CheckBox chkBiayaView 
               Appearance      =   0  'Flat
               Caption         =   "&Biaya"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   450
               TabIndex        =   32
               Top             =   3060
               Width           =   2500
            End
            Begin VB.CheckBox chkBiayaAdd 
               Appearance      =   0  'Flat
               Caption         =   "&Tambah"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   3660
               TabIndex        =   31
               Top             =   3060
               Width           =   1155
            End
            Begin VB.CheckBox chkBiayaEdit 
               Appearance      =   0  'Flat
               Caption         =   "&Ubah"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   4980
               TabIndex        =   30
               Top             =   3060
               Width           =   975
            End
            Begin VB.CheckBox chkBiayaDelete 
               Appearance      =   0  'Flat
               Caption         =   "&Hapus"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   6000
               TabIndex        =   29
               Top             =   3060
               Width           =   1155
            End
            Begin VB.CheckBox chkInvView 
               Appearance      =   0  'Flat
               Caption         =   "&Inventaris"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   450
               TabIndex        =   28
               Top             =   3420
               Width           =   2500
            End
            Begin VB.CheckBox chkInvAdd 
               Appearance      =   0  'Flat
               Caption         =   "&Tambah"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   3660
               TabIndex        =   27
               Top             =   3420
               Width           =   1155
            End
            Begin VB.CheckBox chkInvEdit 
               Appearance      =   0  'Flat
               Caption         =   "&Ubah"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   4980
               TabIndex        =   26
               Top             =   3420
               Width           =   975
            End
            Begin VB.CheckBox chkInvDelete 
               Appearance      =   0  'Flat
               Caption         =   "&Hapus"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   6000
               TabIndex        =   25
               Top             =   3420
               Width           =   1155
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackColor       =   &H00404040&
               Caption         =   "NAMA FORM"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   120
               TabIndex        =   62
               Top             =   120
               Width           =   3315
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               BackColor       =   &H00404040&
               Caption         =   "FITUR"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   3420
               TabIndex        =   61
               Top             =   120
               Width           =   5055
            End
            Begin VB.Shape Shape1 
               BorderColor     =   &H00404040&
               Height          =   3735
               Left            =   120
               Top             =   120
               Width           =   8355
            End
            Begin VB.Line Line1 
               X1              =   120
               X2              =   7320
               Y1              =   420
               Y2              =   420
            End
            Begin VB.Line Line2 
               BorderColor     =   &H00404040&
               X1              =   3420
               X2              =   3420
               Y1              =   3840
               Y2              =   420
            End
         End
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   3975
            Left            =   -74940
            ScaleHeight     =   3975
            ScaleWidth      =   8595
            TabIndex        =   11
            Top             =   360
            Width           =   8595
            Begin VB.CheckBox chkUserPass 
               Appearance      =   0  'Flat
               Caption         =   "&Password"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   7140
               TabIndex        =   12
               Top             =   540
               Width           =   1155
            End
            Begin VB.CheckBox chkLevelDelete 
               Appearance      =   0  'Flat
               Caption         =   "&Hapus"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   6030
               TabIndex        =   20
               Top             =   900
               Width           =   1155
            End
            Begin VB.CheckBox chkLevelEdit 
               Appearance      =   0  'Flat
               Caption         =   "&Ubah"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   5010
               TabIndex        =   19
               Top             =   900
               Width           =   975
            End
            Begin VB.CheckBox chkLevelAdd 
               Appearance      =   0  'Flat
               Caption         =   "&Tambah"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   3690
               TabIndex        =   18
               Top             =   900
               Width           =   1155
            End
            Begin VB.CheckBox chkLevelView 
               Appearance      =   0  'Flat
               Caption         =   "&Manajemen Level"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   390
               TabIndex        =   17
               Top             =   900
               Width           =   2500
            End
            Begin VB.CheckBox chkUserDelete 
               Appearance      =   0  'Flat
               Caption         =   "&Hapus"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   6030
               TabIndex        =   16
               Top             =   540
               Width           =   1065
            End
            Begin VB.CheckBox chkUserEdit 
               Appearance      =   0  'Flat
               Caption         =   "&Ubah"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   5010
               TabIndex        =   15
               Top             =   540
               Width           =   975
            End
            Begin VB.CheckBox chkUserAdd 
               Appearance      =   0  'Flat
               Caption         =   "&Tambah"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   3690
               TabIndex        =   14
               Top             =   540
               Width           =   1155
            End
            Begin VB.CheckBox chkUserView 
               Appearance      =   0  'Flat
               Caption         =   "&Manajemen Pengguna"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   390
               TabIndex        =   13
               Top             =   540
               Width           =   2500
            End
            Begin VB.Line Line5 
               BorderColor     =   &H00404040&
               X1              =   3420
               X2              =   3420
               Y1              =   3840
               Y2              =   420
            End
            Begin VB.Shape Shape3 
               BorderColor     =   &H00404040&
               Height          =   3735
               Left            =   120
               Top             =   120
               Width           =   8355
            End
            Begin VB.Label Label6 
               Alignment       =   2  'Center
               BackColor       =   &H00404040&
               Caption         =   "FITUR"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   3420
               TabIndex        =   22
               Top             =   120
               Width           =   5055
            End
            Begin VB.Line Line6 
               X1              =   120
               X2              =   7320
               Y1              =   420
               Y2              =   420
            End
            Begin VB.Label Label7 
               Alignment       =   2  'Center
               BackColor       =   &H00404040&
               Caption         =   "NAMA FORM"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   120
               TabIndex        =   21
               Top             =   120
               Width           =   3315
            End
         End
         Begin VB.PictureBox PictureFrame 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   4005
            Left            =   60
            ScaleHeight     =   4005
            ScaleWidth      =   8595
            TabIndex        =   10
            Top             =   360
            Width           =   8595
         End
      End
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
      Top             =   7560
      Width           =   12765
      Begin Project1.isButton mLevel 
         Height          =   435
         Index           =   1
         Left            =   2520
         TabIndex        =   1
         Top             =   0
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   767
         Icon            =   "FormLevel.frx":0060
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
      Begin Project1.isButton mLevel 
         Height          =   435
         Index           =   2
         Left            =   4680
         TabIndex        =   2
         Top             =   0
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   767
         Icon            =   "FormLevel.frx":0A72
         Style           =   8
         Caption         =   "&Batal"
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
      Begin Project1.isButton mLevel 
         Height          =   435
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   0
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   767
         Icon            =   "FormLevel.frx":1484
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
         TabIndex        =   4
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
      Begin Project1.isButton mLevel 
         Height          =   435
         Index           =   3
         Left            =   6780
         TabIndex        =   66
         Top             =   0
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   767
         Icon            =   "FormLevel.frx":1F96
         Style           =   8
         Caption         =   "&Simpan"
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
      Begin Project1.isButton mLevel 
         Height          =   435
         Index           =   4
         Left            =   8880
         TabIndex        =   67
         Top             =   0
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   767
         Icon            =   "FormLevel.frx":2330
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
   End
   Begin VSFlex8Ctl.VSFlexGrid GridLevel 
      Height          =   5445
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   4515
      _cx             =   7964
      _cy             =   9604
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
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FormLevel.frx":28CA
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
      Caption         =   "DATA USER LEVEL"
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
      TabIndex        =   7
      Top             =   60
      Width           =   4815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kelola Informasi Data User Level Yang Tersedia."
      Height          =   195
      Left            =   195
      TabIndex        =   6
      Top             =   300
      Width           =   3450
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
Attribute VB_Name = "FormLevel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RSLevel As New ADODB.Recordset
Dim RSDepartemen As New ADODB.Recordset

Public blnPilih As Boolean

Dim ClassLevel As ClassLevel

'Menampilkan Data Ke GridUser
Sub TampilGrid()
'On Error Resume Next
    Baris = 0
    If RSLevel.EOF Then
        Exit Sub
    Else
        With RSLevel
            .MoveFirst
            Do Until .EOF
                Baris = Baris + 1
                GridLevel.Rows = Baris
                GridLevel.AddItem RSLevel!Kode & vbTab & RSLevel!Nama
                .MoveNext
            Loop
            GridLevel.Select 1, 1, 1, GridLevel.Cols - 1
        End With
    End If
End Sub

Sub TampilanAwal()

'Menampilkan Data ke Grid
    Set RSLevel = New ADODB.Recordset

    RSLevel.Open "SELECT kode,nama" & _
               " From pos_level WHERE id>000", Conn, adOpenForwardOnly, adLockReadOnly
    Call TampilGrid

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

    Call TampilanAwal
End Sub

Private Sub Form_Resize()
'On Error Resume Next
    Call AutoResize
End Sub

Public Sub AutoResize()
'On Error Resume Next
    shpBar.Width = ScaleWidth

    With GridLevel
        .Left = 120
        .Top = shpBar.Height + 120
        .Height = (FormMain.PictureBeranda.ScaleHeight) - GridLevel.Top - Picture1.Height - 120 - 975 - 120
        .Width = 4515
    End With

    With FrameLevel
        .Left = 120
        .Top = FormMain.PictureBeranda.ScaleHeight - Picture1.Height - 120 - .Height
    End With

    With FrameAkses
        .Left = GridLevel.Width + 240
        .Top = shpBar.Height + 120
        .Width = FormMain.PictureBeranda.ScaleWidth - GridLevel.Width - 360
        .Height = (FormMain.PictureBeranda.ScaleHeight) - GridLevel.Top - Picture1.Height - 120
    End With

    With TabAkses
        .Left = 120
        .Top = 480
        .Width = FrameAkses.Width - 240
        .Height = FrameAkses.Height - 600
    End With

    With PictureFrame
        .Left = 0
        .Top = 330
        .Width = TabAkses.Width
        .Height = TabAkses.Height - 330
    End With

    Picture1.Move 0, FrameAkses.Height + shpBar.Height + 240, shpBar.Width

    ButtonKeluar.Move FormMain.PictureBeranda.ScaleWidth - ButtonKeluar.Width - 180, 90

    For i = 0 To mLevel.Count - 1
        mLevel(i).Move 180 + (i * 1875) + (i * 120), 90, 1875
    Next
End Sub

Private Sub GridLevel_Click()
    Dim MyRS As New ADODB.Recordset
    Set MyRS = New ADODB.Recordset
    MyRS.Open "SELECT * FROM pos_level Where kode Like '" & GridLevel.TextMatrix(GridLevel.Row, 0) & "%'", Conn, adOpenDynamic, adLockOptimistic
    If GridLevel.Rows > 1 Then
        With FormLevel
            blnPilih = True
            .chkKaryView.Value = GantiTF(MyRS.Fields("karyawan_view"))
            .chkKaryAdd.Value = GantiTF(MyRS.Fields("karyawan_create"))
            .chkKaryEdit.Value = GantiTF(MyRS.Fields("karyawan_update"))
            .chkKaryDelete.Value = GantiTF(MyRS.Fields("karyawan_delete"))
            .chkPotView.Value = GantiTF(MyRS.Fields("potongan_view"))
            .chkPotAdd.Value = GantiTF(MyRS.Fields("potongan_create"))
            .chkPotEdit.Value = GantiTF(MyRS.Fields("potongan_update"))
            .chkPotDelete.Value = GantiTF(MyRS.Fields("potongan_delete"))
            .chkTunView.Value = GantiTF(MyRS.Fields("tunjangan_view"))
            .chkTunAdd.Value = GantiTF(MyRS.Fields("tunjangan_create"))
            .chkTunEdit.Value = GantiTF(MyRS.Fields("tunjangan_update"))
            .chkTunDelete.Value = GantiTF(MyRS.Fields("tunjangan_delete"))

            .chkGajView.Value = GantiTF(MyRS.Fields("gaji_view"))
            .chkGajAdd.Value = GantiTF(MyRS.Fields("gaji_create"))
            .chkGajEdit.Value = GantiTF(MyRS.Fields("gaji_update"))
            .chkGajDelete.Value = GantiTF(MyRS.Fields("gaji_delete"))
            .chkGajiGenerate.Value = GantiTF(MyRS.Fields("gaji_generate"))
            .chkPengView.Value = GantiTF(MyRS.Fields("pengeluaran_view"))
            .chkPengAdd.Value = GantiTF(MyRS.Fields("pengeluaran_create"))
            .chkPengEdit.Value = GantiTF(MyRS.Fields("pengeluaran_update"))
            .chkPengDelete.Value = GantiTF(MyRS.Fields("pengeluaran_delete"))
            .chkPemView.Value = GantiTF(MyRS.Fields("pemasukan_view"))
            .chkPemAdd.Value = GantiTF(MyRS.Fields("pemasukan_create"))
            .chkPemEdit.Value = GantiTF(MyRS.Fields("pemasukan_update"))
            .chkPemDelete.Value = GantiTF(MyRS.Fields("pemasukan_delete"))

            .chkKasView.Value = GantiTF(MyRS.Fields("kas_view"))
            .chkKasAdd.Value = GantiTF(MyRS.Fields("kas_create"))
            .chkKasEdit.Value = GantiTF(MyRS.Fields("kas_update"))
            .chkKasDelete.Value = GantiTF(MyRS.Fields("kas_delete"))
            .chkBiayaView.Value = GantiTF(MyRS.Fields("biaya_view"))
            .chkBiayaAdd.Value = GantiTF(MyRS.Fields("biaya_create"))
            .chkBiayaEdit.Value = GantiTF(MyRS.Fields("biaya_update"))
            .chkBiayaDelete.Value = GantiTF(MyRS.Fields("biaya_delete"))
            .chkInvView.Value = GantiTF(MyRS.Fields("inventaris_view"))
            .chkInvAdd.Value = GantiTF(MyRS.Fields("inventaris_create"))
            .chkInvEdit.Value = GantiTF(MyRS.Fields("inventaris_update"))
            .chkInvDelete.Value = GantiTF(MyRS.Fields("inventaris_delete"))

            .chkUserView.Value = GantiTF(MyRS.Fields("user_view"))
            .chkUserAdd.Value = GantiTF(MyRS.Fields("user_create"))
            .chkUserEdit.Value = GantiTF(MyRS.Fields("user_update"))
            .chkUserDelete.Value = GantiTF(MyRS.Fields("user_delete"))
            .chkUserPass.Value = GantiTF(MyRS.Fields("user_change"))
            .chkLevelView.Value = GantiTF(MyRS.Fields("level_view"))
            .chkLevelAdd.Value = GantiTF(MyRS.Fields("level_create"))
            .chkLevelEdit.Value = GantiTF(MyRS.Fields("level_update"))
            .chkLevelDelete.Value = GantiTF(MyRS.Fields("level_delete"))
        End With
    End If
End Sub

Private Sub mLevel_Click(Index As Integer)
    Select Case Index
    Case 0
        blnTambah = True
        mLevel(0).Enabled = False
        mLevel(1).Enabled = False
        mLevel(2).Enabled = True
        mLevel(3).Enabled = True
        mLevel(4).Enabled = False
        TextKode.Enabled = True
        TextNama.Enabled = True
        TextKode.Text = ""
        TextNama.Text = ""
        LabelKode.Caption = ""
        TextKode.SetFocus
    Case 1
        Dim MyRS As New ADODB.Recordset
        Set MyRS = New ADODB.Recordset
        MyRS.Open "SELECT * FROM pos_level Where kode Like '" & GridLevel.TextMatrix(GridLevel.Row, 0) & "%'", Conn, adOpenDynamic, adLockOptimistic
        If GridLevel.Rows > 1 Then
            With FormLevel
                blnTambah = False
                .mLevel(0).Enabled = False
                .mLevel(1).Enabled = False
                .mLevel(2).Enabled = True
                .mLevel(3).Enabled = True
                .mLevel(4).Enabled = False
                .TextKode.Enabled = True
                .TextNama.Enabled = True
                .TextKode.Text = MyRS.Fields("kode")
                .LabelKode.Caption = MyRS.Fields("kode")
                .TextNama.Text = MyRS.Fields("nama")
            End With
        End If
    Case 2
        mLevel(0).Enabled = True
        mLevel(1).Enabled = True
        mLevel(2).Enabled = False
        mLevel(3).Enabled = False
        mLevel(4).Enabled = True
        TextKode.Text = ""
        TextNama.Text = ""
        LabelKode.Caption = ""
        TextKode.Enabled = False
        TextNama.Enabled = False
    Case 3
        If blnTambah = True Then
            If Trim(TextKode.Text) = "" Or Trim(TextNama.Text) = "" Then
                Beep
                Exit Sub
            End If

            Set ClassLevel = New ClassLevel
            Sukses = ClassLevel.AddLevel(TextKode.Text, TextNama.Text, _
                                         Conn)
            If Sukses Then
                FormLevel.TampilanAwal
                FormUser.TampilanAwal
                mLevel(0).Enabled = True
                mLevel(1).Enabled = True
                mLevel(2).Enabled = False
                mLevel(3).Enabled = False
                mLevel(4).Enabled = True
                TextKode.Text = ""
                TextNama.Text = ""
                LabelKode.Caption = ""
                TextKode.Enabled = False
                TextNama.Enabled = False
            Else
                MsgBox "Data Level Gagal Disimpan", vbExclamation, "Peringatan"
            End If
        Else
            If Trim(TextKode.Text) = "" Or Trim(TextNama.Text) = "" Then
                Beep
                Exit Sub
            End If

            Set ClassLevel = New ClassLevel
            Sukses = ClassLevel.UpdateLevel(LabelKode.Caption, TextKode.Text, TextNama.Text, Ganti01(chkKaryView.Value), Ganti01(chkKaryAdd.Value), Ganti01(chkKaryEdit.Value), Ganti01(chkKaryDelete.Value), Ganti01(chkPotView.Value), Ganti01(chkPotAdd.Value), Ganti01(chkPotEdit.Value), Ganti01(chkPotDelete.Value), Ganti01(chkTunView.Value), Ganti01(chkTunAdd.Value), Ganti01(chkTunEdit.Value), Ganti01(chkTunDelete.Value), Ganti01(chkGajView.Value), Ganti01(chkGajAdd.Value), Ganti01(chkGajEdit.Value), Ganti01(chkGajDelete.Value), Ganti01(chkGajiGenerate.Value), _
                                            Ganti01(chkPengView.Value), Ganti01(chkPengAdd.Value), Ganti01(chkPengEdit.Value), Ganti01(chkPengDelete.Value), Ganti01(chkPemView.Value), Ganti01(chkPemAdd.Value), Ganti01(chkPemEdit.Value), Ganti01(chkPemDelete.Value), Ganti01(chkKasView.Value), Ganti01(chkKasAdd.Value), Ganti01(chkKasEdit.Value), Ganti01(chkKasDelete.Value), Ganti01(chkBiayaView.Value), Ganti01(chkBiayaAdd.Value), Ganti01(chkBiayaEdit.Value), Ganti01(chkBiayaDelete.Value), _
                                            Ganti01(chkInvView.Value), Ganti01(chkInvAdd.Value), Ganti01(chkInvEdit.Value), Ganti01(chkInvDelete.Value), Ganti01(chkUserView.Value), Ganti01(chkUserAdd.Value), Ganti01(chkUserEdit.Value), Ganti01(chkUserDelete.Value), Ganti01(chkUserPass.Value), Ganti01(chkLevelView.Value), Ganti01(chkLevelAdd.Value), Ganti01(chkLevelEdit.Value), Ganti01(chkLevelDelete.Value), _
                                            Conn)
            If Sukses Then
                FormLevel.TampilanAwal
                FormUser.TampilanAwal
                mLevel(0).Enabled = True
                mLevel(1).Enabled = True
                mLevel(2).Enabled = False
                mLevel(3).Enabled = False
                mLevel(4).Enabled = True
                TextKode.Text = ""
                TextNama.Text = ""
                LabelKode.Caption = ""
                TextKode.Enabled = False
                TextNama.Enabled = False
            Else
                MsgBox "Data Level Gagal Diperbaharui", vbExclamation, "Peringatan"
            End If
        End If
        Set ClassLevel = Nothing
    Case 4
        If blnPilih = False Then
            MsgBox "Data Level belum dipilih", vbInformation, "Informasi"
            Exit Sub
        End If

        Pesan_Peringatan "Question", "Apakah Data Level dengan Nama " & Chr(34) & GridLevel.TextMatrix(GridLevel.Row, 1) & Chr(34) & " Ingin dihapus ?", "Konfirmasi"
        If Respon = "Iya" Then
            Set ClassLevel = New ClassLevel
            Sukses = ClassLevel.DeleteLevel(GridLevel.TextMatrix(GridLevel.Row, 0), Conn)
            If Sukses = True Then
                FormLevel.TampilanAwal
                FormUser.TampilanAwal
            Else
                MsgBox "Data Level Gagal Dihapus", vbExclamation, "Peringatan"
            End If
        End If
    End Select
End Sub

Private Sub TextKode_KeyPress(KeyAscii As Integer)
    GantiPetik KeyAscii
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TextNama_KeyPress(KeyAscii As Integer)
    GantiPetik KeyAscii
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
