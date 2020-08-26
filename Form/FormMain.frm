VERSION 5.00
Begin VB.Form FormMain 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SIPLITE"
   ClientHeight    =   9000
   ClientLeft      =   1380
   ClientTop       =   1125
   ClientWidth     =   16590
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9402.556
   ScaleMode       =   0  'User
   ScaleWidth      =   16590
   Begin VB.PictureBox PictureSplash 
      Appearance      =   0  'Flat
      BackColor       =   &H00585DF1&
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
      Height          =   5595
      Left            =   2700
      ScaleHeight     =   5595
      ScaleWidth      =   10395
      TabIndex        =   95
      Top             =   1440
      Width           =   10395
      Begin VB.PictureBox PictureTengah 
         Appearance      =   0  'Flat
         BackColor       =   &H00585DF1&
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
         Height          =   3495
         Left            =   780
         ScaleHeight     =   3495
         ScaleWidth      =   8955
         TabIndex        =   96
         Top             =   120
         Width           =   8955
         Begin VB.Label LabelCaption 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "New Experience With Flat User Interface User, Friendly for New User."
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
            Height          =   435
            Left            =   240
            TabIndex        =   97
            Top             =   3720
            Width           =   8475
         End
         Begin VB.Image ImageSlide 
            Appearance      =   0  'Flat
            Height          =   3420
            Left            =   60
            Picture         =   "FormMain.frx":000C
            Stretch         =   -1  'True
            Top             =   60
            Width           =   8835
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00C0C0C0&
            X1              =   8820
            X2              =   120
            Y1              =   3540
            Y2              =   3540
         End
      End
      Begin VB.Timer TimerLoad 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   360
         Top             =   960
      End
      Begin Project1.N_ProgressBar BarLoad 
         Height          =   75
         Left            =   300
         TabIndex        =   98
         Top             =   5340
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   132
      End
      Begin VB.Label LabelStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Membuka Koneksi"
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
         Left            =   5700
         TabIndex        =   100
         Top             =   4560
         Width           =   1545
      End
      Begin VB.Label LabelVer 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Version: 1.0.0.1"
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
         Left            =   8940
         TabIndex        =   99
         Top             =   5280
         Width           =   1275
      End
   End
   Begin VB.PictureBox PictureLogin 
      Appearance      =   0  'Flat
      BackColor       =   &H00312D22&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6855
      Left            =   8640
      ScaleHeight     =   6855
      ScaleWidth      =   11235
      TabIndex        =   73
      Top             =   1680
      Visible         =   0   'False
      Width           =   11235
      Begin VB.Frame FrameLogin 
         Appearance      =   0  'Flat
         BackColor       =   &H00191919&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         ForeColor       =   &H80000008&
         Height          =   4695
         Left            =   1200
         TabIndex        =   74
         Top             =   600
         Width           =   5175
         Begin VB.CheckBox CheckRemember 
            Appearance      =   0  'Flat
            BackColor       =   &H00191919&
            Caption         =   "&Remember Me"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   900
            TabIndex        =   85
            Top             =   3420
            Width           =   3435
         End
         Begin VB.TextBox TextPassword 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1230
            TabIndex        =   84
            Top             =   2940
            Width           =   3105
         End
         Begin VB.TextBox TextUsername 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1230
            TabIndex        =   83
            Top             =   2460
            Width           =   3105
         End
         Begin VB.PictureBox PictureFoto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1335
            Left            =   1920
            ScaleHeight     =   1335
            ScaleWidth      =   1335
            TabIndex        =   75
            Top             =   960
            Width           =   1335
            Begin VB.Image ImageFoto 
               Height          =   1335
               Left            =   0
               Picture         =   "FormMain.frx":9B67
               Stretch         =   -1  'True
               Top             =   0
               Width           =   1335
            End
         End
         Begin Project1.HMCommand ButtonLogin 
            Height          =   495
            Left            =   900
            TabIndex        =   86
            Top             =   3780
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   873
            BackColor       =   16777215
            BackColor       =   16777215
            BackColor       =   16777215
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderColor     =   16777215
            Caption         =   "&Login"
            MouseDownColor  =   14737632
         End
         Begin VB.Frame FrameLog 
            Appearance      =   0  'Flat
            BackColor       =   &H00BC8D3C&
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            ForeColor       =   &H0037C9FD&
            Height          =   795
            Left            =   0
            TabIndex        =   76
            Top             =   0
            Width           =   5175
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "LOGIN FORM"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   345
               Left            =   1665
               TabIndex        =   79
               Top             =   240
               Width           =   1935
            End
         End
         Begin VB.Label LabelForgot 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Forgot Password?"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   2820
            MouseIcon       =   "FormMain.frx":A66B
            MousePointer    =   99  'Custom
            TabIndex        =   78
            Top             =   3780
            Visible         =   0   'False
            Width           =   1290
         End
         Begin VB.Label LabelCreate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Create an Account."
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   2820
            MouseIcon       =   "FormMain.frx":A7BD
            MousePointer    =   99  'Custom
            TabIndex        =   77
            Top             =   4080
            Visible         =   0   'False
            Width           =   1410
         End
         Begin VB.Shape Shape4 
            BorderColor     =   &H00808080&
            Height          =   345
            Left            =   900
            Top             =   2940
            Width           =   345
         End
         Begin VB.Shape Shape3 
            BorderColor     =   &H00808080&
            Height          =   345
            Left            =   900
            Top             =   2460
            Width           =   345
         End
         Begin VB.Image Image5 
            Height          =   345
            Left            =   900
            Picture         =   "FormMain.frx":A90F
            Stretch         =   -1  'True
            Top             =   2940
            Width           =   345
         End
         Begin VB.Image Image4 
            Height          =   345
            Left            =   900
            Picture         =   "FormMain.frx":DDFF
            Stretch         =   -1  'True
            Top             =   2460
            Width           =   345
         End
      End
      Begin Project1.isButton ButtonKeluar 
         Cancel          =   -1  'True
         Height          =   555
         Left            =   180
         TabIndex        =   81
         Top             =   180
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   979
         Style           =   8
         Caption         =   "&Keluar"
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
      Begin VB.Label LabelInfoCopy 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright © 2017 All Rights Reserved | Design by CarlesneoID"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   3165
         TabIndex        =   80
         Top             =   6480
         Width           =   4545
      End
   End
   Begin VB.PictureBox UserProfil 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
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
      Height          =   4280
      Left            =   5580
      Picture         =   "FormMain.frx":1122A
      ScaleHeight     =   4245
      ScaleWidth      =   4215
      TabIndex        =   31
      Top             =   840
      Visible         =   0   'False
      Width           =   4240
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H00F9F9F9&
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
         Height          =   855
         Left            =   0
         ScaleHeight     =   855
         ScaleWidth      =   4215
         TabIndex        =   34
         Top             =   3420
         Width           =   4215
         Begin Project1.HMCommand TombolProfile 
            Height          =   555
            Left            =   180
            TabIndex        =   39
            Top             =   120
            Visible         =   0   'False
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   979
            BackColor       =   -2147483644
            BackColor       =   -2147483644
            BackColor       =   -2147483644
            ForeColor       =   8421504
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderColor     =   14737632
            Caption         =   "Proflie"
            MouseDownColor  =   12632256
         End
         Begin Project1.HMCommand TombolSignout 
            Height          =   555
            Left            =   2940
            TabIndex        =   40
            Top             =   120
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   979
            BackColor       =   -2147483644
            BackColor       =   -2147483644
            BackColor       =   -2147483644
            ForeColor       =   8421504
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderColor     =   14737632
            Caption         =   "Sign out"
            MouseDownColor  =   12632256
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   735
         Left            =   10
         ScaleHeight     =   735
         ScaleWidth      =   4215
         TabIndex        =   32
         Top             =   2640
         Width           =   4215
         Begin VB.Label LabelUser 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Manage User"
            ForeColor       =   &H00444449&
            Height          =   195
            Left            =   240
            MouseIcon       =   "FormMain.frx":4B304
            MousePointer    =   99  'Custom
            TabIndex        =   43
            Top             =   300
            Width           =   945
         End
         Begin VB.Label LabelLevel 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "User Level"
            ForeColor       =   &H00444449&
            Height          =   195
            Left            =   1680
            MouseIcon       =   "FormMain.frx":4B456
            MousePointer    =   99  'Custom
            TabIndex        =   42
            Top             =   300
            Width           =   750
         End
         Begin VB.Label LabelChangePW 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Change Password"
            ForeColor       =   &H00444449&
            Height          =   195
            Left            =   2700
            MouseIcon       =   "FormMain.frx":4B5A8
            MousePointer    =   99  'Custom
            TabIndex        =   41
            Top             =   300
            Width           =   1290
         End
      End
      Begin VB.PictureBox PictureWBG 
         Appearance      =   0  'Flat
         BackColor       =   &H00BC8D3C&
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
         Height          =   2655
         Left            =   0
         ScaleHeight     =   2655
         ScaleWidth      =   4215
         TabIndex        =   33
         Top             =   0
         Width           =   4215
         Begin VB.PictureBox PictureBgUser 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            Height          =   1215
            Left            =   1500
            ScaleHeight     =   1215
            ScaleWidth      =   1215
            TabIndex        =   38
            Top             =   240
            Width           =   1215
            Begin VB.Image ImageBgUser 
               Appearance      =   0  'Flat
               Height          =   1215
               Left            =   0
               Stretch         =   -1  'True
               Top             =   0
               Width           =   1215
            End
         End
         Begin VB.PictureBox PictureBgRound 
            Appearance      =   0  'Flat
            BackColor       =   &H00CAA463&
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
            Height          =   1335
            Left            =   1440
            ScaleHeight     =   1335
            ScaleWidth      =   1335
            TabIndex        =   37
            Top             =   180
            Width           =   1335
         End
         Begin VB.Label LabelAciveSince 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Active since Aug. 2014"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00F0E6CE&
            Height          =   240
            Left            =   1260
            TabIndex        =   36
            Top             =   1980
            Width           =   1845
         End
         Begin VB.Label LabelNameWork 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Administrator - Desktop Developer"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00F0E6CE&
            Height          =   300
            Left            =   255
            TabIndex        =   35
            Top             =   1620
            Width           =   3705
         End
      End
   End
   Begin VB.PictureBox PictureAPP 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00A87F36&
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
      Height          =   765
      Left            =   0
      ScaleHeight     =   765
      ScaleWidth      =   3465
      TabIndex        =   0
      Top             =   0
      Width           =   3465
      Begin VB.Label LabelAPP 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NeoLTE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   870
         TabIndex        =   9
         Top             =   180
         Width           =   1725
      End
   End
   Begin VB.PictureBox PictureSetting 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00312D22&
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
      Height          =   8595
      Left            =   10920
      ScaleHeight     =   8595
      ScaleWidth      =   3465
      TabIndex        =   13
      Top             =   770
      Width           =   3465
      Begin Project1.HMCommand ButtonAdvance 
         Height          =   375
         Left            =   180
         TabIndex        =   72
         Top             =   8160
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   661
         BackColor       =   3222818
         BackColor       =   3222818
         BackColor       =   3222818
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   16777215
         Caption         =   "&Advance [*]"
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00312D22&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   2475
         Left            =   60
         TabIndex        =   65
         Top             =   5160
         Width           =   3315
         Begin Project1.NEOComboFile ComboLanguage 
            Height          =   435
            Left            =   120
            TabIndex        =   67
            Top             =   360
            Width           =   3075
            _ExtentX        =   5424
            _ExtentY        =   767
            Text            =   ""
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
            BorderColor     =   12632256
            BackColor       =   3222818
            ColorSelect     =   4789739
            Theme           =   0
            Text            =   ""
         End
         Begin Project1.isButton BrowseBG 
            Height          =   390
            Left            =   2760
            TabIndex        =   68
            Top             =   1125
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   688
            Style           =   8
            Caption         =   " ..."
            IconAlign       =   3
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
         Begin VB.TextBox TextBackground 
            Appearance      =   0  'Flat
            BackColor       =   &H00312D22&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   180
            TabIndex        =   69
            Top             =   1200
            Width           =   2925
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00C0C0C0&
            Height          =   375
            Left            =   120
            Shape           =   4  'Rounded Rectangle
            Top             =   1140
            Width           =   3075
         End
         Begin VB.Label LabelBackground 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Background"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   120
            TabIndex        =   70
            Top             =   780
            Width           =   1320
         End
         Begin VB.Label LabelBahasa 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Language"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   120
            TabIndex        =   66
            Top             =   0
            Width           =   1080
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H00312D22&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
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
         Height          =   1035
         Left            =   60
         TabIndex        =   62
         Top             =   4140
         Width           =   3315
         Begin VB.CheckBox CheckBadge 
            Appearance      =   0  'Flat
            BackColor       =   &H00312D22&
            Caption         =   "Gunakan 4 Badge Informasi [*]"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   180
            TabIndex        =   71
            Top             =   720
            Width           =   2955
         End
         Begin VB.CheckBox CheckBeranda 
            Appearance      =   0  'Flat
            BackColor       =   &H00312D22&
            Caption         =   "Auto Beranda On Start [*]"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   180
            TabIndex        =   64
            Top             =   420
            Width           =   2955
         End
         Begin VB.Label LabelStart 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Startup"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   120
            TabIndex        =   63
            Top             =   60
            Width           =   840
         End
      End
      Begin VB.PictureBox FrameSkins 
         Appearance      =   0  'Flat
         BackColor       =   &H00312D22&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4215
         Left            =   60
         ScaleHeight     =   4215
         ScaleWidth      =   3375
         TabIndex        =   87
         Top             =   120
         Width           =   3375
         Begin Project1.NEOSkinSelect PictureSkin 
            Height          =   1095
            Index           =   5
            Left            =   1740
            TabIndex        =   88
            Top             =   2820
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   1931
            Color1          =   1660369
            Color2          =   3244004
            Themes          =   6
         End
         Begin Project1.NEOSkinSelect PictureSkin 
            Height          =   1095
            Index           =   3
            Left            =   1740
            TabIndex        =   89
            Top             =   1620
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   1931
            Color1          =   1219827
            Color2          =   1084891
            Themes          =   4
         End
         Begin Project1.NEOSkinSelect PictureSkin 
            Height          =   1095
            Index           =   1
            Left            =   1740
            TabIndex        =   90
            Top             =   420
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   1931
            Color1          =   4818688
            Color2          =   5940736
            Themes          =   2
         End
         Begin Project1.NEOSkinSelect PictureSkin 
            Height          =   1095
            Index           =   4
            Left            =   120
            TabIndex        =   91
            Top             =   2820
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   1931
            Color1          =   11225230
            Color2          =   11885723
            Themes          =   5
         End
         Begin Project1.NEOSkinSelect PictureSkin 
            Height          =   1095
            Index           =   2
            Left            =   120
            TabIndex        =   92
            Top             =   1620
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   1931
            Color1          =   11042614
            Color2          =   12356924
            Themes          =   0
         End
         Begin Project1.NEOSkinSelect PictureSkin 
            Height          =   1095
            Index           =   0
            Left            =   120
            TabIndex        =   93
            Top             =   420
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   1931
            Color1          =   2439639
            Color2          =   3754973
            Themes          =   1
         End
         Begin VB.Label LabelSkins 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Skins"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   120
            TabIndex        =   94
            Top             =   60
            Width           =   615
         End
         Begin VB.Shape ShapeSkin 
            BorderColor     =   &H00FFFF00&
            BorderWidth     =   3
            Height          =   1215
            Left            =   60
            Top             =   360
            Width           =   1635
         End
      End
      Begin VB.Label LabelKoneksi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Database Connect: "
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   780
         TabIndex        =   82
         Top             =   7860
         Width           =   1440
      End
   End
   Begin VB.PictureBox PictureSidebar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00312D22&
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
      Height          =   8595
      Left            =   0
      ScaleHeight     =   8595
      ScaleWidth      =   3465
      TabIndex        =   2
      Top             =   780
      Width           =   3465
      Begin Project1.XPFrame FrameNavigasi 
         Height          =   7455
         Left            =   0
         TabIndex        =   48
         Top             =   1140
         Width           =   3460
         _ExtentX        =   6112
         _ExtentY        =   13150
         HeaderLightColor=   33023
         HeaderDarkColor =   33023
         BackLightColor  =   12640511
         BackDarkColor   =   12640511
         BorderColor     =   16512
         TextColor       =   16777215
         Caption         =   ""
         HeaderHeight    =   40
         Curvature       =   0
         Begin VB.PictureBox PictureLine 
            Appearance      =   0  'Flat
            BackColor       =   &H00004080&
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
            Height          =   7455
            Left            =   3450
            ScaleHeight     =   7455
            ScaleWidth      =   15
            TabIndex        =   56
            Top             =   0
            Width           =   20
         End
         Begin Project1.N_Menu MenuNavigasi 
            Height          =   615
            Index           =   4
            Left            =   15
            TabIndex        =   54
            ToolTipText     =   "Pengeluaran"
            Top             =   3000
            Width           =   3810
            _ExtentX        =   6720
            _ExtentY        =   1085
            Picture_Normal  =   "FormMain.frx":4B6FA
            Picture_Down    =   "FormMain.frx":4BBB6
            Picture_Hover   =   "FormMain.frx":4C428
            Picture_Active  =   "FormMain.frx":4CC9A
            Caption         =   "Pengeluaran"
            BackColorNormal =   12640511
            BackColorHover  =   12091449
            BackColorDown   =   12640511
            BackColorActive =   3244004
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColorNormal =   16576
            ForeColorHover  =   16576
            ForeColorDown   =   16576
            ForeColorActive =   16777215
            Menu_Expandido  =   -1  'True
         End
         Begin Project1.N_Menu MenuNavigasi 
            Height          =   615
            Index           =   10
            Left            =   15
            TabIndex        =   53
            ToolTipText     =   "Tentang"
            Top             =   6600
            Width           =   3810
            _ExtentX        =   6720
            _ExtentY        =   1085
            Picture_Normal  =   "FormMain.frx":4D50C
            Picture_Down    =   "FormMain.frx":4DFF0
            Picture_Hover   =   "FormMain.frx":4E862
            Picture_Active  =   "FormMain.frx":4F0D4
            Caption         =   "Tentang"
            BackColorNormal =   12640511
            BackColorHover  =   12091449
            BackColorDown   =   12640511
            BackColorActive =   3244004
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColorNormal =   16576
            ForeColorHover  =   16576
            ForeColorDown   =   16576
            ForeColorActive =   16777215
            Menu_Expandido  =   -1  'True
         End
         Begin Project1.N_Menu MenuNavigasi 
            Height          =   615
            Index           =   3
            Left            =   15
            TabIndex        =   52
            ToolTipText     =   "Data Gaji"
            Top             =   2400
            Width           =   3810
            _ExtentX        =   6720
            _ExtentY        =   1085
            Picture_Normal  =   "FormMain.frx":4F946
            Picture_Down    =   "FormMain.frx":4FBE4
            Picture_Hover   =   "FormMain.frx":50456
            Picture_Active  =   "FormMain.frx":50CC8
            Caption         =   "Data Gaji"
            BackColorNormal =   12640511
            BackColorHover  =   12091449
            BackColorDown   =   12640511
            BackColorActive =   3244004
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColorNormal =   16576
            ForeColorHover  =   16576
            ForeColorDown   =   16576
            ForeColorActive =   16777215
            Menu_Expandido  =   -1  'True
         End
         Begin Project1.N_Menu MenuNavigasi 
            Height          =   615
            Index           =   2
            Left            =   15
            TabIndex        =   51
            ToolTipText     =   "Data Tunjangan"
            Top             =   1800
            Width           =   3810
            _ExtentX        =   6720
            _ExtentY        =   1085
            Picture_Normal  =   "FormMain.frx":5153A
            Picture_Down    =   "FormMain.frx":51A24
            Picture_Hover   =   "FormMain.frx":52296
            Picture_Active  =   "FormMain.frx":52780
            Caption         =   "Data Tunjangan"
            BackColorNormal =   12640511
            BackColorHover  =   12091449
            BackColorDown   =   12640511
            BackColorActive =   3244004
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColorNormal =   16576
            ForeColorHover  =   16576
            ForeColorDown   =   16576
            ForeColorActive =   16777215
            Menu_Expandido  =   -1  'True
         End
         Begin Project1.N_Menu MenuNavigasi 
            Height          =   615
            Index           =   1
            Left            =   15
            TabIndex        =   50
            ToolTipText     =   "Data Potongan"
            Top             =   1200
            Width           =   3810
            _ExtentX        =   6720
            _ExtentY        =   1085
            Picture_Normal  =   "FormMain.frx":52FF2
            Picture_Down    =   "FormMain.frx":534C3
            Picture_Hover   =   "FormMain.frx":53D35
            Picture_Active  =   "FormMain.frx":545A7
            Caption         =   "Data Potongan"
            BackColorNormal =   12640511
            BackColorHover  =   12091449
            BackColorDown   =   12640511
            BackColorActive =   3244004
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColorNormal =   16576
            ForeColorHover  =   16576
            ForeColorDown   =   16576
            ForeColorActive =   16777215
            Menu_Expandido  =   -1  'True
         End
         Begin Project1.N_Menu MenuNavigasi 
            Height          =   615
            Index           =   0
            Left            =   15
            TabIndex        =   49
            ToolTipText     =   "Data Karyawan"
            Top             =   600
            Width           =   3810
            _ExtentX        =   6720
            _ExtentY        =   1085
            Picture_Normal  =   "FormMain.frx":54E19
            Picture_Down    =   "FormMain.frx":5568B
            Picture_Hover   =   "FormMain.frx":55EFD
            Picture_Active  =   "FormMain.frx":5676F
            Caption         =   "Data Karyawan"
            BackColorNormal =   12640511
            BackColorHover  =   12091449
            BackColorDown   =   12640511
            BackColorActive =   3244004
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColorNormal =   16576
            ForeColorHover  =   16576
            ForeColorDown   =   16576
            ForeColorActive =   16777215
            Menu_Expandido  =   -1  'True
         End
         Begin Project1.N_Menu MenuNavigasi 
            Height          =   615
            Index           =   5
            Left            =   15
            TabIndex        =   57
            ToolTipText     =   "Pemasukan"
            Top             =   3600
            Width           =   3810
            _ExtentX        =   6720
            _ExtentY        =   1085
            Picture_Normal  =   "FormMain.frx":56A4D
            Picture_Down    =   "FormMain.frx":572BF
            Picture_Hover   =   "FormMain.frx":57B31
            Picture_Active  =   "FormMain.frx":583A3
            Caption         =   "Pemasukan"
            BackColorNormal =   12640511
            BackColorHover  =   12091449
            BackColorDown   =   12640511
            BackColorActive =   3244004
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColorNormal =   16576
            ForeColorHover  =   16576
            ForeColorDown   =   16576
            ForeColorActive =   16777215
            Menu_Expandido  =   -1  'True
         End
         Begin Project1.N_Menu MenuNavigasi 
            Height          =   615
            Index           =   6
            Left            =   15
            TabIndex        =   58
            ToolTipText     =   "Kas"
            Top             =   4200
            Width           =   3810
            _ExtentX        =   6720
            _ExtentY        =   1085
            Picture_Normal  =   "FormMain.frx":5887D
            Picture_Down    =   "FormMain.frx":58DAD
            Picture_Hover   =   "FormMain.frx":5961F
            Picture_Active  =   "FormMain.frx":59E91
            Caption         =   "Kas"
            BackColorNormal =   12640511
            BackColorHover  =   12091449
            BackColorDown   =   12640511
            BackColorActive =   3244004
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColorNormal =   16576
            ForeColorHover  =   16576
            ForeColorDown   =   16576
            ForeColorActive =   16777215
            Menu_Expandido  =   -1  'True
         End
         Begin Project1.N_Menu MenuNavigasi 
            Height          =   615
            Index           =   7
            Left            =   15
            TabIndex        =   59
            ToolTipText     =   "Biaya"
            Top             =   4800
            Width           =   3810
            _ExtentX        =   6720
            _ExtentY        =   1085
            Picture_Normal  =   "FormMain.frx":5A703
            Picture_Down    =   "FormMain.frx":5ABBA
            Picture_Hover   =   "FormMain.frx":5B42C
            Picture_Active  =   "FormMain.frx":5BC9E
            Caption         =   "Biaya"
            BackColorNormal =   12640511
            BackColorHover  =   12091449
            BackColorDown   =   12640511
            BackColorActive =   3244004
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColorNormal =   16576
            ForeColorHover  =   16576
            ForeColorDown   =   16576
            ForeColorActive =   16777215
            Menu_Expandido  =   -1  'True
         End
         Begin Project1.N_Menu MenuNavigasi 
            Height          =   615
            Index           =   8
            Left            =   15
            TabIndex        =   60
            ToolTipText     =   "Inventaris"
            Top             =   5400
            Width           =   3810
            _ExtentX        =   6720
            _ExtentY        =   1085
            Picture_Normal  =   "FormMain.frx":5C510
            Picture_Down    =   "FormMain.frx":5CA4E
            Picture_Hover   =   "FormMain.frx":5D2C0
            Picture_Active  =   "FormMain.frx":5DB32
            Caption         =   "Inventaris"
            BackColorNormal =   12640511
            BackColorHover  =   12091449
            BackColorDown   =   12640511
            BackColorActive =   3244004
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColorNormal =   16576
            ForeColorHover  =   16576
            ForeColorDown   =   16576
            ForeColorActive =   16777215
            Menu_Expandido  =   -1  'True
         End
         Begin Project1.N_Menu MenuNavigasi 
            Height          =   615
            Index           =   9
            Left            =   15
            TabIndex        =   61
            ToolTipText     =   "Laporan"
            Top             =   6000
            Width           =   3810
            _ExtentX        =   6720
            _ExtentY        =   1085
            Picture_Normal  =   "FormMain.frx":5E3A4
            Picture_Down    =   "FormMain.frx":5EC16
            Picture_Hover   =   "FormMain.frx":5F488
            Picture_Active  =   "FormMain.frx":5FCFA
            Caption         =   "Laporan"
            BackColorNormal =   12640511
            BackColorHover  =   12091449
            BackColorDown   =   12640511
            BackColorActive =   3244004
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColorNormal =   16576
            ForeColorHover  =   16576
            ForeColorDown   =   16576
            ForeColorActive =   16777215
            Menu_Expandido  =   -1  'True
         End
         Begin VB.Label LabelMenu 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "MENU NAVIGASI"
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Left            =   240
            TabIndex        =   55
            Top             =   240
            Width           =   1365
         End
      End
      Begin VB.PictureBox PictureProfil 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Height          =   800
         Left            =   180
         ScaleHeight     =   795
         ScaleWidth      =   795
         TabIndex        =   21
         Top             =   180
         Width           =   800
         Begin VB.Image ImageProfil 
            Appearance      =   0  'Flat
            Height          =   795
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   795
         End
      End
      Begin VB.PictureBox PictureUser 
         Appearance      =   0  'Flat
         BackColor       =   &H00312D22&
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
         Height          =   735
         Left            =   180
         ScaleHeight     =   735
         ScaleWidth      =   3075
         TabIndex        =   18
         Top             =   180
         Width           =   3075
         Begin VB.Shape Shape2 
            BorderColor     =   &H0000C000&
            FillColor       =   &H0000C000&
            FillStyle       =   0  'Solid
            Height          =   135
            Left            =   960
            Shape           =   3  'Circle
            Top             =   420
            Width           =   135
         End
         Begin VB.Label LabelUsername 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "admin"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   1200
            TabIndex        =   20
            Top             =   375
            Width           =   420
         End
         Begin VB.Label LabelNama 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Administrator"
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
            Height          =   240
            Left            =   960
            TabIndex        =   19
            Top             =   60
            Width           =   1350
         End
      End
   End
   Begin VB.PictureBox PictureHeader 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00BC8D3C&
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
      Height          =   765
      Left            =   3420
      ScaleHeight     =   765
      ScaleWidth      =   13065
      TabIndex        =   1
      Top             =   0
      Width           =   13065
      Begin Project1.N_Menu MenuUser 
         Height          =   855
         Left            =   10860
         TabIndex        =   30
         Top             =   0
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   1508
         Picture_Normal  =   "FormMain.frx":60218
         Picture_Down    =   "FormMain.frx":60234
         Picture_Hover   =   "FormMain.frx":60250
         Picture_Active  =   "FormMain.frx":6026C
         Caption         =   "Administrator"
         BackColorNormal =   12356924
         BackColorHover  =   11042614
         BackColorDown   =   11042614
         BackColorActive =   11042614
         FontBold        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColorNormal =   16777215
         ForeColorHover  =   16777215
         ForeColorDown   =   16777215
         ForeColorActive =   16777215
         Menu_Expandido  =   -1  'True
         Menu_Popup      =   -1  'True
      End
      Begin Project1.N_Image ButtonExpand 
         Height          =   765
         Index           =   2
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   1349
         Picture         =   "FormMain.frx":60288
         PictureHover    =   "FormMain.frx":625EC
         PictureDown     =   "FormMain.frx":64950
      End
      Begin VB.TextBox TS 
         Appearance      =   0  'Flat
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
         Height          =   255
         Left            =   2460
         TabIndex        =   15
         Text            =   "Text2"
         Top             =   120
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.TextBox TX 
         Appearance      =   0  'Flat
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
         Height          =   255
         Left            =   2460
         TabIndex        =   14
         Text            =   "Text3"
         Top             =   360
         Visible         =   0   'False
         Width           =   675
      End
      Begin Project1.N_Image TombolRestore 
         Height          =   285
         Left            =   9660
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   0
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   503
         Picture         =   "FormMain.frx":66CB4
         PictureHover    =   "FormMain.frx":66F2B
         PictureDown     =   "FormMain.frx":67186
      End
      Begin Project1.N_Image TombolMaximize 
         Height          =   285
         Left            =   9120
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   0
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   503
         Picture         =   "FormMain.frx":673F1
         PictureHover    =   "FormMain.frx":67648
         PictureDown     =   "FormMain.frx":6788E
      End
      Begin Project1.N_Image TombolKeluar 
         Height          =   285
         Left            =   10020
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   360
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   503
         Picture         =   "FormMain.frx":67ADE
         PictureHover    =   "FormMain.frx":67EB5
         PictureDown     =   "FormMain.frx":682A5
      End
      Begin Project1.N_Image TombolMinimize 
         Height          =   285
         Left            =   8640
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   0
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   503
         Picture         =   "FormMain.frx":68679
         PictureHover    =   "FormMain.frx":68837
         PictureDown     =   "FormMain.frx":689F5
      End
      Begin Project1.N_Image ButtonExpand 
         Height          =   765
         Index           =   3
         Left            =   1020
         TabIndex        =   25
         Top             =   0
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   1349
         Picture         =   "FormMain.frx":68BB8
         PictureHover    =   "FormMain.frx":6AF1C
         PictureDown     =   "FormMain.frx":6C9BA
      End
      Begin Project1.N_Image ButtonExpand 
         Height          =   765
         Index           =   1
         Left            =   780
         TabIndex        =   24
         Top             =   0
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   1349
         Picture         =   "FormMain.frx":6E458
         PictureHover    =   "FormMain.frx":6FEF6
         PictureDown     =   "FormMain.frx":7225A
      End
      Begin Project1.N_Image ButtonExpand 
         Height          =   765
         Index           =   0
         Left            =   240
         TabIndex        =   23
         Top             =   0
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   1349
         Picture         =   "FormMain.frx":745BE
         PictureHover    =   "FormMain.frx":76922
         PictureDown     =   "FormMain.frx":78C86
      End
      Begin Project1.N_Menu TombolPengaturan 
         Height          =   795
         Index           =   2
         Left            =   7380
         TabIndex        =   26
         Top             =   60
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   1402
         Picture_Normal  =   "FormMain.frx":7A724
         Picture_Down    =   "FormMain.frx":7C7C2
         Picture_Hover   =   "FormMain.frx":7E860
         Picture_Active  =   "FormMain.frx":808FE
         Caption         =   "1"
         BackColorNormal =   12356924
         BackColorHover  =   12356924
         BackColorDown   =   12356924
         BackColorActive =   12356924
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColorNormal =   16777215
         ForeColorHover  =   16777215
         ForeColorDown   =   16777215
         ForeColorActive =   16777215
      End
      Begin Project1.N_Menu TombolPengaturan 
         Height          =   795
         Index           =   0
         Left            =   6300
         TabIndex        =   27
         Top             =   0
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   1402
         Picture_Normal  =   "FormMain.frx":8299C
         Picture_Down    =   "FormMain.frx":84A3A
         Picture_Hover   =   "FormMain.frx":86AD8
         Picture_Active  =   "FormMain.frx":88B76
         Caption         =   "1"
         BackColorNormal =   3754973
         BackColorHover  =   3754973
         BackColorDown   =   3754973
         BackColorActive =   3754973
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColorNormal =   16777215
         ForeColorHover  =   16777215
         ForeColorDown   =   16777215
         ForeColorActive =   16777215
      End
      Begin Project1.N_Menu TombolPengaturan 
         Height          =   795
         Index           =   1
         Left            =   5820
         TabIndex        =   28
         Top             =   0
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   1402
         Picture_Normal  =   "FormMain.frx":8AC14
         Picture_Down    =   "FormMain.frx":8CCB2
         Picture_Hover   =   "FormMain.frx":8ED50
         Picture_Active  =   "FormMain.frx":90DEE
         Caption         =   "1"
         BackColorNormal =   5940736
         BackColorHover  =   5940736
         BackColorDown   =   5940736
         BackColorActive =   5940736
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColorNormal =   16777215
         ForeColorHover  =   16777215
         ForeColorDown   =   16777215
         ForeColorActive =   16777215
      End
      Begin Project1.N_Menu TombolPengaturan 
         Height          =   795
         Index           =   3
         Left            =   5580
         TabIndex        =   29
         Top             =   60
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   1402
         Picture_Normal  =   "FormMain.frx":92E8C
         Picture_Down    =   "FormMain.frx":94F2A
         Picture_Hover   =   "FormMain.frx":96FC8
         Picture_Active  =   "FormMain.frx":99066
         Caption         =   "1"
         BackColorNormal =   1219827
         BackColorHover  =   1219827
         BackColorDown   =   1219827
         BackColorActive =   1219827
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColorNormal =   16777215
         ForeColorHover  =   16777215
         ForeColorDown   =   16777215
         ForeColorActive =   16777215
      End
      Begin Project1.N_Image ButtonExpand 
         Height          =   765
         Index           =   4
         Left            =   1200
         TabIndex        =   44
         Top             =   0
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   1349
         Picture         =   "FormMain.frx":9B104
         PictureHover    =   "FormMain.frx":9D468
         PictureDown     =   "FormMain.frx":9F7CC
      End
      Begin Project1.N_Menu TombolPengaturan 
         Height          =   795
         Index           =   5
         Left            =   4620
         TabIndex        =   45
         Top             =   0
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   1402
         Picture_Normal  =   "FormMain.frx":A1B30
         Picture_Down    =   "FormMain.frx":A4694
         Picture_Hover   =   "FormMain.frx":A71F8
         Picture_Active  =   "FormMain.frx":A9D5C
         Caption         =   "1"
         BackColorNormal =   3244004
         BackColorHover  =   3244004
         BackColorDown   =   3244004
         BackColorActive =   3244004
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColorNormal =   16777215
         ForeColorHover  =   16777215
         ForeColorDown   =   16777215
         ForeColorActive =   16777215
      End
      Begin Project1.N_Image ButtonExpand 
         Height          =   765
         Index           =   5
         Left            =   1500
         TabIndex        =   46
         Top             =   0
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   1349
         Picture         =   "FormMain.frx":AC8C0
         PictureHover    =   "FormMain.frx":AEC24
         PictureDown     =   "FormMain.frx":B0F88
      End
      Begin Project1.N_Menu TombolPengaturan 
         Height          =   795
         Index           =   4
         Left            =   4860
         TabIndex        =   47
         Top             =   0
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   1402
         Picture_Normal  =   "FormMain.frx":B32EC
         Picture_Down    =   "FormMain.frx":B5E50
         Picture_Hover   =   "FormMain.frx":B89B4
         Picture_Active  =   "FormMain.frx":BB518
         Caption         =   "1"
         BackColorNormal =   11885723
         BackColorHover  =   11885723
         BackColorDown   =   11885723
         BackColorActive =   11885723
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColorNormal =   16777215
         ForeColorHover  =   16777215
         ForeColorDown   =   16777215
         ForeColorActive =   16777215
      End
   End
   Begin VB.PictureBox PictureFooter 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00191919&
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
      Height          =   750
      Left            =   3120
      ScaleHeight     =   750
      ScaleWidth      =   11490
      TabIndex        =   3
      Top             =   8460
      Width           =   11490
      Begin VB.Image TombolResize 
         Height          =   105
         Left            =   6120
         Picture         =   "FormMain.frx":BD5B6
         Top             =   480
         Width           =   105
      End
      Begin VB.Label LabelVersion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   5760
         TabIndex        =   17
         Top             =   180
         Width           =   720
      End
      Begin VB.Label LabelNumber 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1.11.17"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   6600
         TabIndex        =   16
         Top             =   180
         Width           =   645
      End
      Begin VB.Label LabelCopyLegal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ". All rights reserved."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   3240
         TabIndex        =   12
         Top             =   240
         Width           =   1860
      End
      Begin VB.Label LabelCopyURL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RiauKode"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00B88039&
         Height          =   270
         Left            =   2460
         MouseIcon       =   "FormMain.frx":BD6A0
         MousePointer    =   99  'Custom
         TabIndex        =   11
         Top             =   300
         Width           =   885
      End
      Begin VB.Label LabelCopy 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright © 2014-2017"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   180
         TabIndex        =   10
         Top             =   240
         Width           =   2100
      End
   End
   Begin VB.PictureBox PictureBeranda 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
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
      Height          =   5415
      Left            =   3480
      ScaleHeight     =   5415
      ScaleWidth      =   9315
      TabIndex        =   22
      Top             =   780
      Width           =   9315
      Begin VB.Image ImageBackground 
         Height          =   735
         Left            =   480
         Stretch         =   -1  'True
         Top             =   480
         Width           =   1875
      End
   End
End
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer

'Sesuaikan Form untuk menampilkan jendela task bar, layar penuh
Private Const SPI_GETWORKAREA = 48
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long

Const WS_CAPTION = &HC00000
Const WS_SYSMENU = &H80000
Const WS_MINIMIZEBOX = &H20000
Const WS_MAXIMIZEBOX = &H10000
Const GWL_STYLE = (-16)
Const WS_THICKFRAME = &H40000

Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Public Sidebar_Mini As Boolean
Dim Sidebar_Setting As Boolean
Dim Layar_Penuh As Boolean

Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32.dll" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

'Awal Deklarasi Shadow Form
Private Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const CS_DROPSHADOW As Long = &H20000
Private Const GCL_STYLE As Long = -26

'----------------------------------- Movable Form ------------------------------
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Const WM_NCLBUTTONDOWN = &HA1
Const HTBOTTOMRIGHT = 17
Const HTCAPTION = 2
'----------------------------------- Movable Form ------------------------------

Dim ClassUser As ClassUser

Private Sub DropShadow(ByVal hWnd As Long)
    Call SetClassLong(hWnd, GCL_STYLE, GetClassLong(hWnd, GCL_STYLE) Or CS_DROPSHADOW)
End Sub
'Akhir Deklarasi Shadow Form

Private Sub CreateRoundPicBox(tmpPic As PictureBox)

    Dim Ix As Long, Iy As Long
    Dim tmpRgn As Long

    Ix = ScaleX(tmpPic.Width, vbTwips, vbPixels) - 1
    Iy = ScaleY(tmpPic.Height, vbTwips, vbPixels) - 1

    tmpRgn = CreateEllipticRgn(0, 0, Ix, Iy)
    SetWindowRgn tmpPic.hWnd, tmpRgn, True

End Sub

Private Sub BrowseBG_Click()
    Dim hImage As Long, tmpPic As StdPicture
    Dim hHandle As Long, imageData() As Byte, bytesRead As Long

    Dim cBrowser As ClassOpenSaveDialog
    Set cBrowser = New ClassOpenSaveDialog
    With cBrowser
        .CancelError = True
        .Filter = "Bitmaps|*.bmp|JPEGs|*.jpg;*jpeg|GIFs|*.gif"
        '.Filter = "Images Files|*.bmp;*.jpg;*.gif"
        .Flags = OFN_FILEMUSTEXIST
        .DialogTitle = "Select Image"
    End With
    On Error GoTo EH
    cBrowser.ShowOpen Me.hWnd
    On Error GoTo 0

    ' Note: You don't need to load bitmaps separately as done in this example.
    ' The LoadImageW call is shown only to let you see that we can pass a unicode filename
    '   to an API.  The API declaration must be tweaked to look for a Long vs String
    If cBrowser.FilterIndex = 1 Then    ' loaded a .bmp file?
        hImage = LoadImageW(0&, StrPtr(cBrowser.FileName), IMAGE_BITMAP, 0&, 0&, LR_LOADFROMFILE)
        If hImage Then Set tmpPic = HandleToStdPicture(hImage, vbPicTypeBitmap)
    Else    ' loaded a gif, jpg or possibly something else?
        hHandle = GetFileHandle(cBrowser.FileName)
        If hHandle <> INVALID_HANDLE_VALUE Then
            If hHandle Then
                bytesRead = GetFileSize(hHandle, ByVal 0&)
                If bytesRead Then
                    ReDim imageData(0 To bytesRead - 1)
                    ReadFile hHandle, imageData(0), bytesRead, bytesRead, ByVal 0&
                    If bytesRead > UBound(imageData) Then
                        Set tmpPic = ArrayToPicture(imageData(), 0, bytesRead)
                    End If
                End If
                CloseHandle hHandle
            End If
        End If
    End If
    If tmpPic Is Nothing Then
        MsgBox "Error loading that file", vbOKOnly + vbExclamation
    Else
        Set ImageBackground.Picture = tmpPic
        TextBackground.Text = cBrowser.FileName
        Call WriteINI("Settings", "Background", TextBackground.Text, (Lokasi_File_Konfigurasi))
    End If
EH:
    If err Then
        If err.Number <> CommonDialogErrorsEnum.CDERR_CANCELED Then
            MsgBox err.Description, vbOKOnly, "Error Encountered"
        End If
        err.Clear
    End If
End Sub

Private Sub ButtonAdvance_Click()
    If VersiPro = True Then
        TombolPengaturan_Click (0)
        Call Form_Diatas(FormAdvance, PictureBeranda)
    ElseIf VersiPro = False Then
        TombolPengaturan_Click (0)
    End If
End Sub

Private Sub ButtonExpand_Click(Index As Integer)
    LockWindowUpdate FormMain.hWnd
    If Sidebar_Mini = True Then
        For i = 810 To 3465 / 100 Step 100
            PictureSidebar.Width = i
            DoEvents
        Next

        For i = PictureSidebar.Width To 3465 Step 100
            PictureSidebar.Width = i
            DoEvents
        Next
        Sidebar_Mini = False
        Sidebar_Minis True
        Call WriteINI("Settings", "Sidebar", "0", (Lokasi_File_Konfigurasi))
        Call ResizeForm
    Else
        PictureSidebar.Width = 810
        Sidebar_Mini = True
        Sidebar_Minis False
        Call WriteINI("Settings", "Sidebar", "1", (Lokasi_File_Konfigurasi))
        Call ResizeForm
    End If
    InitTimer
    LockWindowUpdate 0
End Sub

Private Sub ButtonKeluar_Click()
    Unload Me
End Sub

Private Sub ButtonLogin_Click()
    Dim Remembers As String: Remembers = ReadINI("Settings", "Remember", Lokasi_File_Konfigurasi)
    If Remembers = "1" Then
        Call WriteINI("String", "StringCode", Encrypt(TextUsername.Text), Lokasi_File_Konfigurasi)
        'Else
        '    MsgBox "Error", vbOKOnly, "Error"
    End If
    MulaiLogin
    'PictureLogin.Visible = False
End Sub

Private Sub MulaiLogin()
    On Error GoTo Error_Code
    Dim TeksHasilCipher As String
    Dim Rs As ADODB.Recordset

    Dim strMenuAkses As String
    Dim strSQL As String

    strSQL = "SELECT pword, namauser, kode, jabatan, registered, gender " & _
             "FROM pos_akses " & _
             "WHERE uname = '" & TextUsername.Text & "'"
    Set Rs = Conn.Execute(strSQL)
    If Rs.EOF Then
        'MsgBox "User Anda tidak terdaftar", vbExclamation, "Peringatan"
        Pesan_Peringatan "Information", "User Anda tidak terdaftar!", "Peringatan"
    Else
        Set ClassUser = New ClassUser
        TeksHasilCipher = ClassUser.GetTeksHasilCipher(Encrypt(TextPassword.Text))
        If TeksHasilCipher <> Rs("pword") Then
            'MsgBox "Password Anda salah", vbExclamation, "Peringatan"
            Pesan_Peringatan "Information", "Password Anda salah!", "Peringatan"
        Else
            strUserAktif = TextUsername.Text
            strNamaUserAktif = Rs("namauser")
            strJabatan = Rs("jabatan")
            strRegistered = Rs("registered")
            strKodeUser = Rs("kode")

            LabelUsername.Caption = strUserAktif
            LabelNama.Caption = strNamaUserAktif
            MenuUser.Caption = strNamaUserAktif
            LabelNameWork.Caption = strNamaUserAktif & " - " & strJabatan
            LabelAciveSince.Caption = ReadINI("FormMain", "AktifSejak", Lokasi_File_Bahasa) & " " & Format(strRegistered, "MMMM. yyyy")

            Dim Gender As String: Gender = Rs("gender")
            If Gender = "LAKI-LAKI" Then
                ImageProfil.Picture = LoadPicture(App.Path & "\Resource\PIC\Male.bmp")
                ImageBgUser.Picture = LoadPicture(App.Path & "\Resource\PIC\Male.bmp")
            ElseIf Gender = "PEREMPUAN" Then
                ImageProfil.Picture = LoadPicture(App.Path & "\Resource\PIC\Female.bmp")
                ImageBgUser.Picture = LoadPicture(App.Path & "\Resource\PIC\Female.bmp")
            Else
                ImageProfil.Picture = LoadPicture(App.Path & "\Resource\PIC\Unknown.bmp")
                ImageBgUser.Picture = LoadPicture(App.Path & "\Resource\PIC\Unknown.bmp")
            End If

            Set ClassUser = New ClassUser
            Sukses = ClassUser.HakAkses(strKodeUser)

            PictureLogin.Visible = False
        End If
    End If
    Rs.Close
    Set Rs = Nothing

    Exit Sub
Error_Code:
    MsgBox err.Description & vbCrLf & vbCrLf & strSQL
End Sub

Private Sub CheckBadge_Click()
    If CheckBadge.Value = 0 Then
        Call WriteINI("Settings", "Badge", "False", Lokasi_File_Konfigurasi)
    ElseIf CheckBadge.Value = 1 Then
        Call WriteINI("Settings", "Badge", "True", Lokasi_File_Konfigurasi)
    End If
End Sub

Private Sub CheckBeranda_Click()
    If CheckBeranda.Value = 0 Then
        Call WriteINI("Settings", "Beranda", "0", Lokasi_File_Konfigurasi)
    ElseIf CheckBeranda.Value = 1 Then
        Call WriteINI("Settings", "Beranda", "1", Lokasi_File_Konfigurasi)
    End If
End Sub

Private Sub CheckRemember_Click()
    If CheckRemember.Value = 1 Then
        Call WriteINI("Settings", "Remember", "1", Lokasi_File_Konfigurasi)
    Else
        Call WriteINI("Settings", "Remember", "0", Lokasi_File_Konfigurasi)
    End If
End Sub

Private Sub ComboLanguage_Click()
    On Error GoTo Error_Code

    Call WriteINI("Settings", "Language", ComboLanguage.Text, Lokasi_File_Konfigurasi)
    TerapkanBahasa

    Exit Sub
Error_Code:
    MsgBox err.Description & vbCrLf & vbCrLf & strSQL
End Sub

Private Sub Form_Load()
    On Error GoTo Error_Code
    Dim Lefts As String: Lefts = Me.ScaleWidth - PictureSetting.ScaleWidth
    Dim Leftx As String: Leftx = Me.ScaleWidth

    Call SetIcon(Me.hWnd, "APPICON", True)
    Call SetIcon(Me.hWnd, "FORMICON", False)

    Dim F As Long

    TS.Text = Lefts
    TX.Text = Leftx

    iTPPX& = Screen.TwipsPerPixelX
    iTPPY& = Screen.TwipsPerPixelY

    'Remove Border, Controlbox and Caption
    F = GetWindowLong(Me.hWnd, GWL_STYLE)
    F = F And Not (WS_THICKFRAME)
    F = F Xor WS_CAPTION
    F = SetWindowLong(Me.hWnd, GWL_STYLE, F)

    DropShadow Me.hWnd

    BukaMySQL
    
    If BukaMySQL = True Then LabelStatus.Caption = "Koneksi Database Berhasil"

    CenterForm Me
    Call BacaKonfigurasiForm
    Form_Resize
    TimerLoad.Enabled = True
    CreateRoundPicBox PictureFoto

    Call AutoBeranda
    CRemember
    
    InitTimer

    LabelVer.Caption = "Version: " & App.Major & "." & App.Minor & "." & App.Revision
    LabelNumber.Caption = App.Major & "." & App.Minor & "." & App.Revision
    LabelKoneksi.Caption = "Koneksi Database: " & Decrypt(ReadINI("ConnMySQL", "Server", Konfigurasi))

    Exit Sub
Error_Code:
    MsgBox err.Description & vbCrLf & vbCrLf & strSQL
End Sub

Sub AutoBeranda()
    Dim Beranda As String: Beranda = ReadINI("Settings", "Beranda", Lokasi_File_Konfigurasi)
    If Beranda = "0" Then
        CheckBeranda.Value = 0
    ElseIf Beranda = "1" Then
        CheckBeranda.Value = 1
        Call Form_Diatas(FormBeranda, PictureBeranda)
    End If

    Dim BadgeTrue As String: BadgeTrue = ReadINI("Settings", "Badge", Lokasi_File_Konfigurasi)
    If BadgeTrue = "False" Then
        CheckBadge.Value = 0
    ElseIf BadgeTrue = "True" Then
        CheckBadge.Value = 1
    End If
End Sub

Private Sub CRemember()
    Dim Remembers As String: Remembers = ReadINI("Settings", "Remember", Lokasi_File_Konfigurasi)
    If Remembers = "1" Then
        CheckRemember.Value = 1
        Me.TextUsername.Text = Decrypt(ReadINI("String", "StringCode", Lokasi_File_Konfigurasi))
        TextPassword.Text = ""
    Else
        CheckRemember.Value = 0
        Me.TextUsername.Text = ""
        TextPassword.Text = ""
    End If
End Sub

Private Sub LabelChangePW_Click()
    UserProfil.Visible = False
    FormPassword.Show vbModal
End Sub

Private Sub LabelLevel_Click()
    UserProfil.Visible = False
    Call Form_Diatas(FormLevel, PictureBeranda)
End Sub

Private Sub LabelUser_Click()
    UserProfil.Visible = False
    Call Form_Diatas(FormUser, PictureBeranda)
End Sub

Private Sub TextPassword_Change()
    With TextPassword
        .FontName = "Wingdings"
        .FontSize = 9
        .PasswordChar = "l"
    End With
End Sub

Private Sub TextPassword_KeyPress(KeyAscii As Integer)
    GantiPetik KeyAscii
    If KeyAscii = 13 Then ButtonLogin_Click
End Sub

Private Sub TextUsername_KeyPress(KeyAscii As Integer)
    GantiPetik KeyAscii
    If KeyAscii = 13 Then TextPassword.SetFocus
End Sub

Private Sub TimerLoad_Timer()
    BarLoad.Value = BarLoad.Value + 1
    If BarLoad.Value = BarLoad.Max Then
        TimerLoad.Enabled = False
        PictureSplash.Visible = False
        PictureLogin.Visible = True
    End If
End Sub

Private Sub ResizeForm()
    On Error Resume Next
    Dim i As Integer

    If Sidebar_Mini = True Then
        With PictureAPP
            .Move 0, 0, 810, 765
        End With

        With LabelAPP
            .Caption = "SIP"
            .Font = "Cornerstone"
            .FontBold = True
            .FontSize = "16"
            .Left = (PictureAPP.ScaleWidth * 0.5) - (.Width * 0.5)
            .Top = (PictureAPP.ScaleHeight * 0.5) - (.Height * 0.5)
        End With

        With LabelMenu
            .Caption = "NAV"
        End With

        With PictureProfil
            .Move 120, 180, 615, 615
        End With

        With ImageProfil
            .Move 0, 0, 615, 615
        End With

        With FrameNavigasi
            .Top = 950
            .Left = 0
            .Width = 3460
            .Height = Me.Height - PictureAPP.Height - 950
        End With

        With PictureLine
            .Left = 800
            .Height = Me.Height - PictureAPP.Height - 950
        End With

        CreateRoundPicBox PictureProfil
        CreateRoundPicBox PictureBgUser
        CreateRoundPicBox PictureBgRound
    Else
        With PictureAPP
            .Top = 0
            .Left = 0
            .Height = 765
            .Width = 3465
        End With

        With LabelAPP
            .Caption = "SIPLite"
            .Font = "Cornerstone"
            .FontBold = True
            .FontSize = "24"
            .Left = (PictureAPP.ScaleWidth * 0.5) - (.Width * 0.5)
            .Top = (PictureAPP.ScaleHeight * 0.5) - (.Height * 0.5)
        End With

        With LabelMenu
            .Caption = ReadINI("FormMain", "MenuNavigasi", Lokasi_File_Bahasa)    '"MENU NAVIGASI"
        End With

        With PictureProfil
            .Move 180, 180, 800, 800
        End With

        With ImageProfil
            .Move 0, 0, 795, 795
        End With

        With FrameNavigasi
            .Top = 1155
            .Left = 0
            .Width = 3460
            .Height = Me.Height - PictureAPP.Height - 1155
        End With

        With PictureLine
            .Left = 3450
            .Height = Me.Height - PictureAPP.Height - 1155
        End With

        CreateRoundPicBox PictureProfil
        CreateRoundPicBox PictureBgUser
        CreateRoundPicBox PictureBgRound
    End If

    With PictureSidebar
        .Top = PictureAPP.Top + PictureAPP.Height
        .Left = 0
        .Width = PictureAPP.Width
        .Height = Me.ScaleHeight - PictureAPP.Height
    End With

    With PictureHeader
        .Top = 0
        .Left = PictureAPP.Left + PictureAPP.Width
        .Width = Me.Width - PictureAPP.Width
        .Height = 765
    End With

    With TombolKeluar
        .Move PictureHeader.Width - .Width - 120, 0
    End With

    With TombolMaximize
        .Move TombolKeluar.Left - .Width, 0
    End With

    With TombolRestore
        .Move TombolMaximize.Left, TombolMaximize.Top
    End With

    With TombolMinimize
        .Move TombolMaximize.Left - .Width, 0
    End With

    For i = 0 To TombolPengaturan.Count - 1
        TombolPengaturan(i).Top = 0
        TombolPengaturan(i).Left = TombolMinimize.Left - TombolPengaturan(i).Width - 20
    Next i

    For i = 0 To ButtonExpand.Count - 1
        ButtonExpand(i).Top = 0
        ButtonExpand(i).Left = 0
    Next i

    With PictureFooter
        .Top = Me.ScaleHeight - .ScaleHeight
        .Left = PictureSidebar.Left + PictureSidebar.Width
        .Height = 750
        .Width = Me.ScaleWidth - PictureSidebar.Width
    End With

    With TombolResize
        .Move PictureFooter.Width - .Width - 25, PictureFooter.Height - .Height - 25
    End With

    With LabelCopy
        .Left = 180
        .Top = (PictureFooter.ScaleHeight * 0.5) - (.Height * 0.5)
    End With

    With LabelCopyURL
        .Left = LabelCopy.Left + LabelCopy.Width + 60
        .Top = (PictureFooter.ScaleHeight * 0.5) - (.Height * 0.5)
    End With

    With LabelCopyLegal
        .Left = LabelCopyURL.Left + LabelCopyURL.Width + 10
        .Top = (PictureFooter.ScaleHeight * 0.5) - (.Height * 0.5)
    End With

    With LabelNumber
        .Left = PictureFooter.Width - .Width - 180
        .Top = (PictureFooter.ScaleHeight * 0.5) - (.Height * 0.5)
    End With

    With LabelVersion
        .Left = LabelNumber.Left - .Width - 40
        .Top = (PictureFooter.ScaleHeight * 0.5) - (.Height * 0.5)
    End With

    With PictureSetting
        .Top = 770
        .Left = Me.ScaleWidth
        .Height = Me.ScaleHeight - PictureHeader.ScaleHeight
        .Width = 3465
    End With

    With ButtonAdvance
        .Top = PictureSetting.ScaleHeight - ButtonAdvance.Height - 180
        .Left = (PictureSetting.ScaleWidth * 0.5) - (.Width * 0.5)
    End With

    With LabelKoneksi
        .Top = PictureSetting.ScaleHeight - ButtonAdvance.Height - 400
        .Left = (PictureSetting.ScaleWidth * 0.5) - (.Width * 0.5)
    End With

    With PictureBeranda
        .Top = PictureHeader.Top + PictureHeader.Height
        .Left = PictureSidebar.Left + PictureSidebar.Width
        .Height = Me.ScaleHeight - PictureHeader.ScaleHeight - PictureFooter.ScaleHeight
        .Width = Me.Width - PictureSidebar.ScaleWidth
    End With

    With ImageBackground
        .Top = 0
        .Left = 0
        .Height = PictureBeranda.ScaleHeight
        .Width = PictureBeranda.ScaleWidth
    End With

    With ShapeSkin
        .Height = PictureSkin(0).Height + 4
        .Width = PictureSkin(0).Width + 4
    End With

    With MenuUser
        .Left = TombolPengaturan(0).Left - .Width
        .Top = TombolPengaturan(0).Top
    End With

    With UserProfil
        .Visible = False
    End With

    'Splash Screen
    With PictureSplash
        .Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    End With
    With PictureTengah
        .Move (PictureSplash.ScaleWidth * 0.5) - (.Width * 0.5), (PictureSplash.ScaleHeight * 0.5) - (.Height * 0.5)
    End With
    With BarLoad
        .Move PictureSplash.ScaleLeft, PictureSplash.ScaleHeight - .Height, PictureSplash.ScaleWidth, 120
    End With
    With LabelVer
        .Move PictureSplash.Left + 120, BarLoad.Top - .Height - 60
    End With
    With LabelStatus
        .Move PictureSplash.Left + PictureSplash.Width - .Width - 120, BarLoad.Top - .Height - 60
    End With
    'Splash Screen

    'Login Screen
    With PictureLogin
        .Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    End With

    With FrameLogin
        .Move (PictureLogin.ScaleWidth * 0.5) - (.Width * 0.5), (PictureLogin.ScaleHeight * 0.5) - (.Height * 0.5)
    End With

    With LabelInfoCopy
        .Move (PictureLogin.ScaleWidth * 0.5) - (.Width * 0.5), FrameLogin.Top + FrameLogin.Height + .Height + 240
    End With

    With ButtonKeluar
        .Move (PictureLogin.ScaleWidth * 0.5) - (.Width * 0.5), (PictureLogin.ScaleHeight * 0.5) - (.Height * 0.5)
    End With
    'Login Form
End Sub

Private Sub BacaKonfigurasiForm()
'Prosedur untuk melakukan verifikasi pengaturan form
    On Error GoTo Error_Code

    FormUI.Text_Layar_Penuh.Text = ReadINI("Settings", "FullScreen", Lokasi_File_Konfigurasi)
    FormUI.Label_Top.Caption = ReadINI("Settings", "Form_Top", Lokasi_File_Konfigurasi)
    FormUI.Label_Left.Caption = ReadINI("Settings", "Form_Left", Lokasi_File_Konfigurasi)
    FormUI.Label_Height.Caption = ReadINI("Settings", "Form_Height", Lokasi_File_Konfigurasi)
    FormUI.Label_Width.Caption = ReadINI("Settings", "Form_Width", Lokasi_File_Konfigurasi)

    TextBackground.Text = ReadINI("Settings", "Background", Lokasi_File_Konfigurasi)
    ComboLanguage.Text = ReadINI("Settings", "Language", Lokasi_File_Konfigurasi)

    Dim Exist As String: Exist = FileExists(ReadINI("Settings", "Background", Lokasi_File_Konfigurasi))
    If Exist = True Then
        ImageBackground.Picture = LoadPicture(ReadINI("Settings", "Background", Lokasi_File_Konfigurasi))
    Else
        ImageBackground.Picture = LoadPicture(App.Path & "\Resource\BG\BG.jpg")
    End If

    'Verificar as definições do programa
    If FormUI.Text_Layar_Penuh.Text <> Empty Then
        If FormUI.Text_Layar_Penuh.Text = "True" Then
            TombolMaximize_Click
        Else
            TombolRestore_Click
        End If
    End If


    Dim Sidebar_Menu As Integer: Sidebar_Menu = ReadINI("Settings", "Sidebar", Lokasi_File_Konfigurasi)
    If Sidebar_Menu = "0" Then
        Sidebar_Minis True
        Sidebar_Mini = False
    Else
        Sidebar_Mini = True
        Sidebar_Minis False
    End If

    Dim color As Integer: color = ReadINI("Settings", "Color", Lokasi_File_Konfigurasi)

    If color = 0 Then
        PictureSkin_Click (0)
    ElseIf color = 1 Then
        PictureSkin_Click (1)
    ElseIf color = 2 Then
        PictureSkin_Click (2)
    ElseIf color = 3 Then
        PictureSkin_Click (3)
    ElseIf color = 4 Then
        PictureSkin_Click (4)
    ElseIf color = 5 Then
        PictureSkin_Click (5)
    Else
        PictureSkin_Click (5)
    End If

    Call TerapkanBahasa
    Exit Sub
Error_Code:
    MsgBox err.Description & vbCrLf & vbCrLf & strSQL
End Sub

Public Function PosFormRelativeTaskBar(F As Form)
'Fungsi untuk memaksimalkan tampilan form tanpa menghilangkan taskbar menu
'Put the WindowsState=0 normal
    On Error Resume Next
    Dim WindowRect As RECT
    SystemParametersInfo SPI_GETWORKAREA, 0, WindowRect, 0
    SetWindowPos hWnd, 0, WindowRect.Left, WindowRect.Top, WindowRect.Right - WindowRect.Left, WindowRect.Bottom - WindowRect.Top, 0
    F.Top = WindowRect.Bottom * Screen.TwipsPerPixelY - F.Height
    F.Left = WindowRect.Right * Screen.TwipsPerPixelX - F.Width
End Function

Private Sub Sidebar_Minis(Value As Boolean)
'Procedimento para desactivar todos os menus
'On Error GoTo Corrige_Erro
    Dim i As Integer: For i = 0 To MenuNavigasi.Count - 1
        MenuNavigasi(i).Menu_Expandido = Value
    Next

    Exit Sub
Error_Code:
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.Height < 9810 Then
        Me.Height = "9810"    '"8805"
    End If
    Call ResizeForm
    If Me.Width < 19215 Then
        Me.Width = "19215"    '"14670"
    End If
    Call ResizeForm
    InitTimer
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Pesan_Peringatan "Question", "Yakin Ingin Keluar Program Sekarang?", "Keluar dari Program"
    If Respon = "Iya" Then
        Call WriteINI("Settings", "FullScreen", FormUI.Text_Layar_Penuh.Text, (Lokasi_File_Konfigurasi))
        Call WriteINI("Settings", "Form_Top", FormUI.Label_Top.Caption, (Lokasi_File_Konfigurasi))
        Call WriteINI("Settings", "Form_Left", FormUI.Label_Left.Caption, (Lokasi_File_Konfigurasi))
        Call WriteINI("Settings", "Form_Height", FormUI.Label_Height.Caption, (Lokasi_File_Konfigurasi))
        Call WriteINI("Settings", "Form_Width", FormUI.Label_Width.Caption, (Lokasi_File_Konfigurasi))
        Cancel = 0
        End
        Exit Sub
Error_Code:
    Else
        Cancel = 1
    End If
End Sub

Private Sub ImageProfil_Click()
'Code Goes Here
End Sub

Private Sub LabelCopyURL_Click()
    Pesan_Peringatan "Question", "Anda akan dialihkan ke laman Website" & vbNewLine & " Menggunakan Browser Default.", "Buka Link"
    If Respon = "Iya" Then
        ShellExecute hWnd, "open", "http://blog.carlesneo.id/", _
                     vbNullString, vbNullString, 1
    End If
End Sub

Private Sub MenuNavigasi_Click(Index As Integer)
    Dim A As Integer: For A = 0 To MenuNavigasi.Count - 1
        MenuNavigasi(A).Menu_Activo = False
    Next
    MenuNavigasi(Index).Menu_Activo = True

    Select Case Index
    Case 0
        Call Form_Diatas(FormKaryawan, PictureBeranda)
    Case 1
        Call Form_Diatas(FormPotongan, PictureBeranda)
    Case 2
        Call Form_Diatas(FormTunjangan, PictureBeranda)
    Case 3
        Call Form_Diatas(FormPenggajian, PictureBeranda)
    Case 4
        Call Form_Diatas(FormPengeluaran, PictureBeranda)
    Case 5
        Call Form_Diatas(FormPemasukan, PictureBeranda)
    Case 6
        Call Form_Diatas(FormKas, PictureBeranda)
    Case 7
        Call Form_Diatas(FormBiaya, PictureBeranda)
    Case 8
        Call Form_Diatas(FormInventaris, PictureBeranda)
    Case 9
        PopupMenu FormUI.mLaporan, , (MenuNavigasi(9).Left + MenuNavigasi(9).Width - 340)
    Case 10
        Call Form_Diatas(FormTentang, PictureBeranda)
    End Select
    
    InitTimer

    Set ClassUser = New ClassUser
    Dim User As String: User = ClassUser.HakAkses(strKodeUser)
End Sub

Private Sub MenuUser_Click()
    If UserProfil.Visible = False Then
        If Sidebar_Mini = False Then
            With UserProfil
                .Visible = True
                .Top = 770
                .Left = TombolPengaturan(0).Left - TombolPengaturan(0).Width + 40
            End With
        Else
            With UserProfil
                .Visible = True
                .Top = 770
                .Left = TombolPengaturan(0).Left - 2655 - TombolPengaturan(0).Width + 40
            End With
        End If
    Else
        With UserProfil
            .Visible = False
        End With
    End If
End Sub

Private Sub PictureHeader_DblClick()
    If Layar_Penuh = True Then
        TombolRestore_Click
    Else
        TombolMaximize_Click
    End If
End Sub

Private Sub PictureHeader_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then Capture_Posisi_Form FormMain
    If Layar_Penuh = False Then Me.MousePointer = 15
End Sub

Private Sub PictureHeader_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Layar_Penuh = False Then Move_Form FormMain
End Sub

Private Sub PictureHeader_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'Meletakkan form pada posisi akhir
    Large_Form FormMain
    Capture_Dimensi_Posisi_Form
    Me.MousePointer = 0
End Sub

Public Sub Capture_Dimensi_Posisi_Form()
'Prosedur untuk memperbaharui nilai-nilai preferensi program
'On Error GoTo Error_Code
    If Layar_Penuh = False Then
        FormUI.Label_Top.Caption = Me.Top
        FormUI.Label_Left.Caption = Me.Left
        FormUI.Label_Height.Caption = Me.Height
        FormUI.Label_Width.Caption = Me.Width
    End If

    Exit Sub
Error_Code:
End Sub

Private Sub PictureSkin_Click(Index As Integer)
    On Error Resume Next
    Dim i As Integer
    LockWindowUpdate FormMain.hWnd
    Select Case Index
    Case 0
        For i = 0 To TombolPengaturan.Count - 1
            TombolPengaturan(i).Visible = False
        Next i
        TombolPengaturan(Index).Visible = True

        For i = 0 To ButtonExpand.Count - 1
            ButtonExpand(i).Visible = False
        Next i
        ButtonExpand(Index).Visible = True

        With MenuUser
            .BackColorActive = &H2539D
            .BackColorDown = &H2539D7
            .BackColorHover = &H2539D7
            .BackColorNormal = &H394BDD
        End With

        Theme vbWhite, &H2539D7, &H394BDD, &H312D22, &H312D22, &H2539D7, &H616FE4, &H394BDD, &HD3D7F5, &H191919, &H25221A, &H312D22, &H312D22, &H312D22, &H25221A, &H25221A, &HFFFFFF

        Call WriteINI("Settings", "Color", "0", (Lokasi_File_Konfigurasi))
    Case 1
        For i = 0 To TombolPengaturan.Count - 1
            TombolPengaturan(i).Visible = False
        Next i
        TombolPengaturan(Index).Visible = True

        For i = 0 To ButtonExpand.Count - 1
            ButtonExpand(i).Visible = False
        Next i
        ButtonExpand(Index).Visible = True

        With MenuUser
            .BackColorActive = &H498700
            .BackColorDown = &H498700
            .BackColorHover = &H498700
            .BackColorNormal = &H5AA600
        End With

        Theme vbWhite, &H498700, &H5AA600, &H312D22, &H312D22, &H498700, &H7BB833, &H5AA600, &HDAEBBF, &H191919, &H25221A, &H312D22, &H312D22, &H312D22, &H25221A, &H25221A, &HFFFFFF

        Call WriteINI("Settings", "Color", "1", (Lokasi_File_Konfigurasi))
    Case 2
        For i = 0 To TombolPengaturan.Count - 1
            TombolPengaturan(i).Visible = False
        Next i
        TombolPengaturan(Index).Visible = True

        For i = 0 To ButtonExpand.Count - 1
            ButtonExpand(i).Visible = False
        Next i
        ButtonExpand(Index).Visible = True

        With MenuUser
            .BackColorActive = &HA87F36
            .BackColorDown = &HA87F36
            .BackColorHover = &HA87F36
            .BackColorNormal = &HBC8D3C
        End With

        Theme vbWhite, &HA87F36, &HBC8D3C, &H312D22, &H312D22, &HA87F36, &HCAA463, &HBC8D3C, &HF0E6CE, &H191919, &H25221A, &H312D22, &H312D22, &H312D22, &H25221A, &H25221A, &HFFFFFF

        Call WriteINI("Settings", "Color", "2", (Lokasi_File_Konfigurasi))
    Case 3
        For i = 0 To TombolPengaturan.Count - 1
            TombolPengaturan(i).Visible = False
        Next i
        TombolPengaturan(Index).Visible = True

        For i = 0 To ButtonExpand.Count - 1
            ButtonExpand(i).Visible = False
        Next i
        ButtonExpand(Index).Visible = True

        With MenuUser
            .BackColorActive = &HB8BDB
            .BackColorDown = &HB8BDB
            .BackColorHover = &HB8BDB
            .BackColorNormal = &H129CF3
        End With

        Theme vbWhite, &HB8BDB, &H129CF3, &H312D22, &H312D22, &HB8BDB, &H41B0F6, &H129CF3, &HB4E9FC, &H191919, &H25221A, &H312D22, &H312D22, &H312D22, &H25221A, &H25221A, &HFFFFFF

        Call WriteINI("Settings", "Color", "3", (Lokasi_File_Konfigurasi))
    Case 4
        For i = 0 To TombolPengaturan.Count - 1
            TombolPengaturan(i).Visible = False
        Next i
        TombolPengaturan(Index).Visible = True

        For i = 0 To ButtonExpand.Count - 1
            ButtonExpand(i).Visible = False
        Next i
        ButtonExpand(Index).Visible = True

        Theme vbWhite, &HAB488E, &HB55C9B, &H312D22, &H312D22, &HAB488E, &HE6C8DD, &HB55C9B, &HE6C8DD, &H191919, &H25221A, &H312D22, &H312D22, &H312D22, &H25221A, &H25221A, &HFFFFFF

        With MenuUser
            .BackColorActive = &HAB488E
            .BackColorDown = &HAB488E
            .BackColorHover = &HAB488E
            .BackColorNormal = &HB55C9B
        End With

        Call WriteINI("Settings", "Color", "4", (Lokasi_File_Konfigurasi))
    Case 5
        For i = 0 To TombolPengaturan.Count - 1
            TombolPengaturan(i).Visible = False
        Next i
        TombolPengaturan(Index).Visible = True

        For i = 0 To ButtonExpand.Count - 1
            ButtonExpand(i).Visible = False
        Next i
        ButtonExpand(Index).Visible = True

        Theme vbWhite, &H1955D1, &H317FE4, &H312D22, &H312D22, &H1955D1, &HBAD4F6, &H317FE4, &HBAD4F6, &H191919, &H317FE4, &HC0E0FF, &HC0E0FF, &HC0E0FF, &H4080&, &H80FF&, &H40C0&

        With MenuUser
            .BackColorActive = &H1955D1
            .BackColorDown = &H1955D1
            .BackColorHover = &H1955D1
            .BackColorNormal = &H317FE4
        End With

        Call WriteINI("Settings", "Color", "5", (Lokasi_File_Konfigurasi))
    End Select
    ShapeSkin.Left = PictureSkin(Index).Left - 2
    ShapeSkin.Top = PictureSkin(Index).Top - 2
    UserProfil.Visible = False
    LockWindowUpdate 0
End Sub

Private Sub PictureSplash_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then Capture_Posisi_Form FormMain
    If Layar_Penuh = False Then Me.MousePointer = 15
End Sub

Private Sub PictureSplash_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Layar_Penuh = False Then Move_Form FormMain
End Sub

Private Sub PictureSplash_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'Meletakkan form pada posisi akhir
    Large_Form FormMain
    Capture_Dimensi_Posisi_Form
    Me.MousePointer = 0
End Sub

Private Sub TombolKeluar_Click()
    Unload Me
End Sub

Private Sub TombolMaximize_Click()
    On Error Resume Next
    If Layar_Penuh = True Then
        With FormMain
            If FormUI.Label_Height.Caption <> "" Then .Height = FormUI.Label_Height.Caption Else .Height = 9420
            If FormUI.Label_Width.Caption <> "" Then .Width = FormUI.Label_Width.Caption Else .Width = 15000    '13380 '15585
            If FormUI.Label_Top.Caption <> "" Then .Top = FormUI.Label_Top.Caption Else .Top = (Screen.Height - Me.Height) / 2
            If FormUI.Label_Left.Caption <> "" Then .Left = FormUI.Label_Left.Caption Else .Left = (Screen.Width - Me.Width) / 2
        End With
        CenterForm Me
        Layar_Penuh = False
        FormUI.Text_Layar_Penuh.Text = "False"
        TombolMaximize.Visible = True
        TombolRestore.Visible = False
        'Call AutoResize
    Else
        PosFormRelativeTaskBar Me
        Layar_Penuh = True
        FormUI.Text_Layar_Penuh.Text = "True"
        TombolMaximize.Visible = False
        TombolRestore.Visible = True
        'Call AutoResize
    End If
    InitTimer
End Sub

Private Sub TombolMinimize_Click()
    Me.WindowState = 1
End Sub

Private Sub TombolPengaturan_Click(Index As Integer)
    Dim Lefts As String: Lefts = Me.Width - PictureSetting.Width
    Dim Leftx As String: Leftx = Me.Width + 160
    Dim i As Integer
    TS.Text = Lefts
    TX.Text = Leftx
    If PictureSetting.Left = TS.Text Then
        For i = TS.Text To TX.Text / 50 Step 5
            PictureSetting.Left = i
            DoEvents
        Next

        For i = TS.Text To TX.Text Step 100
            PictureSetting.Left = i
            DoEvents
        Next
        On Error Resume Next
        PictureSetting.SetFocus
        Me.SetFocus
    Else
        PictureSetting.Left = TS.Text
    End If

    For i = 0 To TombolPengaturan.Count - 1
        TombolPengaturan(i).Left = TombolMinimize.Left - TombolPengaturan(i).Width - 20
    Next i
End Sub

Private Sub TombolRestore_Click()
'Restore Form
'On Error GoTo Code_Error
    If Me.WindowState = 1 Or Me.WindowState = 2 Then Exit Sub
    With FormMain
        If FormUI.Label_Height.Caption <> "" Then .Height = FormUI.Label_Height.Caption Else .Height = 9420
        If FormUI.Label_Width.Caption <> "" Then .Width = FormUI.Label_Width.Caption Else .Width = 15000    '13380 '15585
        If FormUI.Label_Top.Caption <> "" Then .Top = FormUI.Label_Top.Caption Else .Top = (Screen.Height - Me.Height) / 2
        If FormUI.Label_Left.Caption <> "" Then .Left = FormUI.Label_Left.Caption Else .Left = (Screen.Width - Me.Width) / 2
    End With
    Layar_Penuh = False
    FormUI.Text_Layar_Penuh.Text = "False"
    TombolMaximize.Visible = True
    TombolRestore.Visible = False

    InitTimer

    Exit Sub
Code_Error:
End Sub

Private Sub TombolResize_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'Redimensionar o formulário conforme as dimensões pretendidas
    If Button = vbLeftButton Then
        If Layar_Penuh = False Then
            ReleaseCapture
            SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTBOTTOMRIGHT, 0

            ResizeForm
            Capture_Dimensi_Posisi_Form
            InitTimer
        End If
    End If
End Sub

Private Sub TombolResize_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'Alterar o mousepointer
    If Layar_Penuh = True Then
        TombolResize.MousePointer = vbDefault
    Else
        TombolResize.MousePointer = 8
    End If
End Sub

Private Sub TombolResize_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Capture_Dimensi_Posisi_Form
    InitTimer
End Sub

Private Sub InitTimer()
    FormKaryawan.AutoResize
    FormBeranda.AutoResize
    FormPotongan.AutoResize
    FormTunjangan.AutoResize
    FormBiaya.AutoResize
    FormPengeluaran.AutoResize
    FormPemasukan.AutoResize
    FormKas.AutoResize
    FormTentang.AutoResize
    FormInventaris.AutoResize
    FormPenggajian.AutoResize
    FormUser.AutoResize
    FormLevel.AutoResize
    FormDepartemen.AutoResize
    Cetak_Karyawan.AutoResize
    Cetak_Potongan.AutoResize
    Cetak_Tunjangan.AutoResize
    Cetak_Pemasukan.AutoResize
    Cetak_Pengeluaran.AutoResize
    Cetak_Inventaris.AutoResize
    Cetak_DataGaji.AutoResize
    Cetak_SlipGaji.AutoResize
End Sub

Private Sub TombolSignout_Click()
    Pesan_Peringatan "Question", "Yakin Ingin Logout dari Program Sekarang?", "Logout"
    If Respon = "Iya" Then
        MenuUser_Click
        PictureLogin.Visible = True
        TextPassword.Text = ""
    End If
End Sub

