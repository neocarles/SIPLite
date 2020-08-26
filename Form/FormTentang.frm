VERSION 5.00
Begin VB.Form FormTentang 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form Tentang"
   ClientHeight    =   8250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13665
   Icon            =   "FormTentang.frx":0000
   LinkTopic       =   "FormTentang"
   ScaleHeight     =   8250
   ScaleWidth      =   13665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PictureAbout 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   240
      ScaleHeight     =   3735
      ScaleWidth      =   7695
      TabIndex        =   1
      Top             =   240
      Width           =   7695
      Begin Project1.isButton ButtonLicense 
         Height          =   555
         Left            =   1860
         TabIndex        =   6
         Top             =   1140
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   979
         Style           =   8
         Caption         =   "License"
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
      Begin VB.Label lblOSVer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%%OSVersion%%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   1800
         Width           =   1320
      End
      Begin VB.Label LabelCopyURL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "https://blog.carlesneo.id/siplite"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   360
         MouseIcon       =   "FormTentang.frx":000C
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   2580
         Width           =   2655
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Web:"
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
         Left            =   120
         TabIndex        =   9
         Top             =   2340
         Width           =   420
      End
      Begin VB.Label LabelReg 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "% Name %"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2940
         TabIndex        =   8
         Top             =   360
         Width           =   825
      End
      Begin VB.Label LabelEd 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "% Edition %"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   2760
         TabIndex        =   7
         Top             =   120
         Width           =   1185
      End
      Begin VB.Label lblTrial 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Copyrigth © 2017 MPC IT"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1860
         TabIndex        =   5
         Top             =   840
         Width           =   1860
      End
      Begin VB.Label lblLic 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Registered to: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1860
         TabIndex        =   4
         Top             =   360
         Width           =   1080
      End
      Begin VB.Label lblVer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "v1.11.17"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1860
         TabIndex        =   3
         Top             =   600
         Width           =   660
      End
      Begin VB.Label lblAPPS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SIPLite - "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1860
         TabIndex        =   2
         Top             =   120
         Width           =   885
      End
      Begin VB.Image imgAPP 
         Height          =   1530
         Left            =   60
         Picture         =   "FormTentang.frx":015E
         Stretch         =   -1  'True
         Top             =   60
         Width           =   1665
      End
   End
   Begin Project1.isButton ButtonKeluar 
      Cancel          =   -1  'True
      Height          =   435
      Left            =   10560
      TabIndex        =   0
      Top             =   6180
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   767
      Style           =   8
      Caption         =   "&OK"
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
Attribute VB_Name = "FormTentang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objOperSys  As ClassOSVer

Sub GetVersionOS()
Set objOperSys = New ClassOSVer
        With objOperSys
            ' Capture information about this operating system
            gstrOperSystem = .VersionName & vbNewLine & TrimStr("Ver " & .VersionData) & " " & _
                             .ProcessArchitecture & " (" & .BaseBitStructure & ")"
            gblnWin8or81 = .bCaptionIsCentered   ' If TRUE then this is Win 8 or 8.1
            lngMajorVer = CLng(.MajorVersion)    ' Capture OS major version
        End With
lblOSVer.Caption = "OS : " & gstrOperSystem
End Sub

Private Sub ButtonKeluar_Click()
    Unload Me
End Sub

Private Sub ButtonLicense_Click()
    FormLicense.Show vbModal
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
    Call GetVersionOS
    
    Call GetReg
    
    lblVer.Caption = "v" & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Public Sub GetReg()
If Registered = True Then
    LabelReg.Caption = RegisterAs
Else
    LabelReg.Caption = "UNREGISTERED"
End If

If VersiPro = True Then
    LabelEd.Caption = "LICENSE EDITION"
Else
    LabelEd.Caption = "DEMO EDITION"
End If
End Sub

Private Sub Form_Resize()
    Call AutoResize
End Sub

Public Sub AutoResize()
If Me.WindowState = 1 Then Exit Sub
    With ButtonKeluar
        .Move FormMain.PictureBeranda.ScaleWidth - ButtonKeluar.Width - 180, FormMain.PictureBeranda.ScaleHeight - ButtonKeluar.Height - 180
    End With
End Sub

Private Sub LabelCopyURL_Click()
    Pesan_Peringatan "Question", "Anda akan dialihkan ke laman Official Website" & vbNewLine & " Menggunakan Browser Default.", "Buka Link"
    If Respon = "Iya" Then
        ShellExecute hWnd, "open", "https://blog.carlesneo.id/siplite", _
                     vbNullString, vbNullString, 1
    End If
End Sub
