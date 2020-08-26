VERSION 5.00
Begin VB.Form FormLicense 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4650
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6345
   Icon            =   "FormLicense.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4650
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Project1.isButton ButtonApply 
      Height          =   435
      Left            =   1740
      TabIndex        =   13
      Top             =   4020
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   767
      Style           =   8
      Caption         =   "Apply"
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
   Begin Project1.isButton ButtonCancel 
      Height          =   435
      Left            =   4680
      TabIndex        =   15
      Top             =   4020
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   767
      Style           =   8
      Caption         =   "Cancel"
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
   Begin Project1.isButton ButtonPurchase 
      Height          =   435
      Left            =   3180
      TabIndex        =   14
      Top             =   4020
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   767
      Style           =   8
      Caption         =   "Purchase"
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
   Begin VB.TextBox TextKey 
      Appearance      =   0  'Flat
      Height          =   1275
      Left            =   1740
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   2580
      Width           =   4335
   End
   Begin VB.TextBox TextSerial 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1740
      TabIndex        =   11
      Top             =   2100
      Width           =   4335
   End
   Begin VB.TextBox TextAddress 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1740
      TabIndex        =   10
      Top             =   1620
      Width           =   4335
   End
   Begin VB.TextBox TextCompany 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1740
      TabIndex        =   9
      Top             =   1140
      Width           =   4335
   End
   Begin VB.TextBox TextNama 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1740
      TabIndex        =   8
      Top             =   660
      Width           =   4335
   End
   Begin VB.PictureBox Header 
      Appearance      =   0  'Flat
      BackColor       =   &H00191919&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   0
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   421
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   6315
      Begin Project1.N_Image TombolKeluar 
         Height          =   285
         Left            =   5520
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   503
         Picture         =   "FormLicense.frx":000C
         PictureHover    =   "FormLicense.frx":03E3
         PictureDown     =   "FormLicense.frx":07D3
      End
      Begin VB.Label LabelTitle 
         AutoSize        =   -1  'True
         BackColor       =   &H00272727&
         BackStyle       =   0  'Transparent
         Caption         =   "SIPLite Registration"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   1410
      End
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   420
      Picture         =   "FormLicense.frx":0BA7
      Top             =   3180
      Width           =   960
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Registration Key"
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
      Left            =   240
      TabIndex        =   7
      Top             =   2700
      Width           =   1410
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Serial Number"
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
      Left            =   240
      TabIndex        =   6
      Top             =   2160
      Width           =   1185
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
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
      Left            =   240
      TabIndex        =   5
      Top             =   1680
      Width           =   690
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Company"
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
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   795
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
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
      Left            =   240
      TabIndex        =   3
      Top             =   780
      Width           =   480
   End
   Begin VB.Shape ShapeBorder 
      BorderColor     =   &H00191919&
      BorderWidth     =   2
      Height          =   4635
      Left            =   15
      Top             =   0
      Width           =   6315
   End
End
Attribute VB_Name = "FormLicense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ButtonApply_Click()

End Sub

Private Sub ButtonCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Call SetIcon(Me.hWnd, "FORMICON", False)
Call AutoResize
CenterForm Me
End Sub

Private Sub Form_Resize()
Call AutoResize
End Sub

Public Sub AutoResize()
With Me
    .Height = ShapeBorder.Height
    .Width = ShapeBorder.Width + 15
End With
End Sub

Private Sub TombolKeluar_Click()
    Unload Me
End Sub

Private Sub Header_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Capture_Posisi_Form FormLicense
End Sub

Private Sub Header_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Move_Form FormLicense
End Sub

Private Sub Header_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Large_Form FormLicense
End Sub

Private Sub LabelTitle_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Capture_Posisi_Form FormLicense
End Sub

Private Sub LabelTitle_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Move_Form FormLicense
End Sub

Private Sub LabelTitle_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Large_Form FormLicense
End Sub
