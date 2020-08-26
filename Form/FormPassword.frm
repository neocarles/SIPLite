VERSION 5.00
Begin VB.Form FormPassword 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Password"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5775
   Icon            =   "FormPassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   5775
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PictureFooter 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F5F5&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   5775
      TabIndex        =   8
      Top             =   3000
      Width           =   5775
      Begin Project1.isButton ButtonSave 
         Height          =   420
         Left            =   2760
         TabIndex        =   9
         Top             =   120
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   741
         Icon            =   "FormPassword.frx":000C
         Style           =   8
         Caption         =   "&Save"
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
      Begin Project1.isButton ButtonCancel 
         Height          =   420
         Left            =   4320
         TabIndex        =   10
         Top             =   120
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   741
         Icon            =   "FormPassword.frx":03A6
         Style           =   8
         Caption         =   "&Cancel"
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
   End
   Begin VB.TextBox TextNewPassword2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3180
      TabIndex        =   7
      Top             =   2160
      Width           =   2295
   End
   Begin VB.TextBox TextNewPassword 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3180
      TabIndex        =   6
      Top             =   1620
      Width           =   2295
   End
   Begin VB.TextBox TextOldPassword 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3180
      TabIndex        =   5
      Top             =   1080
      Width           =   2295
   End
   Begin VB.PictureBox PictureUser 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2000
      Left            =   240
      ScaleHeight     =   1995
      ScaleWidth      =   1995
      TabIndex        =   4
      Top             =   720
      Width           =   2000
      Begin VB.Image ImageUser 
         Height          =   1995
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1995
      End
   End
   Begin VB.PictureBox PictureHeader 
      Appearance      =   0  'Flat
      BackColor       =   &H00F7EDD9&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   0
      ScaleHeight     =   450
      ScaleWidth      =   5775
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Change Password"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00725628&
         Height          =   240
         Left            =   480
         TabIndex        =   11
         Top             =   90
         Width           =   1725
      End
      Begin VB.Label L3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00725628&
         Height          =   345
         Left            =   120
         TabIndex        =   3
         Top             =   60
         Width           =   270
      End
      Begin VB.Label L2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00725628&
         Height          =   345
         Left            =   120
         TabIndex        =   2
         Top             =   -30
         Width           =   270
      End
      Begin VB.Label L1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00725628&
         Height          =   345
         Left            =   120
         TabIndex        =   1
         Top             =   -120
         Width           =   270
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000A&
      X1              =   2460
      X2              =   2460
      Y1              =   540
      Y2              =   2880
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H8000000A&
      Height          =   375
      Left            =   3060
      Shape           =   4  'Rounded Rectangle
      Top             =   2100
      Width           =   2475
   End
   Begin VB.Image Image3 
      Height          =   240
      Left            =   2805
      Picture         =   "FormPassword.frx":0740
      Top             =   2160
      Width           =   240
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H8000000A&
      Height          =   375
      Left            =   3060
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   2475
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   2805
      Picture         =   "FormPassword.frx":081F
      Top             =   1620
      Width           =   240
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000A&
      Height          =   375
      Left            =   3060
      Shape           =   4  'Rounded Rectangle
      Top             =   1020
      Width           =   2475
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   2805
      Picture         =   "FormPassword.frx":08FE
      Top             =   1080
      Width           =   240
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H8000000A&
      FillColor       =   &H8000000A&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   2700
      Shape           =   4  'Rounded Rectangle
      Top             =   1020
      Width           =   435
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H8000000A&
      FillColor       =   &H8000000A&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   2700
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   435
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H8000000A&
      FillColor       =   &H8000000A&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   2700
      Shape           =   4  'Rounded Rectangle
      Top             =   2100
      Width           =   435
   End
End
Attribute VB_Name = "FormPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ClassUser As ClassUser
Dim RsGantiPass As ADODB.Recordset

Private Sub ButtonCancel_Click()
    Unload Me
End Sub

Private Sub ButtonSave_Click()
    Dim TeksHasilCipher As String
    '
    If TextNewPassword.Text <> TextNewPassword2.Text Then
        MsgBox "Password baru yang diisi tidak sama dengan konfirmasinya", vbExclamation, "Peringatan"
        Exit Sub
    End If
    '
    If Len(TextNewPassword.Text) < 5 Then
        MsgBox "Minimal Password 5 karakter", vbExclamation, "Peringantan"
        Exit Sub
    End If
    '
    Set ClassUser = New ClassUser
    TeksHasilCipher = ClassUser.GetTeksHasilCipher(Encrypt(TextOldPassword.Text))
    Sukses = ClassUser.OldPassword(FormMain.LabelUsername.Caption, TeksHasilCipher, Conn)
    If Sukses = False Then    'password lama salah
        MsgBox "Password lama tidak benar", vbExclamation, "Peringatan"

    Else
        TeksHasilCipher = ClassUser.GetTeksHasilCipher(Encrypt(TextNewPassword.Text))
        Sukses = ClassUser.UpdatePassword(FormMain.LabelUsername.Caption, TeksHasilCipher, Conn)
        If Sukses Then
            Unload Me
        Else
            MsgBox "Password baru user gagal diupdate", vbExclamation, "Peringatan"
        End If
    End If
    Set ClassUser = Nothing
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case Is = vbKeyEscape
        ButtonCancel_Click
    Case Is = vbKeyF8
        ButtonSave_Click
    End Select
End Sub

Private Sub Form_Load()
    CenterForm Me
    ImageUser.Picture = FormMain.ImageProfil.Picture
End Sub

Private Sub PictureFooter_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case Is = vbKeyEscape
        ButtonCancel_Click
    Case Is = vbKeyF8
        ButtonSave_Click
    End Select
End Sub

Private Sub PictureHeader_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case Is = vbKeyEscape
        ButtonCancel_Click
    Case Is = vbKeyF8
        ButtonSave_Click
    End Select
End Sub

Private Sub PictureUser_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case Is = vbKeyEscape
        ButtonCancel_Click
    Case Is = vbKeyF8
        ButtonSave_Click
    End Select
End Sub

Private Sub TextNewPassword_Change()
    With TextNewPassword
        .FontName = "Wingdings"
        .FontSize = 9
        .PasswordChar = "l"
    End With
End Sub

Private Sub TextNewPassword2_Change()
    With TextNewPassword2
        .FontName = "Wingdings"
        .FontSize = 9
        .PasswordChar = "l"
    End With
End Sub

Private Sub TextOldPassword_Change()
    With TextOldPassword
        .FontName = "Wingdings"
        .FontSize = 9
        .PasswordChar = "l"
    End With
End Sub
