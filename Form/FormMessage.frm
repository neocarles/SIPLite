VERSION 5.00
Begin VB.Form FormMessage 
   Appearance      =   0  'Flat
   BackColor       =   &H00F9F9F9&
   BorderStyle     =   0  'None
   ClientHeight    =   3810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6765
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormMessage.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   254
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   451
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox FrameCenter 
      Appearance      =   0  'Flat
      BackColor       =   &H00F9F9F9&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   0
      ScaleHeight     =   177
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   449
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   465
      Width           =   6735
      Begin VB.Label LabelMessage 
         AutoSize        =   -1  'True
         BackColor       =   &H00313131&
         BackStyle       =   0  'Transparent
         Caption         =   "Message"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1440
         TabIndex        =   4
         Top             =   480
         Width           =   735
      End
      Begin VB.Image PictureMessage 
         Enabled         =   0   'False
         Height          =   720
         Left            =   360
         Top             =   240
         Width           =   840
      End
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
      ScaleWidth      =   417
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   6255
      Begin Project1.N_Image TombolKeluar 
         Height          =   285
         Left            =   5460
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   0
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   503
         Picture         =   "FormMessage.frx":000C
         PictureHover    =   "FormMessage.frx":03E3
         PictureDown     =   "FormMessage.frx":07D3
      End
      Begin VB.Label LabelTitle 
         AutoSize        =   -1  'True
         BackColor       =   &H00272727&
         BackStyle       =   0  'Transparent
         Caption         =   "SIPLite"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   75
         TabIndex        =   1
         Top             =   120
         Width           =   600
      End
   End
   Begin VB.PictureBox FrameTombol 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFF1F2&
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
      Height          =   615
      Left            =   0
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   441
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3120
      Width           =   6615
      Begin Project1.isButton TombolIya 
         Height          =   375
         Left            =   1380
         TabIndex        =   7
         Top             =   120
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   661
         Style           =   8
         Caption         =   "Yes"
         Object.ToolTipText     =   ""
         ToolTipTitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Project1.isButton TombolOk 
         Height          =   375
         Left            =   2820
         TabIndex        =   8
         Top             =   120
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   661
         Style           =   8
         Caption         =   "Ok"
         Object.ToolTipText     =   ""
         ToolTipTitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Project1.isButton TombolNo 
         Height          =   375
         Left            =   4260
         TabIndex        =   9
         Top             =   120
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   661
         Style           =   8
         Caption         =   "No"
         Object.ToolTipText     =   ""
         ToolTipTitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label_Linha 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   15
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   2055
      End
   End
   Begin VB.Shape ShapeControl 
      BorderColor     =   &H00C0C0C0&
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "FormMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Nikyts Player
'   Copyright © 2011-2013 Nikyts software ™ - Informática e tecnologia
'   www.nikyts.net / nikyts@hotmail.com
'   Desenvolvido por: Nelson do Carmo
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub Form_Initialize()
    CenterForm Me
End Sub

Private Sub TombolIya_Click()
    Respon = "Iya"
    Unload Me
End Sub

Private Sub TombolKeluar_Click()
    Unload Me
    Respon = "No"
End Sub

Private Sub TombolNo_Click()
    Respon = "No"
    Unload Me
End Sub

Private Sub TombolOk_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
'Arredondar os cantos do formulário
    RoundedSideForm Me, True
    CenterForm Me
End Sub

Private Sub Form_Load()
'Definir os valores de x e y para poder mover o formulário
    iTPPX& = Screen.TwipsPerPixelX
    iTPPY& = Screen.TwipsPerPixelY
    
    Call SetIcon(Me.hWnd, "FORMICON", False)

    Resize_Form
    'LabelTitle.Caption = App.ProductName
    RoundedSideForm Me, True
    CenterForm Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'Teclas de atalho
    If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Resize()
'Chamar o procedimento
    Resize_Form
End Sub

Public Sub Resize_Form()
'Procedimento para ajustar os objectos
    If Me.WindowState = 1 Then Exit Sub
    With Me
        'Descrição dos 3 ultimos valores. Espaço entre: borda do form -> icon, icon -> label messagem, LabelMessage -> borda do form
        .Width = Screen.TwipsPerPixelX * (LabelMessage.Width + PictureMessage.Width + 40 + 17 + 50)
        .Height = Screen.TwipsPerPixelY * (Header.ScaleHeight + FrameTombol.ScaleHeight + 36 + LabelMessage.Height + 30 + 10 + 20)
    End With

    SesuaikanForm FormMessage, False, False, True, True

    With TombolIya
        '.Top = (FrameTombol.ScaleHeight - .Height) / 2
        '.Left = (FrameTombol.ScaleWidth / 2) - .Width - (.Top / 2)
        .Move (FrameTombol.ScaleWidth / 2) - .Width - (.Top / 2), (FrameTombol.ScaleHeight - .Height) / 2
    End With

    With TombolOk
        '.Top = (FrameTombol.ScaleHeight - .Height) / 2
        '.Left = (FrameTombol.ScaleWidth - .Width) / 2
        .Move (FrameTombol.ScaleWidth - .Width) / 2, (FrameTombol.ScaleHeight - .Height) / 2
    End With

    With TombolNo
        '.Top = (FrameTombol.ScaleHeight - .Height) / 2
        '.Left = (FrameTombol.ScaleWidth / 2) + (.Top / 2)
        .Move (FrameTombol.ScaleWidth / 2) + (.Top / 2), (FrameTombol.ScaleHeight - .Height) / 2
    End With

    With PictureMessage
        '.Top = 20
        '.Left = 20
        .Move 20, 20
    End With

    With LabelMessage
        '.Top = PictureMessage.Top + 10
        '.Left = PictureMessage.Left + PictureMessage.Width + 20
        .Move PictureMessage.Left + PictureMessage.Width + 20, PictureMessage.Top + 10
    End With

    With Label_Linha
        '.Top = 0
        '.Left = 0
        '.Width = FrameTombol.ScaleWidth
        .Move 0, 0, FrameTombol.ScaleWidth
    End With

    'Ajustar os objectos depois de arredondar os cantos do formulário
    ShapeControl.Left = 0
    ShapeControl.Width = Me.ScaleWidth - 1
    ShapeControl.Height = Me.ScaleHeight - 1
    FrameTombol.Width = FrameTombol.ScaleWidth - 1
End Sub

Private Sub Header_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Capture_Posisi_Form FormMessage
End Sub

Private Sub Header_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Move_Form FormMessage
End Sub

Private Sub Header_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Large_Form FormMessage
End Sub

Private Sub LabelTitle_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Capture_Posisi_Form FormMessage
End Sub

Private Sub LabelTitle_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Move_Form FormMessage
End Sub

Private Sub LabelTitle_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Large_Form FormMessage
End Sub

Public Sub Carregar_Idioma()
''Procedimento para carregar o idioma selecionado
'Localizacao_Ficheiro_Lingua = App.Path & "\Languages\" & Form_Principal.Text_Lingua.Text & ".lng"
'
' TombolKeluar.ToolTipText = ReadINI("Message", "Button_Close", Localizacao_Ficheiro_Lingua)
' TombolIya.Caption = ReadINI("Message", "Label_Yes", Localizacao_Ficheiro_Lingua)
' TombolOk.Caption = ReadINI("Message", "Label_Ok", Localizacao_Ficheiro_Lingua)
'TombolNo.Caption = ReadINI("Message", "Label_No", Localizacao_Ficheiro_Lingua)
End Sub

