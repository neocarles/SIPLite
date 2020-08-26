VERSION 5.00
Begin VB.UserControl NEOSkinSelect 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "NEOSkin.ctx":0000
   Begin VB.PictureBox PictureBG 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   1515
      TabIndex        =   0
      Top             =   0
      Width           =   1515
      Begin VB.PictureBox S 
         Appearance      =   0  'Flat
         BackColor       =   &H00312D22&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   0
         ScaleHeight     =   855
         ScaleWidth      =   495
         TabIndex        =   1
         Top             =   260
         Width           =   495
      End
      Begin VB.Image ImageSelect 
         Height          =   1095
         Left            =   0
         Top             =   0
         Width           =   1515
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00A87F36&
         FillColor       =   &H00A87F36&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   0
         Top             =   0
         Width           =   495
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00BC8D3C&
         FillColor       =   &H00BC8D3C&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   480
         Top             =   0
         Width           =   1035
      End
   End
End
Attribute VB_Name = "NEOSkinSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Enum S_Themes
    [Red] = 1
    [Green] = 2
    [Blue] = 3
    [Yellow] = 4
    [Purple] = 5
    [Brown] = 6
    [Custom] = 7
End Enum

Private S_Theme As S_Themes

Event Click()
Event DblClick()
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseEnter()
Event MouseLeave()

Public Property Let Themes(ByVal NewValue As S_Themes)

    S_Theme = NewValue
    PropertyChanged "Themes"
Bar_Draw
End Property

Public Property Get Themes() As S_Themes
    Themes = S_Theme
End Property

Private Sub Bar_Draw()
On Error Resume Next
CheckTheme

UserControl.Cls
End Sub

Private Sub CheckTheme()
If Themes = 1 Then
'BACK
Shape1.FillColor = RGB(215, 57, 37)
Shape1.BorderColor = RGB(215, 57, 37)
Shape2.FillColor = RGB(221, 75, 57)
Shape2.BorderColor = RGB(221, 75, 57)
ElseIf Themes = 2 Then
'BACK
Shape1.FillColor = RGB(0, 135, 73)
Shape1.BorderColor = RGB(0, 135, 73)
Shape2.FillColor = RGB(0, 166, 90)
Shape2.BorderColor = RGB(0, 166, 90)
ElseIf Themes = 3 Then
'BACK
Shape1.FillColor = RGB(54, 127, 168)
Shape1.BorderColor = RGB(54, 127, 168)
Shape2.FillColor = RGB(60, 141, 188)
Shape2.BorderColor = RGB(60, 141, 188)
ElseIf Themes = 4 Then
'BACK
Shape1.FillColor = RGB(243, 156, 18)
Shape1.BorderColor = RGB(243, 156, 18)
Shape2.FillColor = RGB(219, 141, 16)
Shape2.BorderColor = RGB(219, 141, 16)
ElseIf Themes = 5 Then
'BACK
Shape1.FillColor = RGB(142, 72, 171)
Shape1.BorderColor = RGB(142, 72, 171)
Shape2.FillColor = RGB(155, 92, 181)
Shape2.BorderColor = RGB(155, 92, 181)
ElseIf Themes = 6 Then
'BACK
Shape1.FillColor = RGB(209, 85, 25)
Shape1.BorderColor = RGB(209, 85, 25)
Shape2.FillColor = RGB(228, 127, 49)
Shape2.BorderColor = RGB(228, 127, 49)
ElseIf Themes = 7 Then
'BACK
Shape1.FillColor = Color1
Shape1.BorderColor = Color1
Shape2.FillColor = Color2
Shape2.BorderColor = Color2
End If
End Sub

Private Sub ImageSelect_Click()
RaiseEvent Click
End Sub

Private Sub S_Click()
RaiseEvent Click
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    'Ler as propriedades do control
    'On Error GoTo Baca_Propertis_Kontrol
    With PropBag
        Color1 = .ReadProperty("Color1", "&H00A87F36&")
        Color2 = .ReadProperty("Color2", "&H00BC8D3C&")
        Themes = .ReadProperty("Themes", "")
    End With
    
Baca_Propertis_Kontrol:
End Sub

Private Sub UserControl_Resize()
Call Resize_Control
End Sub

Private Sub Resize_Control()
UserControl.Width = PictureBG.Width '3555
UserControl.Height = PictureBG.Height '1650
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    'Escrever as propriedades do control
    'On Error GoTo Baca_Propertis_Kontrol
    With PropBag
        Call .WriteProperty("Color1", Color1, "&H00A87F36&")
        Call .WriteProperty("Color2", Color2, "&H00BC8D3C&")
        .WriteProperty "Themes", S_Theme, ""
    End With
    
Baca_Propertis_Kontrol:
End Sub


Public Property Get Color1() As OLE_COLOR
    'Escolher a cor inicial de fundo do botão
    Color1 = Shape1.FillColor
    Color1 = Shape1.BorderColor
End Property

Public Property Let Color1(New_Value As OLE_COLOR)
    'Alterar a cor inicial de fundo do botão
    Shape1.FillColor = New_Value
    Shape1.BorderColor = New_Value
    PropertyChanged "Color1"
End Property

Public Property Get Color2() As OLE_COLOR
    'Escolher a cor inicial de fundo do botão
    Color2 = Shape2.FillColor
    Color2 = Shape2.BorderColor
End Property

Public Property Let Color2(New_Value As OLE_COLOR)
    'Alterar a cor inicial de fundo do botão
    Shape2.FillColor = New_Value
    Shape2.BorderColor = New_Value
    PropertyChanged "Color2"
End Property


