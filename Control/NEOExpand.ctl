VERSION 5.00
Begin VB.UserControl NEOExpand 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   600
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00BC8D3C&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   0
      Picture         =   "NEOExpand.ctx":0000
      ScaleHeight     =   765
      ScaleWidth      =   660
      TabIndex        =   0
      Top             =   0
      Width           =   660
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   230
         X2              =   430
         Y1              =   420
         Y2              =   420
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   230
         X2              =   430
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   230
         X2              =   420
         Y1              =   300
         Y2              =   300
      End
   End
End
Attribute VB_Name = "NEOExpand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Dim m_Color
Dim m_ColorHover
Dim m_ColorDown

Dim isOver As Boolean
Dim m_State As Integer
Event Click()
Event DblClick()
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseEnter()
Event MouseLeave()

Public Property Get Color() As OLE_COLOR
    'Escolher a cor inicial de fundo do botão
    'Color = Shape1.FillColor
    'Color = Shape1.BorderColor
    Color = Picture1.BackColor
End Property

Public Property Set Color(m_New_Color As OLE_COLOR)
    'Alterar a cor inicial de fundo do botão
    Set m_Color = m_New_Color
    DrawState
    PropertyChanged "Color"
End Property

Public Property Let Color(m_New_Color As OLE_COLOR)
    'Alterar a cor inicial de fundo do botão
    'Shape1.FillColor = m_New_Color
    'Shape1.BorderColor = m_New_Color
    Picture1.BackColor = m_New_Color
    PropertyChanged "Color"
End Property

Public Property Get ColorHover() As OLE_COLOR
    'Escolher a cor inicial de fundo do botão
    'ColorHover = Shape1.FillColor
    'ColorHover = Shape1.BorderColor
    ColorHover = Picture1.BackColor
End Property

Public Property Set ColorHover(m_New_ColorHover As OLE_COLOR)
    'Alterar a cor inicial de fundo do botão
    Set m_ColorHover = m_New_ColorHover
    DrawState
    PropertyChanged "ColorHover"
End Property

Public Property Let ColorHover(m_New_ColorHover As OLE_COLOR)
    'Alterar a cor inicial de fundo do botão
    'Shape1.FillColor = m_New_ColorHover
    'Shape1.BorderColor = m_New_ColorHover
    PropertyChanged "ColorHover"
End Property

Public Property Get ColorDown() As OLE_COLOR
    'Escolher a cor inicial de fundo do botão
    'ColorDown = Shape1.FillColor
    'ColorDown = Shape1.BorderColor
    ColorDown = Picture1.BackColor
End Property

Public Property Set ColorDown(m_New_ColorDown As OLE_COLOR)
    'Alterar a cor inicial de fundo do botão
    Set m_ColorDown = m_New_ColorDown
    DrawState
    PropertyChanged "ColorDown"
End Property

Public Property Let ColorDown(m_New_ColorDown As OLE_COLOR)
    'Alterar a cor inicial de fundo do botão
    'Shape1.FillColor = m_New_ColorDown
    'Shape1.BorderColor = m_New_ColorDown
    PropertyChanged "ColorDown"
End Property

Private Sub DrawState()
    On Error Resume Next
    If m_State = 1 Then 'mouse hover
        UserControl.Cls
        'UserControl.PaintPicture m_PictureHover, 0, 0
        'Picture1.Picture = m_ColorHover
        Picture1.BackColor = m_ColorHover
        Shape1.FillColor = m_ColorHover
        Shape1.BorderColor = m_ColorHover
        UserControl.Width = Picture1.Width
        UserControl.Height = Picture1.Height
    ElseIf m_State = 2 Then 'mouse down
        UserControl.Cls
        'UserControl.PaintPicture m_PictureDown, 0, 0
        'Picture1.Picture = m_ColorDown
        Picture1.BackColor = m_ColorDown
        Shape1.FillColor = m_ColorDown
        Shape1.BorderColor = m_ColorDown
        UserControl.Width = Picture1.Width
        UserControl.Height = Picture1.Height
    Else
        UserControl.Cls
        'UserControl.PaintPicture m_Picture, 0, 0
        'Picture1.Picture = m_Color
        Picture1.BackColor = m_Color
        'Shape1.FillColor = m_Color
        'Shape1.BorderColor = m_Color
        UserControl.Width = Picture1.Width
        UserControl.Height = Picture1.Height
    End If
End Sub

Private Sub Timer1_Timer()
    If Not CheckMouseOver Then
        Timer1.Enabled = False
        isOver = False
        RaiseEvent MouseLeave
        m_State = 0
        Call DrawState
    End If
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_InitProperties()
    Set m_Color = Nothing
    Set m_ColorHover = Nothing
    Set m_ColorDown = Nothing
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
    If Button = 1 Then
        m_State = 2
        Call DrawState
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
    If Button < 2 Then
        If Not CheckMouseOver Then
            isOver = False
            m_State = 0
            Call DrawState
        Else
            If Button = 0 And Not isOver Then
                Timer1.Enabled = True
                isOver = True
                RaiseEvent MouseEnter
                m_State = 1
                Call DrawState
            ElseIf Button = 1 Then
                isOver = True
                m_State = 2
                Call DrawState
                isOver = False
            End If
        End If
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
    If CheckMouseOver Then
        m_State = 1
    Else
        m_State = 0
    End If
    Call DrawState
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        'Set m_Color = .ReadProperty("Color", Nothing)
        'Set m_ColorHover = .ReadProperty("ColorHover", Nothing)
        'Set m_ColorDown = .ReadProperty("ColorDown", Nothing)
        m_Color = .ReadProperty("Color", "&H00394BDD&")
        m_ColorHover = .ReadProperty("ColorHover", "&H003344C7&")
        m_ColorDown = .ReadProperty("ColorDown", "&H003344C7&")
    End With
End Sub

Private Sub UserControl_Resize()
    DrawState
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        'Call .WriteProperty("Color", m_Color, Nothing)
        'Call .WriteProperty("ColorHover", m_ColorHover, Nothing)
        'Call .WriteProperty("ColorDown", m_ColorDown, Nothing)
        Call .WriteProperty("Color", m_Color, "&H00394BDD&")
        Call .WriteProperty("ColorHover", m_ColorHover, "&H003344C7&")
        Call .WriteProperty("ColorDown", m_ColorDown, "&H003344C7&")
    End With
End Sub

Private Function CheckMouseOver() As Boolean
    Dim pt As POINTAPI
    GetCursorPos pt
    CheckMouseOver = (WindowFromPoint(pt.X, pt.Y) = UserControl.hWnd)
End Function
