VERSION 5.00
Begin VB.UserControl NEOBadge 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BackStyle       =   0  'Transparent
   ClientHeight    =   1650
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3555
   DataSourceBehavior=   1  'vbDataSource
   ScaleHeight     =   1650
   ScaleWidth      =   3555
   ToolboxBitmap   =   "NEOBadge.ctx":0000
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1860
      Top             =   240
   End
   Begin VB.Image Image1 
      Enabled         =   0   'False
      Height          =   915
      Left            =   2580
      Picture         =   "NEOBadge.ctx":0312
      Stretch         =   -1  'True
      Top             =   180
      Width           =   855
   End
   Begin VB.Image Image4 
      Height          =   1275
      Left            =   0
      Top             =   0
      Width           =   3555
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   255
      Left            =   1680
      MouseIcon       =   "NEOBadge.ctx":0E16
      MousePointer    =   99  'Custom
      Picture         =   "NEOBadge.ctx":0F68
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   255
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   0
      MouseIcon       =   "NEOBadge.ctx":4458
      MousePointer    =   99  'Custom
      Top             =   1260
      Width           =   3555
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "190"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   705
      Left            =   180
      TabIndex        =   1
      Top             =   120
      Width           =   1125
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Barang"
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
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   1050
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H003344C7&
      FillColor       =   &H003344C7&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   0
      Top             =   1260
      Width           =   3555
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H003344C7&
      FillColor       =   &H00394BDD&
      FillStyle       =   0  'Solid
      Height          =   1275
      Left            =   0
      Top             =   0
      Width           =   3555
   End
End
Attribute VB_Name = "NEOBadge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Enum D_Themes
    [Red] = 1
    [Green] = 2
    [Blue] = 3
    [Yellow] = 4
    [Custom] = 5
End Enum

'Deklarasi Variabel
Private Const LOGPIXELSY = 90
Private Const LF_FACESIZE = 32
Private Const DT_BOTTOM = &H8
Private Const DT_CENTER = &H1
Private Const DT_LEFT = &H0
Private Const DT_RIGHT = &H2
Private Const DT_TOP = &H0
Private Const DT_VCENTER = &H4
Private Const DT_SINGLELINE = &H20
Private Const DT_WORDBREAK = &H10
Private Const DT_CALCRECT = &H400
Private Const DT_END_ELLIPSIS = &H8000
Private Const DT_MODIFYSTRING = &H10000
Private Const DT_WORD_ELLIPSIS = &H40000

'Warna
Private Const COLOR_BTNHIGHLIGHT = 20
Private Const COLOR_BTNSHADOW = 16
Private Const COLOR_BTNFACE = 15
Private Const COLOR_HIGHLIGHT = 13
Private Const COLOR_ACTIVEBORDER = 10
Private Const COLOR_WINDOWFRAME = 6
Private Const COLOR_3DDKSHADOW = 21
Private Const COLOR_3DLIGHT = 22
Private Const COLOR_INFOTEXT = 23
Private Const COLOR_INFOBK = 24
Private Const PATCOPY = &HF00021
Private Const SRCCOPY = &HCC0020
Private Const PS_SOLID = 0
Private Const PS_DASHDOT = 3
Private Const PS_DASHDOTDOT = 4
Private Const PS_DOT = 2
Private Const PS_DASH = 1
Private Const PS_ENDCAP_FLAT = &H200

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

'Api Untuk Komponen
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
        X As Long
        Y As Long
End Type
Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hDc As Long, ByVal hRgn As Long) As Long
Private Declare Function SelectClipPath Lib "gdi32" (ByVal hDc As Long, ByVal iMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDc As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDc As Long, ByVal nIndex As Long) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
'Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long ' api repetida
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDc As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDc As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDc As Long, ByVal hObject As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hDc As Long, lpRect As RECT) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hDc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long

Private D_Theme As D_Themes

Dim m_Picture1 As New StdPicture
Dim m_Picture2 As New StdPicture

Dim isOver As Boolean
Dim m_State As Integer
Event Click()
Event DblClick()
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseEnter()
Event MouseLeave()

Public Property Get Caption1() As String
    'Escolher o nome do botão
    Caption1 = Label1.Caption
End Property

Public Property Let Caption1(New_Value As String)
    'Alterar o caption para o novo texto
    Label1.Caption = New_Value
    PropertyChanged "Caption1"
End Property

Public Property Get Caption2() As String
    'Escolher o nome do botão
    Caption2 = Label2.Caption
End Property

Public Property Let Caption2(New_Value As String)
    'Alterar o caption para o novo texto
    Label2.Caption = New_Value
    PropertyChanged "Caption2"
End Property

Public Property Get FontCap1() As Font
    Set FontCap1 = Label1.Font
End Property

Public Property Set FontCap1(ByRef NewFont As Font)
    Set Label1.Font = NewFont
    PropertyChanged "FontCap1"
End Property

Public Property Get FontCap2() As Font
    Set FontCap2 = Label2.Font
End Property

Public Property Set FontCap2(ByRef NewFont As Font)
    Set Label2.Font = NewFont
    PropertyChanged "FontCap2"
End Property

Private Sub Image2_Click()
RaiseEvent Click
End Sub

Private Sub Image2_DblClick()
RaiseEvent DblClick
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
    If Button = 1 Then
        m_State = 2
        'Call DrawState
    End If
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
    If Button < 2 Then
        If Not CheckMouseOver Then
            isOver = False
            m_State = 0
            'Call DrawState
        Else
            If Button = 0 And Not isOver Then
                Timer1.Enabled = True
                isOver = True
                RaiseEvent MouseEnter
                m_State = 1
                'Call DrawState
            ElseIf Button = 1 Then
                isOver = True
                m_State = 2
                'Call DrawState
                isOver = False
            End If
        End If
    End If
End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
    If CheckMouseOver Then
        m_State = 1
    Else
        m_State = 0
    End If
    'Call DrawState
End Sub

Private Sub Image3_Click()
RaiseEvent Click
End Sub

Private Sub Image3_DblClick()
RaiseEvent DblClick
End Sub

Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
    If Button = 1 Then
        m_State = 2
        'Call DrawState
    End If
End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
    If Button < 2 Then
        If Not CheckMouseOver Then
            isOver = False
            m_State = 0
            'Call DrawState
        Else
            If Button = 0 And Not isOver Then
                Timer1.Enabled = True
                isOver = True
                RaiseEvent MouseEnter
                m_State = 1
                'Call DrawState
            ElseIf Button = 1 Then
                isOver = True
                m_State = 2
                'Call DrawState
                isOver = False
            End If
        End If
    End If
End Sub

Private Sub Image3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
    If CheckMouseOver Then
        m_State = 1
    Else
        m_State = 0
    End If
    'Call DrawState
End Sub

Private Function CheckMouseOver() As Boolean
    Dim pt As POINTAPI
    GetCursorPos pt
    CheckMouseOver = (WindowFromPoint(pt.X, pt.Y) = UserControl.hWnd)
End Function

'Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'With Image1
'    .Left = .Left - 10
'    .Top = .Top - 10
'    .Width = .Width - 20
'    .Height = .Height - 20
'End With
'End Sub

Private Sub UserControl_Paint()
    'Escolher a imagem normal para o estado normal do botao
    Image1.Picture = Picture1
    Image2.Picture = Picture2
End Sub

Public Property Let Themes(ByVal NewValue As D_Themes)

    D_Theme = NewValue
    PropertyChanged "Themes"
Bar_Draw
End Property

Public Property Get Themes() As D_Themes
    Themes = D_Theme
End Property

Private Sub Bar_Draw()
On Error Resume Next
CheckTheme

UserControl.Cls
End Sub

Private Sub CheckTheme()
If Themes = 1 Then
'BACK
Shape1.FillColor = RGB(221, 75, 57)
Shape1.BorderColor = RGB(221, 75, 57)
Shape2.FillColor = RGB(199, 68, 51)
Shape2.BorderColor = RGB(199, 68, 51)
ElseIf Themes = 2 Then
'BACK
Shape1.FillColor = RGB(0, 166, 90)
Shape1.BorderColor = RGB(0, 166, 90)
Shape2.FillColor = RGB(0, 150, 81)
Shape2.BorderColor = RGB(0, 150, 81)
ElseIf Themes = 3 Then
'BACK
Shape1.FillColor = RGB(0, 192, 239)
Shape1.BorderColor = RGB(0, 192, 239)
Shape2.FillColor = RGB(0, 173, 216)
Shape2.BorderColor = RGB(0, 173, 216)
ElseIf Themes = 4 Then
'BACK
Shape1.FillColor = RGB(243, 156, 18)
Shape1.BorderColor = RGB(243, 156, 18)
Shape2.FillColor = RGB(219, 141, 16)
Shape2.BorderColor = RGB(219, 141, 16)
ElseIf Themes = 5 Then
'BACK
Shape1.FillColor = Color1
Shape1.BorderColor = Color1
Shape2.FillColor = Color2
Shape2.BorderColor = Color2
End If
End Sub

Private Sub Timer1_Timer()
    If Not CheckMouseOver Then
        Timer1.Enabled = False
        isOver = False
        RaiseEvent MouseLeave
        m_State = 0
        'Call DrawState
    End If
End Sub

Private Sub Resize_Control()
'UserControl.Width = Shape1.Width '3555
UserControl.Height = Shape1.Height + Shape2.Height '1650

With Shape1
    .Top = 0
    .Left = 0
    .Width = UserControl.ScaleWidth
End With

With Shape2
    .Top = Shape1.Height
    .Left = 0
    .Width = UserControl.ScaleWidth
End With

With Image1
    .Left = UserControl.Width - .Width - 120
    .Top = (Shape1.Height * 0.5) - (.Height * 0.5)
End With

With Image2
    .Left = (Shape2.Width * 0.5) - (.Width * 0.5)
    .Top = 1320 '(Shape2.Height * 0.5) - (.Height * 0.5)
End With

With Image3
    .Top = Shape1.Height
    .Left = 0
    .Width = UserControl.ScaleWidth
End With

With Image4
    .Top = 0
    .Left = 0
    .Width = UserControl.ScaleWidth
End With
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    'Ler as propriedades do control
    'On Error GoTo Baca_Propertis_Kontrol
    With PropBag
        Set m_Picture1 = .ReadProperty("Picture1", Nothing)
        Set m_Picture2 = .ReadProperty("Picture2", Nothing)
        'Set m_Picture_Hover = .ReadProperty("Picture_Hover", Nothing)
        'Image1.Stretch = .ReadProperty("Stretch", False)
        
        Set Label1.Font = PropBag.ReadProperty("FontCap1", "Tahoma")
        Set Label2.Font = PropBag.ReadProperty("FontCap2", "Tahoma")
        
        Caption1 = .ReadProperty("Caption1", "000")
        Caption2 = .ReadProperty("Caption2", "Nilai Data")
        Color1 = .ReadProperty("Color1", "&H00394BDD&")
        Color2 = .ReadProperty("Color2", "&H003344C7&")
        Themes = .ReadProperty("Themes", "")
    End With
    
Baca_Propertis_Kontrol:
End Sub

Private Sub UserControl_Resize()
Call Resize_Control
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    'Escrever as propriedades do control
    'On Error GoTo Baca_Propertis_Kontrol
    With PropBag
        Call .WriteProperty("Picture1", m_Picture1, Nothing)
        Call .WriteProperty("Picture2", m_Picture2, Nothing)
        '.WriteProperty "Picture_Hover", m_Picture_Hover, 0
        '.WriteProperty "Stretch", Image1.Stretch
        
        PropBag.WriteProperty "FontCap1", Label1.Font
        PropBag.WriteProperty "FontCap2", Label2.Font
        
        Call .WriteProperty("Caption1", Caption1, "")
        Call .WriteProperty("Caption2", Caption2, "")
        Call .WriteProperty("Color1", Color1, "&H00394BDD&")
        Call .WriteProperty("Color2", Color2, "&H003344C7&")
        .WriteProperty "Themes", D_Theme, ""
    End With
    
Baca_Propertis_Kontrol:
End Sub

Public Property Get Picture1() As StdPicture
    'Obter a imagem normal
    Set Picture1 = m_Picture1
End Property

Public Property Set Picture1(vNewPic As StdPicture)
    'Escolher a imagem normal
    Set m_Picture1 = vNewPic
    PropertyChanged "Picture1"
    Image1.Picture = Picture1
End Property

Public Property Get Picture2() As StdPicture
    'Obter a imagem normal
    Set Picture2 = m_Picture2
End Property

Public Property Set Picture2(vNewPic As StdPicture)
    'Escolher a imagem normal
    Set m_Picture2 = vNewPic
    PropertyChanged "Picture2"
    Image2.Picture = Picture2
End Property

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
