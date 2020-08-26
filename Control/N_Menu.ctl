VERSION 5.00
Begin VB.UserControl N_Menu 
   BackColor       =   &H00585DF1&
   ClientHeight    =   615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9450
   ScaleHeight     =   41
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   630
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4440
      Top             =   0
   End
   Begin VB.Shape Image_Apontador 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFF00&
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   4080
      Top             =   60
      Width           =   75
   End
   Begin VB.Image Seta_Menu 
      Enabled         =   0   'False
      Height          =   90
      Left            =   3720
      Picture         =   "N_Menu.ctx":0000
      Top             =   240
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Image Icon_User 
      Height          =   375
      Left            =   3120
      Picture         =   "N_Menu.ctx":0049
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label_Linha1 
      BackColor       =   &H0025221A&
      Enabled         =   0   'False
      Height          =   15
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Label Label_Linha2 
      BackColor       =   &H0025221A&
      Enabled         =   0   'False
      Height          =   15
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Image Botao_Expandir 
      Height          =   405
      Left            =   8280
      Top             =   120
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image Image_Apontadors 
      Enabled         =   0   'False
      Height          =   480
      Left            =   1680
      Picture         =   "N_Menu.ctx":01BC
      Top             =   0
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image Fundo_Menu 
      Height          =   600
      Left            =   5880
      Top             =   0
      Visible         =   0   'False
      Width           =   3525
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0042362A&
      BackStyle       =   0  'Transparent
      Caption         =   "NShape1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   765
   End
   Begin VB.Image Image_Icon 
      Enabled         =   0   'False
      Height          =   390
      Left            =   60
      Top             =   45
      Width           =   390
   End
   Begin VB.Label Fundo_Menu_Activo 
      Appearance      =   0  'Flat
      BackColor       =   &H009AB01C&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   5160
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "N_Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'Component: N_Menu
'Copyright (c) 2014 Nikyts Software - Informatic and thecnologies
'Developed by Nelson do Carmo
'Contact: nikyts@hotmail.com
'Web: www.nikyts.net
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'Declaração de variáveis
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

'Cores
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

'Api's utilizadas pelo componente
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
        X As Long
        Y As Long
End Type
Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long) As Long
Private Declare Function SelectClipPath Lib "gdi32" (ByVal hDC As Long, ByVal iMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
'Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long ' api repetida
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long

'Para animar o icon
Enum PicStates
    picNothing = 0
    picDown = 1
    picHover = 2
    picNorm = 3
    picActive = 4
End Enum
Dim PicState As PicStates
Dim m_Picture_Normal As New StdPicture
Dim m_Picture_Hover As New StdPicture
Dim m_Picture_Down As New StdPicture
Dim m_Picture_Active As New StdPicture
Dim m_MouseInside As Boolean

'Dim m_Picture_Pointer As New StdPicture

'Animar o botao/cor da shape
Dim isOver As Boolean
Dim m_State As Integer
Event Click()
Event DblClick()
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseEnter()
Event MouseLeave()
Public Event MouseOver()
Public Event MouseOut()
Public Event Resize()

'Eventos para as teclas
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)

'Variável para saber se o botao tem icon
Dim m_Sub_Menu As Boolean
Dim m_Menu_Expandido As Boolean
Dim m_Menu_Activo As Boolean
Dim m_Sub_Menu_Activo As Boolean
Dim m_Menu_Popup As Boolean

'Cores utilizadas pelos eventos do botao
Dim cor_fundo_normal As OLE_COLOR
Dim cor_fundo_hover As OLE_COLOR
Dim cor_fundo_down As OLE_COLOR
Dim cor_fundo_ativo As OLE_COLOR
Dim m_BackColorActive As OLE_COLOR

Dim cor_letra_normal As OLE_COLOR
Dim cor_letra_hover As OLE_COLOR
Dim cor_letra_down As OLE_COLOR
Dim cor_letra_active As OLE_COLOR

Dim m_Linha_Divisoria As Boolean
Dim m_Primeiro_Menu As Boolean

'Carregar PNG
'Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private Sub Label1_Change()
    'Ajustar o botao ao alterar o texto da label
    Call Ajustar_Botao
End Sub

Private Sub Label1_Click()
    'Atalho para
    UserControl_Click
End Sub

Private Sub Label1_DblClick()
    'Atalho para
    UserControl_DblClick
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Evento mousedown
    RaiseEvent MouseDown(Button, Shift, X, Y)
    If Button = 1 Then
        m_State = 2
        Call DrawState
    End If
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Evento mousemove
    RaiseEvent MouseMove(Button, Shift, X, Y)
    
    If Menu_Activo = True And Menu_Popup = False Then Exit Sub 'Or Sub_Menu_Activo = True
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

    'Icon do botão
    If Button = vbLeftButton Then
        Exit Sub
    End If

    If GetCapture() <> UserControl.hwnd Then
        SetCapture (UserControl.hwnd)
        If Not Image_Icon.Picture = Picture_Hover Then
            Image_Icon.Picture = Picture_Hover
            m_MouseInside = True
        End If
    Else
        Dim pt As POINTAPI
        pt.X = X
        pt.Y = Y
        ClientToScreen UserControl.hwnd, pt
        If WindowFromPoint(pt.X, pt.Y) <> UserControl.hwnd Then
            Refresh
            If Button <> vbLeftButton Then
                ReleaseCapture
                Image_Icon.Picture = Picture_Normal
                m_MouseInside = False
                RaiseEvent MouseOut
            End If
            End If
    End If
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Evento mouseup
    RaiseEvent MouseUp(Button, Shift, X, Y)
    If CheckMouseOver Then
        m_State = 1
    Else
        m_State = 0
    End If
    
    RaiseEvent Click
'    Menu_Activo = True
'    Call DrawState
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    'Evento keypress
    LastButton = 1
    Call UserControl_Click
End Sub

Private Sub UserControl_Click()
    'Evento click
    'RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    'Evento duploclick
    RaiseEvent DblClick
End Sub

Private Sub UserControl_Initialize()
    'Iniciando o componente
    'On Error GoTo Corrige_Erro
    Call Ajustar_Botao
    If Not Picture_Normal Is Nothing Then
        Set Image_Icon.Picture = Picture_Normal
    End If
    
Exit Sub
Corrige_Erro:
End Sub

Public Property Get Caption() As String
    'Escolher o nome do botão
    Caption = Label1.Caption
End Property

Public Property Let Caption(New_Value As String)
    'Alterar o caption para o novo texto
    Label1.Caption = New_Value
    PropertyChanged "Caption"
End Property

Private Sub UserControl_InitProperties()
    'Ler as propriedades do botão
    'On Error GoTo Corrige_Erro
    Caption = UserControl.Extender.Name '"N_Menu1"
    Enabled = True
    BackColorNormal = RGB(54, 65, 80)
    BackColorHover = RGB(44, 53, 66) 'RGB(62, 75, 92)
    BackColorDown = RGB(28, 175, 154)
    BackColorActive = RGB(28, 175, 154)
    FontSize = 8
    FontBold = False
    Font = "Verdana" '"MS Sans Serif"
    ForeColorNormal = RGB(255, 255, 255)
    ForeColorHover = RGB(255, 255, 255)
    ForeColorDown = RGB(255, 255, 255)
    ForeColorActive = RGB(255, 255, 255)
    Set Image_Icon.Picture = Picture_Normal
'    Set Image_Apontador.Picture = Picture_Pointer
    Call Ajustar_Botao
    Sub_Menu = False
    Menu_Expandido = True
    Menu_Activo = False
    Sub_Menu_Activo = False
    Menu_Popup = False
    Linha_Divisoria = False
    Primeiro_Menu = False
    
Exit Sub
Corrige_Erro:
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    'Evento keydown, atalho para as teclas
    Select Case KeyCode
        Case vbKeyRight, vbKeyDown
            Call SendKeys("{TAB}")
            Case vbKeyLeft, vbKeyUp
            Call SendKeys("+{TAB}")
        
        Case vbKeyReturn
            UserControl_Click
    End Select
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    'Evento keypress
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    'Evento keyup
    RaiseEvent KeyUp(KeyCode, Shift)
    UserControl.Refresh
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Evento mousedown
    RaiseEvent MouseDown(Button, Shift, X, Y)
    If Button = 1 Then
        m_State = 2
        Call DrawState
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Evento mousemove
    RaiseEvent MouseMove(Button, Shift, X, Y)
    
    If Menu_Activo = True And Menu_Popup = False Then Exit Sub ' Or Sub_Menu_Activo = True
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

    'Icon do botão
    If Button = vbLeftButton Then
        Exit Sub
    End If

    If GetCapture() <> UserControl.hwnd Then
        SetCapture (UserControl.hwnd)
        If Not Image_Icon.Picture = Picture_Hover Then
            Image_Icon.Picture = Picture_Hover
            m_MouseInside = True
        End If
    Else
        Dim pt As POINTAPI
        pt.X = X
        pt.Y = Y
        ClientToScreen UserControl.hwnd, pt
        If WindowFromPoint(pt.X, pt.Y) <> UserControl.hwnd Then
            Refresh
            If Button <> vbLeftButton Then
                ReleaseCapture
                Image_Icon.Picture = Picture_Normal
                m_MouseInside = False
                RaiseEvent MouseOut
            End If
            End If
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Evento mouseup
    RaiseEvent MouseUp(Button, Shift, X, Y)
    
    If CheckMouseOver Then
        m_State = 1
    Else
        m_State = 0
    End If
    
    RaiseEvent Click
'    Menu_Activo = True
'    Call DrawState
End Sub

Private Function CheckMouseOver() As Boolean
    'Efectuar o over do botao
    Dim pt As POINTAPI
    GetCursorPos pt
    CheckMouseOver = (WindowFromPoint(pt.X, pt.Y) = UserControl.hwnd)
End Function

Private Sub DrawState()
    On Error Resume Next
    If Menu_Activo = False Then
        If m_State = 1 Then 'mouse hover
            UserControl.BackColor = cor_fundo_hover
            Label1.ForeColor = cor_letra_hover
            Set Image_Icon.Picture = Picture_Hover
            Image_Apontador.Visible = False
        
        ElseIf m_State = 2 Then 'mouse down
            UserControl.BackColor = cor_fundo_down
            Label1.ForeColor = cor_letra_down
            Set Image_Icon.Picture = Picture_Down
            Call Ajustar_Botao
            'If Sub_Menu = False Then Image_Apontador.Visible = True
            
        Else 'normal 'If m_State = 3 Then
            UserControl.BackColor = cor_fundo_normal
            Label1.ForeColor = cor_letra_normal
            Set Image_Icon.Picture = Picture_Normal
            Image_Apontador.Visible = False
'
'        ElseIf m_State = 4 Then 'activo
'            UserControl.BackColor = cor_fundo_ativo
'            Label1.ForeColor = cor_letra_normal
'            Set Image_Icon.Picture = Picture_Active
'            Image_Apontador.Visible = True
        End If
    End If
End Sub

Private Sub Timer1_Timer()
    'Animar o botão
    If Not CheckMouseOver Then
        Timer1.Enabled = False
        isOver = False
        RaiseEvent MouseLeave
        m_State = 0
        Call DrawState
    End If
End Sub

Public Property Get FontSize() As Integer
    'Escolher o tamanho da letra
    FontSize = Label1.FontSize
End Property

Public Property Let FontSize(New_Value As Integer)
    'Alterar o tamanho da letra
    Label1.FontSize = New_Value
    PropertyChanged "FontSize"
End Property

Public Property Get FontBold() As Boolean
    'Indicar se que ou não a letra em negrito
    FontBold = Label1.FontBold
End Property

Public Property Let FontBold(New_Value As Boolean)
    'Alterar a letra para negrito se for o caso
    Label1.FontBold = New_Value
    PropertyChanged "FontBold"
End Property

Public Property Get Enabled() As Boolean
    'Escolher se o botão fica activo ou não
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(New_Value As Boolean)
    'Verificar se o botão fica activo ou não
    UserControl.Enabled = New_Value
    PropertyChanged "Enabled"

'    If New_Value = True Then
'        Label1.ForeColor = cor_letra_normal
'    Else
'        Label1.ForeColor = &HB3B3B3  'Cinzento escuro
'    End If
    UserControl.Refresh
End Property

Public Property Get Picture_Normal() As StdPicture
    'Obter a imagem normal
    Set Picture_Normal = m_Picture_Normal
End Property

Public Property Set Picture_Normal(vNewPic As StdPicture)
    'Escolher a imagem normal
    Set m_Picture_Normal = vNewPic
    PropertyChanged "Picture_Normal"
    Image_Icon.Picture = Picture_Normal
End Property

Public Property Get Picture_Down() As StdPicture
    'Obter a imagem down
    Set Picture_Down = m_Picture_Down
End Property

Public Property Set Picture_Down(vNewPic As StdPicture)
    'Escolher a imagem down
    Set m_Picture_Down = vNewPic
    PropertyChanged "Picture_Down"
End Property

Public Property Get Picture_Hover() As StdPicture
    'Obter a imagem over
    Set Picture_Hover = m_Picture_Hover
End Property

Public Property Set Picture_Hover(vNewPic As StdPicture)
    'Escolher a imagem over
    Set m_Picture_Hover = vNewPic
    PropertyChanged "Picture_Hover"
End Property

Public Property Get Picture_Active() As StdPicture
    'Obter a imagem down
    Set Picture_Active = m_Picture_Active
End Property

Public Property Set Picture_Active(vNewPic As StdPicture)
    'Escolher a imagem down
    Set m_Picture_Active = vNewPic
    PropertyChanged "Picture_Active"
End Property

Private Sub UserControl_Paint()
    'Escolher a imagem normal para o estado normal do botao
    'Image_Icon.Picture = Picture_Normal
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    'Ler as propriedades do control
    'On Error GoTo Corrige_Erro
    With PropBag
        Set m_Picture_Normal = .ReadProperty("Picture_Normal", Nothing)
        Set m_Picture_Down = .ReadProperty("Picture_Down", Nothing)
        Set m_Picture_Hover = .ReadProperty("Picture_Hover", Nothing)
        Set m_Picture_Active = .ReadProperty("Picture_Active", Nothing)
        Caption = .ReadProperty("Caption", "N_Menu1")
        BackColorNormal = .ReadProperty("BackColorNormal", "&H8080FF")
        BackColorHover = .ReadProperty("BackColorHover", "&H80C0FF")
        BackColorDown = .ReadProperty("BackColorDown", "&H80FFFF")
        BackColorActive = .ReadProperty("BackColorActive", "&H80FFFF")
        FontSize = .ReadProperty("FontSize", "8")
        FontBold = .ReadProperty("FontBold", "False")
        Font = .ReadProperty("Font", "MS Sans Serif")
        Enabled = .ReadProperty("Enabled", "True")
        Sub_Menu = .ReadProperty("Sub_Menu", "False")
        ForeColorNormal = .ReadProperty("ForeColorNormal", "&H8080FF")
        ForeColorHover = .ReadProperty("ForeColorHover", "&H8080FF")
        ForeColorDown = .ReadProperty("ForeColorDown", "&H8080FF")
        ForeColorActive = .ReadProperty("ForeColorActive", "&H8080FF")
        novo_icon = .ReadProperty("List_Icons", 1)
        Menu_Expandido = .ReadProperty("Menu_Expandido", "False")
        Menu_Activo = .ReadProperty("Menu_Activo", "False")
        Sub_Menu_Activo = .ReadProperty("Sub_Menu_Activo", "False")
'        Set m_Picture_Pointer = .ReadProperty("Picture_Pointer", Nothing)
        Menu_Popup = .ReadProperty("Menu_Popup", "False")
        Linha_Divisoria = .ReadProperty("Linha_Divisoria", "False")
        Primeiro_Menu = .ReadProperty("Primeiro_Menu", "False")
    End With
    
Corrige_Erro:
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    'Escrever as propriedades do control
    'On Error GoTo Corrige_Erro
    With PropBag
        .WriteProperty "Picture_Normal", m_Picture_Normal, 0
        .WriteProperty "Picture_Down", m_Picture_Down, 0
        .WriteProperty "Picture_Hover", m_Picture_Hover, 0
        .WriteProperty "Picture_Active", m_Picture_Active, 0
        Call .WriteProperty("Caption", Caption, "")
        Call .WriteProperty("BackColorNormal", BackColorNormal, "&H8080FF")
        Call .WriteProperty("BackColorHover", BackColorHover, "&H80C0FF")
        Call .WriteProperty("BackColorDown", BackColorDown, "&H80FFFF")
        Call .WriteProperty("BackColorActive", BackColorActive, "&H80FFFF")
        Call .WriteProperty("FontSize", FontSize, "8")
        Call .WriteProperty("FontBold", FontBold, "False")
        Call .WriteProperty("Font", Font, "MS Sans Serif")
        Call .WriteProperty("Enabled", Enabled, "True")
        Call .WriteProperty("Sub_Menu", Sub_Menu, "False")
        Call .WriteProperty("ForeColorNormal", ForeColorNormal, "&HFF8080")
        Call .WriteProperty("ForeColorHover", ForeColorHover, "&HFF8080")
        Call .WriteProperty("ForeColorDown", ForeColorDown, "&HFF8080")
        Call .WriteProperty("ForeColorActive", ForeColorActive, "&HFF8080")
        Call .WriteProperty("List_Icons", novo_icon, 1)
        Call .WriteProperty("Menu_Expandido", Menu_Expandido, "False")
        Call .WriteProperty("Menu_Activo", Menu_Activo, "False")
        Call .WriteProperty("Sub_Menu_Activo", Sub_Menu_Activo, "False")
'        .WriteProperty "Picture_Pointer", m_Picture_Pointer, 0
        Call .WriteProperty("Menu_Popup", Menu_Popup, "False")
        Call .WriteProperty("Linha_Divisoria", Linha_Divisoria, "False")
        Call .WriteProperty("Primeiro_Menu", Primeiro_Menu, "False")
    End With
    
Corrige_Erro:
End Sub

Public Property Get Sub_Menu() As Boolean
    'Escolher se o botão fica activo ou não
    Sub_Menu = m_Sub_Menu
End Property

Public Property Let Sub_Menu(New_Value As Boolean)
    'Verificar se o botão fica activo ou não
    m_Sub_Menu = New_Value
    PropertyChanged "Sub_Menu"

    If m_Sub_Menu = True Then
        Image_Icon.Visible = False
    Else
        Image_Icon.Visible = True
    End If
    
    Call Ajustar_Botao
End Property

Public Property Get Menu_Expandido() As Boolean
    'Escolher se o botão fica activo ou não
    Menu_Expandido = m_Menu_Expandido
End Property

Public Property Let Menu_Expandido(New_Value As Boolean)
    'Verificar se o botão fica activo ou não
    m_Menu_Expandido = New_Value
    PropertyChanged "Menu_Expandido"

    If m_Menu_Expandido = True Then
        Label1.Visible = True
    Else
        Label1.Visible = False
    End If
    
    Call Ajustar_Botao
End Property

Public Property Get Menu_Activo() As Boolean
    'Escolher se o botão fica activo ou não
    Menu_Activo = m_Menu_Activo
End Property

Public Property Let Menu_Activo(New_Value As Boolean)
    'Verificar se o botão fica activo ou não
    m_Menu_Activo = New_Value
    PropertyChanged "Menu_Activo"

    If Sub_Menu = False Then
        If m_Menu_Activo = True Then
            Timer1.Enabled = False
            
            Set Image_Icon.Picture = Picture_Active
            'Carregar a imagem PNG do apontador, quando o menu está ativo
            If Menu_Popup = False Then
                Dim Token As Long
                Dim c As Long
                c = UserControl.BackColor
                If c < 0 Then c = GetSysColor(c - &H80000000)
                Token = InitGDIPlus
'                Set Image_Apontador.Picture = LoadPictureGDIPlus(App.Path & "\Images\Icon_Menu_Pointer.png", , , m_BackColorActive)
                FreeGDIPlus Token
                Call Ajustar_Botao
                Fundo_Menu_Activo.Visible = True
                Image_Apontador.Visible = True
                Label1.ForeColor = ForeColorActive
            End If
        Else
            Timer1.Enabled = True
            Set Image_Icon.Picture = Picture_Normal
            Call Ajustar_Botao
            Image_Apontador.Visible = False
            Fundo_Menu_Activo.Visible = False
            Label1.ForeColor = ForeColorNormal
        End If
    End If
End Property

Private Sub UserControl_Resize()
    'Desenhar o botão ajustando os controles do form
    RaiseEvent Resize
    Ajustar_Botao
End Sub

Public Property Get CurPicture() As PicStates
    'Verificar qual é a imagem a ser vista
    If Image_Icon.Picture = 0 Then
        CurPicture = picNothing
    ElseIf Image_Icon.Picture = Picture_Normal Then
        CurPicture = picNorm
    ElseIf Image_Icon.Picture = Picture_Down Then
        CurPicture = picDown
    ElseIf Image_Icon.Picture = Picture_Hover Then
        CurPicture = picHover
    ElseIf Image_Icon.Picture = Picture_Active Then
        CurPicture = picActive
    End If
End Property

Public Property Get MouseInside() As Boolean
    'Mouse em cima do cntrol
    MouseInside = m_MouseInside
End Property

Public Sub Ajustar_Botao()
    'Actualizar o botão
    On Error Resume Next
    With UserControl
        'If Sub_Menu = False Then
        '    .Height = Screen.TwipsPerPixelY * Fundo_Menu.Height
        'Else
        '    .Height = Screen.TwipsPerPixelY * Fundo_Sub_Menu.Height
        'End If
        If Menu_Popup = False Then
            If Sub_Menu = False Then
                If m_Menu_Expandido = True Then
                    .Width = Screen.TwipsPerPixelX * (Fundo_Menu.Width + 19)
                Else
                    .Width = Screen.TwipsPerPixelX * (5 + Botao_Expandir.Width + 19)
                End If
            End If
        
        Else
            .Width = Screen.TwipsPerPixelX * (7 + Icon_User.Width + 5 + Label1.Width + 7 + 20) 'Seta_Menu.Width + 10)
        End If
    End With
    
    With Image_Icon
        .Top = (UserControl.ScaleHeight - .Height) / 2
        If m_Menu_Expandido = True Then
            .Left = 18
        Else
            .Left = ((UserControl.ScaleWidth) - .Width) / 2
        End If
    End With

    With Icon_User
        .Top = (UserControl.ScaleHeight - .Height) / 2
        .Left = 18
    End With
    
    With Label1
        .Top = (UserControl.ScaleHeight - .Height) / 2
        If Menu_Popup = False Then
            If Sub_Menu = False Then
                .Left = 55
            Else
                .Left = 25
            End If
        Else
            .Left = 7 + Icon_User.Width + 15
        End If
    End With
    
    With Image_Apontador
        .Top = (UserControl.ScaleHeight - .Height) / 2
        .Left = UserControl.ScaleLeft - .Width + 3 'UserControl.ScaleWidth - .Width + 1
        .Height = UserControl.ScaleHeight
    End With
    
    With Fundo_Menu_Activo
        .Top = 0
        .Height = UserControl.ScaleHeight
        .Left = 0
        .Width = UserControl.ScaleWidth
    End With
    
    'With Seta_Menu
    '    .Top = ((UserControl.ScaleHeight - .Height) / 2) + 2
    '    .Left = Label1.Left + Label1.Width + 7
    'End With
    
    With Label_Linha1
        .Top = 0
        .Left = 0
        .Width = UserControl.ScaleWidth
    End With
    
    With Label_Linha2
        .Top = UserControl.ScaleHeight - .Height
        .Left = 0
        .Width = UserControl.ScaleWidth
    End With
End Sub

Public Property Get Linha_Divisoria() As Boolean
    'Escolher se o botão fica activo ou não
    Linha_Divisoria = m_Linha_Divisoria
End Property

Public Property Let Linha_Divisoria(New_Value As Boolean)
    'Verificar se o botão fica activo ou não
    m_Linha_Divisoria = New_Value
    PropertyChanged "Linha_Divisoria"

    If m_Linha_Divisoria = True Then
        Label_Linha2.Visible = True
    Else
        Label_Linha2.Visible = False
    End If
End Property

Public Property Get ForeColorNormal() As OLE_COLOR
    'Escolher a cor inicial de fundo do botão
    ForeColorNormal = Label1.ForeColor
End Property

Public Property Let ForeColorNormal(New_Value As OLE_COLOR)
    'Alterar a cor inicial de fundo do botão
    cor_letra_normal = New_Value
    Label1.ForeColor = New_Value
    PropertyChanged "ForeColorNormal"
End Property

Public Property Get ForeColorHover() As OLE_COLOR
    'Escolher a cor inicial da letra do botão
    ForeColorHover = cor_letra_hover
End Property

Public Property Let ForeColorHover(New_Value As OLE_COLOR)
    'Alterar a cor inicial da letra do botão
    cor_letra_hover = New_Value
    'Label1.ForeColor = New_Value
    PropertyChanged "ForeColorHover"
End Property

Public Property Get ForeColorDown() As OLE_COLOR
    'Escolher a cor inicial da letra do botão
    ForeColorDown = cor_letra_down
End Property

Public Property Let ForeColorDown(New_Value As OLE_COLOR)
    'Alterar a cor inicial da letra do botão
    cor_letra_down = New_Value
    'Label1.ForeColor = New_Value
    PropertyChanged "ForeColorDown"
End Property

Public Property Get ForeColorActive() As OLE_COLOR
    'Escolher a cor inicial da letra do botão
    ForeColorActive = cor_letra_active
End Property

Public Property Let ForeColorActive(New_Value As OLE_COLOR)
    'Alterar a cor inicial da letra do botão
    cor_letra_active = New_Value
    'Label1.ForeColor = New_Value
    PropertyChanged "ForeColorActive"
End Property

Public Property Get BackColorNormal() As OLE_COLOR
    'Escolher a cor inicial de fundo do botão
    BackColorNormal = UserControl.BackColor
End Property

Public Property Let BackColorNormal(New_Value As OLE_COLOR)
    'Alterar a cor inicial de fundo do botão
    'On Error GoTo Corrige_Erro
    cor_fundo_normal = New_Value
    UserControl.BackColor = New_Value
    PropertyChanged "BackColorNormal"

Corrige_Erro:
End Property

Public Property Get BackColorHover() As OLE_COLOR
    'Escolher a cor inicial de fundo do botão
    BackColorHover = cor_fundo_hover
End Property

Public Property Let BackColorHover(New_Value As OLE_COLOR)
    'Alterar a cor inicial de fundo do botão
    'On Error GoTo Corrige_Erro
    cor_fundo_hover = New_Value
    'UserControl.BackColor = new_Value
    PropertyChanged "BackColorHover"
    
Corrige_Erro:
End Property

Public Property Get BackColorDown() As OLE_COLOR
    'Escolher a cor inicial de fundo do botão
    BackColorDown = cor_fundo_down
End Property

Public Property Let BackColorDown(New_Value As OLE_COLOR)
    'Alterar a cor inicial de fundo do botão
    'On Error GoTo Corrige_Erro
    cor_fundo_down = New_Value
    'UserControl.BackColor = new_Value
    PropertyChanged "BackColorDown"
    
Corrige_Erro:
End Property
'
'Public Property Get Picture_Pointer() As StdPicture
'    'Obter a imagem normal
'    Set Picture_Pointer = m_Picture_Pointer
'End Property
'
'Public Property Set Picture_Pointer(vNewPic As StdPicture)
'    'Escolher a imagem Pointer
'    Set m_Picture_Pointer = vNewPic
'    PropertyChanged "Picture_Pointer"
''    Image_Apontador.Picture = Picture_Pointer
'End Property

Public Property Get BackColorActive() As OLE_COLOR
    'Escolher a cor inicial de fundo do botão
    BackColorActive = m_BackColorActive
End Property

Public Property Let BackColorActive(New_Value As OLE_COLOR)
    'Alterar a cor inicial de fundo do botão
    'On Error GoTo Corrige_Erro
    m_BackColorActive = New_Value
    Fundo_Menu_Activo.BackColor = New_Value
    PropertyChanged "BackColorActive"
    
Corrige_Erro:
End Property

Public Property Get Sub_Menu_Activo() As Boolean
    'Escolher se o botão fica activo ou não
    Sub_Menu_Activo = m_Sub_Menu_Activo
End Property

Public Property Let Sub_Menu_Activo(New_Value As Boolean)
    'Verificar se o botão fica activo ou não
    m_Sub_Menu_Activo = New_Value
    PropertyChanged "Sub_Menu_Activo"
    
'    If Sub_Menu = False Then
        If m_Sub_Menu_Activo = True Then
            Timer1.Enabled = False
'            Set Image_Icon.Picture = Picture_Active
'            'Carregar a imagem PNG do apontador, quando o menu está ativo
'            Dim Token As Long
'            Dim c As Long
'            c = UserControl.BackColor
'            If c < 0 Then c = GetSysColor(c - &H80000000)
'            Token = InitGDIPlus
'            Set Image_Apontador.Picture = LoadPictureGDIPlus(App.Path & "\Images\Icon_Menu_Pointer.png", , , m_BackColorActive)
'            FreeGDIPlus Token
'            Call Ajustar_Botao
'            Image_Apontador.Visible = True
            Fundo_Menu_Activo.Visible = True
            Label1.ForeColor = ForeColorActive
        Else
            Timer1.Enabled = True
'            Set Image_Icon.Picture = Picture_Normal
'            Call Ajustar_Botao
'            Image_Apontador.Visible = False
            Fundo_Menu_Activo.Visible = False
            Label1.ForeColor = ForeColorNormal
        End If
'    End If
End Property

Public Property Get Menu_Popup() As Boolean
    'Escolher se o botão fica activo ou não
    Menu_Popup = m_Menu_Popup
End Property

Public Property Let Menu_Popup(New_Value As Boolean)
    'Verificar se o botão fica activo ou não
    m_Menu_Popup = New_Value
    PropertyChanged "Menu_Popup"
    
    If m_Menu_Popup = True Then
        'Seta_Menu.Visible = True
        Icon_User.Visible = True
        Image_Icon.Visible = False
    Else
        Seta_Menu.Visible = False
        Icon_User.Visible = False
        Image_Icon.Visible = True
    End If
    Call Ajustar_Botao
End Property

Public Property Get Primeiro_Menu() As Boolean
    'Escolher se o botão fica activo ou não
    Primeiro_Menu = m_Primeiro_Menu
End Property

Public Property Let Primeiro_Menu(New_Value As Boolean)
    'Verificar se o botão fica activo ou não
    m_Primeiro_Menu = New_Value
    PropertyChanged "Primeiro_Menu"

    If m_Primeiro_Menu = True Then
        Label_Linha1.Visible = True
    Else
        Label_Linha1.Visible = False
    End If
End Property
