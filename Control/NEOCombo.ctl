VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.UserControl NEOComboFile 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BackStyle       =   0  'Transparent
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5490
   ScaleHeight     =   375
   ScaleWidth      =   5490
   ToolboxBitmap   =   "NEOCombo.ctx":0000
   Begin VB.PictureBox Bar_Teks_Bahasa 
      Appearance      =   0  'Flat
      BackColor       =   &H00F9F9F9&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   5475
      TabIndex        =   3
      Top             =   0
      Width           =   5475
      Begin VB.PictureBox Set_Bahasa 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   5160
         Picture         =   "NEOCombo.ctx":0312
         ScaleHeight     =   345
         ScaleWidth      =   285
         TabIndex        =   5
         Top             =   15
         Width           =   285
         Begin VB.Image Icon_Set 
            Enabled         =   0   'False
            Height          =   450
            Left            =   0
            Picture         =   "NEOCombo.ctx":0A5C
            Top             =   30
            Width           =   285
         End
      End
      Begin VB.TextBox Text_Bahasa 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   195
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   85
         Width           =   5280
      End
      Begin VB.Shape Line_Bahasa 
         BackColor       =   &H00F9F9F9&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   375
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   5475
      End
   End
   Begin MSFlexGridLib.MSFlexGrid List_Bahasa 
      Height          =   1515
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   370
      Visible         =   0   'False
      Width           =   5455
      _ExtentX        =   9631
      _ExtentY        =   2672
      _Version        =   393216
      Rows            =   19
      Cols            =   16
      BackColor       =   16777215
      BackColorFixed  =   16777215
      BackColorSel    =   4789739
      BackColorBkg    =   16777215
      GridColor       =   16777215
      GridColorFixed  =   16777215
      FocusRect       =   0
      GridLines       =   0
      GridLinesFixed  =   0
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
   End
   Begin VB.FileListBox File_Bahasa 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF80FF&
      Height          =   810
      Left            =   1980
      Pattern         =   "*.lng"
      TabIndex        =   2
      Top             =   1740
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.DirListBox Dir_Bahasa 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF80FF&
      Height          =   765
      Left            =   -60
      TabIndex        =   1
      Top             =   1740
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   4140
      TabIndex        =   6
      Top             =   2460
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "NEOComboFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Enum C_Themes
    [Red] = 1
    [Green] = 2
    [Blue] = 3
    [Yellow] = 4
    [Purple] = 5
    [Pink] = 6
    [Custom] = 7
End Enum

Private U_Theme As C_Themes
Event Click()

Private Sub List_Bahasa_Click()
    'Selecionar o dia
    Text_Bahasa.Text = List_Bahasa.TextMatrix(List_Bahasa.Row, 1)
    List_Bahasa.Visible = False
    
'    Text_Bahasa.SetFocus
End Sub

Private Sub List_Bahasa_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Selecionar as linhas da lista com o mouse
    If List_Bahasa.Rows > 1 Then
        If List_Bahasa.Row <> List_Bahasa.MouseRow And List_Bahasa.MouseRow > 0 Then
            List_Bahasa.Col = 0
            List_Bahasa.Row = List_Bahasa.MouseRow
            List_Bahasa.ColSel = List_Bahasa.Cols - 1
        End If
    End If
End Sub

Private Sub Set_Bahasa_Click()
    'Ver/ocultar lista
    If List_Bahasa.Visible = True Then
        List_Bahasa.Visible = False
        UserControl.Height = UserControl.Height - List_Bahasa.Height + 10
    Else
        List_Bahasa.Visible = True
        UserControl.Height = UserControl.Height + List_Bahasa.Height - 10
    End If
End Sub

Public Property Get Text() As String
    Text = Text_Bahasa.Text
End Property

Public Property Let Text(ByVal newText As String)
    Text_Bahasa.Text = newText
    Label1.Caption = newText
    PropertyChanged "Text"
End Property

Public Property Get Font() As Font
    Set Font = Text_Bahasa.Font
End Property

Public Property Set Font(ByRef NewFont As Font)
    Set Text_Bahasa.Font = NewFont
    Set Label1.Font = Text_Bahasa.Font
    'SetAlign ctlVAlign
    PropertyChanged "FONT"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = Bar_Teks_Bahasa.BackColor
End Property

Public Property Let BackColor(ByVal theCol As OLE_COLOR)
    Bar_Teks_Bahasa.BackColor = theCol
    'txt.BackColor = theCol
    PropertyChanged "BackColor"
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = Text_Bahasa.ForeColor
End Property

Public Property Let ForeColor(ByVal theCol As OLE_COLOR)
    Text_Bahasa.ForeColor = theCol
    Label1.ForeColor = theCol
    PropertyChanged "ForeColor"
End Property

Public Property Get BorderColor() As OLE_COLOR
    BorderColor = Line_Bahasa.BorderColor
End Property

Public Property Let BorderColor(ByVal theCol As OLE_COLOR)
    Line_Bahasa.BorderColor = theCol
    PropertyChanged "BorderColor"
End Property

Public Property Get Directory() As String
    'Escolher o nome do botão
    Directory = Dir_Bahasa.Path
End Property

Public Property Let Directory(New_Value As String)
    'Alterar o caption para o novo texto
    Label1.Caption = New_Value
    PropertyChanged "Directory"
End Property

Public Property Get FilePath() As String
    'Escolher o nome do botão
    FilePath = File_Bahasa.Path
End Property

Public Property Let FilePath(New_Value As String)
    'Alterar o caption para o novo texto
    File_Bahasa.Path = New_Value
    PropertyChanged "FilePath"
End Property

Public Property Get FilePattern() As String
    'Escolher o nome do botão
    FilePattern = File_Bahasa.Pattern
End Property

Public Property Let FilePattern(New_Value As String)
    'Alterar o caption para o novo texto
    File_Bahasa.Pattern = New_Value
    PropertyChanged "FilePattern"
End Property

Private Sub Text_Bahasa_Change()
RaiseEvent Click
End Sub

Private Sub Text_Bahasa_DblClick()
Set_Bahasa_Click
End Sub

Private Sub UserControl_Initialize()
    Dir_Bahasa.Path = App.Path & "\Language\"
    File_Bahasa.Path = Dir_Bahasa.Path
    File_Bahasa.Pattern = "*.lng"
    
    With List_Bahasa
        .ColWidth(0) = 0
        .ColWidth(1) = 10000
        .RowHeight(0) = 0
        .Rows = 1
    End With
    Dim i As Integer: i = 1
    Dim Z As Integer
    
    File_Bahasa.ListIndex = 0
    For Z = 0 To File_Bahasa.ListCount - 1
        List_Bahasa.Rows = List_Bahasa.Rows + 1
        List_Bahasa.TextMatrix(i, 1) = Left$(File_Bahasa.List(Z), InStr(File_Bahasa.List(Z), ".") - (1)) 'Retirar a extensão do ficheiro ".lng"
        i = i + 1
    Next Z
    
    'Selecionar a 1ª linha da combo lingua
    If List_Bahasa.Rows > 0 Then List_Bahasa.Row = 1
End Sub

Public Property Let Theme(ByVal NewValue As C_Themes)

    U_Theme = NewValue
    PropertyChanged "Theme"
Bar_Draw
End Property

Public Property Get Theme() As C_Themes
    Theme = U_Theme
End Property

Private Sub Bar_Draw()
On Error Resume Next
CheckTheme

UserControl.Cls
End Sub

Private Sub CheckTheme()
If Theme = 1 Then
'BACK
List_Bahasa.BackColorSel = RGB(215, 57, 37)
ElseIf Theme = 2 Then
'BACK
List_Bahasa.BackColorSel = RGB(91, 221, 21)
ElseIf Theme = 3 Then
'BACK
List_Bahasa.BackColorSel = RGB(0, 176, 223)
ElseIf Theme = 4 Then
'BACK
List_Bahasa.BackColorSel = RGB(255, 128, 0)
ElseIf Theme = 5 Then
'BACK
List_Bahasa.BackColorSel = RGB(142, 72, 171)
ElseIf Theme = 6 Then
'BACK
List_Bahasa.BackColorSel = RGB(237, 20, 73)
ElseIf Theme = 7 Then
'BACK
ColorSelect = &H312D22
List_Bahasa.BackColorSel = ColorSelect
End If
End Sub

Public Property Get ColorSelect() As OLE_COLOR
    'Escolher a cor inicial de fundo do botão
    ColorSelect = List_Bahasa.BackColorSel
End Property

Public Property Let ColorSelect(New_Value As OLE_COLOR)
    'Alterar a cor inicial de fundo do botão
    List_Bahasa.BackColorSel = New_Value
    PropertyChanged "ColorSelect"
End Property


Private Sub UserControl_Paint()
    'Escolher a imagem normal para o estado normal do botao
    Text_Bahasa.Text = Text
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    'Baca Properties Kontrol
    On Error GoTo Baca_Propertis_Kontrol
    With PropBag
        'Set m_Picture1 = .ReadProperty("Picture1", Nothing)
        'Set m_Picture2 = .ReadProperty("Picture2", Nothing)
        'Set m_Picture_Hover = .ReadProperty("Picture_Hover", Nothing)
        'Image1.Stretch = .ReadProperty("Stretch", False)
        
        Text_Bahasa.Text = .ReadProperty("Text", "")
        Label1.Caption = Text_Bahasa.Text
        Set Text_Bahasa.Font = PropBag.ReadProperty("FONT", "Arial")
        Set Label1.Font = Text_Bahasa.Font
        Text_Bahasa.ForeColor = PropBag.ReadProperty("ForeColor", &HD05C28)
        Label1.ForeColor = Text_Bahasa.ForeColor
        Line_Bahasa.BorderColor = PropBag.ReadProperty("BorderColor", &HD05C28)
        
        Bar_Teks_Bahasa.BackColor = PropBag.ReadProperty("BackColor", &HF9F9F9)
        ColorSelect = .ReadProperty("ColorSelect", "&H004915EB&")
        Theme = .ReadProperty("Theme", "")
        
        Directory = .ReadProperty("Directory", Dir_Bahasa.Path)
        FilePath = .ReadProperty("FilePath", File_Bahasa.Path)
        FilePattern = .ReadProperty("FilePattern", File_Bahasa.Pattern)
        Text = .ReadProperty("Text", "NEOCombo")
    End With
    
Baca_Propertis_Kontrol:
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    'Tulis Properties Kontrol
    On Error GoTo Baca_Propertis_Kontrol
    With PropBag
        'Call .WriteProperty("Picture1", m_Picture1, Nothing)
        'Call .WriteProperty("Picture2", m_Picture2, Nothing)
        '.WriteProperty "Picture_Hover", m_Picture_Hover, 0
        '.WriteProperty "Stretch", Image1.Stretch
        
        PropBag.WriteProperty "Text", Text_Bahasa.Text
        PropBag.WriteProperty "FONT", Text_Bahasa.Font
        PropBag.WriteProperty "ForeColor", Text_Bahasa.ForeColor
        PropBag.WriteProperty "BorderColor", Line_Bahasa.BorderColor
        
        PropBag.WriteProperty "BackColor", Bar_Teks_Bahasa.BackColor
        Call .WriteProperty("ColorSelect", ColorSelect, "&H004915EB&")
        .WriteProperty "Theme", U_Theme, ""
        
        Call .WriteProperty("Directory", Directory, Directory)
        Call .WriteProperty("FilePath", FilePath, FilePath)
        Call .WriteProperty("FilePattern", FilePattern, FilePattern)
        Call .WriteProperty("Text", Text, "NEOCombo")
    End With
    
Baca_Propertis_Kontrol:
End Sub


Private Sub UserControl_Resize()
Resize_Control
End Sub

Private Sub Resize_Control()
With Bar_Teks_Bahasa
    .Width = UserControl.Width
    .Height = 375
    .Left = 0
    .Top = 0
End With

With Line_Bahasa
    .Width = Bar_Teks_Bahasa.Width - 15
    .Height = Bar_Teks_Bahasa.Height - 15
    .Left = Bar_Teks_Bahasa.Left + 15
    .Top = Bar_Teks_Bahasa.Top
End With

With Text_Bahasa
    .Left = 120
    .Top = 85
    .Width = Bar_Teks_Bahasa.Width - 180
    .Height = 195 'Bar_Teks_Bahasa.Height - 120
End With

With Set_Bahasa
    .Left = Bar_Teks_Bahasa.Left + Bar_Teks_Bahasa.Width - .Width - 30
    .Top = 15
    .Height = Bar_Teks_Bahasa.Height - 50
End With

With List_Bahasa
    .Top = Bar_Teks_Bahasa.Top + Bar_Teks_Bahasa.Height - 10
    .Width = Bar_Teks_Bahasa.Width
    .Left = Bar_Teks_Bahasa.Left
End With
End Sub
