VERSION 5.00
Begin VB.UserControl HMCommand 
   ClientHeight    =   630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1350
   DefaultCancel   =   -1  'True
   FillColor       =   &H00FFFFFF&
   ScaleHeight     =   630
   ScaleWidth      =   1350
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "HMcommand"
      Height          =   195
      Left            =   180
      TabIndex        =   1
      Top             =   225
      Width           =   960
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   90
      TabIndex        =   0
      Top             =   105
      Width           =   1155
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   45
      Top             =   75
      Width           =   1245
   End
End
Attribute VB_Name = "HMCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim posTop As Integer
Dim posLeft As Integer
'Default Property Values:

Const m_def_BackStyle = 0
'Const m_def_MouseDownColor = 0

'Property Variables:
Dim m_BackStyle As Integer
Dim m_MouseDownColor As OLE_COLOR
Dim temp As OLE_COLOR
'Event Declarations:
Event Click() 'MappingInfo=Label2,Label2,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=Label2,Label2,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=Label2,Label2,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."



'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label2,Label2,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = Label2.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    Label2.BackColor() = New_BackColor
    Label1.BackColor = New_BackColor
    Shape1.BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label2,Label2,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = Label2.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Label2.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label2,Label2,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = Label2.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set Label2.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = m_BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    m_BackStyle = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Shape1,Shape1,-1,BorderStyle
Public Property Get BorderStyle() As BorderStyleConstants
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = Shape1.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleConstants)
    Shape1.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Private Sub Label1_Click()
RaiseEvent Click
'MsgBox AccessKeys
End Sub

Private Sub Label2_Click()
    RaiseEvent Click
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
    posLeft = Label2.Left
 posTop = Label2.Top
 Label2.Move Label2.Left + 17, Label2.Top - 17
 temp = Shape1.BackColor
 Shape1.BackColor = MouseDownColor
 Label1.BackColor = MouseDownColor
 Label2.BackColor = MouseDownColor

End Sub

Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
    Label2.Move posLeft, posTop
    Shape1.BackColor = temp
 Label1.BackColor = temp
 Label2.BackColor = temp
End Sub
Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
    posLeft = Label2.Left
 posTop = Label2.Top
 Label2.Move Label2.Left + 16, Label2.Top - 16
 temp = Shape1.BackColor
 Shape1.BackColor = MouseDownColor
 Label1.BackColor = MouseDownColor
 Label2.BackColor = MouseDownColor

End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
    Label2.Move posLeft, posTop
    Shape1.BackColor = temp
 Label1.BackColor = temp
 Label2.BackColor = temp
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label2,Label2,-1,AutoSize
Public Property Get AutoSize() As Boolean
Attribute AutoSize.VB_Description = "Determines whether a control is automatically resized to display its entire contents."
    AutoSize = Label2.AutoSize
End Property

Public Property Let AutoSize(ByVal New_AutoSize As Boolean)
    Label2.AutoSize() = New_AutoSize
    PropertyChanged "AutoSize"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Shape1,Shape1,-1,BorderColor
Public Property Get BorderColor() As OLE_COLOR
Attribute BorderColor.VB_Description = "Returns/sets the color of an object's border."
    BorderColor = Shape1.BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
    Shape1.BorderColor() = New_BorderColor
    PropertyChanged "BorderColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Shape1,Shape1,-1,BorderWidth
Public Property Get BorderWidth() As Integer
Attribute BorderWidth.VB_Description = "Returns or sets the width of a control's border."
    BorderWidth = Shape1.BorderWidth
End Property

Public Property Let BorderWidth(ByVal New_BorderWidth As Integer)
    Shape1.BorderWidth() = New_BorderWidth
    PropertyChanged "BorderWidth"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label2,Label2,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = Label2.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Label2.Caption() = New_Caption
    PropertyChanged "Caption"
'    UserControl.AccessKeys = Left$(New_Caption, 1)
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get MouseDownColor() As OLE_COLOR
    MouseDownColor = m_MouseDownColor
End Property

Public Property Let MouseDownColor(ByVal New_MouseDownColor As OLE_COLOR)
    m_MouseDownColor = New_MouseDownColor
    PropertyChanged "MouseDownColor"
End Property

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
RaiseEvent Click
End Sub

Private Sub UserControl_GotFocus()
Dim pos As Integer
Label1.BorderStyle = 1
'pos = InStr("&", Caption)
'MsgBox Caption
    AccessKeys = Mid(Caption, 1, 1)

End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_BackStyle = m_def_BackStyle
    m_MouseDownColor = m_def_MouseDownColor
End Sub

Private Sub UserControl_LostFocus()
Label1.BorderStyle = 0
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Label2.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Label1.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Shape1.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Label2.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set Label2.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
    Shape1.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    Label2.AutoSize = PropBag.ReadProperty("AutoSize", True)
    Shape1.BorderColor = PropBag.ReadProperty("BorderColor", -2147483640)
    Shape1.BorderWidth = PropBag.ReadProperty("BorderWidth", 1)
    Label2.Caption = PropBag.ReadProperty("Caption", "HMcommand")
    m_MouseDownColor = PropBag.ReadProperty("MouseDownColor", m_def_MouseDownColor)
End Sub

Private Sub UserControl_Resize()
Shape1.Move 0, 0, ScaleWidth, ScaleHeight
Label1.Move 20, 20, Shape1.Width - 30, Shape1.Height - 30
Label2.Move (Shape1.Width \ 2) - (Label2.Width \ 2), (Shape1.Height \ 2) - (Label2.Height \ 2)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", Label2.BackColor, &H8000000F)
    Call PropBag.WriteProperty("BackColor", Label1.BackColor, &H8000000F)
    Call PropBag.WriteProperty("BackColor", Shape1.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", Label2.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", Label2.Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
    Call PropBag.WriteProperty("BorderStyle", Shape1.BorderStyle, 1)
    Call PropBag.WriteProperty("AutoSize", Label2.AutoSize, True)
    Call PropBag.WriteProperty("BorderColor", Shape1.BorderColor, -2147483640)
    Call PropBag.WriteProperty("BorderWidth", Shape1.BorderWidth, 1)
    Call PropBag.WriteProperty("Caption", Label2.Caption, "HMcommand")
    Call PropBag.WriteProperty("MouseDownColor", m_MouseDownColor, m_def_MouseDownColor)
End Sub

