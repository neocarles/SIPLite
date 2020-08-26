VERSION 5.00
Begin VB.Form Add_Kas 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Input Data Kas"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5955
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Add_Kas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   5955
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PictureBG 
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
      Height          =   2175
      Left            =   120
      ScaleHeight     =   2175
      ScaleWidth      =   5715
      TabIndex        =   0
      Top             =   120
      Width           =   5715
      Begin VB.TextBox TextSaldo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1800
         TabIndex        =   8
         Top             =   1560
         Width           =   1755
      End
      Begin VB.TextBox TextNama 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1800
         TabIndex        =   6
         Top             =   1080
         Width           =   3195
      End
      Begin VB.TextBox TextKode 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1800
         TabIndex        =   4
         Top             =   600
         Width           =   1755
      End
      Begin VB.Label LabelKas 
         Height          =   315
         Left            =   3660
         TabIndex        =   10
         Top             =   600
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Saldo Kas"
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
         Left            =   120
         TabIndex        =   9
         Top             =   1620
         Width           =   1515
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Nama Kas"
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
         Left            =   120
         TabIndex        =   7
         Top             =   1140
         Width           =   1515
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Kode Kas"
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
         Left            =   120
         TabIndex        =   5
         Top             =   660
         Width           =   1515
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Informasi Kas"
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
         TabIndex        =   3
         Top             =   210
         Width           =   1170
      End
      Begin VB.Shape ShapeBorder 
         Height          =   2055
         Left            =   60
         Top             =   60
         Width           =   5595
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000010&
         FillColor       =   &H00404040&
         Height          =   255
         Left            =   180
         Top             =   180
         Width           =   5355
      End
   End
   Begin Project1.isButton ButtonBatal 
      Cancel          =   -1  'True
      Height          =   435
      Left            =   3420
      TabIndex        =   1
      Top             =   2520
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   767
      Style           =   8
      Caption         =   "&Batal [Esc]"
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   0
      RoundedBordersByTheme=   0   'False
   End
   Begin Project1.isButton ButtonSimpan 
      Height          =   435
      Left            =   480
      TabIndex        =   2
      Top             =   2520
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   767
      Style           =   8
      Caption         =   "&Simpan [F8]"
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   0
      RoundedBordersByTheme=   0   'False
   End
End
Attribute VB_Name = "Add_Kas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ClassKas As ClassKas

Private Sub ButtonBatal_Click()
    Unload Me
End Sub

Private Sub ButtonSimpan_Click()
    If blnTambah = True Then    'data baru
        If Trim(TextKode.Text) = "" Or Trim(TextNama.Text) = "" Or Trim(TextSaldo.Text) = "" Then
            Beep
            Exit Sub
        End If

        TextSaldo.Alignment = 1
        TextSaldo.Text = Format(TextSaldo.Text, "#,##0")

        Set ClassKas = New ClassKas

        Sukses = ClassKas.AddKas(TextKode.Text, TextNama.Text, Format(TextSaldo.Text, "###0"), _
                                 Conn)

        If Sukses Then
            FormKas.TampilanAwal
            Unload Me
        Else
            MsgBox "Data Kas gagal disimpan", vbExclamation, "Peringatan"
        End If

    Else
        If Trim(TextKode.Text) = "" Or Trim(TextNama.Text) = "" Or Trim(TextSaldo.Text) = "" Then
            Beep
            Exit Sub
        End If

        Set ClassKas = New ClassKas
        Sukses = ClassKas.UpdateKas(TextKode.Text, LabelKas.Caption, TextNama.Text, Format(TextSaldo.Text, "###0"), _
                                    Conn)

        If Sukses Then
            FormKas.TampilanAwal
            Unload Me
        Else
            MsgBox "Data Kas gagal diperbaharui", vbExclamation, "Peringatan"
        End If

    End If

    Set ClassKas = Nothing    '
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case Is = vbKeyEscape
        ButtonBatal_Click
    Case Is = vbKeyF8
        ButtonSimpan_Click
    End Select
End Sub

Private Sub Form_Load()
    Call SetIcon(Me.hWnd, "FORMICON", False)
    CenterForm Me
End Sub

Private Sub PictureBG_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case Is = vbKeyEscape
        ButtonBatal_Click
    Case Is = vbKeyF8
        ButtonSimpan_Click
    End Select
End Sub

Private Sub TextKode_Click()
    TextSaldo.Alignment = 1
    TextSaldo.Text = Format(TextSaldo.Text, "#,##0")
End Sub

Private Sub TextKode_KeyPress(KeyAscii As Integer)
    GantiPetik KeyAscii
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TextNama_Click()
    TextSaldo.Alignment = 1
    TextSaldo.Text = Format(TextSaldo.Text, "#,##0")
End Sub

Private Sub TextNama_KeyPress(KeyAscii As Integer)
    GantiPetik KeyAscii
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TextSaldo_Click()
    TextSaldo.Alignment = 0
    TextSaldo.Text = Format(TextSaldo.Text, "###0")
End Sub

Private Sub TextSaldo_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then KeyAscii = 0
    If KeyAscii = 13 Then
        TextSaldo.Alignment = 1
        TextSaldo.Text = Format(TextSaldo.Text, "#,##0")
    End If
End Sub
