VERSION 5.00
Begin VB.Form Add_User 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Input Data User"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5955
   Icon            =   "Add_User.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   5955
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PictureBG 
      BorderStyle     =   0  'None
      Height          =   3375
      Left            =   120
      ScaleHeight     =   3375
      ScaleWidth      =   5715
      TabIndex        =   0
      Top             =   120
      Width           =   5715
      Begin VB.ComboBox cmbGender 
         Height          =   315
         ItemData        =   "Add_User.frx":000C
         Left            =   1800
         List            =   "Add_User.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   2520
         Width           =   2355
      End
      Begin VB.ComboBox ComboLevel 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2040
         Width           =   2355
      End
      Begin VB.TextBox TextUser 
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
         TabIndex        =   3
         Top             =   600
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
         TabIndex        =   2
         Top             =   1080
         Width           =   3555
      End
      Begin VB.TextBox TextJabatan 
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
         TabIndex        =   1
         Top             =   1560
         Width           =   3195
      End
      Begin VB.Label LabelWarn 
         AutoSize        =   -1  'True
         Caption         =   "* Harap Teliti, Tidak Bisa Diubah Setelah di Set Hanya Dapat di Hapus"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   360
         TabIndex        =   15
         Top             =   3000
         Width           =   4980
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Jenis Kelamin"
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
         Left            =   180
         TabIndex        =   14
         Top             =   2580
         Width           =   1395
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Kode Level"
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
         Left            =   180
         TabIndex        =   12
         Top             =   2100
         Width           =   1395
      End
      Begin VB.Shape ShapeBorder 
         Height          =   3255
         Left            =   60
         Top             =   60
         Width           =   5595
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Informasi User"
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
         TabIndex        =   8
         Top             =   210
         Width           =   1260
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "User Login*"
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
         Top             =   660
         Width           =   1515
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Nama User"
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
         TabIndex        =   6
         Top             =   1140
         Width           =   1515
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Jabatan"
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
         Top             =   1620
         Width           =   1515
      End
      Begin VB.Label LabelKas 
         Height          =   315
         Left            =   3660
         TabIndex        =   4
         Top             =   600
         Visible         =   0   'False
         Width           =   1815
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
      TabIndex        =   9
      Top             =   3660
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
      TabIndex        =   10
      Top             =   3660
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
Attribute VB_Name = "Add_User"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ClassUser As ClassUser

Private Sub ButtonBatal_Click()
    Unload Me
End Sub

Private Sub ButtonSimpan_Click()
    If blnTambah = True Then    'data baru
        If Trim(TextUser.Text) = "" Or Trim(TextNama.Text) = "" Or Trim(TextJabatan.Text) = "" Or Trim(ComboLevel.Text) = "" Or Trim(cmbGender.Text) = "" Then
            Beep
            Exit Sub
        End If

        Set ClassUser = New ClassUser

        Sukses = ClassUser.AddUser(TextUser.Text, TextNama.Text, TextJabatan.Text, ComboLevel.Text, cmbGender.Text, _
                                   Conn)

        If Sukses Then
            FormUser.TampilanAwal
            Unload Me
        Else
            MsgBox "Data User Gagal Disimpan", vbExclamation, "Peringatan"
        End If

    Else
        If Trim(TextUser.Text) = "" Or Trim(TextNama.Text) = "" Or Trim(TextJabatan.Text) = "" Or Trim(ComboLevel.Text) = "" Or Trim(cmbGender.Text) = "" Then
            Beep
            Exit Sub
        End If

        Set ClassUser = New ClassUser
        Sukses = ClassUser.UpdateUser(TextUser.Text, TextNama.Text, TextJabatan.Text, ComboLevel.Text, cmbGender.Text, _
                                      Conn)

        If Sukses Then
            FormUser.TampilanAwal
            Unload Me
            FormMain.UserProfil.Visible = False
            FormMain.TextPassword.Text = ""
            FormMain.PictureLogin.Visible = True
        Else
            MsgBox "Data User Gagal Disimpan", vbExclamation, "Peringatan"
        End If

    End If

    Set ClassUser = Nothing    '
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case Is = vbKeyEscape
        Unload Me
    Case Is = vbKeyF8
        ButtonSimpan_Click
    End Select
End Sub

Private Sub Form_Load()
    CenterForm Me

    Call initCombo
End Sub

Private Sub initCombo()
    Dim RSKode As New ADODB.Recordset
    Set RSKode = New ADODB.Recordset
    Dim Combo1 As String
    Combo1 = ComboListSingle(ComboLevel, RSKode, "pos_level", "kode", "WHERE id>000 ORDER BY kode ASC", "kode")
End Sub

Private Sub TextKode_KeyPress(KeyAscii As Integer)
    GantiPetik KeyAscii
End Sub

Private Sub TextNama_KeyPress(KeyAscii As Integer)
    GantiPetik KeyAscii
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TextSaldo_KeyPress(KeyAscii As Integer)
    GantiPetik KeyAscii
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
