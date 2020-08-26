VERSION 5.00
Begin VB.Form Add_Karyawan 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Input Data Karyawan"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8370
   Icon            =   "Add_Karyawan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   8370
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PictureBG 
      BorderStyle     =   0  'None
      Height          =   4275
      Left            =   120
      ScaleHeight     =   4275
      ScaleWidth      =   8115
      TabIndex        =   0
      Top             =   120
      Width           =   8115
      Begin VB.TextBox TextKodeLama 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   5580
         TabIndex        =   21
         Top             =   1080
         Width           =   1275
      End
      Begin VB.TextBox TextGaji 
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
         Left            =   1860
         TabIndex        =   20
         Top             =   3480
         Width           =   2775
      End
      Begin VB.TextBox TextHK 
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
         Left            =   1860
         TabIndex        =   19
         Top             =   3000
         Width           =   1575
      End
      Begin VB.TextBox TextID 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7020
         TabIndex        =   17
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox TextAuto 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5580
         TabIndex        =   16
         Top             =   600
         Width           =   1275
      End
      Begin VB.ComboBox ComboDepartemen 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   600
         Width           =   2115
      End
      Begin VB.TextBox TextKontak 
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
         Left            =   1860
         TabIndex        =   5
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox TextAlamat 
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
         Left            =   1860
         TabIndex        =   4
         Top             =   1560
         Width           =   4935
      End
      Begin VB.TextBox TextStatus 
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
         Left            =   1860
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   2520
         Width           =   4935
      End
      Begin VB.TextBox TextKode 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Left            =   4080
         TabIndex        =   2
         Top             =   600
         Width           =   1335
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
         Left            =   1860
         TabIndex        =   1
         Top             =   1080
         Width           =   3555
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Basic Gaji"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   18
         Top             =   3540
         Width           =   1515
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Informasi Karyawan"
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
         Left            =   300
         TabIndex        =   12
         Top             =   210
         Width           =   1710
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Basic HK"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   11
         Top             =   3060
         Width           =   1515
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Status Karyawan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   10
         Top             =   2580
         Width           =   1515
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Kontak"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   9
         Top             =   2100
         Width           =   1515
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Alamat"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   8
         Top             =   1620
         Width           =   1515
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Nama"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   7
         Top             =   1140
         Width           =   1515
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Departemen"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   6
         Top             =   660
         Width           =   1515
      End
      Begin VB.Shape ShapeBorder 
         Height          =   4095
         Left            =   60
         Top             =   60
         Width           =   7995
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000010&
         FillColor       =   &H00404040&
         Height          =   255
         Left            =   180
         Top             =   180
         Width           =   7755
      End
   End
   Begin Project1.isButton ButtonBatal 
      Cancel          =   -1  'True
      Height          =   435
      Left            =   5340
      TabIndex        =   13
      Top             =   4620
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
      Left            =   1020
      TabIndex        =   14
      Top             =   4620
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
Attribute VB_Name = "Add_Karyawan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ClassKaryawan As ClassKaryawan

Private Sub ButtonBatal_Click()
    Unload Me
End Sub

Private Sub ButtonSimpan_Click()
    If blnTambah = True Then    'data baru
        If Trim(TextKode.Text) = "" Or Trim(TextNama.Text) = "" Or Trim(TextHK.Text) = "" Or Trim(TextGaji.Text) = "" Then
            Beep
            Exit Sub
        End If

        Dim TextLen As String: TextLen = Len(TextGaji.Text)
        If TextLen > 15 Then
            Beep
            Exit Sub
        End If

        TextGaji.Alignment = 1
        TextGaji.Text = toMoney(TextGaji.Text) 'Format(TextGaji.Text, "#,##0")

        Set ClassKaryawan = New ClassKaryawan

        Sukses = ClassKaryawan.AddKaryawan(TextAuto.Text, TextID, TextNama.Text, TextAlamat.Text, TextKontak.Text, TextStatus, TextHK.Text, Format(TextGaji.Text, "###0"), _
                                           Conn)

        If Sukses Then
            FormKaryawan.TampilanAwal
            Unload Me
        Else
            MsgBox "Data Karyawan Gagal Disimpan ke Database", vbExclamation, "Peringatan"
        End If

    Else
        If Trim(TextKode.Text) = "" Or Trim(TextNama.Text) = "" Or Trim(TextHK.Text) = "" Or Trim(TextGaji.Text) = "" Then
            Beep
            Exit Sub
        End If

        Dim TextLen2 As String: TextLen2 = Len(TextGaji.Text)
        If TextLen2 > 15 Then
            Beep
            Exit Sub
        End If

        TextGaji.Alignment = 1
        TextGaji.Text = toMoney(TextGaji.Text) 'Format(TextGaji.Text, "#,##0")

        Set ClassKaryawan = New ClassKaryawan
        Sukses = ClassKaryawan.UpdateKaryawan(TextKodeLama.Text, TextID, TextNama.Text, TextAlamat.Text, TextKontak.Text, TextStatus, TextHK.Text, Format(TextGaji.Text, "###0"), _
                                              Conn)

        If Sukses Then
            FormKaryawan.TampilanAwal
            Unload Me
        Else
            MsgBox "Data Karyawan Gagal diperbaharui", vbExclamation, "Peringatan"
        End If

    End If

    Set ClassKaryawan = Nothing    '
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case Is = vbKeyEscape
        ButtonBatal_Click
    Case Is = vbKeyF8
        ButtonSimpan_Click
    End Select
End Sub

Private Sub ComboDepartemen_Click()
    Dim RSID As New ADODB.Recordset
    Set RSID = New ADODB.Recordset
    RSID.Open "SELECT * FROM departemen WHERE id='" & Left(ComboDepartemen.Text, 2) & "'", _
              Conn, adOpenDynamic, adLockOptimistic
    RSID.Requery
    With RSID
        If .EOF And .BOF Then
            MsgBox "ID TIDAK DITEMUKAN", _
                   vbOKOnly + vbCritical, "Error"
            Exit Sub
        Else
            TextKode.Text = !autokode_karyawan
            TextID.Text = !Id
        End If
    End With
    RSID.Close
    Call AutoKey
End Sub

Private Sub AutoKey()
    Dim RSAuto As New ADODB.Recordset
    RSAuto.Open ("SELECT * FROM departemen_karyawan WHERE id in(select max(id) from departemen_karyawan WHERE departemen_id=" & TextID.Text & ") order by id desc"), Conn
    RSAuto.Requery
    Dim Urut As String * 5
    Dim Hitung As Long
    With RSAuto
        If .EOF Then
            Urut = TextKode.Text
            TextAuto.Text = Urut
        Else
            Hitung = Right(!Id, 5) + 1
            Urut = Right(TextKode.Text & Hitung, 5)
        End If
        TextAuto.Text = Urut
    End With
End Sub

Private Sub Form_Load()
    Call SetIcon(Me.hWnd, "FORMICON", False)
    CenterForm Me
    Call initCombo
End Sub

Private Sub PictureBG_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case Is = vbKeyEscape
        ButtonBatal_Click
    Case Is = vbKeyF8
        ButtonSimpan_Click
    End Select
End Sub

Private Sub initCombo()
    Dim RSDepartemen As New ADODB.Recordset
    Set RSDepartemen = New ADODB.Recordset
    Dim Combo1 As String
    Combo1 = ComboList(ComboDepartemen, RSDepartemen, "departemen", "*", "WHERE id>1 ORDER BY id ASC", "id", "keterangan")
End Sub

Private Sub TextAlamat_Click()
    TextGaji.Alignment = 1
    TextGaji.Text = toMoney(TextGaji.Text) 'Format(TextGaji.Text, "#,##0")
End Sub

Private Sub TextAlamat_KeyPress(KeyAscii As Integer)
    GantiPetik KeyAscii
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TextGaji_Click()
    TextGaji.Alignment = 0
    TextGaji.Text = toMoney(TextGaji.Text) 'Format(TextGaji.Text, "###0")
End Sub

Private Sub TextGaji_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then KeyAscii = 0
    If KeyAscii = 13 Then
        TextGaji.Alignment = 1
        TextGaji.Text = toMoney(TextGaji.Text) 'Format(TextGaji.Text, "#,##0")
    End If
End Sub

Private Sub TextHK_Click()
    TextGaji.Alignment = 1
    TextGaji.Text = toMoney(TextGaji.Text) 'Format(TextGaji.Text, "#,##0")
End Sub

Private Sub TextJumlah_Click()
    TextGaji.Alignment = 1
    TextGaji.Text = toMoney(TextGaji.Text) 'Format(TextGaji.Text, "#,##0")
End Sub

Private Sub TextNama_Click()
    TextGaji.Alignment = 1
    TextGaji.Text = toMoney(TextGaji.Text) 'Format(TextGaji.Text, "#,##0")
End Sub

Private Sub TextNama_KeyPress(KeyAscii As Integer)
    GantiPetik KeyAscii
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TextStatus_Click()
    TextGaji.Alignment = 1
    TextGaji.Text = toMoney(TextGaji.Text) 'Format(TextGaji.Text, "#,##0")
End Sub

Private Sub TextStatus_KeyPress(KeyAscii As Integer)
    GantiPetik KeyAscii
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
