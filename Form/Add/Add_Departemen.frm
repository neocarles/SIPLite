VERSION 5.00
Begin VB.Form Add_Departemen 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Input Data Departemen"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8355
   Icon            =   "Add_Departemen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   8355
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PictureBG 
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   120
      ScaleHeight     =   1695
      ScaleWidth      =   8115
      TabIndex        =   0
      Top             =   120
      Width           =   8115
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
         Left            =   2100
         TabIndex        =   2
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox TextDepartemen 
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
         Left            =   2100
         TabIndex        =   1
         Top             =   1140
         Width           =   5115
      End
      Begin VB.Label LabelDepartemen 
         Height          =   315
         Left            =   5280
         TabIndex        =   6
         Top             =   600
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Informasi Departemen"
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
         TabIndex        =   5
         Top             =   210
         Width           =   1935
      End
      Begin VB.Shape ShapeBorder 
         Height          =   1575
         Left            =   60
         Top             =   60
         Width           =   7995
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Kode Departemen"
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
         TabIndex        =   4
         Top             =   660
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Nama Departemen"
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
         TabIndex        =   3
         Top             =   1200
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
         Width           =   7755
      End
   End
   Begin Project1.isButton ButtonSimpan 
      Height          =   435
      Left            =   600
      TabIndex        =   7
      Top             =   2040
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
   Begin Project1.isButton ButtonBatal 
      Cancel          =   -1  'True
      Height          =   435
      Left            =   5640
      TabIndex        =   8
      Top             =   2040
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
End
Attribute VB_Name = "Add_Departemen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ClassDepartemen As ClassDepartemen
Dim RsCekKode As ADODB.Recordset

Private Sub ButtonBatal_Click()
    Unload Me
End Sub

Private Sub ButtonSimpan_Click()
    If blnTambah = True Then    'data baru
        If Trim(TextKode.Text) = "" Or Trim(TextDepartemen.Text) = "" Then
            Beep
            Exit Sub
        End If

        Set RsCekKode = New ADODB.Recordset
        RsCekKode.Open "SELECT autokode_karyawan FROM departemen WHERE autokode_karyawan='" & TextKode.Text & "'", Conn
        If Not RsCekKode.EOF Then
            MsgBox "Data dengan Nama " & Chr(34) & TextKode.Text & Chr(34) & " sudah ada pada Database", vbInformation, "Informasi"
            Exit Sub
        End If
        Set RsCekKode = Nothing

        Set ClassDepartemen = New ClassDepartemen
        Sukses = ClassDepartemen.AddDepartemen(TextKode.Text, TextDepartemen.Text, _
                                     Conn)

        If Sukses Then
            FormBiaya.TampilanAwal
            Unload Me
        Else
            MsgBox "Data Biaya gagal disimpan", vbExclamation, "Peringatan"
        End If

    Else
        If Trim(TextKode.Text) = "" Or Trim(TextDepartemen.Text) = "" Then
            Beep
            Exit Sub
        End If

        If CStr(LabelDepartemen.Caption) <> CStr(TextDepartemen.Text) Then
            Set RsCekKode = New ADODB.Recordset
            RsCekKode.Open "SELECT autokode_karyawan FROM departemen WHERE autokode_karyawan='" & TextKode.Text & "'", Conn
            If Not RsCekKode.EOF Then
                MsgBox "Data dengan Nama " & Chr(34) & TextKode.Text & Chr(34) & " sudah ada pada Database", vbInformation, "Informasi"
                Exit Sub
            End If
            Set RsCekKode = Nothing
        Else
            Set ClassDepartemen = New ClassDepartemen
            Sukses = ClassDepartemen.UpdateDepartemen(TextDepartemen.Text, LabelDepartemen.Caption, TextKode.Text, _
                                            Conn)
        End If

        If Sukses Then
            FormBiaya.TampilanAwal
            Unload Me
        Else
            MsgBox "Data Biaya gagal diperbaharui", vbExclamation, "Peringatan"
        End If

    End If

    Set ClassDepartemen = Nothing    '
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

Private Sub TextKode_KeyPress(KeyAscii As Integer)
    GantiPetik KeyAscii
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TextDepartemen_KeyPress(KeyAscii As Integer)
    GantiPetik KeyAscii
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


