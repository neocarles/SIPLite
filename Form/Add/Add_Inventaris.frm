VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Add_Inventaris 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Input Data Inventaris"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8385
   Icon            =   "Add_Inventaris.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   8385
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
      Begin VB.TextBox TextTanggal 
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
         Left            =   1860
         TabIndex        =   7
         Top             =   1080
         Width           =   1755
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
         Left            =   1860
         TabIndex        =   6
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox TextPerolehan 
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
         Height          =   975
         Left            =   1860
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   2520
         Width           =   4935
      End
      Begin VB.ComboBox ComboKategori 
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
         Height          =   360
         Left            =   1860
         TabIndex        =   3
         Text            =   "-- Silahkan Pilih Kategori --"
         Top             =   3660
         Width           =   2775
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
         TabIndex        =   2
         Top             =   1560
         Width           =   4935
      End
      Begin VB.TextBox TextJumlah 
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
         Top             =   2040
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker Tanggal 
         Height          =   330
         Left            =   3600
         TabIndex        =   5
         Top             =   1080
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/mm/yyyy"
         Format          =   149356547
         CurrentDate     =   42998
      End
      Begin VB.Shape ShapeBorder 
         Height          =   4095
         Left            =   60
         Top             =   60
         Width           =   7995
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Kode Pengeluaran"
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
         TabIndex        =   14
         Top             =   660
         Width           =   1515
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Tanggal"
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
         TabIndex        =   13
         Top             =   1140
         Width           =   1515
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Nama Barang"
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
         TabIndex        =   12
         Top             =   1620
         Width           =   1515
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Jumlah"
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
         Top             =   2100
         Width           =   1515
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Perolehan"
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
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Kategori"
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
         Top             =   3660
         Width           =   1515
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Informasi Inventaris"
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
         TabIndex        =   8
         Top             =   210
         Width           =   1755
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
      TabIndex        =   15
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
      TabIndex        =   16
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
Attribute VB_Name = "Add_Inventaris"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RSInventaris As New ADODB.Recordset
Dim RSKategori As New ADODB.Recordset

Dim ClassInventaris As ClassInventaris

Private Sub ButtonBatal_Click()
    Unload Me
End Sub

Private Sub ButtonSimpan_Click()
    If blnTambah = True Then    'data baru
        If Trim(TextNama.Text) = "" Or Trim(TextJumlah.Text) = "" Or Trim(TextPerolehan.Text) = "" Or ComboKategori.Text = "-- Silahkan Pilih Kategori --" Then
            Beep
            Exit Sub
        End If

        Set ClassInventaris = New ClassInventaris
        Sukses = ClassInventaris.AddInventaris(TextKode.Text, TextTanggal.Text, TextNama.Text, TextJumlah.Text, _
                                               TextPerolehan.Text, ComboKategori.Text, FormMain.LabelNama.Caption, _
                                               Conn)

        If Sukses Then
            FormInventaris.TampilanAwal
            Unload Me
        Else
            MsgBox "Data Inventaris Gagal Disimpan", vbExclamation, "Peringatan"
        End If

    Else
        If Trim(TextNama.Text) = "" Or Trim(TextJumlah.Text) = "" Or Trim(TextPerolehan.Text) = "" Or ComboKategori.Text = "-- Silahkan Pilih Kategori --" Then
            Beep
            Exit Sub
        End If

        Set ClassInventaris = New ClassInventaris
        Sukses = ClassInventaris.UpdateInventaris(TextKode.Text, TextTanggal.Text, TextNama.Text, TextJumlah.Text, _
                                                  TextPerolehan.Text, ComboKategori.Text, FormMain.LabelNama.Caption, _
                                                  Conn)

        If Sukses Then
            FormInventaris.TampilanAwal
            Unload Me
        Else
            MsgBox "Data Inventaris Gagal Disimpan", vbExclamation, "Peringatan"
        End If

    End If

    Set ClassInventaris = Nothing
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
    Call SetIcon(Me.hWnd, "FORMICON", False)
    CenterForm Me
    TextTanggal.Text = Format(Now, "dd/mm/yyyy")

    Call initCombo
End Sub

Private Sub initCombo()
    Set RSKategori = New ADODB.Recordset
    Dim Combo1 As String: Combo1 = ComboListSingle(ComboKategori, RSKategori, "jenis_inventaris", "nama", "order by nama asc", "nama")
End Sub

Public Sub AutoNumberINT()
    Set RSInventaris = New ADODB.Recordset
    RSInventaris.Open ("SELECT * FROM inventaris WHERE id in(select max(id) from inventaris)order by id DESC"), Conn
    RSInventaris.Requery
    Dim Urut As String * 9
    Dim Hitung As Long
    With RSInventaris
        If .EOF Then
            Urut = "INT" + "000001"
            TextKode = Urut
        Else
            Hitung = Right(!Id, 6) + 1
            Urut = "INT" + Right("000000" & Hitung, 6)
        End If
        TextKode = Urut
    End With
End Sub

Private Sub TextJumlah_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then KeyAscii = 0
End Sub

Private Sub TextNama_KeyPress(KeyAscii As Integer)
    GantiPetik KeyAscii
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TextPerolehan_KeyPress(KeyAscii As Integer)
    GantiPetik KeyAscii
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub PictureBG_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case Is = vbKeyEscape
        ButtonBatal_Click
    Case Is = vbKeyF8
        ButtonSimpan_Click
    End Select
End Sub
