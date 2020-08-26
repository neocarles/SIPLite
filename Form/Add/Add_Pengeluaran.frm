VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Add_Pengeluaran 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Input Data Pengeluaran"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8370
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Add_Pengeluaran.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   8370
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
      Height          =   4275
      Left            =   120
      ScaleHeight     =   4275
      ScaleWidth      =   8115
      TabIndex        =   0
      Top             =   120
      Width           =   8115
      Begin VB.TextBox TextJumlahAwal 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   6360
         TabIndex        =   23
         Top             =   480
         Visible         =   0   'False
         Width           =   1575
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
         TabIndex        =   15
         Top             =   3660
         Width           =   2475
      End
      Begin VB.ComboBox ComboKas 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1860
         TabIndex        =   14
         Text            =   "-- Silahkan Pilih Kode Kas --"
         Top             =   3180
         Width           =   2475
      End
      Begin VB.TextBox TextKeterangan 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   1860
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   2040
         Width           =   4935
      End
      Begin VB.ComboBox ComboBiaya 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1860
         TabIndex        =   9
         Text            =   "-- Silahkan Pilih Kode Biaya --"
         Top             =   1560
         Width           =   3315
      End
      Begin MSComCtl2.DTPicker Tanggal 
         Height          =   325
         Left            =   3300
         TabIndex        =   6
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
         Format          =   149422083
         CurrentDate     =   42998
      End
      Begin VB.TextBox txtKode 
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
         TabIndex        =   3
         Top             =   600
         Width           =   1755
      End
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
         Width           =   1455
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Informasi Pengeluaran"
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
         TabIndex        =   22
         Top             =   210
         Width           =   1935
      End
      Begin VB.Label LabelAkhirKas 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   6180
         TabIndex        =   21
         Top             =   3660
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.Label LabelJumlahKas 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   6180
         TabIndex        =   20
         Top             =   3180
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.Label LabelKasAkhir 
         Height          =   315
         Left            =   4440
         TabIndex        =   19
         Top             =   3660
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label LabelKas 
         Caption         =   "%Kas%"
         Height          =   315
         Left            =   180
         TabIndex        =   18
         Top             =   3180
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label LabelJumlah 
         Caption         =   "%Jumlah%"
         Height          =   315
         Left            =   4440
         TabIndex        =   17
         Top             =   3180
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label LabelBiaya 
         Caption         =   "%Biaya%"
         Height          =   315
         Left            =   5280
         TabIndex        =   16
         Top             =   1560
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Jumlah"
         Height          =   255
         Left            =   180
         TabIndex        =   13
         Top             =   3720
         Width           =   1515
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Kas"
         Height          =   255
         Left            =   180
         TabIndex        =   12
         Top             =   3240
         Width           =   1515
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Keterangan"
         Height          =   255
         Left            =   180
         TabIndex        =   10
         Top             =   2100
         Width           =   1515
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Biaya"
         Height          =   255
         Left            =   180
         TabIndex        =   8
         Top             =   1620
         Width           =   1515
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Tanggal"
         Height          =   255
         Left            =   180
         TabIndex        =   5
         Top             =   1140
         Width           =   1515
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Kode Pengeluaran"
         Height          =   255
         Left            =   180
         TabIndex        =   4
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
      TabIndex        =   1
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
      TabIndex        =   2
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
Attribute VB_Name = "Add_Pengeluaran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RSPengeluaran As New ADODB.Recordset

Dim RSPotongan As New ADODB.Recordset
Dim RSBiaya As New ADODB.Recordset
Dim RSKas As New ADODB.Recordset
Dim RSIDKas As New ADODB.Recordset
Dim RSIDBiaya As New ADODB.Recordset

Dim ClassPengeluaran As ClassPengeluaran
Dim ClassKas As ClassKas

Private Sub ButtonBatal_Click()
    Unload Me
End Sub

Private Sub ButtonSimpan_Click()
    If blnTambah = True Then    'data baru
        If Trim(TextKeterangan.Text) = "" Or Trim(TextJumlah.Text) = "" Or ComboBiaya.Text = "-- Silahkan Pilih Kode Biaya --" Or ComboKas.Text = "-- Silahkan Pilih Kode Kas --" Then
            Beep
            Exit Sub
        End If

        TextJumlah.Alignment = 1
        TextJumlah.Text = Format(TextJumlah.Text, "#,##0")

        Set ClassPengeluaran = New ClassPengeluaran
        Set ClassKas = New ClassKas
        If Val(Format(LabelJumlah.Caption, "###0")) < Val(Format(TextJumlah.Text, "###0")) Then
            MsgBox "Maaf, Saldo Kas Tidak Mencukupi", vbExclamation, "Peringatan"
            Exit Sub
        Else
            Sukses = ClassPengeluaran.AddPengeluaran(txtKode.Text, TextTanggal.Text, LabelBiaya.Caption, TextKeterangan.Text, _
                                                     LabelKas.Caption, TextJumlah.Text, FormMain.LabelNama.Caption, _
                                                     Conn)
            Sukses = ClassKas.UpdateJumlahKas(LabelKas.Caption, Val(Format(LabelKasAkhir.Caption, "###0")), Conn, ComboKas.Text)
        End If

        If Sukses Then
            FormPengeluaran.TampilanAwal
            FormKas.TampilanAwal
            Unload Me
        Else
            MsgBox "Data Pengeluaran Asal gagal disimpan", vbExclamation, "Peringatan"
        End If

    Else
        If Trim(TextKeterangan.Text) = "" Or Trim(TextJumlah.Text) = "" Or ComboBiaya.Text = "-- Silahkan Pilih Kode Biaya --" Or ComboKas.Text = "-- Silahkan Pilih Kode Kas --" Then
            Beep
            Exit Sub
        End If

        Set ClassPengeluaran = New ClassPengeluaran
        Set ClassKas = New ClassKas
        If Val(Format(LabelJumlah.Caption, "###0")) < Val(Format(TextJumlah.Text, "###0")) Then
            MsgBox "Maaf, Saldo Kas Tidak Mencukupi", vbExclamation, "Peringatan"
            Exit Sub
        Else
            LabelAkhirKas.Caption = Val(Format(LabelJumlahKas.Caption, "###0")) - Val(Format(TextJumlah.Text, "###0"))
            Sukses = ClassPengeluaran.UpdatePengeluaran(txtKode.Text, TextTanggal.Text, LabelBiaya.Caption, TextKeterangan.Text, _
                                                        LabelKas.Caption, TextJumlah.Text, FormMain.LabelNama.Caption, _
                                                        Conn)
            Sukses = ClassKas.UpdateJumlahKas(LabelKas.Caption, Val(Format(LabelAkhirKas.Caption, "###0")), Conn, ComboKas.Text)
        End If

        If Sukses Then
            FormPengeluaran.TampilanAwal
            FormKas.TampilanAwal
            Unload Me
        Else
            MsgBox "Data Pengeluaran Asal gagal disimpan", vbExclamation, "Peringatan"
        End If

    End If

    Set ClassPengeluaran = Nothing
End Sub

Private Sub ComboBiaya_Click()
    Set RSIDBiaya = New ADODB.Recordset
    RSIDBiaya.Open "SELECT * FROM biaya WHERE nama='" & ComboBiaya.Text & "'", _
                   Conn, adOpenDynamic, adLockOptimistic
    RSIDBiaya.Requery
    With RSIDBiaya
        If .EOF And .BOF Then
            MsgBox "ID TIDAK DITEMUKAN", _
                   vbOKOnly + vbCritical, "Error"
            Exit Sub
        Else
            LabelBiaya.Caption = !Kode
        End If
    End With
    RSIDBiaya.Close
End Sub

Private Sub ComboKas_Click()
    Set RSIDKas = New ADODB.Recordset
    RSIDKas.Open "SELECT * FROM kas WHERE nama='" & ComboKas.Text & "'", _
                 Conn, adOpenDynamic, adLockOptimistic
    RSIDKas.Requery
    With RSIDKas
        If .EOF And .BOF Then
            MsgBox "ID TIDAK DITEMUKAN", _
                   vbOKOnly + vbCritical, "Error"
            Exit Sub
        Else
            LabelJumlah.Caption = Format(Val(!Saldo), "Rp ###,##,0")
            LabelKas.Caption = !Kode
            LabelJumlah.Visible = True
        End If
        LabelJumlahKas.Caption = Val(Format(!Saldo, "###0")) + Val(Format(TextJumlah.Text, "###0"))
    End With
    LabelKasAkhir.Caption = Val(Format(LabelJumlah.Caption, "###0")) - Val(Format(TextJumlah.Text, "###0"))
    LabelAkhirKas.Caption = Val(Format(LabelJumlahKas.Caption, "###0")) - Val(Format(TextJumlah.Text, "###0"))
    RSIDKas.Close
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
    Call AutoNumber
    Tanggal.Value = Now
    TextTanggal.Text = Format(Tanggal.Value, "dd/mm/yyyy")

    Call initCombo
End Sub

Private Sub initCombo()
    Set RSBiaya = New ADODB.Recordset
    Set RSKas = New ADODB.Recordset
    Dim Combo1 As String: Combo1 = ComboListSingle(ComboBiaya, RSBiaya, "biaya", "*", "order by nama asc", "nama")
    Dim Combo2 As String: Combo2 = ComboListSingle(ComboKas, RSKas, "kas", "*", "order by nama asc", "nama")
End Sub

Private Sub PictureBG_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case Is = vbKeyEscape
        ButtonBatal_Click
    Case Is = vbKeyF8
        ButtonSimpan_Click
    End Select
End Sub

Private Sub AutoNumber()
    Set RSPengeluaran = New ADODB.Recordset
    RSPengeluaran.Open "SELECT * FROM pengeluaran WHERE id in(select max(id) from pengeluaran)order by id desc", Conn
    RSPengeluaran.Requery
    Dim Urut As String * 10
    Dim Hitung As Long
    With RSPengeluaran
        If .EOF Then
            Urut = "PG" + Format(Date, "ddmmyy") + "01"
        Else
            If Mid(!Id, 3, 6) <> Format(Date, "ddmmyy") Then
                Urut = "PG" + Format(Date, "ddmmyy") + "01"
            Else
                Hitung = Right$(!Id, 2) + 1
                Urut = "PG" + Format(Date, "ddmmyy") + Right("00" & Hitung, 2)
            End If
        End If
        txtKode.Text = Urut
    End With
End Sub

Private Sub TextJumlah_Click()
    TextJumlah.Alignment = 0
    TextJumlah.Text = Format(TextJumlah.Text, "###0")
End Sub

Private Sub TextJumlah_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then KeyAscii = 0
    If KeyAscii = 13 Then
        TextJumlah.Alignment = 1
        TextJumlah.Text = Format(TextJumlah.Text, "#,##0")
        LabelKasAkhir.Caption = Val(Format(LabelJumlah.Caption, "###0")) - Val(Format(TextJumlah.Text, "###0"))
        LabelAkhirKas.Caption = Val(Format(LabelJumlahKas.Caption, "###0")) - Val(Format(TextJumlah.Text, "###0"))
    End If
End Sub

Private Sub TextKeterangan_KeyPress(KeyAscii As Integer)
    GantiPetik KeyAscii
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
