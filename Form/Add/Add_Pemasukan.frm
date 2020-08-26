VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Add_Pemasukan 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Input Data Pemasukan"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8370
   Icon            =   "Add_Pemasukan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   8370
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PictureBG 
      BorderStyle     =   0  'None
      Height          =   3795
      Left            =   120
      ScaleHeight     =   3795
      ScaleWidth      =   8115
      TabIndex        =   0
      Top             =   120
      Width           =   8115
      Begin VB.TextBox TextTanggal 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1860
         TabIndex        =   6
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox TextKode 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1860
         TabIndex        =   5
         Top             =   600
         Width           =   1755
      End
      Begin VB.TextBox TextKeterangan 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   1860
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   1560
         Width           =   4935
      End
      Begin VB.ComboBox ComboKas 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1860
         TabIndex        =   2
         Text            =   "-- Silahkan Pilih Kode Kas --"
         Top             =   2700
         Width           =   2715
      End
      Begin VB.TextBox TextJumlah 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1860
         TabIndex        =   1
         Top             =   3180
         Width           =   2715
      End
      Begin MSComCtl2.DTPicker Tanggal 
         Height          =   325
         Left            =   3300
         TabIndex        =   4
         Top             =   1080
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   582
         _Version        =   393216
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
         Format          =   149291011
         CurrentDate     =   42998
      End
      Begin VB.Label LabelKasTukar 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   4980
         TabIndex        =   22
         Top             =   900
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label LabelSaldoAwal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   4980
         TabIndex        =   21
         Top             =   480
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label LabelKodeAwal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   6660
         TabIndex        =   20
         Top             =   480
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label LabelKasUpdate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   4680
         TabIndex        =   19
         Top             =   3180
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label LabelJumlahKas 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   6360
         TabIndex        =   18
         Top             =   2700
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.Label LabelAkhirKas 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   6360
         TabIndex        =   17
         Top             =   3180
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.Shape ShapeBorder 
         Height          =   3675
         Left            =   60
         Top             =   60
         Width           =   7995
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Kode Pemasukan"
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
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Keterangan"
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
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Kas"
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
         Top             =   2760
         Width           =   1515
      End
      Begin VB.Label Label7 
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
         TabIndex        =   10
         Top             =   3240
         Width           =   1515
      End
      Begin VB.Label LabelKode 
         Caption         =   "%KodeKas%"
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
         Left            =   4680
         TabIndex        =   9
         Top             =   2700
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label LabelKasAkhir 
         Height          =   315
         Left            =   4440
         TabIndex        =   8
         Top             =   3660
         Width           =   1815
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Informasi Pemasukan"
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
         TabIndex        =   7
         Top             =   210
         Width           =   1860
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
      Left            =   5400
      TabIndex        =   15
      Top             =   4140
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
      Left            =   1080
      TabIndex        =   16
      Top             =   4140
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
Attribute VB_Name = "Add_Pemasukan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RSIDKas As New ADODB.Recordset
Dim RSKas As New ADODB.Recordset

Dim ClassPemasukan As ClassPemasukan
Dim ClassKas As ClassKas

Private Sub ButtonBatal_Click()
    Unload Me
End Sub

Private Sub ButtonSimpan_Click()
    If blnTambah = True Then    'data baru
        If Trim(TextKeterangan.Text) = "" Or Trim(TextJumlah.Text) = "" Or TextTanggal.Text = "" Or ComboKas.Text = "-- Silahkan Pilih Kode Kas --" Then
            Beep
            Exit Sub
        End If

        TextJumlah.Alignment = 1
        TextJumlah.Text = Format(TextJumlah.Text, "#,##0")

        Set ClassPemasukan = New ClassPemasukan
        Set ClassKas = New ClassKas
        Sukses = ClassPemasukan.AddPemasukan(TextKode.Text, TextTanggal.Text, TextKeterangan.Text, _
                                             LabelKode.Caption, TextJumlah.Text, FormMain.LabelNama.Caption, _
                                             Conn)
        Sukses = ClassKas.UpdateKasPemasukan(LabelKode.Caption, Val(Format(TextJumlah.Text, "###0")), Conn, ComboKas.Text)

        If Sukses Then
            FormPemasukan.TampilanAwal
            FormKas.TampilanAwal
            Unload Me
        Else
            MsgBox "Data Pemasukan Gagal Disimpan", vbExclamation, "Peringatan"
        End If

    Else
        If Trim(TextKeterangan.Text) = "" Or Trim(TextJumlah.Text) = "" Or TextTanggal.Text = "" Or ComboKas.Text = "-- Silahkan Pilih Kode Kas --" Then
            Beep
            Exit Sub
        End If

        Set ClassPemasukan = New ClassPemasukan
        Set ClassKas = New ClassKas
        LabelAkhirKas.Caption = Val(Format(LabelJumlahKas.Caption, "###0")) - Val(Format(TextJumlah.Text, "###0"))
        LabelKasUpdate.Caption = Val(Format(LabelJumlahKas.Caption, "###0")) + Val(Format(TextJumlah.Text, "###0"))
        LabelKasTukar.Caption = Val(Format(LabelKasUpdate.Caption, "###0")) + Val(Format(TextJumlah.Text, "###0"))
        Sukses = ClassPemasukan.UpdatePemasukan(TextKode.Text, TextTanggal.Text, TextKeterangan.Text, _
                                                LabelKode.Caption, TextJumlah.Text, FormMain.LabelNama.Caption, _
                                                Conn)
        If LabelKode.Caption = LabelKodeAwal.Caption Then
            Sukses = ClassKas.UpdateJumlahKas(LabelKode.Caption, Val(Format(LabelKasUpdate.Caption, "###0")), Conn)
        Else
            Sukses = ClassKas.UpdateJumlahKas(LabelKode.Caption, Val(Format(LabelKasTukar.Caption, "###0")), Conn)
            Sukses = ClassKas.UpdateJumlahKas(LabelKodeAwal.Caption, Val(Format(LabelSaldoAwal.Caption, "###0")), Conn)
        End If

        If Sukses Then
            FormPemasukan.TampilanAwal
            FormKas.TampilanAwal
            Unload Me
        Else
            MsgBox "Data Pemasukan Gagal Diperbaharui", vbExclamation, "Peringatan"
        End If

    End If

    Set ClassPemasukan = Nothing
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
            LabelKode.Caption = !Kode
        End If
        LabelJumlahKas.Caption = Val(Format(!Saldo, "###0")) - Val(Format(TextJumlah.Text, "###0"))
    End With
    LabelAkhirKas.Caption = Val(Format(LabelJumlahKas.Caption, "###0")) - Val(Format(TextJumlah.Text, "###0"))
    LabelKasUpdate.Caption = Val(Format(LabelJumlahKas.Caption, "###0")) + Val(Format(TextJumlah.Text, "###0"))
    LabelKasTukar.Caption = Val(Format(LabelKasUpdate.Caption, "###0")) + Val(Format(TextJumlah.Text, "###0"))
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

    Call initCombo
End Sub

Private Sub initCombo()
    Set RSKas = New ADODB.Recordset
    Dim Combo1 As String: Combo1 = ComboListSingle(ComboKas, RSKas, "kas", "*", "order by nama asc", "nama")
End Sub

Private Sub PictureBG_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case Is = vbKeyEscape
        ButtonBatal_Click
    Case Is = vbKeyF8
        ButtonSimpan_Click
    End Select
End Sub

Public Sub AutoNumber()
    Set RSPengeluaran = New ADODB.Recordset
    RSPengeluaran.Open "SELECT * FROM pemasukan WHERE kode in(select max(kode) from pemasukan)order by kode desc", Conn
    RSPengeluaran.Requery
    Dim Urut As String * 10
    Dim Hitung As Long
    With RSPengeluaran
        If .EOF Then
            Urut = "PM" + Format(Date, "ddmmyy") + "01"
        Else
            If Mid(!Kode, 3, 6) <> Format(Date, "ddmmyy") Then
                Urut = "PM" + Format(Date, "ddmmyy") + "01"
            Else
                Hitung = Right$(!Kode, 2) + 1
                Urut = "PM" + Format(Date, "ddmmyy") + Right("00" & Hitung, 2)
            End If
        End If
        TextKode.Text = Urut
    End With
End Sub

Private Sub Tanggal_Change()
    TextTanggal.Text = Format(Tanggal.Value, "dd/mm/yyyy")
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
        LabelAkhirKas.Caption = Val(Format(LabelJumlahKas.Caption, "###0")) - Val(Format(TextJumlah.Text, "###0"))
        LabelKasUpdate.Caption = Val(Format(LabelJumlahKas.Caption, "###0")) + Val(Format(TextJumlah.Text, "###0"))
        LabelKasTukar.Caption = Val(Format(LabelKasUpdate.Caption, "###0")) + Val(Format(TextJumlah.Text, "###0"))
    End If
End Sub

Private Sub TextKeterangan_KeyPress(KeyAscii As Integer)
    GantiPetik KeyAscii
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
