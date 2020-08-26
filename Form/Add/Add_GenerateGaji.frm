VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Add_GenerateGaji 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generate Data Gaji Kosong"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4950
   Icon            =   "Add_GenerateGaji.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   4950
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   120
      ScaleHeight     =   1935
      ScaleWidth      =   4695
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      Begin Project1.isButton ButtonGenerate 
         Height          =   555
         Left            =   2880
         TabIndex        =   6
         Top             =   1260
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   979
         Style           =   8
         Caption         =   "&Generate"
         iNonThemeStyle  =   0
         Object.ToolTipText     =   ""
         ToolTipTitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         RoundedBordersByTheme=   0   'False
      End
      Begin VB.TextBox TextTanggal 
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
         TabIndex        =   4
         Top             =   660
         Width           =   2115
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
         Left            =   2100
         TabIndex        =   1
         Top             =   180
         Width           =   2415
      End
      Begin MSComCtl2.DTPicker Tanggal 
         Height          =   315
         Left            =   4200
         TabIndex        =   3
         Top             =   660
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         _Version        =   393216
         Format          =   95092737
         CurrentDate     =   43008
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Penggajian"
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
         Left            =   60
         TabIndex        =   5
         Top             =   720
         Width           =   1755
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Penggajian Awal"
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
         TabIndex        =   2
         Top             =   240
         Width           =   1635
      End
   End
End
Attribute VB_Name = "Add_GenerateGaji"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ClassPenggajian As ClassPenggajian

Private Sub ButtonGenerate_Click()
    Set ClassPenggajian = New ClassPenggajian
    Sukses = ClassPenggajian.GeneratePenggajian("0", "Rp 0", "Rp 0", "Rp 0", "Rp 0", _
                                                "Rp 0", Tanggal.Value, Tanggal.Value, FormMain.LabelNama.Caption, Conn)

    If Sukses Then
        FormPenggajian.TampilanAwal
        Unload Me
    Else
        MsgBox "Data Gaji Karyawan Mungkin Telah Dibuat", vbExclamation, "Peringatan"
    End If
End Sub

Public Sub AutoNumber()
    Set RSPenggajian = New ADODB.Recordset
    RSPenggajian.Open "SELECT id FROM karyawan_gaji WHERE id in(select max(id) from karyawan_gaji)order by id desc", Conn
    RSPenggajian.Requery
    Dim Urut As String * 12
    Dim Hitung As Long
    With RSPenggajian
        If .EOF Then
            Urut = "GK" + Format(Date, "yymmdd") + "0001"
        Else
            If Mid(!Id, 3, 6) <> Format(Date, "yymmdd") Then
                Urut = "GK" + Format(Date, "yymmdd") + "0001"
            Else
                Hitung = Right$(!Id, 4) + 1
                Urut = "GK" + Format(Date, "yymmdd") + Right("0000" & Hitung, 4)
            End If
        End If
        TextKode.Text = Urut
    End With
End Sub

Private Sub Form_Load()
    CenterForm Me
End Sub
