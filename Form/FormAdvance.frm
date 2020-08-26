VERSION 5.00
Begin VB.Form FormAdvance 
   Appearance      =   0  'Flat
   BorderStyle     =   0  'None
   Caption         =   "Pengaturan Lanjutan"
   ClientHeight    =   8895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13470
   Icon            =   "FormAdvance.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8895
   ScaleWidth      =   13470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Image ButtonClose 
      Appearance      =   0  'Flat
      Height          =   450
      Left            =   12900
      Picture         =   "FormAdvance.frx":000C
      Top             =   120
      Width           =   450
   End
End
Attribute VB_Name = "FormAdvance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ButtonClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call SetIcon(Me.hWnd, "FORMICON", False)
    With Me
        .Top = 0
        .Height = Screen.Height
        .Left = 0
        .Width = Screen.Width
    End With

    AutoResize
End Sub

Public Sub AutoResize()
    With ButtonClose
        .Left = FormMain.TombolKeluar.Left + 220
    End With
End Sub
