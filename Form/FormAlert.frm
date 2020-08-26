VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form FormAlert 
   BorderStyle     =   0  'None
   Caption         =   "Alert"
   ClientHeight    =   5850
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11055
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5850
   ScaleWidth      =   11055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TimerAlert 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   6300
      Top             =   1200
   End
   Begin SHDocVwCtl.WebBrowser WebAlert 
      Height          =   3255
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5115
      ExtentX         =   9022
      ExtentY         =   5741
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "FormAlert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Call AutoResize
WebAlert.Navigate (App.Path & "/Resource/index.html")
End Sub

Private Sub Form_Resize()
Call AutoResize
End Sub

Sub AutoResize()
    With Me
    '    .Top
    '    .Left
    '    .Width
    '    .Height
    End With
    With WebAlert
        .Top = 0
        .Left = 0
        .Width = Me.Width
        .Height = Me.Height
    End With
End Sub

Private Sub TimerAlert_Timer()
Unload Me
TimerAlert.Enabled = False
End Sub

Private Sub WebAlert_DocumentComplete(ByVal pDisp As Object, URL As Variant)
TimerAlert.Enabled = True
End Sub
