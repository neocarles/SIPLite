Attribute VB_Name = "ModuleForm"
Option Explicit
'Api untuk memindahkan form
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetCapture Lib "user32" () As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long

'Posisi X dan Y
Global iTPPY As Long
Global iTPPX As Long

'Variabel untuk memindahkan form
Dim bMoveFrom As Boolean, LastPoint As POINTAPI

Public Type POINTAPI
    x As Long
    y As Long
End Type

'API para o procedimento alway's on top
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'Para ver o relat´rio através do browser
Public Const SW_SHOWNORMAL = 1
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOW = 5
Public Const SW_MINIMIZE = 6
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNA = 8
Public Const SW_RESTORE = 9
Public Const SW_SHOWDEFAULT = 10
'Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Variável das msgboxs
Public Respon As String
'Public Localizacao_Ficheiro_Preferencias As String
'Public Localizacao_Ficheiro_Lingua As String

'Memungkinkan Form Tampil di Atas PictureBox
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndParent As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

'Shell Function For Open Webpage
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
                                     (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
                                      ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd _
                                                                                                 As Long) As Long

'Perpindahan Form Tanpa Border
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'Public Declare Sub ReleaseCapture Lib "User32" ()
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2

Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Public Sub Form_Diatas(Nama_Form As Form, Nama_PictureBox As PictureBox)
'Memungkinkan Meletakkan Form Di Dalam PictureBox
    On Error GoTo ERROR_CODE
    Dim ret As Long
    Call SetParent(Nama_Form.hwnd, Nama_PictureBox.hwnd)
    Nama_Form.Show

    Exit Sub
ERROR_CODE:
End Sub

'Memposisikan Form Pada Tengah Layar
Public Sub CenterForm(ByRef srcForm As Form)
    srcForm.Left = (Screen.Width - srcForm.Width) \ 2
    srcForm.Top = (Screen.Height - srcForm.Height) \ 2
End Sub

Public Sub Move_Form(Form As Form)
'Prosedur untuk memindahkan form
    If Form.WindowState = 0 Then
        Dim iDX As Long, iDY As Long
        Dim POINT As POINTAPI
        If Not bMoveFrom Then Exit Sub
        GetCursorPos POINT
        iDX& = (POINT.x - LastPoint.x) * iTPPX&
        iDY& = (POINT.y - LastPoint.y) * iTPPY&
        LastPoint.x = POINT.x
        LastPoint.y = POINT.y
        Form.Move Form.Left + iDX&, Form.Top + iDY&
    End If
End Sub

Public Sub Capture_Posisi_Form(Form As Form)
'Menangkap posisi x dan y
    Dim POINT As POINTAPI
    GetCursorPos POINT
    LastPoint.x = POINT.x
    LastPoint.y = POINT.y
    bMoveFrom = True
End Sub

Public Sub Large_Form(Form As Form)
'Letakkan Form Untuk Posisi Akhir
    bMoveFrom = False
End Sub

Public Sub Pesan_Peringatan(Peringatan As String, pesan As String, Title As String)
'Procedimento para mostrar uma mensagem de aviso
    With FormMessage    'FormMessageBox
        If Peringatan = "Information" Then
            .PictureMessage.Picture = FormUI.Icon_Info.Picture
            .TombolOk.Visible = True
            Beep
        ElseIf Peringatan = "Error" Then
            .PictureMessage.Picture = FormUI.Icon_Error.Picture
            .TombolOk.Visible = True

        ElseIf Peringatan = "Question" Then
            .PictureMessage.Picture = FormUI.Icon_Quest.Picture
            .TombolIya.Visible = True
            .TombolNo.Visible = True
            Beep
        End If

        .LabelMessage.Caption = pesan
        .LabelTitle.Caption = Title
        .Show vbModal
    End With
End Sub

Public Sub SesuaikanForm(Form As Form, Icon_Visible As Boolean, Form_Adjustable As Boolean, Frame_Center_Visible As Boolean, _
                         Frame_Button_Visible As Boolean)
'Procedimento para ajustar os componentes dos formulários
    If Form.WindowState = 1 Then Exit Sub
    With Form.ShapeControl
        .Height = Form.Height
        .Top = 0
        .Width = Form.Width
        .Left = 0
    End With

    With Form.Header
        .Height = 26
        .Top = 0    ' 1
        .Width = Form.ScaleWidth    '- 2
        .Left = 0    ' 1
    End With

    With Form.LabelTitle
        .Top = (Form.Header.ScaleHeight - .Height) / 2
        If Icon_Visible = False Then .Left = 10 Else: .Left = 26
    End With

    'Botões do controlbox
    Dim Adjust_Buttons As String
    Adjust_Buttons = "False"    'ReadINI("Dimensions", "Adjust_Button_ControlBox", Localizacao_Ficheiro_Skin)

    With Form.TombolKeluar
        If Adjust_Buttons = "False" Then
            .Top = 0
        Else
            .Top = 0
        End If
        .Left = Form.Header.Width - .Width - 6
    End With

    If Form_Adjustable = True Then
        With Form.TombolMaximize
            .Top = Form.TombolKeluar.Top
            If Adjust_Buttons = "False" Then
                .Left = Form.TombolKeluar.Left - .Width - 8
            Else
                .Left = Form.TombolKeluar.Left - .Width
            End If
        End With

        With Form.TombolRestore
            .Top = Form.TombolKeluar.Top
            .Left = Form.TombolMaximize.Left
        End With

        With Form.TombolMinimize
            .Top = Form.TombolKeluar.Top
            If Adjust_Buttons = "False" Then
                .Left = Form.TombolMinimize.Left - .Width - 8
            Else
                .Left = Form.TombolMaximize.Left - .Width
            End If
        End With
    End If

    If Frame_Button_Visible = True Then
        With Form.FrameTombol
            .Height = 41
            .Top = Form.ScaleHeight - .ScaleHeight - 2
            .Width = Form.ScaleWidth - 2
            .Left = 1
        End With
    End If

    If Frame_Center_Visible = True Then
        With Form.FrameCenter
            .Height = Form.ScaleHeight - Form.Header.ScaleHeight - Form.FrameTombol.ScaleHeight - 2
            .Top = Form.Header.Top + Form.Header.ScaleHeight
            .Width = Form.ScaleWidth - 20
            .Left = 10
        End With
    End If
End Sub




