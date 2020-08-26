Attribute VB_Name = "ModuleUI"
Option Explicit
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
' Load BMP in resource ( *.dll ) by Muhammad Rifqi
' http://www.indramayu-pc.com/ dan http://www.avansi.net/
'---------------------------------------------------------
' Thanks gan udah kasih SC ini.

Private Type GUID
     Data1 As Long
     Data2 As Integer
     Data3 As Integer
     Data4(7) As Byte
End Type

Private Type PicBmp
     size As Long
     Type As Long
     hBmp As Long
     hPal As Long
     Reserved As Long
End Type
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Const WS_EX_LAYERED = &H80000
Const GWL_EXSTYLE = (-20)
Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1

Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnHandle As Long, IPic As IPicture) As Long
Private Declare Function LoadBitmap Lib "user32.dll" Alias "LoadBitmapA" (ByVal hInstance As Long, ByVal lpBitmapID As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
               
Public Function LoadPictureDLL(ByVal lResourceId As Long) As Picture
Dim hInst As Long
Dim hBmp  As Long
Dim Pic As PicBmp
Dim IPic As IPicture
Dim IID_IDispatch As GUID
Dim lRC As Long
hInst = LoadLibrary(StrPtr(App.Path & "\Resource\Resource.dll"))
If hInst <> 0 Then
    hBmp = LoadBitmap(hInst, lResourceId)
    If hBmp <> 0 Then
        IID_IDispatch.Data1 = &H20400
        IID_IDispatch.Data4(0) = &HC0
        IID_IDispatch.Data4(7) = &H46
        Pic.size = Len(Pic)
        Pic.Type = vbPicTypeBitmap
        Pic.hBmp = hBmp
        Pic.hPal = 0
        lRC = OleCreatePictureIndirect(Pic, IID_IDispatch, 1, IPic)
        If lRC = 0 Then
            Set LoadPictureDLL = IPic
            Set IPic = Nothing
        Else
            DeleteObject hBmp
        End If
    End If
    FreeLibrary hInst
    hInst = 0
End If
End Function
        
Sub Theme(FontColor As String, H1 As String, H2 As String, S1 As String, ButtonNormal As String, ButtonHover As String, UserRoundBG As String, UserBG As String, LabelUser As String, Footer As String, Ba As String, Bd As String, Bn As String, Back As String, Border As String, Header As String, FontNormal As String)
On Error Resume Next
Dim i As Integer
For i = 0 To FormMain.MenuNavigasi.Count - 1
    FormMain.MenuNavigasi(i).ForeColorHover = FontColor
    FormMain.MenuNavigasi(i).BackColorHover = ButtonHover
    FormMain.LabelCopyURL.ForeColor = ButtonHover
    FormMain.PictureAPP.BackColor = H1
    FormMain.PictureHeader.BackColor = H2
    FormMain.PictureSidebar.BackColor = S1
    FormMain.PictureBgRound.BackColor = UserRoundBG
    FormMain.PictureWBG.BackColor = UserBG
    FormMain.LabelNameWork.ForeColor = LabelUser
    FormMain.LabelAciveSince.ForeColor = LabelUser
    FormMain.PictureFooter.BackColor = Footer
    FormMain.MenuNavigasi(i).BackColorActive = Ba
    FormMain.MenuNavigasi(i).BackColorDown = Bd
    FormMain.MenuNavigasi(i).BackColorNormal = Bn
    FormMain.FrameNavigasi.BackDarkColor = Back
    FormMain.FrameNavigasi.BackLightColor = Back
    FormMain.FrameNavigasi.BorderColor = Border
    FormMain.PictureLine.BackColor = Border
    FormMain.FrameNavigasi.HeaderDarkColor = Header
    FormMain.FrameNavigasi.HeaderLightColor = Header
    FormMain.MenuNavigasi(i).ForeColorNormal = FontNormal
    FormMain.MenuNavigasi(i).ForeColorDown = FontNormal
    FormMain.MenuNavigasi(i).ForeColorHover = FontNormal
    FormMain.MenuNavigasi(i).ForeColorActive = FontNormal
Next i
End Sub
