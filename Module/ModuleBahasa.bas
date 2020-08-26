Attribute VB_Name = "ModuleBahasa"
Public Sub TerapkanBahasa()
Lokasi_File_Bahasa = App.Path & "\Language\" & ReadINI("Settings", "Language", Lokasi_File_Konfigurasi) & ".lng"
With FormMain
    If FormMain.Sidebar_Mini = True Then
    .LabelMenu.Caption = "N"
    Else
    .LabelMenu.Caption = ReadINI("FormMain", "MenuNavigasi", Lokasi_File_Bahasa)
    End If
    .MenuNavigasi(0).Caption = ReadINI("FormMain", "MenuNavigasi(0)", Lokasi_File_Bahasa)
    .MenuNavigasi(1).Caption = ReadINI("FormMain", "MenuNavigasi(1)", Lokasi_File_Bahasa)
    .MenuNavigasi(2).Caption = ReadINI("FormMain", "MenuNavigasi(2)", Lokasi_File_Bahasa)
    .MenuNavigasi(3).Caption = ReadINI("FormMain", "MenuNavigasi(3)", Lokasi_File_Bahasa)
    .MenuNavigasi(4).Caption = ReadINI("FormMain", "MenuNavigasi(4)", Lokasi_File_Bahasa)
    .MenuNavigasi(5).Caption = ReadINI("FormMain", "MenuNavigasi(5)", Lokasi_File_Bahasa)
    .MenuNavigasi(6).Caption = ReadINI("FormMain", "MenuNavigasi(6)", Lokasi_File_Bahasa)
    .MenuNavigasi(7).Caption = ReadINI("FormMain", "MenuNavigasi(7)", Lokasi_File_Bahasa)
    .MenuNavigasi(8).Caption = ReadINI("FormMain", "MenuNavigasi(8)", Lokasi_File_Bahasa)
    .MenuNavigasi(9).Caption = ReadINI("FormMain", "MenuNavigasi(9)", Lokasi_File_Bahasa)
    .MenuNavigasi(10).Caption = ReadINI("FormMain", "MenuNavigasi(10)", Lokasi_File_Bahasa)
    .LabelSkins.Caption = ReadINI("FormMain", "LabelSkins", Lokasi_File_Bahasa)
    .LabelStart.Caption = ReadINI("FormMain", "LabelStartup", Lokasi_File_Bahasa)
    .LabelBahasa.Caption = ReadINI("FormMain", "LabelBahasa", Lokasi_File_Bahasa)
    .LabelBackground.Caption = ReadINI("FormMain", "LabelBackground", Lokasi_File_Bahasa)
    .LabelVersion.Caption = ReadINI("FormMain", "LabelVersion", Lokasi_File_Bahasa)
    
    If VersiPro = False Then
        .CheckBeranda.Caption = ReadINI("FormMain", "CheckBeranda", Lokasi_File_Bahasa)
        .CheckBadge.Caption = ReadINI("FormMain", "CheckBadge", Lokasi_File_Bahasa)
        .ButtonAdvance.Caption = ReadINI("FormMain", "ButtonAdvance", Lokasi_File_Bahasa)
    ElseIf VersiPro = True Then
        .CheckBeranda.Caption = ReadINI("FormMain", "CheckBeranda2", Lokasi_File_Bahasa)
        .CheckBadge.Caption = ReadINI("FormMain", "CheckBadge2", Lokasi_File_Bahasa)
        .ButtonAdvance.Caption = ReadINI("FormMain", "ButtonAdvance2", Lokasi_File_Bahasa)
    End If
End With
End Sub
