' ======================================================================
' 1. AYARLARI KAYDET & TEMA MOTORU (DARK MODE BEYAZLIK VE KAYMA KORUMALI)
' ======================================================================
Private Sub btnAyarlariKaydet_Click()
    Dim wsAyar As Worksheet
    Dim ctrl As Control
    Dim arkaPlanRengi As Long, yaziRengi As Long
    Dim kutuArkaRengi As Long, cerceveRengi As Long
    Dim secilenTema As String

    ' --- BUTON ÇİVİLEME (DPI Ölçeklendirme Kayma Koruması) ---
    On Error Resume Next
    btnAyarlariKaydet.Top = 210: btnAyarlariKaydet.Left = 380
    btnDegistir.Top = 210: btnDegistir.Left = 380
    btnYedekAl.Top = 210: btnYedekAl.Left = 260
    On Error GoTo 0

    On Error Resume Next
    Set wsAyar = ThisWorkbook.Sheets("Ayarlar")
    On Error GoTo 0

    If wsAyar Is Nothing Then
        MsgBox "Lütfen Excel'de 'Ayarlar' adında bir sayfa oluşturun!", vbCritical, "Veritabanı Eksik"
        Exit Sub
    End If

    ' --- VERİLERİ EXCEL'E KALICI YAZMA (Çift İsim Desteğiyle) ---
    On Error Resume Next
    wsAyar.Cells(1, 1).Value = "Kullanıcı Adı"
    If txtAyarKullanici.Text <> "" Then wsAyar.Cells(2, 1).Value = txtAyarKullanici.Text
    If txtYeniKullaniciAdi.Text <> "" Then wsAyar.Cells(2, 1).Value = txtYeniKullaniciAdi.Text

    wsAyar.Cells(1, 2).Value = "Şifre"
    If txtAyarSifre.Text <> "" Then wsAyar.Cells(2, 2).Value = txtAyarSifre.Text
    If txtYeniSifre.Text <> "" Then wsAyar.Cells(2, 2).Value = txtYeniSifre.Text

    wsAyar.Cells(1, 3).Value = "Gizli Cevap"
    If txtAyarGuvenlik.Text <> "" Then wsAyar.Cells(2, 3).Value = txtAyarGuvenlik.Text
    
    wsAyar.Cells(1, 4).Value = "Tema"
    If cmbTema.Value <> "" Then wsAyar.Cells(2, 4).Value = cmbTema.Value
    If cmbAyarTema.Value <> "" Then wsAyar.Cells(2, 4).Value = cmbAyarTema.Value
    
    wsAyar.Cells(1, 5).Value = "Font"
    If cmbFont.Value <> "" Then wsAyar.Cells(2, 5).Value = cmbFont.Value
    If cmbAyarFont.Value <> "" Then wsAyar.Cells(2, 5).Value = cmbAyarFont.Value
    On Error GoTo 0

    ' --- DİNAMİK TEMA MOTORU ---
    On Error Resume Next
    secilenTema = wsAyar.Cells(2, 4).Value
    On Error GoTo 0
    
    If secilenTema = "Dark Mode (Koyu)" Or secilenTema = "Koyu Tema (Dark Mode)" Then
        arkaPlanRengi = RGB(33, 37, 41)   ' Form Ana Arka Planı
        yaziRengi = RGB(240, 240, 240)    ' Genel Yazı Rengi
        kutuArkaRengi = RGB(33, 37, 41)   ' Kutuların İçindeki Beyazlığı Kaldıran Renk
        cerceveRengi = RGB(44, 47, 51)    ' Çerçeve (Frame) Rengi
    Else
        arkaPlanRengi = RGB(245, 245, 245): yaziRengi = RGB(30, 30, 30)
        kutuArkaRengi = RGB(255, 255, 255): cerceveRengi = RGB(220, 220, 220)
    End If

    ' --- ARAYÜZÜ BOYAMA MOTORU ---
    Me.BackColor = arkaPlanRengi
    For Each ctrl In Me.Controls
        On Error Resume Next
        If TypeName(ctrl) <> "CommandButton" Then
            If wsAyar.Cells(2, 5).Value <> "" Then ctrl.Font.Name = wsAyar.Cells(2, 5).Value
        End If

        Select Case TypeName(ctrl)
            Case "Frame"
                ctrl.BackColor = cerceveRengi: ctrl.ForeColor = yaziRengi
            Case "Label", "CheckBox", "OptionButton"
                ctrl.BackStyle = 0 ' Şeffaf Arka Plan
                ctrl.ForeColor = yaziRengi
            Case "TextBox", "ComboBox", "ListBox"
                ctrl.BackColor = kutuArkaRengi: ctrl.ForeColor = yaziRengi
                ctrl.BorderStyle = 1
                ctrl.BorderColor = RGB(100, 100, 100)
            Case "CommandButton"
                ctrl.BackColor = RGB(52, 152, 219): ctrl.ForeColor = RGB(255, 255, 255)
        End Select
        On Error GoTo 0
    Next ctrl
End Sub

Private Sub btnDegistir_Click()
    Call btnAyarlariKaydet_Click
End Sub

Private Sub cmbTema_Change()
    Call btnAyarlariKaydet_Click
End Sub

Private Sub cmbAyarTema_Change()
    Call btnAyarlariKaydet_Click
End Sub

' ======================================================================
' 2. VERİTABANI TAM YEDEKLEME (TARİHLİ BACKUP SİSTEMİ)
' ======================================================================
Private Sub btnYedekAl_Click()
    Dim yedekKlasor As String, yedekDosyaAdi As String
    
    If MsgBox("Mevcut veritabanının ve sistemin tam yedeği alınsın mı?", vbQuestion + vbYesNo, "Sistem Yedekleme") = vbNo Then
        Exit Sub
    End If
    
    yedekKlasor = ThisWorkbook.Path & "\Yedekler\"
    If Len(Dir(yedekKlasor, vbDirectory)) = 0 Then
        MkDir yedekKlasor
    End If
    
    yedekDosyaAdi = yedekKlasor & "IT_Sistem_Yedek_" & Format(Now, "dd_mm_yyyy_hh_mm_ss") & ".xlsm"
    
    On Error Resume Next
    ThisWorkbook.SaveCopyAs yedekDosyaAdi
    On Error GoTo 0
    
    Call LogYaz("Sistem yedeklemesi çalıştırıldı. Yedek: " & Dir(yedekDosyaAdi))
    MsgBox "Tam veritabanı yedeklemesi başarıyla tamamlandı!" & vbCrLf & vbCrLf & "Konum: " & yedekDosyaAdi, vbInformation, "Yedek Alındı"
End Sub

' ======================================================================
' 3. KALICI LOG KAYIT MOTORU VE TEXTBOX'A BASMA
' ======================================================================
Sub LogYaz(IslemDetayi As String)
    Dim wsAyar As Worksheet
    Dim sonSatir As Long
    
    On Error Resume Next
    Set wsAyar = ThisWorkbook.Sheets("Ayarlar")
    On Error GoTo 0
    
    If wsAyar Is Nothing Then Exit Sub
    
    sonSatir = wsAyar.Cells(Rows.Count, "G").End(xlUp).Row + 1
    If sonSatir < 2 Then
        wsAyar.Cells(1, "G").Value = "ZAMAN"
        wsAyar.Cells(1, "H").Value = "İŞLEM"
        sonSatir = 2
    End If
    
    wsAyar.Cells(sonSatir, "G").Value = Now
    wsAyar.Cells(sonSatir, "H").Value = IslemDetayi
    
    Call LoglariKutuyaDoldur
End Sub

Sub LoglariKutuyaDoldur()
    Dim wsAyar As Worksheet
    Dim i As Long, sonSatir As Long
    Dim tumLoglar As String
    
    On Error Resume Next
    Set wsAyar = ThisWorkbook.Sheets("Ayarlar")
    sonSatir = wsAyar.Cells(Rows.Count, "G").End(xlUp).Row
    On Error GoTo 0
    
    If wsAyar Is Nothing Or sonSatir < 2 Then Exit Sub
    tumLoglar = ""
    
    For i = sonSatir To 2 Step -1
        If wsAyar.Cells(i, "G").Value <> "" Then
            tumLoglar = tumLoglar & wsAyar.Cells(i, "G").Value & " -> " & wsAyar.Cells(i, "H").Value & vbCrLf
        End If
    Next i
    
    On Error Resume Next
    txtSonGirisler.Text = tumLoglar
    txtLogPaneli.Text = tumLoglar
    On Error GoTo 0
End Sub

' ======================================================================
' 4. GİRİŞ (LOGIN) SİSTEMİ
' ======================================================================
Private Sub btnGiris_Click()
    Dim wsAyar As Worksheet
    Dim kAyar As String, sAyar As String
    
    On Error Resume Next
    Set wsAyar = ThisWorkbook.Sheets("Ayarlar")
    If Not wsAyar Is Nothing Then
        kAyar = wsAyar.Cells(2, 1).Value
        sAyar = wsAyar.Cells(2, 2).Value
    End If
    On Error GoTo 0
    
    If kAyar = "" Then kAyar = "admin": sAyar = "1234"
    
    If txtKullanici.Text = kAyar And txtSifre.Text = sAyar Then
        MsgBox "Giriş Başarılı! Otomasyon paneli yükleniyor.", vbInformation, "Hoş Geldiniz"
        Call LogYaz("Sisteme başarılı admin girişi yapıldı.")
        Call KritikStokKontrol
        
        ' Giriş başarılı olunca sekmeleri aç ve ana panele yönlendir
        MultiPage1.Style = fmTabStyleTabs
        MultiPage1.Pages(0).Visible = False
        MultiPage1.Value = 1
        Me.Width = 530: Me.Height = 360
    Else
        MsgBox "Kullanıcı adı veya şifre hatalı!", vbCritical, "Erişim Engellendi"
        Call LogYaz("Sisteme hatalı giriş denemesi! Kullanıcı: " & txtKullanici.Text)
        txtSifre.Text = "": txtSifre.SetFocus
    End If
End Sub

' ======================================================================
' 5. PERSONEL / ZİMMET KAYIT & SORGULAMA MOTORU
' ======================================================================
Private Sub btnDosyaSec_Click()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Lütfen yüklenecek fotoğrafı seçin"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Resim Dosyaları", "*.jpg; *.jpeg; *.png"
        If .Show = -1 Then
            txtDosyaYolu.Text = .SelectedItems(1)
            MsgBox "Fotoğraf seçildi.", vbInformation, "Bilgi"
        End If
    End With
    Set fd = Nothing
End Sub

Private Sub btnKayitOlustur_Click()
    Dim wsCalisan As Worksheet, wsTeslim As Worksheet, wsIade As Worksheet
    Dim sonSatir As Long, i As Integer
    Dim secilenYanCihazlar As String
    
    On Error Resume Next
    Set wsCalisan = ThisWorkbook.Sheets("Calisanlar")
    Set wsTeslim = ThisWorkbook.Sheets("Teslimat")
    Set wsIade = ThisWorkbook.Sheets("Geri_Alim")
    On Error GoTo 0
    
    If txtId.Text <> "" And txtAd.Text <> "" Then
        sonSatir = wsCalisan.Cells(Rows.Count, 1).End(xlUp).Row + 1
        wsCalisan.Cells(sonSatir, 1).Value = txtId.Text
        wsCalisan.Cells(sonSatir, 2).Value = txtAd.Text
        wsCalisan.Cells(sonSatir, 3).Value = txtSoyad.Text
        wsCalisan.Cells(sonSatir, 4).Value = txtTelNo.Text
        wsCalisan.Cells(sonSatir, 5).Value = txtBirim.Text
        wsCalisan.Cells(sonSatir, 6).Value = txtUnvan.Text
    End If
    
    If txtTeslimId.Text <> "" Then
        sonSatir = wsTeslim.Cells(Rows.Count, 1).End(xlUp).Row + 1
        secilenYanCihazlar = ""
        For i = 0 To lstTeslimYanCihaz.ListCount - 1
            If lstTeslimYanCihaz.Selected(i) = True Then secilenYanCihazlar = secilenYanCihazlar & lstTeslimYanCihaz.List(i) & ", "
        Next i
        If Len(secilenYanCihazlar) > 0 Then secilenYanCihazlar = Left(secilenYanCihazlar, Len(secilenYanCihazlar) - 2) Else secilenYanCihazlar = "Yok"
        
        wsTeslim.Cells(sonSatir, 1).Value = txtTeslimId.Text
        wsTeslim.Cells(sonSatir, 2).Value = txtId.Text
        wsTeslim.Cells(sonSatir, 3).Value = txtCihazTuru.Text
        wsTeslim.Cells(sonSatir, 4).Value = txtTeslimMarka.Text
        wsTeslim.Cells(sonSatir, 5).Value = txtTeslimSeriNo.Text
        wsTeslim.Cells(sonSatir, 6).Value = txtTeslimTarihi.Text
        wsTeslim.Cells(sonSatir, 7).Value = txtTeslimEden.Text
        wsTeslim.Cells(sonSatir, 8).Value = secilenYanCihazlar
    End If
    
    If txtGeriAlimId.Text <> "" Then
        sonSatir = wsIade.Cells(Rows.Count, 1).End(xlUp).Row + 1
        secilenYanCihazlar = ""
        For i = 0 To lstGeriYanCihaz.ListCount - 1
            If lstGeriYanCihaz.Selected(i) = True Then secilenYanCihazlar = secilenYanCihazlar & lstGeriYanCihaz.List(i) & ", "
        Next i
        If Len(secilenYanCihazlar) > 0 Then secilenYanCihazlar = Left(secilenYanCihazlar, Len(secilenYanCihazlar) - 2) Else secilenYanCihazlar = "Yok"
        
        wsIade.Cells(sonSatir, 1).Value = txtGeriAlimId.Text
        wsIade.Cells(sonSatir, 2).Value = txtId.Text
        wsIade.Cells(sonSatir, 3).Value = cmbCihazDurumu.Value
        wsIade.Cells(sonSatir, 4).Value = txtGeriMarka.Text
        wsIade.Cells(sonSatir, 5).Value = txtGeriSeriNo.Text
        wsIade.Cells(sonSatir, 6).Value = txtGeriAlimTarihi.Text
        wsIade.Cells(sonSatir, 7).Value = txtTeknikNot.Text
        wsIade.Cells(sonSatir, 8).Value = secilenYanCihazlar
    End If
    
    If txtDosyaYolu.Text <> "" And txtId.Text <> "" Then
        Dim fso As Object, klasorYolu As String
        Set fso = CreateObject("Scripting.FileSystemObject")
        klasorYolu = ThisWorkbook.Path & "\Fotograflar\"
        If Not fso.FolderExists(klasorYolu) Then fso.CreateFolder (klasorYolu)
        On Error Resume Next
        fso.CopyFile Source:=txtDosyaYolu.Text, Destination:=klasorYolu & txtId.Text & ".jpg", OverWriteFiles:=True
        On Error GoTo 0
        txtDosyaYolu.Text = ""
    End If
    
    MsgBox "İşlem başarıyla kaydedildi!", vbInformation
End Sub

Private Sub Sorgubtn_Click()
    Dim wsC As Worksheet, wsT As Worksheet, wsI As Worksheet
    Dim tc As String, i As Long, son As Long, bul As Boolean
    
    On Error Resume Next
    Set wsC = ThisWorkbook.Sheets("Calisanlar")
    Set wsT = ThisWorkbook.Sheets("Teslimat")
    Set wsI = ThisWorkbook.Sheets("Geri_Alim")
    On Error GoTo 0
    
    tc = Trim(SorguTc.Text): bul = False
    If tc = "" Then Exit Sub
    
    On Error Resume Next
    SorguAd.Value = "": SorguSoyad.Value = "": SorguBirim.Value = "": SorguÜnvan.Value = "": SorguTelNo.Value = ""
    SorguTarih.Value = "": SoruTeslim.Value = "": SorguCihaz.Value = "": SorguMarka.Value = "": SorguSerino.Value = ""
    SorguYancihaz.Value = "": SorguTekniknot.Value = "": SorguTeslimGeri.Value = "": SorguDurumcmb.Value = ""
    On Error GoTo 0

    If Not wsC Is Nothing Then
        son = wsC.Cells(Rows.Count, 1).End(xlUp).Row
        For i = 2 To son
            If CStr(wsC.Cells(i, 1).Value) = tc Then
                On Error Resume Next
                SorguAd.Value = wsC.Cells(i, 2).Value: SorguSoyad.Value = wsC.Cells(i, 3).Value
                SorguTelNo.Value = wsC.Cells(i, 4).Value: SorguBirim.Value = wsC.Cells(i, 5).Value
                SorguÜnvan.Value = wsC.Cells(i, 6).Value
                On Error GoTo 0
                bul = True: Exit For
            End If
        Next i
    End If

    If Not wsT Is Nothing Then
        son = wsT.Cells(Rows.Count, 2).End(xlUp).Row
        For i = son To 2 Step -1
            If CStr(wsT.Cells(i, 2).Value) = tc Then
                On Error Resume Next
                SorguTarih.Value = wsT.Cells(i, 6).Value: SoruTeslim.Value = wsT.Cells(i, 7).Value
                SorguCihaz.Value = wsT.Cells(i, 3).Value: SorguMarka.Value = wsT.Cells(i, 4).Value
                SorguSerino.Value = wsT.Cells(i, 5).Value: SorguYancihaz.Value = wsT.Cells(i, 8).Value
                SorguTeslimGeri.Value = "Teslim Edildi"
                On Error GoTo 0
                bul = True: Exit For
            End If
        Next i
    End If

    If Not wsI Is Nothing Then
        son = wsI.Cells(Rows.Count, 2).End(xlUp).Row
        For i = son To 2 Step -1
            If CStr(wsI.Cells(i, 2).Value) = tc Then
                On Error Resume Next
                SorguTeslimGeri.Value = "Geri Alındı": SorguTekniknot.Value = wsI.Cells(i, 7).Value
                SorguDurumcmb.Value = wsI.Cells(i, 3).Value: SorguTarih.Value = wsI.Cells(i, 6).Value
                SorguYancihaz.Value = wsI.Cells(i, 8).Value
                On Error GoTo 0
                Exit For
            End If
        Next i
    End If

    Dim yol As String: yol = ThisWorkbook.Path & "\Fotograflar\" & tc & ".jpg"
    On Error Resume Next
    If Dir(yol) <> "" Then
        imgPersonel.PictureSizeMode = fmPictureSizeModeStretch
        imgPersonel.Picture = LoadPicture(yol)
    Else
        Set imgPersonel.Picture = Nothing
    End If
    On Error GoTo 0
    
    If Not bul Then MsgBox "Aranan kriterlere uygun kayıt bulunamadı!", vbInformation, "Sonuç Yok"
End Sub

Private Sub txtId_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Dim fotoYolu As String
    If Trim(txtId.Text) = "" Then
        Set imgPersonel.Picture = Nothing: Exit Sub
    End If
    fotoYolu = ThisWorkbook.Path & "\Fotograflar\" & txtId.Text & ".jpg"
    On Error Resume Next
    If Dir(fotoYolu) <> "" Then imgPersonel.Picture = LoadPicture(fotoYolu) Else Set imgPersonel.Picture = Nothing
    On Error GoTo 0
End Sub

' ======================================================================
' 6. ARIZA TAKİP VE KAYIT MODÜLÜ
' ======================================================================
Private Sub txtArizaTc_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Dim wsCalisan As Worksheet, wsTeslim As Worksheet
    Dim arananTC As String, i As Long, sonSatir As Long, calisanBulundu As Boolean
    
    arananTC = Trim(txtArizaTc.Text)
    If arananTC = "" Then Exit Sub
    
    On Error Resume Next
    Set wsCalisan = ThisWorkbook.Sheets("Calisanlar")
    Set wsTeslim = ThisWorkbook.Sheets("Teslimat")
    txtArizaAdSoyad.Text = "": txtArizaDepartman.Text = "": txtArizaTel.Text = ""
    txtArizaCihazTuru.Text = "": txtArizaCihazSeri.Text = ""
    On Error GoTo 0
    
    calisanBulundu = False
    If Not wsCalisan Is Nothing Then
        sonSatir = wsCalisan.Cells(Rows.Count, "A").End(xlUp).Row
        For i = 2 To sonSatir
            If CStr(wsCalisan.Cells(i, 1).Value) = arananTC Then
                txtArizaAdSoyad.Text = wsCalisan.Cells(i, 2).Value & " " & wsCalisan.Cells(i, 3).Value
                txtArizaTel.Text = wsCalisan.Cells(i, 4).Value
                txtArizaDepartman.Text = wsCalisan.Cells(i, 5).Value
                calisanBulundu = True: Exit For
            End If
        Next i
    End If

    If calisanBulundu And Not wsTeslim Is Nothing Then
        sonSatir = wsTeslim.Cells(Rows.Count, "B").End(xlUp).Row
        For i = sonSatir To 2 Step -1
            If CStr(wsTeslim.Cells(i, 2).Value) = arananTC Then
                txtArizaCihazTuru.Text = wsTeslim.Cells(i, 3).Value & " - " & wsTeslim.Cells(i, 4).Value
                txtArizaCihazSeri.Text = wsTeslim.Cells(i, 5).Value
                Exit For
            End If
        Next i
    ElseIf Not calisanBulundu Then
        MsgBox "Bu TC/ID numarasına ait bir çalışan bulunamadı!", vbExclamation, "Kayıt Yok"
    End If
End Sub

Private Sub btnArizaKaydet_Click()
    Dim wsAriza As Worksheet
    Dim sonSatir As Long, yeniArizaId As String

    On Error Resume Next
    Set wsAriza = ThisWorkbook.Sheets("Arizalar")
    On Error GoTo 0

    If wsAriza Is Nothing Then
        MsgBox "Lütfen Excel'de 'Arizalar' adında yeni bir sayfa oluşturun!", vbCritical
        Exit Sub
    End If

    If txtArizaTc.Text = "" Or txtArizaSorun.Text = "" Or cmbArizaKategori.Value = "" Then
        MsgBox "Lütfen Personel ID, Kategori ve Sorun Açıklaması alanlarını doldurun!", vbExclamation
        Exit Sub
    End If

    sonSatir = wsAriza.Cells(Rows.Count, 1).End(xlUp).Row + 1
    If sonSatir = 2 Then
        yeniArizaId = "ARZ-1001"
    Else
        yeniArizaId = "ARZ-" & (Val(Mid(wsAriza.Cells(sonSatir - 1, 1).Value, 5)) + 1)
    End If

    wsAriza.Cells(sonSatir, 1).Value = yeniArizaId
    wsAriza.Cells(sonSatir, 2).Value = txtArizaTarih.Text
    wsAriza.Cells(sonSatir, 3).Value = txtArizaTc.Text
    wsAriza.Cells(sonSatir, 4).Value = txtArizaAdSoyad.Text
    wsAriza.Cells(sonSatir, 5).Value = txtArizaDepartman.Text
    wsAriza.Cells(sonSatir, 6).Value = txtArizaTel.Text
    wsAriza.Cells(sonSatir, 7).Value = txtArizaKonum.Text
    wsAriza.Cells(sonSatir, 8).Value = txtArizaCihazSeri.Text
    wsAriza.Cells(sonSatir, 9).Value = txtArizaCihazTuru.Text
    wsAriza.Cells(sonSatir, 10).Value = cmbArizaOncelik.Value
    wsAriza.Cells(sonSatir, 11).Value = cmbArizaKategori.Value
    wsAriza.Cells(sonSatir, 12).Value = txtArizaSorun.Text
    If cmbArizaDurum.Value = "" Then wsAriza.Cells(sonSatir, 13).Value = "Açık (Bekliyor)" Else wsAriza.Cells(sonSatir, 13).Value = cmbArizaDurum.Value
    wsAriza.Cells(sonSatir, 14).Value = cmbArizaITPersonel.Value
    wsAriza.Cells(sonSatir, 15).Value = txtArizaCozum.Text

    MsgBox "Arıza kaydı başarıyla oluşturuldu! Kayıt No: " & yeniArizaId, vbInformation, "Başarılı"

    On Error Resume Next
    txtArizaTc.Text = "": txtArizaAdSoyad.Text = "": txtArizaDepartman.Text = ""
    txtArizaTel.Text = "": txtArizaKonum.Text = "": txtArizaCihazSeri.Text = "": txtArizaCihazTuru.Text = ""
    cmbArizaKategori.ListIndex = -1: cmbArizaOncelik.ListIndex = -1: txtArizaSorun.Text = ""
    cmbArizaITPersonel.ListIndex = -1: cmbArizaDurum.ListIndex = -1: txtArizaCozum.Text = ""
    On Error GoTo 0
End Sub

Private Sub btnArizaAra_Click()
    Dim wsAriza As Worksheet
    Dim sonSatir As Long, i As Long, j As Integer
    Dim arananKelime As String, satirEklendi As Boolean
    
    On Error Resume Next
    Set wsAriza = ThisWorkbook.Sheets("Arizalar")
    arananKelime = LCase(Trim(txtArizaArama.Text))
    lstArizaSonuclar.Clear
    On Error GoTo 0
    
    If arananKelime = "" Then MsgBox "Lütfen aramak için kelime yazın!", vbExclamation: Exit Sub
    
    lstArizaSonuclar.AddItem "ARIZA ID": lstArizaSonuclar.List(0, 1) = "TARİH"
    lstArizaSonuclar.List(0, 2) = "TC / ID": lstArizaSonuclar.List(0, 3) = "AD SOYAD"
    lstArizaSonuclar.List(0, 4) = "DEPARTMAN": lstArizaSonuclar.List(0, 5) = "KATEGORİ"
    lstArizaSonuclar.List(0, 6) = "SORUN": lstArizaSonuclar.List(0, 7) = "DURUM"
    
    If wsAriza Is Nothing Then Exit Sub
    sonSatir = wsAriza.Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To sonSatir
        satirEklendi = False
        For j = 1 To 15
            If InStr(1, LCase(wsAriza.Cells(i, j).Value), arananKelime, vbTextCompare) > 0 Then
                lstArizaSonuclar.AddItem wsAriza.Cells(i, 1).Value
                lstArizaSonuclar.List(lstArizaSonuclar.ListCount - 1, 1) = wsAriza.Cells(i, 2).Value
                lstArizaSonuclar.List(lstArizaSonuclar.ListCount - 1, 2) = wsAriza.Cells(i, 3).Value
                lstArizaSonuclar.List(lstArizaSonuclar.ListCount - 1, 3) = wsAriza.Cells(i, 4).Value
                lstArizaSonuclar.List(lstArizaSonuclar.ListCount - 1, 4) = wsAriza.Cells(i, 5).Value
                lstArizaSonuclar.List(lstArizaSonuclar.ListCount - 1, 5) = wsAriza.Cells(i, 11).Value
                lstArizaSonuclar.List(lstArizaSonuclar.ListCount - 1, 6) = wsAriza.Cells(i, 12).Value
                lstArizaSonuclar.List(lstArizaSonuclar.ListCount - 1, 7) = wsAriza.Cells(i, 13).Value
                satirEklendi = True: Exit For
            End If
        Next j
    Next i
    
    If lstArizaSonuclar.ListCount = 1 Then
        MsgBox "Kritere uygun arıza kaydı bulunamadı.", vbInformation
        lstArizaSonuclar.Clear
    End If
End Sub

' ======================================================================
' 7. STOK VE DEPO YÖNETİM MODÜLÜ
' ======================================================================
Private Sub btnStokKaydet_Click()
    Dim wsStok As Worksheet
    Dim sonSatir As Long, i As Long, stokBulundu As Boolean
    
    On Error Resume Next
    Set wsStok = ThisWorkbook.Sheets("Stok_Depo")
    On Error GoTo 0
    
    If txtStokKodu.Text = "" Or txtUrunAdi.Text = "" Or txtMevcutMiktar.Text = "" Then
        MsgBox "Stok Kodu, Ürün Adı ve Miktar alanlarını doldurun!", vbExclamation: Exit Sub
    End If
    
    stokBulundu = False
    sonSatir = wsStok.Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To sonSatir
        If LCase(wsStok.Cells(i, 1).Value) = LCase(txtStokKodu.Text) Then
            wsStok.Cells(i, 2).Value = cmbStokKategori.Value
            wsStok.Cells(i, 3).Value = txtUrunAdi.Text
            wsStok.Cells(i, 4).Value = Val(txtMevcutMiktar.Text)
            wsStok.Cells(i, 5).Value = Val(txtKritikEsik.Text)
            stokBulundu = True
            MsgBox "Stok güncellendi!", vbInformation
            Exit For
        End If
    Next i
    
    If stokBulundu = False Then
        wsStok.Cells(sonSatir + 1, 1).Value = txtStokKodu.Text
        wsStok.Cells(sonSatir + 1, 2).Value = cmbStokKategori.Value
        wsStok.Cells(sonSatir + 1, 3).Value = txtUrunAdi.Text
        wsStok.Cells(sonSatir + 1, 4).Value = Val(txtMevcutMiktar.Text)
        wsStok.Cells(sonSatir + 1, 5).Value = Val(txtKritikEsik.Text)
        MsgBox "Yeni ürün eklendi!", vbInformation
    End If
    
    Call LogYaz("Stok hareket kaydı yapıldı. Kod: " & txtStokKodu.Text)
    On Error Resume Next
    txtStokKodu.Text = "": cmbStokKategori.ListIndex = -1: txtUrunAdi.Text = ""
    txtMevcutMiktar.Text = "": txtKritikEsik.Text = ""
    On Error GoTo 0
    Call StokListesiniYenile
End Sub

Sub StokListesiniYenile()
    Dim ws As Worksheet: Dim i As Long, son As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Stok_Depo")
    son = ws.Cells(Rows.Count, 1).End(xlUp).Row
    lstStokDurumu.Clear: lstStokDurumu.ColumnCount = 5
    lstStokDurumu.ColumnWidths = "60 pt; 90 pt; 130 pt; 50 pt; 50 pt"
    lstStokDurumu.AddItem "STOK KODU": lstStokDurumu.List(0, 1) = "KATEGORİ": lstStokDurumu.List(0, 2) = "ÜRÜN ADI"
    lstStokDurumu.List(0, 3) = "MİKTAR": lstStokDurumu.List(0, 4) = "KRİTİK EŞİK"
    On Error GoTo 0
    If ws Is Nothing Or son < 2 Then Exit Sub
    For i = 2 To son
        lstStokDurumu.AddItem ws.Cells(i, 1).Value
        lstStokDurumu.List(lstStokDurumu.ListCount - 1, 1) = ws.Cells(i, 2).Value
        lstStokDurumu.List(lstStokDurumu.ListCount - 1, 2) = ws.Cells(i, 3).Value
        lstStokDurumu.List(lstStokDurumu.ListCount - 1, 3) = ws.Cells(i, 4).Value
        lstStokDurumu.List(lstStokDurumu.ListCount - 1, 4) = ws.Cells(i, 5).Value
    Next i
End Sub

Private Sub txtStokAra_Change()
    Dim wsStok As Worksheet, sonSatir As Long, i As Long, aranan As String
    On Error Resume Next
    Set wsStok = ThisWorkbook.Sheets("Stok_Depo")
    aranan = LCase(Trim(txtStokAra.Text))
    lstStokDurumu.Clear
    lstStokDurumu.AddItem "STOK KODU": lstStokDurumu.List(0, 1) = "KATEGORİ"
    lstStokDurumu.List(0, 2) = "ÜRÜN ADI": lstStokDurumu.List(0, 3) = "MİKTAR": lstStokDurumu.List(0, 4) = "KRİTİK EŞİK"
    sonSatir = wsStok.Cells(Rows.Count, 1).End(xlUp).Row
    On Error GoTo 0
    
    If aranan = "" Then Call StokListesiniYenile: Exit Sub
    If wsStok Is Nothing Then Exit Sub
    
    For i = 2 To sonSatir
        If InStr(1, LCase(wsStok.Cells(i, 3).Value), aranan, vbTextCompare) > 0 Or _
           InStr(1, LCase(wsStok.Cells(i, 1).Value), aranan, vbTextCompare) > 0 Then
            lstStokDurumu.AddItem wsStok.Cells(i, 1).Value
            lstStokDurumu.List(lstStokDurumu.ListCount - 1, 1) = wsStok.Cells(i, 2).Value
            lstStokDurumu.List(lstStokDurumu.ListCount - 1, 2) = wsStok.Cells(i, 3).Value
            lstStokDurumu.List(lstStokDurumu.ListCount - 1, 3) = wsStok.Cells(i, 4).Value
            lstStokDurumu.List(lstStokDurumu.ListCount - 1, 4) = wsStok.Cells(i, 5).Value
        End If
    Next i
End Sub

Sub KritikStokKontrol()
    Dim ws As Worksheet
    Dim i As Long, son As Long
    Dim say As Integer
    say = 0
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Stok_Depo")
    son = ws.Cells(Rows.Count, 1).End(xlUp).Row
    On Error GoTo 0
    If ws Is Nothing Or son < 2 Then Exit Sub
    For i = 2 To son
        If Val(ws.Cells(i, 4).Value) <= Val(ws.Cells(i, 5).Value) Then say = say + 1
    Next i
    If say > 0 Then MsgBox "Dikkat: Depoda kritik stok seviyesinde " & say & " adet ürün var!", vbExclamation, "Stok Uyarısı"
End Sub

' ======================================================================
' 8. FORM BAŞLATICI (INITIALIZE - GÜVENLİK VE SEKMELER KİLİTLENDİ)
' ======================================================================
Private Sub UserForm_Initialize()
    On Error Resume Next
    
    ' --- SİSTEM GÜVENLİK KİLİDİ (GİRİŞ EKRANI SIKILAŞTIRMA) ---
    MultiPage1.Style = fmTabStyleNone  ' Üstteki tüm sekmeleri tamamen gizler
    MultiPage1.Pages(0).Visible = True ' Sadece login sayfasını görünür kılar
    MultiPage1.Value = 0               ' Form açıldığında kesin olarak Admin Giriş ekranına kilitler
    txtSifre.PasswordChar = "*"
    
    ' ComboBox statik dolguları
    cmbTema.List = Array("Light Mode (Açık)", "Dark Mode (Koyu)")
    cmbAyarTema.List = Array("Light Mode (Açık)", "Dark Mode (Koyu)")
    cmbFont.List = Array("Segoe UI", "Calibri", "Tahoma")
    cmbAyarFont.List = Array("Segoe UI", "Calibri", "Tahoma")
    
    lstTeslimYanCihaz.List = Array("Adaptör", "Çanta", "Mouse", "Klavye")
    lstGeriYanCihaz.List = Array("Adaptör", "Çanta", "Mouse", "Klavye", "Eksik/Yok")
    cmbCihazDurumu.List = Array("Sağlam", "Arızalı", "Kullanılamaz/Hurda")
    cmbArizaKategori.List = Array("Donanım", "Yazılım", "Ağ / İnternet", "Yazıcı", "Diğer")
    cmbArizaDurum.List = Array("Açık (Bekliyor)", "İşlemde (İnceleniyor)", "Parça Bekliyor", "Çözüldü")
    cmbArizaOncelik.List = Array("Düşük", "Normal", "Yüksek", "Kritik")
    cmbArizaITPersonel.List = Array("IT - Ahmet", "IT - Semih", "IT - Ayşe")
    cmbStokKategori.List = Array("Toner / Kartuş", "Çevre Birimi (Mouse, Klavye)", "Kablo / Adaptör", "Ağ Cihazı", "Yedek Parça", "Diğer")
    
    txtTeslimTarihi.Text = Format(Date, "dd.mm.yyyy")
    txtGeriAlimTarihi.Text = Format(Date, "dd.mm.yyyy")
    txtArizaTarih.Text = Format(Date, "dd.mm.yyyy")
    
    lstArizaSonuclar.ColumnCount = 8
    imgPersonel.PictureSizeMode = fmPictureSizeModeStretch
    imgPersonel.BorderStyle = fmBorderStyleSingle
    On Error GoTo 0
    
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Ayarlar")
    On Error GoTo 0
    
    ' 424 Nesne Bulunamadı Hatasını Kökten Kesen Initialize Bloğu
    On Error Resume Next
    If Not ws Is Nothing Then
        txtAyarKullanici.Text = ws.Cells(2, 1).Value
        txtYeniKullaniciAdi.Text = ws.Cells(2, 1).Value
        txtAyarSifre.Text = ws.Cells(2, 2).Value
        txtYeniSifre.Text = ws.Cells(2, 2).Value
        txtAyarGuvenlik.Text = ws.Cells(2, 3).Value
        
        cmbTema.Value = ws.Cells(2, 4).Value
        cmbAyarTema.Value = ws.Cells(2, 4).Value
        cmbFont.Value = ws.Cells(2, 5).Value
        cmbAyarFont.Value = ws.Cells(2, 5).Value
    End If
    
    ' Eğer veritabanı daha önce hiç ayar görmediyse default atamalar
    If txtAyarKullanici.Text = "" Then txtAyarKullanici.Text = "admin"
    If txtAyarSifre.Text = "" Then txtAyarSifre.Text = "1234"
    If txtYeniKullaniciAdi.Text = "" Then txtYeniKullaniciAdi.Text = "admin"
    If txtYeniSifre.Text = "" Then txtYeniSifre.Text = "1234"
    If cmbTema.Value = "" Then cmbTema.Value = "Light Mode (Açık)"
    If cmbAyarTema.Value = "" Then cmbAyarTema.Value = "Light Mode (Açık)"
    On Error GoTo 0
    
    Me.Width = 530: Me.Height = 360
    
    ' Alt sistem motorlarını uyandırıyoruz
    Call btnAyarlariKaydet_Click
    Call StokListesiniYenile
    Call LoglariKutuyaDoldur
End Sub

' ======================================================================
' 9. ŞİFRE RECOVERY (GÜVENLİK SORUSU SİSTEMİ)
' ======================================================================
Private Sub lblSifreUnuttum_Click()
    Dim cvp As String, gizliCevap As String
    
    On Error Resume Next
    gizliCevap = ThisWorkbook.Sheets("Ayarlar").Cells(2, 3).Value
    On Error GoTo 0
    
    If gizliCevap = "" Then gizliCevap = "semih seker"
    
    cvp = InputBox("Sistemi kurtarmak için güvenlik sorusunu yanıtlayın:" & vbCrLf & vbCrLf & _
                   "Bu otomasyonu geliştiren IT Teknisyeninin adı nedir?", "Şifre Kurtarma Paneli")
                   
    If LCase(Trim(cvp)) = LCase(Trim(gizliCevap)) Then
        Call LogYaz("Şifre kurtarma sorusu doğru yanıtlandı. Kimlik doğrulandı.")
        MsgBox "Kimlik Doğrulandı!" & vbCrLf & vbCrLf & _
               "Sistem Giriş Adı: " & ThisWorkbook.Sheets("Ayarlar").Cells(2, 1).Value & vbCrLf & _
               "Sistem Giriş Şifresi: " & ThisWorkbook.Sheets("Ayarlar").Cells(2, 2).Value, vbInformation, "Erişim Onayı"
    Else
        If cvp <> "" Then
            Call LogYaz("DİKKAT! Şüpheli şifre kurtarma denemesi yapıldı! Girilen Cevap: " & cvp)
            MsgBox "Hatalı cevap! Güvenlik protokolü nedeniyle işlem reddedildi.", vbCritical, "Erişim Reddedildi"
        End If
    End If
End Sub