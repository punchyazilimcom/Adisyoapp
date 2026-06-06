Attribute VB_Name = "BasakAdisyo"
'==============================================================================
' BASAK - Adisyo -> Excel (VBA, Python YOK)
'==============================================================================
' Bu modul, bulundugu .xlsm dosyasinin KENDI ICINDE calisir. Adisyo API'sinden
' veriyi WinHTTP ile ceker, JSON'u kendi parser'iyla ayristirir ve secilen
' tarihin "HAZIRAN (N)" sayfasinda SADECE izinli hucrelere yazar.
'
' KURULUM (her sube dosyasi icin bir kez):
'   1) Excel'i ac > Alt+F11 (VBA editoru) > Dosya menusu > "Dosya Al/Import" >
'      bu BasakAdisyo.bas dosyasini sec.
'   2) Asagidaki CONFIG bolumunde X_API_SECRET ve X_API_CONSUMER'i O SUBEYE gore
'      doldur. X_API_KEY tum subelerde ayni.
'   3) Bir sayfaya buton ekle (Gelistirici > Ekle > Dugme) ve "BasakDoldur_Dun"
'      makrosuna bagla. Ya da tarih sorsun istersen "BasakDoldur_Sor".
'   4) Dosyayi makro etkin (.xlsm) olarak kaydet. Makrolar engellenirse:
'      Dosya > Secenekler > Guven Merkezi > Makro Ayarlari > etkinlestir.
'
' SUBE BILGILERI (X_API_CONSUMER) - secret'lari guvenli yerden (config) girin:
'   Demetevler   -> "Punch - Demetevler"
'   Bahcelievler -> "Punch - Bahcelievler"
'   Etlik        -> "Punch - Etlik"
'   Batikent     -> "Punch - Batikent"
'==============================================================================
Option Explicit

'------------------------------ CONFIG ----------------------------------------
Private Const BASE_URL As String = "https://ext.adisyo.com/api/External/v2"
Private Const X_API_KEY As String = "BURAYA_ORTAK_API_KEY"        ' tum subelerde ayni
Private Const X_API_SECRET As String = "BURAYA_SUBE_API_SECRET"   ' BU dosyanin subesi
Private Const X_API_CONSUMER As String = "Punch - Demetevler"     ' BU dosyanin subesi

Private Const DAY_START_HOUR As Integer = 0   ' gunluk pencere baslangici
Private Const UTC_OFFSET As Integer = 3       ' Istanbul = UTC+3 (sabit)
Private Const PAGE_DELAY_SEC As Long = 42     ' sayfalar arasi bekleme
Private Const API_OK As Long = 100
Private Const KORUMA_SIFRE As String = "571632"   ' gecmis gun sayfa korumasi sifresi
'------------------------------------------------------------------------------

' Izinli hucreler (bunlarin DISINA asla yazilmaz)
Private Const ALLOWED As String = "|L7|L11|L17|L19|E16|E18|E20|R9|Q10|R11|Q12|R13|Q14|R15|Q16|"

' JSON parser durum degiskenleri
Private gJson As String
Private gPos As Long
Private gAsama As String   ' teshis: hata aninda hangi adimda oldugumuz

'==============================================================================
' GIRIS NOKTALARI (butona baglanacak makrolar)
'==============================================================================
Public Sub BasakDoldur_Dun()
    BasakDoldur Date - 1
End Sub

Public Sub BasakDoldur_Sor()
    Dim s As String
    s = InputBox("Hangi gun islensin? (yyyy-aa-gg)", "Basak Adisyo", Format(Date - 1, "yyyy-mm-dd"))
    If s = "" Then Exit Sub
    BasakDoldur DateSerial(CInt(Mid(s, 1, 4)), CInt(Mid(s, 6, 2)), CInt(Mid(s, 9, 2)))
End Sub

' Aktif sayfanin korumasini sifreyle kaldirir (gecmis bir gunu elle duzeltmek icin)
Public Sub BasakKilitAc()
    On Error Resume Next
    ActiveSheet.Unprotect Password:=KORUMA_SIFRE
    On Error GoTo 0
    MsgBox "Koruma kaldirildi: " & ActiveSheet.Name & vbCrLf & _
           "Duzeltme bitince tekrar kilitlemek icin BasakKilitle calistirin.", _
           vbInformation, "Basak Adisyo"
End Sub

' Aktif sayfayi tekrar sifreyle kilitler (tum hucreler)
Public Sub BasakKilitle()
    On Error Resume Next
    ActiveSheet.Unprotect Password:=KORUMA_SIFRE
    ActiveSheet.Cells.Locked = True
    ActiveSheet.Protect Password:=KORUMA_SIFRE, DrawingObjects:=True, Contents:=True, Scenarios:=True
    On Error GoTo 0
    MsgBox "Kilitlendi: " & ActiveSheet.Name, vbInformation, "Basak Adisyo"
End Sub

'==============================================================================
' ANA AKIS
'==============================================================================
Public Sub BasakDoldur(ByVal isleGun As Date)
    On Error GoTo Hata
    Application.StatusBar = "Adisyo: urunler cekiliyor..."

    ' 1) Pencere: yerel [gun+DAY_START_HOUR, +24h) -> UTC startDate
    Dim yerelBas As Date, yerelBit As Date, startUtc As String
    yerelBas = isleGun + TimeSerial(DAY_START_HOUR, 0, 0)
    yerelBit = yerelBas + 1
    startUtc = FmtTs(DateAdd("h", -UTC_OFFSET, yerelBas))

    ' UTC pencere (insertDate filtresi icin)
    Dim utcBas As Date, utcBit As Date
    utcBas = DateAdd("h", -UTC_OFFSET, yerelBas)
    utcBit = DateAdd("h", -UTC_OFFSET, yerelBit)

    ' 1b) Islenen gunden ONCEKI gunleri kilitle (COM bozulmadan, makro basinda)
    gAsama = "Onceki gunler kilitleniyor"
    Dim kilitSayisi As Long
    kilitSayisi = KilitleOncekiGunler(Day(isleGun))

    ' 2) Urunler -> kir pidesi / kutu icecek isim kumeleri
    Dim prodResp As Object, rawProducts As String
    gAsama = "Products HTTP"
    rawProducts = HttpGet(BASE_URL & "/Products")
    gAsama = "Products JSON parse (uzunluk=" & Len(rawProducts) & ")"
    Set prodResp = ParseJson(rawProducts)
    If Nz(prodResp("status")) <> API_OK Then Err.Raise vbObjectError + 1, , _
        "Products status=" & Nz(prodResp("status")) & " " & Nz(prodResp("message"))

    Dim kir As Object, kutu As Object
    Set kir = CreateObject("Scripting.Dictionary")
    Set kutu = CreateObject("Scripting.Dictionary")
    gAsama = "Urun haritasi kuruluyor"
    BuildIndex prodResp("data"), kir, kutu

    ' 3) Siparisleri sayfalayarak cek
    Application.StatusBar = "Adisyo: siparisler cekiliyor..."
    Dim orders As Collection
    Set orders = New Collection
    Dim page As Long, pageCount As Long
    page = 1
    Do
        Dim url As String, rawOrders As String
        url = BASE_URL & "/CompletedOrders?page=" & page & _
              "&startDate=" & UrlEnc(startUtc) & "&includeCancelled=false"
        gAsama = "Siparis HTTP sayfa " & page
        rawOrders = HttpGet(url)
        gAsama = "Siparis JSON parse sayfa " & page & " (uzunluk=" & Len(rawOrders) & ")"
        Dim resp As Object
        Set resp = ParseJson(rawOrders)
        If Nz(resp("status")) <> API_OK Then Err.Raise vbObjectError + 2, , _
            "CompletedOrders status=" & Nz(resp("status")) & " " & Nz(resp("message"))

        gAsama = "Siparis sayfa " & page & " isleniyor"
        Dim arr As Object, i As Long
        Set arr = Nothing
        If resp.Exists("orders") Then Set arr = resp("orders")
        If Not arr Is Nothing Then
            For i = 1 To arr.Count
                orders.Add arr(i)
            Next i
        End If

        pageCount = CLng(NzNum(resp("pageCount"), 1))
        If page >= pageCount Then Exit Do
        page = page + 1
        Application.StatusBar = "Adisyo: rate limit, " & PAGE_DELAY_SEC & " sn bekleniyor..."
        GuvenliBekle PAGE_DELAY_SEC
    Loop

    ' 4) insertDate ile pencereye filtrele
    gAsama = "Pencere filtreleme (" & orders.Count & " siparis)"
    Dim pencere As Collection
    Set pencere = New Collection
    Dim o As Object, ins As Date
    For i = 1 To orders.Count
        Set o = orders(i)
        If o.Exists("insertDate") Then
            ins = ParseUtc(CStr(o("insertDate")))
            If ins >= utcBas And ins < utcBit Then pencere.Add o
        End If
    Next i

    ' 5) Hucre degerlerini hesapla
    gAsama = "Hesaplama (" & pencere.Count & " siparis)"
    Dim c As Object
    Set c = CreateObject("Scripting.Dictionary")
    c("L7") = QtyMatch(pencere, "kucuk ayran", False, kir, kutu, "contains")
    c("L11") = QtyMatch(pencere, "cay", False, kir, kutu, "equals")
    c("L17") = QtyKir(pencere, kir)
    c("L19") = QtyKutu(pencere, kutu)
    c("E16") = QtyMatch(pencere, "kurye menu", False, kir, kutu, "contains")
    c("E18") = QtyMatch(pencere, "kutu icecek menu", False, kir, kutu, "contains")
    c("E20") = QtyMatch(pencere, "buyuk ayran menu", False, kir, kutu, "contains")
    c("R9") = PayNet(pencere, "YS Online"):       c("Q10") = PayDisc(pencere, "YS Online")
    c("R11") = PayNet(pencere, "Trendyol Online"): c("Q12") = PayDisc(pencere, "Trendyol Online")
    c("R13") = PayNet(pencere, "Getir Online"):    c("Q14") = PayDisc(pencere, "Getir Online")
    c("R15") = PayNet(pencere, "Migros Online"):   c("Q16") = PayDisc(pencere, "Migros Online")

    ' 6) Hedef sayfaya YAZ (sadece izinli hucreler)
    ' Uzun makro + WinHTTP/ADODB sonrasi Excel COM durumunu normallestir
    gAsama = "Excel toparlaniyor"
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    DoEvents
    On Error GoTo Hata

    gAsama = "Hedef sayfa araniyor"
    Dim sheetName As String
    sheetName = "HAZIRAN (" & Day(isleGun) & ")"   ' NOT: sablon Turkce "HAZIRAN" / "HAZIRAN" olabilir
    Dim ws As Worksheet, deneme As Long
    Set ws = Nothing
    For deneme = 1 To 3
        On Error Resume Next
        Set ws = BulSayfa(Day(isleGun))
        On Error GoTo Hata
        If Not ws Is Nothing Then Exit For
        DoEvents
        Application.Wait Now + TimeSerial(0, 0, 2)
    Next deneme

    If ws Is Nothing Then
        Application.StatusBar = False
        Dim sliste As String, ix As Long
        On Error Resume Next
        For ix = 1 To ThisWorkbook.Worksheets.Count
            sliste = sliste & "[" & ThisWorkbook.Worksheets(ix).Name & "] "
        Next ix
        sliste = sliste & " | Aktif: [" & ActiveSheet.Name & "]"
        On Error GoTo 0
        MsgBox "Sayfa bulunamadi: '" & sheetName & "'." & vbCrLf & vbCrLf & _
               "Mevcut sayfalar:" & vbCrLf & sliste, _
               vbExclamation, "Basak Adisyo"
        Exit Sub
    End If

    ' Aktif (islenen) sayfa onceden kilitlenmis olabilir -> sifreyle ac, sonra yaz
    gAsama = "Aktif sayfa korumasi aciliyor"
    On Error Resume Next
    ws.Unprotect Password:=KORUMA_SIFRE
    On Error GoTo Hata

    gAsama = "Hucrelere yaziliyor"
    Dim k As Variant
    For Each k In c.Keys
        If InStr(ALLOWED, "|" & k & "|") = 0 Then Err.Raise vbObjectError + 3, , "Izinsiz hucre: " & k
        ws.Range(CStr(k)).Value = c(k)   ' SADECE value; bicim/format korunur
    Next k

    Application.StatusBar = False
    MsgBox sheetName & " dolduruldu (" & pencere.Count & " siparis)." & vbCrLf & _
           "L17 pide=" & c("L17") & "  L19 kutu=" & c("L19") & "  YS net=" & c("R9") & vbCrLf & _
           "Onceki gunler kilitlendi: " & kilitSayisi & " sayfa", _
           vbInformation, "Basak Adisyo"
    Exit Sub

Hata:
    Application.StatusBar = False
    MsgBox "HATA [" & gAsama & "]" & vbCrLf & _
           Err.Description & vbCrLf & _
           "No=" & Err.Number & "  Kaynak=" & Err.Source, _
           vbCritical, "Basak Adisyo"
End Sub

'==============================================================================
' HESAP KURALLARI
'==============================================================================
Private Sub BuildIndex(ByVal dataArr As Object, ByRef kir As Object, ByRef kutu As Object)
    If dataArr Is Nothing Then Exit Sub
    Dim i As Long, j As Long, cat As Object, prods As Object, p As Object
    Dim nb As String, n As String
    For i = 1 To dataArr.Count
        Set cat = dataArr(i)
        nb = ClassifyCat(NormTr(CStr(Nz(cat("categoryName")))))
        If cat.Exists("products") Then
            On Error Resume Next
            Set prods = cat("products")
            On Error GoTo 0
            If Not prods Is Nothing Then
                For j = 1 To prods.Count
                    Set p = prods(j)
                    n = NormTr(CStr(Nz(p("productName"))))
                    If Len(n) > 0 Then
                        If nb = "kir" Then If Not kir.Exists(n) Then kir.Add n, True
                        If nb = "kutu" Then If Not kutu.Exists(n) Then kutu.Add n, True
                    End If
                Next j
            End If
            Set prods = Nothing
        End If
    Next i
End Sub

Private Function ClassifyCat(ByVal nc As String) As String
    If InStr(nc, "kir pide") > 0 Then ClassifyCat = "kir": Exit Function
    If InStr(nc, "kutu icecek") > 0 Then ClassifyCat = "kutu": Exit Function
    If InStr(nc, "menu") > 0 Then ClassifyCat = "menu": Exit Function
    If InStr(nc, "icecek") > 0 Then ClassifyCat = "icecek": Exit Function
    ClassifyCat = "other"
End Function

' Genel adet toplayici. mode="contains" veya "equals". paketOnly=True ise sadece paket.
Private Function QtyMatch(ByVal ords As Collection, ByVal needle As String, _
        ByVal paketOnly As Boolean, kir As Object, kutu As Object, ByVal mode As String) As Double
    Dim t As Double, i As Long, j As Long, o As Object, prods As Object, p As Object, n As String
    For i = 1 To ords.Count
        Set o = ords(i)
        If (Not paketOnly) Or IsPaket(o) Then
            If o.Exists("products") Then
                Set prods = o("products")
                For j = 1 To prods.Count
                    Set p = prods(j)
                    n = NormTr(CStr(Nz(p("productName"))))
                    If (mode = "equals" And n = needle) Or (mode = "contains" And InStr(n, needle) > 0) Then
                        t = t + Val(CStr(NzNum(p("quantity"), 0)))
                    End If
                Next j
            End If
        End If
    Next i
    QtyMatch = t
End Function

' L17: KIR PIDELERI (Kusbasili haric) - sadece paket
Private Function QtyKir(ByVal ords As Collection, kir As Object) As Double
    Dim t As Double, i As Long, j As Long, o As Object, prods As Object, p As Object, n As String
    For i = 1 To ords.Count
        Set o = ords(i)
        If IsPaket(o) And o.Exists("products") Then
            Set prods = o("products")
            For j = 1 To prods.Count
                Set p = prods(j)
                n = NormTr(CStr(Nz(p("productName"))))
                If kir.Exists(n) And InStr(n, "kusbasi") = 0 Then t = t + Val(CStr(NzNum(p("quantity"), 0)))
            Next j
        End If
    Next i
    QtyKir = t
End Function

' L19: KUTU ICECEKLER (kanal kopyalari dahil) - sadece paket
Private Function QtyKutu(ByVal ords As Collection, kutu As Object) As Double
    Dim t As Double, i As Long, j As Long, o As Object, prods As Object, p As Object, n As String
    For i = 1 To ords.Count
        Set o = ords(i)
        If IsPaket(o) And o.Exists("products") Then
            Set prods = o("products")
            For j = 1 To prods.Count
                Set p = prods(j)
                n = NormTr(CStr(Nz(p("productName"))))
                If kutu.Exists(n) Then t = t + Val(CStr(NzNum(p("quantity"), 0)))
            Next j
        End If
    Next i
    QtyKutu = t
End Function

Private Function PayNet(ByVal ords As Collection, ByVal payName As String) As Double
    Dim t As Double, i As Long, j As Long, o As Object, pays As Object, p As Object
    For i = 1 To ords.Count
        Set o = ords(i)
        If o.Exists("payments") Then
            Set pays = o("payments")
            For j = 1 To pays.Count
                Set p = pays(j)
                If CStr(Nz(p("paymentName"))) = payName Then t = t + Val(CStr(NzNum(p("amount"), 0)))
            Next j
        End If
    Next i
    PayNet = Round(t, 2)
End Function

Private Function PayDisc(ByVal ords As Collection, ByVal payName As String) As Double
    Dim t As Double, i As Long, j As Long, o As Object, pays As Object, p As Object, found As Boolean
    For i = 1 To ords.Count
        Set o = ords(i)
        found = False
        If o.Exists("payments") Then
            Set pays = o("payments")
            For j = 1 To pays.Count
                Set p = pays(j)
                If CStr(Nz(p("paymentName"))) = payName Then found = True
            Next j
        End If
        If found Then t = t + Val(CStr(NzNum(o("discountAmount"), 0)))
    Next i
    PayDisc = Round(t, 2)
End Function

Private Function IsPaket(ByVal o As Object) As Boolean
    IsPaket = (CStr(Nz(o("orderType"))) = "Paket Sipari" & ChrW(351) & "i")   ' Paket Siparisi
End Function

'==============================================================================
' YARDIMCILAR - HTTP / sayfa / tarih / normalize
'==============================================================================
Private Function HttpGet(ByVal url As String) As String
    Dim http As Object, tries As Long, body As String, compact As String
    For tries = 0 To 2
        Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
        http.SetTimeouts 30000, 30000, 30000, 60000
        http.Open "GET", url, False
        http.SetRequestHeader "x-api-key", X_API_KEY
        http.SetRequestHeader "x-api-secret", X_API_SECRET
        http.SetRequestHeader "x-api-consumer", X_API_CONSUMER
        http.SetRequestHeader "Accept", "application/json"
        http.Send

        body = ""
        On Error Resume Next
        body = ResponseUtf8(http)
        On Error GoTo 0
        compact = Replace(body, " ", "")

        ' Adisyo istek limiti: HTTP 429 VEYA HTTP 400 + body status=601
        If http.Status = 429 Or InStr(compact, """status"":601") > 0 Then
            If tries < 2 Then
                BekleLimit url      ' bekle ve son bir kez daha dene
            Else
                Err.Raise vbObjectError + 12, , _
                    "Adisyo istek limiti (601) hala acilmadi." & vbCrLf & _
                    "Cok sik denendi. Lutfen 15-20 dakika HIC denemeden bekleyip" & vbCrLf & _
                    "tek seferde tekrar calistirin." & vbCrLf & url
            End If
        ElseIf http.Status >= 200 And http.Status < 300 Then
            HttpGet = body
            Set http = Nothing
            Exit Function
        Else
            Err.Raise vbObjectError + 10, , "HTTP " & http.Status & " - " & url & vbCrLf & Left$(body, 600)
        End If
    Next tries
    Err.Raise vbObjectError + 11, , "Istek limiti: tekrar deneyin (birkac dakika sonra)."
End Function

' Adisyo rate limit beklemesi: geri sayimli, Esc ile kesilebilir, donuk gorunmez.
' /Products 3 dk, diger uclar ~40 sn.
Private Sub BekleLimit(ByVal url As String)
    Dim w As Long, i As Long
    If InStr(url, "/Products") > 0 Then w = 185 Else w = 45
    On Error Resume Next
    For i = w To 1 Step -1
        Application.StatusBar = "Adisyo istek limiti (601) - kalan " & i & " sn (durdurmak: Esc)"
        DoEvents
        Application.Wait Now + TimeSerial(0, 0, 1)
    Next i
    On Error GoTo 0
End Sub

' Application.Wait bazi durumlarda 1004 verir; bu beklemeyi hataya dayanikli yapar.
Private Sub GuvenliBekle(ByVal saniye As Long)
    Dim hedef As Date
    hedef = Now + TimeSerial(0, 0, saniye)
    On Error Resume Next
    Do While Now < hedef
        DoEvents
        Application.Wait Now + TimeSerial(0, 0, 1)
    Loop
    On Error GoTo 0
End Sub

Private Function ResponseUtf8(ByVal http As Object) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Open
    st.Type = 1                ' binary
    st.Write http.ResponseBody
    st.Position = 0
    st.Type = 2                ' text
    st.Charset = "utf-8"
    ResponseUtf8 = st.ReadText
    st.Close
    Set st = Nothing
End Function

' Sablon "HAZIRAN (N)" / "HAZIRAN (N)" gibi farkli yazimlari da bulur.
Private Function BulSayfa(ByVal gun As Integer) As Worksheet
    Dim hedefSade As String, bulunan As Worksheet, idx As Long, cnt As Long
    hedefSade = "haziran" & gun               ' sadece harf+rakam: "haziran3"
    Set bulunan = Nothing

    ' 1) AKTIF sayfa zaten dogruysa onu kullan (koleksiyon enumeratorune DOKUNMA)
    '    Uzun makro sonrasi Worksheets taramasi cokebildigi icin once bu denenir.
    On Error Resume Next
    If SadeAd(NormTr(CStr(ActiveSheet.Name))) = hedefSade Then Set bulunan = ActiveSheet
    On Error GoTo 0

    ' 2) Degilse INDEKSLI ara (For Each YOK -> daha guvenli)
    If bulunan Is Nothing Then
        On Error Resume Next
        cnt = ThisWorkbook.Worksheets.Count
        For idx = 1 To cnt
            If SadeAd(NormTr(CStr(ThisWorkbook.Worksheets(idx).Name))) = hedefSade Then
                Set bulunan = ThisWorkbook.Worksheets(idx)
                Exit For
            End If
        Next idx
        On Error GoTo 0
    End If

    Set BulSayfa = bulunan
End Function

' Islenen gunden ONCEKI tum Haziran sayfalarini sifreyle korur (duzeltilemez yapar).
' TUM hucreleri kilitler (Locked=True) sonra Protect eder. Kilitlenen sayfa sayisini doner.
Private Function KilitleOncekiGunler(ByVal islenenGun As Integer) As Long
    Dim idx As Long, cnt As Long, g As Integer, n As Long
    On Error Resume Next
    cnt = ThisWorkbook.Worksheets.Count
    For idx = 1 To cnt
        g = HaziranGunu(CStr(ThisWorkbook.Worksheets(idx).Name))
        If g > 0 And g < islenenGun Then
            ThisWorkbook.Worksheets(idx).Unprotect Password:=KORUMA_SIFRE
            ThisWorkbook.Worksheets(idx).Cells.Locked = True
            ThisWorkbook.Worksheets(idx).Protect Password:=KORUMA_SIFRE, _
                DrawingObjects:=True, Contents:=True, Scenarios:=True
            n = n + 1
        End If
    Next idx
    On Error GoTo 0
    KilitleOncekiGunler = n
End Function

' "Haziran (3)" -> 3 ; Haziran sayfasi degilse 0
Private Function HaziranGunu(ByVal ad As String) As Integer
    Dim s As String, r As String
    HaziranGunu = 0
    s = SadeAd(NormTr(ad))            ' "haziran3"
    If Len(s) > 7 Then
        If Left(s, 7) = "haziran" Then
            r = Mid(s, 8)
            If IsNumeric(r) Then HaziranGunu = CInt(r)
        End If
    End If
End Function

' Bir metinden yalnizca a-z ve 0-9 karakterlerini birakir (digerlerini atar)
Private Function SadeAd(ByVal s As String) As String
    Dim i As Long, ch As String, o As String
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Then o = o & ch
    Next i
    SadeAd = o
End Function

Private Function FmtTs(ByVal dt As Date) As String
    FmtTs = Format(Year(dt), "0000") & "-" & Format(Month(dt), "00") & "-" & Format(Day(dt), "00") & _
            " " & Format(Hour(dt), "00") & ":" & Format(Minute(dt), "00") & ":" & Format(Second(dt), "00")
End Function

' "2026-06-03T10:00:00" / "...Z" / "... +03:00" -> UTC Date (ilk 19 karakter, UTC kabul)
Private Function ParseUtc(ByVal s As String) As Date
    s = Replace(s, "T", " ")
    ParseUtc = DateSerial(CInt(Mid(s, 1, 4)), CInt(Mid(s, 6, 2)), CInt(Mid(s, 9, 2))) + _
               TimeSerial(CInt(Mid(s, 12, 2)), CInt(Mid(s, 15, 2)), CInt(Mid(s, 18, 2)))
End Function

Private Function NormTr(ByVal s As String) As String
    s = Replace(s, ChrW(304), "i"): s = Replace(s, "I", "i"): s = Replace(s, ChrW(305), "i")  ' I-noktali / I / i-noktasiz
    s = Replace(s, ChrW(350), "s"): s = Replace(s, ChrW(351), "s")  ' S / s (cedilla)
    s = Replace(s, ChrW(286), "g"): s = Replace(s, ChrW(287), "g")  ' G / g (breve)
    s = Replace(s, ChrW(220), "u"): s = Replace(s, ChrW(252), "u")  ' U / u (umlaut)
    s = Replace(s, ChrW(214), "o"): s = Replace(s, ChrW(246), "o")  ' O / o (umlaut)
    s = Replace(s, ChrW(199), "c"): s = Replace(s, ChrW(231), "c")  ' C / c (cedilla)
    s = LCase(s)
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    NormTr = Trim(s)
End Function

Private Function UrlEnc(ByVal s As String) As String
    s = Replace(s, " ", "%20")
    s = Replace(s, ":", "%3A")
    UrlEnc = s
End Function

' Null/eksik -> "" ; sayisal Null/eksik -> varsayilan
Private Function Nz(ByVal v As Variant) As Variant
    If IsNull(v) Then Nz = "" Else Nz = v
End Function
Private Function NzNum(ByVal v As Variant, ByVal def As Double) As Double
    If IsNull(v) Then
        NzNum = def
    ElseIf IsNumeric(v) Then
        NzNum = CDbl(v)
    Else
        NzNum = def
    End If
End Function

'==============================================================================
' JSON PARSER (index tabanli recursive-descent; Python aynasiyla dogrulandi)
'  Obje -> Scripting.Dictionary | Dizi -> Collection | string/sayi/bool/null
'==============================================================================
Public Function ParseJson(ByVal s As String) As Object
    gJson = s
    gPos = 1
    Dim v As Variant
    AssignParse v
    Set ParseJson = v
End Function

Private Sub AssignParse(ByRef out As Variant)
    SkipWs
    Dim ch As String
    ch = Mid(gJson, gPos, 1)
    Select Case ch
        Case "{": Set out = ParseObject
        Case "[": Set out = ParseArray
        Case """": out = ParseString
        Case "t", "f": out = ParseBool
        Case "n": gPos = gPos + 4: out = Null
        Case Else: out = ParseNumber
    End Select
End Sub

Private Function ParseObject() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    gPos = gPos + 1            ' {
    SkipWs
    If Mid(gJson, gPos, 1) = "}" Then gPos = gPos + 1: Set ParseObject = d: Exit Function
    Do
        SkipWs
        Dim key As String
        key = ParseString
        SkipWs
        gPos = gPos + 1        ' :
        Dim v As Variant
        AssignParse v
        If Not d.Exists(key) Then d.Add key, v
        SkipWs
        Dim c As String
        c = Mid(gJson, gPos, 1): gPos = gPos + 1
        If c = "}" Then Exit Do
    Loop
    Set ParseObject = d
End Function

Private Function ParseArray() As Collection
    Dim col As Collection
    Set col = New Collection
    gPos = gPos + 1            ' [
    SkipWs
    If Mid(gJson, gPos, 1) = "]" Then gPos = gPos + 1: Set ParseArray = col: Exit Function
    Do
        Dim v As Variant
        AssignParse v
        col.Add v
        SkipWs
        Dim c As String
        c = Mid(gJson, gPos, 1): gPos = gPos + 1
        If c = "]" Then Exit Do
    Loop
    Set ParseArray = col
End Function

Private Function ParseString() As String
    gPos = gPos + 1            ' acilis "
    Dim sb As String, ch As String, e As String, hex As String, cp As Long
    Do
        ch = Mid(gJson, gPos, 1)
        If ch = """" Then gPos = gPos + 1: Exit Do
        If ch = "\" Then
            gPos = gPos + 1
            e = Mid(gJson, gPos, 1)
            Select Case e
                Case """": sb = sb & """"
                Case "\": sb = sb & "\"
                Case "/": sb = sb & "/"
                Case "b": sb = sb & Chr(8)
                Case "f": sb = sb & Chr(12)
                Case "n": sb = sb & vbLf
                Case "r": sb = sb & vbCr
                Case "t": sb = sb & vbTab
                Case "u"
                    hex = Mid(gJson, gPos + 1, 4)
                    cp = CLng("&H" & hex)
                    If cp > 32767 Then cp = cp - 65536   ' ChrW Integer tasmasini onle
                    sb = sb & ChrW(cp)
                    gPos = gPos + 4
            End Select
            gPos = gPos + 1
        Else
            sb = sb & ch
            gPos = gPos + 1
        End If
    Loop
    ParseString = sb
End Function

Private Function ParseNumber() As Double
    Dim st As Long
    st = gPos
    Do While gPos <= Len(gJson)
        If InStr("0123456789+-.eE", Mid(gJson, gPos, 1)) > 0 Then gPos = gPos + 1 Else Exit Do
    Loop
    ParseNumber = Val(Mid(gJson, st, gPos - st))   ' Val => locale-bagimsiz, nokta ondalik
End Function

Private Function ParseBool() As Boolean
    If Mid(gJson, gPos, 4) = "true" Then
        gPos = gPos + 4: ParseBool = True
    Else
        gPos = gPos + 5: ParseBool = False
    End If
End Function

Private Sub SkipWs()
    Do While gPos <= Len(gJson)
        Select Case Mid(gJson, gPos, 1)
            Case " ", vbTab, vbCr, vbLf: gPos = gPos + 1
            Case Else: Exit Do
        End Select
    Loop
End Sub
