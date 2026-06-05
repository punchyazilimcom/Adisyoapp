# Excel'in İçinde Çalışan Sürüm (VBA — Python YOK)

`BasakAdisyo.bas`, Adisyo verisini **doğrudan `.xlsm` dosyasının içinden** çeker.
Python, kurulum, sanal ortam **gerekmez**. Dosyayı aç → butona bas → o günün
verisi `HAZİRAN (N)` sayfasına yazılır. Yalnızca izinli hücrelere yazar, makroları
ve formülleri korur.

## Kurulum (her şube dosyası için bir kez)

1. Şube dosyasını aç (ör. `DEMETEVLER_HAZİRAN_HESAP.xlsm`).
2. **Alt + F11** → VBA editörü açılır.
3. **Dosya (File) → Dosyayı İçe Aktar (Import File…)** → `BasakAdisyo.bas` seç.
4. Modülün en üstündeki **CONFIG** bölümünü doldur:
   - `X_API_KEY` → ortak anahtar (tüm şubelerde aynı).
   - `X_API_SECRET` → **bu şubenin** secret'ı.
   - `X_API_CONSUMER` → **bu şubenin** consumer'ı (ör. `Punch - Demetevler`).

   | Şube | x-api-consumer |
   |------|----------------|
   | Demetevler | `Punch - Demetevler` |
   | Bahcelievler | `Punch - Bahcelievler` |
   | Etlik | `Punch - Etlik` |
   | Batikent | `Punch - Batikent` |

5. Bir sayfaya buton ekle: **Geliştirici (Developer) → Ekle → Düğme (Form Control)**,
   çizince açılan listede **`BasakDoldur_Dun`** makrosunu seç.
   (Geliştirici sekmesi yoksa: Dosya → Seçenekler → Şeridi Özelleştir → Geliştirici ✔)
6. Dosyayı **makro etkin** (`.xlsm`) kaydet.

> Makro engellenirse: **Dosya → Seçenekler → Güven Merkezi → Güven Merkezi Ayarları
> → Makro Ayarları** → makrolara izin ver. İnternetten inen dosyada **Özellikler →
> "Engellemeyi Kaldır (Unblock)"** gerekebilir.

## Kullanım

- **`BasakDoldur_Dun`** → dünü işler (butona bağlamak için ideal).
- **`BasakDoldur_Sor`** → tarih sorar (`yyyy-aa-gg`), o günü işler.

Çalışınca: ürünleri ve siparişleri çeker, gün penceresine (yerel 00:00–24:00,
UTC+3 ile API'ye uyarlanır) filtreler, hücreleri hesaplar ve `HAZİRAN (N)`
sayfasına yazıp özet kutusu gösterir.

## Yazılan hücreler (yalnızca bunlar)

`L7, L11, L17, L19` (CİRO adet) · `E16, E18, E20` (menü adet) ·
`R9/Q10, R11/Q12, R13/Q14, R15/Q16` (ödeme net/indirim).
Modüldeki `ALLOWED` sabiti bunun dışına yazmayı engeller; sadece `.Value` atanır,
hücre biçimi/formülleri korunur.

## Her gece otomatik (isteğe bağlı)

VBA bir butona basınca çalışır. "Her gece kendiliğinden" için Excel'in açılması
gerekir. İki yol:

1. **Basit:** Sabah dosyayı açıp butona basın (10 saniye).
2. **Tam otomatik:** Windows Görev Zamanlayıcı her gece Excel'i `/e` ile açar,
   `Workbook_Open` olayı `BasakDoldur_Dun`'u çağırır, kaydeder, kapatır. Bu yöntem
   biraz kırılgandır; tam otomasyon istiyorsanız `basak_adisyo_bot` Python sürümü
   + `register_scheduler.bat` daha sağlamdır (aynı sonucu üretir).

`Workbook_Open` için `ThisWorkbook` modülüne:

```vba
Private Sub Workbook_Open()
    ' Sadece zamanlayicidan acildiginda calissin istersen bir bayrak/saat kontrolu ekleyin
    On Error Resume Next
    BasakDoldur_Dun
    ThisWorkbook.Save
End Sub
```

## Notlar

- JSON ayrıştırma modülün içinde gömülüdür (harici kütüphane yok); algoritması
  Python referansıyla doğrulanmıştır.
- Türkçe karakterler için yanıt **UTF-8** olarak çözülür (ADODB.Stream).
- Rate limit: sayfalar arası 42 sn bekler; 429'da üstel backoff (2→64 sn).
- Sayfa bulunamazsa uyarı verir, **çökmeden** çıkar.
