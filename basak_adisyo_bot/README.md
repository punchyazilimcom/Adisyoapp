# Başak Adisyo Bot

Başak Kır Pidesi'nin **GÜNLÜK HESAP** Excel şablonlarını (`.xlsm`), Adisyo POS
resmî API'sinden çektiği veriyle **her gece otomatik** dolduran bir Python botu.

> ⚠️ **Yeni rapor üretmez.** Var olan şablonun yalnızca belirli hücrelerine
> yazar; başka hiçbir hücreye/sayfaya/makroya dokunmaz. (`keep_vba=True`)

---

## Ne yapar?

1. İşlenecek tarih (varsayılan: **dün**) için yerel gün penceresini hesaplar
   (`day_start_hour`'dan +24 saat) ve UTC `startDate`'e çevirir.
2. Her şube için Adisyo `/CompletedOrders` uç noktasını **sayfalayarak** çeker
   (sayfalar arası 42 sn bekler, 429'da üstel backoff), ham sayfaları `cache/`'e
   yazar (tekrar çalıştırmada hızlı resume).
3. `insertDate`'i yerele çevirip gün penceresine düşmeyen siparişleri eler.
4. `/Products`'tan **dinamik** kategori→ürün haritası kurar.
5. Aşağıdaki kurallara göre adetleri/tutarları hesaplar.
6. Şablonu `openpyxl(keep_vba=True)` ile açar, `HAZİRAN (N)` sayfasında
   **yalnızca izinli hücrelere** yazar ve kaydeder.

---

## Yazılan hücreler (SADECE bunlar)

| Hücre | Anlam | Filtre |
|-------|-------|--------|
| `L7`  | KÜÇÜK AYRAN toplam adet | tüm tipler |
| `L11` | ÇAY toplam adet | tüm tipler |
| `L17` | KIR PİDELERİ adedi (Kuşbaşılı içerenler hariç) | sadece paket |
| `L19` | KUTU İÇECEKLER adedi (kanal kopyaları dahil) | sadece paket |
| `E16` | KURYE MENÜ adet | tüm tipler |
| `E18` | KUTU İÇECEK MENÜ adet | tüm tipler |
| `E20` | BÜYÜK AYRAN MENÜ adet | tüm tipler |
| `R9`  / `Q10` | YS Online net / indirim | — |
| `R11` / `Q12` | Trendyol Online net / indirim | — |
| `R13` / `Q14` | Getir Online net / indirim | — |
| `R15` / `Q16` | Migros Online net / indirim | — |

Bu liste `src/excel_writer.py` içindeki **`ALLOWED_CELLS`** sabitiyle korunur;
bot bu kümenin dışına **tek bir hücreye bile** yazamaz. Yalnızca `.value` atanır,
hücre biçimi/formatı korunur. (Stok/formül/personel/masraf hücrelerine dokunulmaz.)

- **Net tahsilat** = `Σ payments.amount` (indirim hariç) — `orderTotal` ile eşit.
- **İndirim** = ilgili ödeme adıyla kapanan siparişlerin `discountAmount` toplamı (platform indirimi).

---

## Kurulum (Windows)

```bat
cd basak_adisyo_bot

REM 1) Sanal ortam
python -m venv venv
call venv\Scripts\activate.bat

REM 2) Bağımlılıklar
pip install -r requirements.txt

REM 3) Yapılandırma
copy config.example.yaml config.yaml
REM config.yaml'ı açıp anahtarları ve excel_path'leri doldurun.
```

> `config.yaml` canlı API anahtarları içerir ve `.gitignore`'dadır — **commit etmeyin.**

Linux/Mac için: `python3 -m venv venv && source venv/bin/activate && pip install -r requirements.txt`

---

## Çalıştırma (CLI)

```bash
# Dünü, tüm şubeler için işle (varsayılan)
python -m src.main --branch all

# Belirli tarih + belirli şube
python -m src.main --date 2026-06-06 --branch Demetevler

# Excel'e yazmadan sadece hesaplananı göster (test)
python -m src.main --dry-run --date 2026-06-06 --branch all
```

| Argüman | Açıklama |
|---------|----------|
| `--date YYYY-MM-DD` | İşlenecek tarih. Verilmezse **dün**. |
| `--branch all\|<Ad>` | Şube seçimi (`Demetevler`, `Bahcelievler`, `Etlik`, `Batikent`). |
| `--dry-run` | Excel'e yazmadan hesaplananı loglar. |
| `--config <yol>` | Alternatif `config.yaml` yolu. |

Hedef sayfa, işlenen tarihin ay+gününe göre seçilir: 6 Haziran → **`HAZİRAN (6)`**.
Sayfa yoksa loglanır, o şube atlanır; **bot çökmez.**

---

## Otomasyon (Windows Görev Zamanlayıcı)

`run_daily.bat` → venv'i aktive eder, `python -m src.main --branch all` çalıştırır,
çıktıyı `logs\run_YYYY-MM-DD.log` dosyasına yazar.

Görevi her gece **02:00**'de kuran komut (yönetici cmd):

```bat
register_scheduler.bat
```

Bu, `schtasks` ile `BasakAdisyoBot` görevini oluşturur (dünü işler).

```bat
REM Hemen test et
schtasks /Run /TN "BasakAdisyoBot"

REM Görevi görüntüle
schtasks /Query /TN "BasakAdisyoBot"

REM Görevi İPTAL et / kaldır
schtasks /Delete /TN "BasakAdisyoBot" /F
```

---

## Tek dosya .exe (PyInstaller)

```bat
call venv\Scripts\activate.bat
pip install pyinstaller
pyinstaller --onefile --name BasakAdisyoBot src/main.py
```

Çıktı: `dist\BasakAdisyoBot.exe`. Çalıştırma:

```bat
dist\BasakAdisyoBot.exe --branch all
dist\BasakAdisyoBot.exe --dry-run --date 2026-06-06
```

> `.exe`, `config.yaml`'ı çalışma dizininde arar. `--config "C:\yol\config.yaml"`
> ile açıkça belirtebilirsiniz. `.exe` kullanıyorsanız `run_daily.bat` içindeki
> `python -m src.main` satırını `BasakAdisyoBot.exe` ile değiştirin.

---

## Yapılandırma (`config.yaml`)

```yaml
api:
  base_url: "https://ext.adisyo.com/api/External/v2"
  x_api_key: "<ortak anahtar>"          # tüm şubelerde aynı
restaurants:
  - name: "Demetevler"
    x_api_secret: "<şube secret>"        # şube bazlı
    x_api_consumer: "Punch - Demetevler" # şube bazlı
    excel_path: "C:\\BasakHesap\\DEMETEVLER_HAZİRAN_HESAP.xlsm"
  # ... diğer şubeler
settings:
  timezone: "Europe/Istanbul"   # API UTC döner; +3 ile yerele çevrilir
  day_start_hour: 0             # 0 = gece yarısı; 7 yaparsan 07:00–07:00 penceresi
  page_delay_seconds: 42        # sayfalar arası bekleme (rate limit 40 sn/çağrı)
  cache_dir: "./cache"
  log_dir: "./logs"
```

---

## Proje yapısı

```
basak_adisyo_bot/
├─ config.yaml              # DOLU (gizli — .gitignore'da)
├─ config.example.yaml      # şablon
├─ requirements.txt
├─ README.md
├─ .gitignore
├─ run_daily.bat            # günlük çalıştırıcı (loglar logs/'a)
├─ register_scheduler.bat   # schtasks ile 02:00 görevi
└─ src/
   ├─ __init__.py
   ├─ adisyo_client.py      # header auth, pagination, rate limit, retry, cache
   ├─ categorize.py         # /Products'tan kategori→ürün haritası + sınıflama
   ├─ compute.py            # hücre değerlerini üreten saf fonksiyonlar (kurallar)
   ├─ excel_writer.py       # şablona izinli hücreleri yazar (ALLOWED_CELLS guard, keep_vba)
   └─ main.py               # CLI, akış, rich log
```

---

## Notlar / davranış

- **Rate limit:** `/CompletedOrders` 40 sn/çağrı (sayfalar arası 42 sn beklenir),
  `/Products` 3 dk/çağrı (günde bir kez çağrılır, cache'lenir). 429'da tenacity
  üstel backoff (2→4→8…120 sn, 6 deneme).
- **Cache:** Ham sayfalar `cache/<Şube>_<tarih>_orders_p<N>.json` ve ürünler
  `cache/<Şube>_<tarih>_products.json` olarak saklanır. Aynı gün tekrar
  çalıştırırsanız API'ye gitmeden cache'ten okunur (resume). Yeniden çekmek için
  ilgili cache dosyalarını silin.
- **Kategori eşleme:** İsimler sabit yazılmaz; harita `/Products`'tan kurulur.
  Kanal kopyaları (MigrosHemen/DeliveryHero) farklı kategori altında olsa bile
  ürün **adı** kanonik kümede ise (ör. kutu içecek) doğru sayılır.
- **Hata yönetimi:** `status != 100`, boş yanıt veya beklenmeyen hata → anlamlı
  Türkçe log, ilgili şube atlanır, diğer şubeler işlenmeye devam eder (çökme yok).
- **Çıkış kodu:** Tüm şubeler başarılıysa `0`, en az biri atlandı/hatalıysa `1`.
```
