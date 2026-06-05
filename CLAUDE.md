# Adisyoapp

Restoran/kafe için **adisyon & POS (sipariş yönetim)** uygulaması.

> Bu depo şu an bir iskelet aşamasındadır. Teknoloji yığını netleştikçe bu
> dosyayı güncelleyin; Claude her oturumda bu dosyayı okur.

## Proje Amacı

Adisyoapp; masaların açılması, sipariş alma, ürün/menü yönetimi, adisyon
birleştirme/bölme, ödeme alma ve raporlama gibi temel POS akışlarını sağlamayı
hedefler.

## Teknoloji Yığını

> **TODO:** Yığın seçildiğinde doldurun. Aday seçenekler:
> - Web: Node.js + React/Next.js + TypeScript
> - Backend: .NET (ASP.NET Core) veya Node.js
> - Mobil/Masaüstü: Flutter
>
> Yığını belirledikten sonra `.claude/hooks/session-start.sh` zaten otomatik
> algılayıp bağımlılıkları kuracaktır.

## Geliştirme Ortamı

Claude Code (web) oturumu başladığında `.claude/hooks/session-start.sh`
çalışır ve mevcut manifest dosyalarına göre bağımlılıkları kurar
(npm / dotnet / flutter / pip).

### Sık Kullanılan Komutlar

Yığına göre tipik komutlar (manifest eklendiğinde geçerli olur):

| İş        | Node            | .NET            | Flutter           |
|-----------|-----------------|-----------------|-------------------|
| Kurulum   | `npm install`   | `dotnet restore`| `flutter pub get` |
| Çalıştır  | `npm run dev`   | `dotnet run`    | `flutter run`     |
| Test      | `npm test`      | `dotnet test`   | `flutter test`    |
| Lint      | `npm run lint`  | `dotnet format` | `flutter analyze` |
| Build     | `npm run build` | `dotnet build`  | `flutter build`   |

## Çalışma Kuralları (Claude için)

- **Dil:** Kod yorumları ve commit mesajları Türkçe veya İngilizce olabilir;
  mevcut dosyaların stiline uyun. Kullanıcıyla iletişim Türkçe.
- **Dal (branch):** Doğrudan `main`'e push yapmayın; özellik dalları kullanın.
- **Sırlar:** `.env`, anahtarlar veya kimlik bilgileri asla commit edilmez.
  POS uygulaması olduğu için ödeme/entegrasyon anahtarlarına dikkat edin.
- **Para/hesap mantığı:** Tutar hesaplarında kuruş hatalarından kaçınmak için
  ondalık (decimal) tipler kullanın; float ile para tutmayın.
- **Test:** Davranış değiştiren her değişiklik için test ekleyin/güncelleyin.
- **Küçük commit'ler:** Net ve açıklayıcı commit mesajları yazın.

## Dizin Yapısı

> Yığın eklendikçe güncelleyin.

```
.
├── .claude/
│   ├── settings.json          # izinler + SessionStart hook kaydı
│   ├── hooks/
│   │   └── session-start.sh   # bağımlılık kurulumu (stack auto-detect)
│   └── skills/                # projeye özel slash komutları
├── CLAUDE.md
└── README.md
```
