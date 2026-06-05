---
name: test
description: Adisyoapp testlerini çalıştır. Yığını (Node / .NET / Flutter / Python) otomatik algılar ve uygun test komutunu çalıştırır. Kullanıcı "testleri çalıştır", "test et", "run tests" dediğinde kullan.
---

# Adisyoapp — Test Çalıştırma

Amaç: Projenin teknoloji yığınını algıla ve doğru test komutunu çalıştır.

## Adımlar

1. Depo kökünde manifest dosyalarını kontrol et ve yığını belirle:
   - `package.json`  → Node:  `npm test` (yoksa `npm run test`)
   - `*.sln` / `*.csproj` → .NET: `dotnet test`
   - `pubspec.yaml` → Flutter: `flutter test`
   - `pyproject.toml` / `requirements.txt` → Python: `pytest`
2. İlgili komutu çalıştır. Kullanıcı belirli bir test/dosya verdiyse sadece onu çalıştır.
3. Hata çıkarsa çıktıyı özetle ve olası nedeni belirt. Geçtiyse net biçimde "geçti" de.
4. Hiç test yoksa kullanıcıyı bilgilendir ve test eklemeyi öner.

## Notlar
- Tüm test takımını gereksiz yere çalıştırma; kullanıcı odak isterse dar kapsamda çalış.
- Para/tutar mantığı içeren testlerde ondalık (decimal) doğruluğuna özellikle dikkat et.
