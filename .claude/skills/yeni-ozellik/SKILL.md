---
name: yeni-ozellik
description: Adisyoapp'e yeni bir özellik eklemek için uçtan uca iş akışı (özellik dalı aç, küçük adımlarla geliştir, test ekle, commit'le). Kullanıcı "yeni özellik ekle", "şunu geliştir", "new feature" dediğinde kullan.
---

# Adisyoapp — Yeni Özellik İş Akışı

Amaç: Yeni bir özelliği güvenli ve test edilebilir biçimde eklemek.

## Adımlar

1. **Anla:** Özelliğin amacını ve kabul kriterlerini netleştir. Belirsizse kullanıcıya sor.
2. **Dal aç:** `main` üzerinde çalışma. Mevcut çalışma dalında kal veya
   `feature/<kısa-ad>` biçiminde yeni bir dal aç.
3. **Tasarla:** Etkilenecek dosyaları ve veri modelini belirle. POS akışlarında
   (masa, sipariş, adisyon, ödeme) tutarlılığı koru.
4. **Küçük adımlarla geliştir:** Mevcut kod stiline uy. Para tutarlarında
   ondalık (decimal) tip kullan, float kullanma.
5. **Test ekle:** Davranışı doğrulayan test(ler) yaz ve `/test` ile çalıştır.
6. **Doğrula:** Lint/analiz komutunu çalıştır (varsa).
7. **Commit'le:** Açıklayıcı bir mesajla commit yap. Sırları (.env, anahtarlar)
   asla ekleme.
8. **Özetle:** Kullanıcıya yapılan değişiklikleri ve test sonuçlarını bildir.

## Kontrol Listesi
- [ ] Doğru dalda mıyım? (`main` değil)
- [ ] Test eklendi/güncellendi mi?
- [ ] Lint/analiz temiz mi?
- [ ] Sır/anahtar sızıntısı yok mu?
