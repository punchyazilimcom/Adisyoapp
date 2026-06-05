---
name: iskelet-kur
description: Adisyoapp için seçilen teknoloji yığınına göre başlangıç proje iskeletini oluşturur (Node/React, .NET, veya Flutter). Kullanıcı "projeyi kur", "iskeleti oluştur", "scaffold" dediğinde kullan.
---

# Adisyoapp — Proje İskeleti Kurulumu

Amaç: Boş depoyu seçilen yığına göre çalışır bir başlangıç projesine dönüştürmek.

## Adımlar

1. **Yığını teyit et:** Kullanıcıya hangi yığın olduğunu sor (henüz `CLAUDE.md`'de
   belirtilmemişse): Node+React/Next, .NET (ASP.NET Core), veya Flutter.
2. **İskeleti oluştur** (yığına göre tipik komutlar):
   - **Node + Next.js:** `npx create-next-app@latest . --ts`
   - **.NET API:** `dotnet new webapi -o src/Adisyoapp.Api`
   - **Flutter:** `flutter create .`
3. **Temel POS modüllerini planla:** masa yönetimi, menü/ürün, sipariş/adisyon,
   ödeme, raporlama. Klasör yapısını buna göre oluştur.
4. **CLAUDE.md güncelle:** Teknoloji yığını ve komut tablosunu doldur.
5. **Hook'u doğrula:** `CLAUDE_CODE_REMOTE=true ./.claude/hooks/session-start.sh`
   çalıştırarak bağımlılıkların kurulduğunu teyit et.
6. **İlk commit'i yap** ve dala push'la.

## Notlar
- İskeleti depo kökünde oluştururken mevcut `.claude/`, `CLAUDE.md`, `README.md`
  dosyalarını ezme.
- Kurulumdan sonra basit bir "çalışıyor" testi ekle.
