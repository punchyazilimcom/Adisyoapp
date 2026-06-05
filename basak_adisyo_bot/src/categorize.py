"""Adisyo /Products yanıtından kategori→ürün haritası ve sınıflama kuralları.

Kategori adları sabit yazılmaz; harita /Products'tan DİNAMİK kurulur.
Kanal kopyaları (MigrosHemen, DeliveryHero vb.) farklı kategori adı altında
aynı ürünü içerebilir → ürünleri kategori ADIYLA değil, isimden türetilen
NORMALİZE küme üyeliğiyle eşleriz.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from typing import Any

log = logging.getLogger("basak.categorize")

# Türkçe karakter → ASCII katlama (büyük/küçük fark etmeksizin)
_TR_MAP = str.maketrans(
    {
        "İ": "i",
        "I": "i",
        "ı": "i",
        "Ş": "s",
        "ş": "s",
        "Ğ": "g",
        "ğ": "g",
        "Ü": "u",
        "ü": "u",
        "Ö": "o",
        "ö": "o",
        "Ç": "c",
        "ç": "c",
    }
)


def normalize(text: str | None) -> str:
    """Türkçe duyarlı normalizasyon: ASCII'ye katla, küçült, boşlukları sadeleştir."""
    if not text:
        return ""
    folded = text.translate(_TR_MAP).lower()
    return " ".join(folded.split())


def _classify_category(norm_category: str) -> str:
    """Normalize edilmiş kategori adından kovayı belirle.

    Sıra önemli: 'kutu icecek' kontrolü düz 'icecek'ten önce yapılır.
    """
    if "kir pide" in norm_category:
        return "kir"
    if "kutu icecek" in norm_category:
        return "kutu"
    if "menu" in norm_category:
        return "menu"
    if "icecek" in norm_category:
        return "icecek"
    return "other"


@dataclass
class ProductIndex:
    """Normalize edilmiş ürün adı kümeleri (kanonik kategorilerden türetilir)."""

    kir_pideleri: set[str] = field(default_factory=set)
    kutu_icecekler: set[str] = field(default_factory=set)
    icecekler: set[str] = field(default_factory=set)
    menuler: set[str] = field(default_factory=set)
    # norm ad → ilk görülen (orijinal) kategori adı (teşhis/log için)
    product_category: dict[str, str] = field(default_factory=dict)

    def is_kir_pidesi(self, name: str) -> bool:
        return normalize(name) in self.kir_pideleri

    def is_kutu_icecek(self, name: str) -> bool:
        return normalize(name) in self.kutu_icecekler


def build_product_index(products_data: list[dict[str, Any]]) -> ProductIndex:
    """/Products ``data`` listesinden ProductIndex kur."""
    idx = ProductIndex()
    for cat in products_data:
        cat_name = cat.get("categoryName", "") or ""
        bucket = _classify_category(normalize(cat_name))
        for prod in cat.get("products") or []:
            norm_name = normalize(prod.get("productName", ""))
            if not norm_name:
                continue
            if bucket == "kir":
                idx.kir_pideleri.add(norm_name)
            elif bucket == "kutu":
                idx.kutu_icecekler.add(norm_name)
            elif bucket == "icecek":
                idx.icecekler.add(norm_name)
            elif bucket == "menu":
                idx.menuler.add(norm_name)
            idx.product_category.setdefault(norm_name, cat_name)

    log.info(
        "Ürün haritası kuruldu: %s kır pidesi, %s kutu içecek, %s içecek, %s menü ürünü",
        len(idx.kir_pideleri),
        len(idx.kutu_icecekler),
        len(idx.icecekler),
        len(idx.menuler),
    )
    return idx
