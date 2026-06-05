"""Hücre değerlerini üreten saf (yan etkisiz) hesap fonksiyonları.

Tüm iş kuralları burada toplanır. Girdi: gün penceresine filtrelenmiş sipariş
listesi + ProductIndex. Çıktı: {hücre: değer} sözlüğü.

Sipariş tipi:
  "Paket Siparişi"  → paket
  "Gel-Al Satış", "Masa Siparişi" → diğer
"""

from __future__ import annotations

from typing import Any, Callable, Iterator

from .categorize import ProductIndex, normalize

PAKET_TYPE = "Paket Siparişi"

# Ödeme adı → (net tutar hücresi, indirim hücresi)
PAYMENT_CELLS: list[tuple[str, str, str]] = [
    ("YS Online", "R9", "Q10"),
    ("Trendyol Online", "R11", "Q12"),
    ("Getir Online", "R13", "Q14"),
    ("Migros Online", "R15", "Q16"),
]

OrderPredicate = Callable[[dict[str, Any]], bool]
NameMatcher = Callable[[str], bool]


def _is_paket(order: dict[str, Any]) -> bool:
    return (order.get("orderType") or "") == PAKET_TYPE


def _as_number(value: float) -> float | int:
    """Tam sayıysa int döndür (adet hücreleri için), değilse 2 hane yuvarla."""
    if float(value).is_integer():
        return int(value)
    return round(value, 2)


def _iter_products(
    orders: list[dict[str, Any]], predicate: OrderPredicate | None = None
) -> Iterator[tuple[dict[str, Any], dict[str, Any]]]:
    for order in orders:
        if predicate and not predicate(order):
            continue
        for product in order.get("products") or []:
            yield order, product


def _sum_qty(
    orders: list[dict[str, Any]],
    match: NameMatcher,
    predicate: OrderPredicate | None = None,
) -> float | int:
    total = 0.0
    for _order, product in _iter_products(orders, predicate):
        if match(product.get("productName", "")):
            total += float(product.get("quantity") or 0)
    return _as_number(total)


# --------------------------------------------------------------------------- #
# CİRO bloğu — ADET (L sütunu)
# --------------------------------------------------------------------------- #
def qty_kucuk_ayran(orders: list[dict[str, Any]], idx: ProductIndex) -> float | int:
    """L7 — KÜÇÜK AYRAN toplam adet (tüm tipler)."""
    return _sum_qty(orders, lambda n: "kucuk ayran" in normalize(n))


def qty_cay(orders: list[dict[str, Any]], idx: ProductIndex) -> float | int:
    """L11 — ÇAY toplam adet (tüm tipler). Menü/ayran adlarını dışlamak için tam eşleşme."""
    return _sum_qty(orders, lambda n: normalize(n) == "cay")


def qty_kir_pideleri_no_kusbasi(
    orders: list[dict[str, Any]], idx: ProductIndex
) -> float | int:
    """L17 — KIR PİDELERİ adedi, "Kuşbaşılı" içerenler HARİÇ (sadece paket)."""

    def match(name: str) -> bool:
        return idx.is_kir_pidesi(name) and "kusbasi" not in normalize(name)

    return _sum_qty(orders, match, predicate=_is_paket)


def qty_kutu_icecekler(
    orders: list[dict[str, Any]], idx: ProductIndex
) -> float | int:
    """L19 — KUTU İÇECEKLER adedi, kanal kopyaları dahil (sadece paket)."""
    return _sum_qty(orders, idx.is_kutu_icecek, predicate=_is_paket)


# --------------------------------------------------------------------------- #
# MENÜLER — ADET (E sütunu, tüm tipler)
# --------------------------------------------------------------------------- #
def qty_kurye_menu(orders: list[dict[str, Any]], idx: ProductIndex) -> float | int:
    """E16 — KURYE MENÜ adet."""
    return _sum_qty(orders, lambda n: "kurye menu" in normalize(n))


def qty_kutu_icecek_menu(
    orders: list[dict[str, Any]], idx: ProductIndex
) -> float | int:
    """E18 — KUTU İÇECEK MENÜ adet."""
    return _sum_qty(orders, lambda n: "kutu icecek menu" in normalize(n))


def qty_buyuk_ayran_menu(
    orders: list[dict[str, Any]], idx: ProductIndex
) -> float | int:
    """E20 — BÜYÜK AYRAN MENÜ adet."""
    return _sum_qty(orders, lambda n: "buyuk ayran menu" in normalize(n))


# --------------------------------------------------------------------------- #
# ÖDEME bloğu — TUTAR
# --------------------------------------------------------------------------- #
def payment_net(orders: list[dict[str, Any]], payment_name: str) -> float:
    """Σ payments.amount (paymentName eşleşen) — NET tahsilat (indirim hariç)."""
    total = 0.0
    for order in orders:
        for payment in order.get("payments") or []:
            if (payment.get("paymentName") or "") == payment_name:
                total += float(payment.get("amount") or 0)
    return round(total, 2)


def payment_discount(orders: list[dict[str, Any]], payment_name: str) -> float:
    """Σ order.discountAmount — o ödemeyle kapanan siparişler (platform indirimi)."""
    total = 0.0
    for order in orders:
        names = {
            (payment.get("paymentName") or "")
            for payment in order.get("payments") or []
        }
        if payment_name in names:
            total += float(order.get("discountAmount") or 0)
    return round(total, 2)


# --------------------------------------------------------------------------- #
# Toplu hesap
# --------------------------------------------------------------------------- #
def compute_cells(
    orders: list[dict[str, Any]], idx: ProductIndex
) -> dict[str, float | int]:
    """İzinli tüm hücreler için değer üret."""
    cells: dict[str, float | int] = {
        "L7": qty_kucuk_ayran(orders, idx),
        "L11": qty_cay(orders, idx),
        "L17": qty_kir_pideleri_no_kusbasi(orders, idx),
        "L19": qty_kutu_icecekler(orders, idx),
        "E16": qty_kurye_menu(orders, idx),
        "E18": qty_kutu_icecek_menu(orders, idx),
        "E20": qty_buyuk_ayran_menu(orders, idx),
    }
    for payment_name, net_cell, disc_cell in PAYMENT_CELLS:
        cells[net_cell] = payment_net(orders, payment_name)
        cells[disc_cell] = payment_discount(orders, payment_name)
    return cells
