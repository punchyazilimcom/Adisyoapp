"""CLI giriş noktası: akış, zaman penceresi, rich log.

Çalıştırma:
    python -m src.main --branch all
    python -m src.main --date 2026-06-06 --branch Demetevler
    python -m src.main --dry-run            (Excel'e yazmadan hesaplananı logla)
"""

from __future__ import annotations

import argparse
import logging
import sys
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any
from zoneinfo import ZoneInfo

import yaml
from rich.console import Console
from rich.logging import RichHandler
from rich.table import Table

from .adisyo_client import AdisyoApiError, AdisyoClient
from .categorize import build_product_index
from .compute import PAYMENT_CELLS, compute_cells
from .excel_writer import sheet_name_for, write_day_cells

UTC = ZoneInfo("UTC")
log = logging.getLogger("basak.main")
console = Console()

# Hücre → insan-okur açıklama (dry-run / log tablosu için)
CELL_LABELS: dict[str, str] = {
    "L7": "KÜÇÜK AYRAN adet (tüm)",
    "L11": "ÇAY adet (tüm)",
    "L17": "KIR PİDELERİ adet, Kuşbaşılı hariç (paket)",
    "L19": "KUTU İÇECEKLER adet (paket)",
    "E16": "KURYE MENÜ adet (tüm)",
    "E18": "KUTU İÇECEK MENÜ adet (tüm)",
    "E20": "BÜYÜK AYRAN MENÜ adet (tüm)",
    "R9": "YS Online net", "Q10": "YS Online indirim",
    "R11": "Trendyol Online net", "Q12": "Trendyol Online indirim",
    "R13": "Getir Online net", "Q14": "Getir Online indirim",
    "R15": "Migros Online net", "Q16": "Migros Online indirim",
}

# Tablo gösterim sırası
CELL_ORDER = [
    "L7", "L11", "L17", "L19",
    "E16", "E18", "E20",
    "R9", "Q10", "R11", "Q12", "R13", "Q14", "R15", "Q16",
]


# --------------------------------------------------------------------------- #
# Yapılandırma
# --------------------------------------------------------------------------- #
@dataclass
class Settings:
    timezone: str
    day_start_hour: int
    page_delay_seconds: float
    cache_dir: Path
    log_dir: Path


def load_config(path: Path) -> tuple[dict[str, Any], list[dict[str, Any]], Settings]:
    if not path.exists():
        console.print(
            f"[red]config.yaml bulunamadı:[/red] {path}\n"
            "config.example.yaml dosyasını kopyalayıp doldurun."
        )
        sys.exit(2)
    cfg = yaml.safe_load(path.read_text(encoding="utf-8"))
    s = cfg.get("settings", {})
    settings = Settings(
        timezone=s.get("timezone", "Europe/Istanbul"),
        day_start_hour=int(s.get("day_start_hour", 0)),
        page_delay_seconds=float(s.get("page_delay_seconds", 42)),
        cache_dir=Path(s.get("cache_dir", "./cache")),
        log_dir=Path(s.get("log_dir", "./logs")),
    )
    return cfg["api"], cfg["restaurants"], settings


# --------------------------------------------------------------------------- #
# Zaman penceresi
# --------------------------------------------------------------------------- #
def day_window(
    process_date: date, tz_name: str, day_start_hour: int
) -> tuple[datetime, datetime, str]:
    """Yerel gün penceresi [start, start+24h) ve UTC startDate string'i.

    Dönüş: (yerel_start, yerel_end, startDate_utc_str)
    """
    tz = ZoneInfo(tz_name)
    local_start = datetime(
        process_date.year, process_date.month, process_date.day,
        day_start_hour, 0, 0, tzinfo=tz,
    )
    local_end = local_start + timedelta(days=1)
    start_utc = local_start.astimezone(UTC)
    return local_start, local_end, start_utc.strftime("%Y-%m-%d %H:%M:%S")


def _parse_insert_date(raw: str) -> datetime:
    """insertDate'i UTC-aware datetime'a çevir (naive → UTC kabul edilir)."""
    s = (raw or "").strip()
    if s.endswith("Z"):
        s = s[:-1] + "+00:00"
    try:
        dt = datetime.fromisoformat(s)
    except ValueError:
        # Boşlukla ayrılmış format ("YYYY-MM-DD HH:MM:SS")
        dt = datetime.strptime(s[:19], "%Y-%m-%d %H:%M:%S")
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=UTC)
    return dt


def filter_orders_in_window(
    orders: list[dict[str, Any]], local_start: datetime, local_end: datetime
) -> list[dict[str, Any]]:
    """insertDate'i [local_start, local_end) penceresine düşen siparişler."""
    start_utc = local_start.astimezone(UTC)
    end_utc = local_end.astimezone(UTC)
    kept: list[dict[str, Any]] = []
    for order in orders:
        raw = order.get("insertDate")
        if not raw:
            continue
        dt = _parse_insert_date(raw).astimezone(UTC)
        if start_utc <= dt < end_utc:
            kept.append(order)
    return kept


# --------------------------------------------------------------------------- #
# Loglama
# --------------------------------------------------------------------------- #
def setup_logging(log_dir: Path, process_date: date) -> None:
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / f"basak_{process_date.isoformat()}.log"

    root = logging.getLogger("basak")
    root.setLevel(logging.INFO)
    root.handlers.clear()

    rich_handler = RichHandler(
        console=console, show_path=False, rich_tracebacks=True, markup=True
    )
    rich_handler.setFormatter(logging.Formatter("%(message)s", datefmt="%H:%M:%S"))

    file_handler = logging.FileHandler(log_file, encoding="utf-8")
    file_handler.setFormatter(
        logging.Formatter("%(asctime)s %(levelname)s [%(name)s] %(message)s")
    )

    root.addHandler(rich_handler)
    root.addHandler(file_handler)


def render_result_table(branch: str, sheet: str, cells: dict[str, Any], n_orders: int) -> None:
    table = Table(title=f"{branch} — {sheet} ({n_orders} sipariş)", show_lines=False)
    table.add_column("Hücre", style="cyan", no_wrap=True)
    table.add_column("Açıklama")
    table.add_column("Değer", justify="right", style="green")
    for cell in CELL_ORDER:
        table.add_row(cell, CELL_LABELS.get(cell, ""), str(cells.get(cell)))
    console.print(table)


# --------------------------------------------------------------------------- #
# Şube işleme
# --------------------------------------------------------------------------- #
def process_branch(
    restaurant: dict[str, Any],
    api: dict[str, Any],
    settings: Settings,
    process_date: date,
    dry_run: bool,
) -> bool:
    name = restaurant["name"]
    console.rule(f"[bold]{name}[/bold] — {process_date.isoformat()}")

    local_start, local_end, start_utc_str = day_window(
        process_date, settings.timezone, settings.day_start_hour
    )
    log.info(
        "Pencere (yerel): %s → %s | startDate (UTC): %s",
        local_start.strftime("%Y-%m-%d %H:%M"),
        local_end.strftime("%Y-%m-%d %H:%M"),
        start_utc_str,
    )

    client = AdisyoClient(
        base_url=api["base_url"],
        x_api_key=api["x_api_key"],
        x_api_secret=restaurant["x_api_secret"],
        x_api_consumer=restaurant["x_api_consumer"],
        page_delay_seconds=settings.page_delay_seconds,
        cache_dir=settings.cache_dir,
    )
    cache_key = f"{name}_{process_date.isoformat()}"

    try:
        products = client.fetch_products(cache_key)
        idx = build_product_index(products)

        raw_orders = client.fetch_completed_orders(start_utc_str, cache_key)
        log.info("Ham sipariş: %s", len(raw_orders))

        orders = filter_orders_in_window(raw_orders, local_start, local_end)
        log.info("Pencereye düşen sipariş: %s", len(orders))

        if not orders:
            log.warning("Bu gün için pencerede sipariş yok — hücreler 0/boş kalır.")

        cells = compute_cells(orders, idx)
    except AdisyoApiError as exc:
        log.error("API hatası (%s) — şube atlanıyor: %s", name, exc)
        return False
    except Exception as exc:  # noqa: BLE001 — şube bazlı izolasyon, akış devam etmeli
        log.exception("Beklenmeyen hata (%s) — şube atlanıyor: %s", name, exc)
        return False

    sheet = sheet_name_for(process_date.month, process_date.day)
    render_result_table(name, sheet, cells, len(orders))

    ok = write_day_cells(
        restaurant["excel_path"],
        process_date.month,
        process_date.day,
        cells,
        dry_run=dry_run,
    )
    return ok


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #
def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        prog="basak_adisyo_bot",
        description="Adisyo → Başak günlük Excel şablonu doldurucu bot.",
    )
    parser.add_argument(
        "--date",
        help="İşlenecek tarih (YYYY-MM-DD). Verilmezse DÜN işlenir.",
    )
    parser.add_argument(
        "--branch",
        default="all",
        help="Şube adı veya 'all' (varsayılan: all).",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Excel'e yazmadan hesaplananı logla.",
    )
    parser.add_argument(
        "--config",
        default=str(Path(__file__).resolve().parent.parent / "config.yaml"),
        help="config.yaml yolu.",
    )
    return parser.parse_args(argv)


def resolve_process_date(arg_date: str | None, tz_name: str) -> date:
    if arg_date:
        return datetime.strptime(arg_date, "%Y-%m-%d").date()
    # Varsayılan: yerel saate göre DÜN
    today_local = datetime.now(ZoneInfo(tz_name)).date()
    return today_local - timedelta(days=1)


def select_restaurants(
    restaurants: list[dict[str, Any]], branch: str
) -> list[dict[str, Any]]:
    if branch.lower() == "all":
        return restaurants
    selected = [r for r in restaurants if r["name"].lower() == branch.lower()]
    if not selected:
        names = ", ".join(r["name"] for r in restaurants)
        console.print(f"[red]Şube bulunamadı:[/red] {branch}. Mevcut: {names}")
        sys.exit(2)
    return selected


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    api, restaurants, settings = load_config(Path(args.config))
    process_date = resolve_process_date(args.date, settings.timezone)

    setup_logging(settings.log_dir, process_date)
    log.info(
        "Başak Adisyo Bot başlıyor | tarih=%s | şube=%s | dry_run=%s",
        process_date.isoformat(),
        args.branch,
        args.dry_run,
    )

    targets = select_restaurants(restaurants, args.branch)
    results: dict[str, bool] = {}
    for restaurant in targets:
        results[restaurant["name"]] = process_branch(
            restaurant, api, settings, process_date, args.dry_run
        )

    console.rule("[bold]ÖZET[/bold]")
    for name, ok in results.items():
        status = "[green]TAMAM[/green]" if ok else "[red]ATLANDI/HATA[/red]"
        console.print(f"  {name}: {status}")

    # Tümü başarılıysa 0, en az biri başarısızsa 1
    return 0 if all(results.values()) else 1


if __name__ == "__main__":
    raise SystemExit(main())
