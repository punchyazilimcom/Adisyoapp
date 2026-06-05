"""Adisyo POS resmî API istemcisi.

Doküman: https://developers.adisyo.com
Base: https://ext.adisyo.com/api/External/v2

Sorumluluklar:
- Header bazlı kimlik doğrulama (x-api-key / x-api-secret / x-api-consumer)
- /CompletedOrders sayfalama + rate limit (40 sn/çağrı) + 429 backoff
- /Products (rate limit 3 dk/çağrı) ile kategori→ürün haritası kaynağı
- Ham yanıtları cache'e yazma (resume / tekrar çalıştırma için)
"""

from __future__ import annotations

import json
import logging
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import requests
from tenacity import (
    retry,
    retry_if_exception_type,
    stop_after_attempt,
    wait_exponential,
)

log = logging.getLogger("basak.adisyo")

API_STATUS_OK = 100  # Adisyo: status == 100 ise istek başarılı


class RateLimitError(Exception):
    """HTTP 429 — istek limiti aşıldı; tenacity ile yeniden denenir."""


class AdisyoApiError(Exception):
    """status != 100 veya beklenmeyen API yanıtı."""


@dataclass
class AdisyoClient:
    """Tek bir şube (x-api-secret/x-api-consumer) için API istemcisi."""

    base_url: str
    x_api_key: str
    x_api_secret: str
    x_api_consumer: str
    page_delay_seconds: float = 42.0
    cache_dir: Path = field(default_factory=lambda: Path("./cache"))
    timeout: int = 60
    session: requests.Session = field(default_factory=requests.Session)

    def _headers(self) -> dict[str, str]:
        return {
            "x-api-key": self.x_api_key,
            "x-api-secret": self.x_api_secret,
            "x-api-consumer": self.x_api_consumer,
            "Content-Type": "application/json",
            "Accept": "application/json",
        }

    @retry(
        retry=retry_if_exception_type(
            (RateLimitError, requests.exceptions.RequestException)
        ),
        wait=wait_exponential(multiplier=2, min=2, max=120),
        stop=stop_after_attempt(6),
        reraise=True,
    )
    def _get(self, path: str, params: dict[str, Any] | None = None) -> dict[str, Any]:
        url = f"{self.base_url}{path}"
        resp = self.session.get(
            url, headers=self._headers(), params=params, timeout=self.timeout
        )
        if resp.status_code == 429:
            log.warning("429 rate limit — %s; backoff ile yeniden denenecek", path)
            raise RateLimitError(path)
        resp.raise_for_status()
        return resp.json()

    # ------------------------------------------------------------------ #
    # CompletedOrders
    # ------------------------------------------------------------------ #
    def fetch_completed_orders(
        self, start_date_utc: str, cache_key: str
    ) -> list[dict[str, Any]]:
        """Tüm sayfaları çek (startDate UTC, bitiş yok → baştan).

        Sayfalar arası rate limit için ``page_delay_seconds`` kadar beklenir.
        Cache'ten okunan sayfalar için bekleme yapılmaz.
        """
        all_orders: list[dict[str, Any]] = []
        page = 1
        while True:
            cache_file = self.cache_dir / f"{cache_key}_orders_p{page}.json"
            data = self._load_cache(cache_file)
            from_network = data is None
            if from_network:
                params = {
                    "page": page,
                    "startDate": start_date_utc,
                    "includeCancelled": "false",
                }
                data = self._get("/CompletedOrders", params)
                self._save_cache(cache_file, data)

            status = data.get("status")
            if status != API_STATUS_OK:
                raise AdisyoApiError(
                    f"CompletedOrders status={status} "
                    f"mesaj={data.get('message')!r} (sayfa {page})"
                )

            orders = data.get("orders") or []
            all_orders.extend(orders)
            page_count = int(data.get("pageCount") or 1)
            total_count = data.get("totalCount")
            log.info(
                "Sayfa %s/%s alındı — %s sipariş (toplam=%s)",
                page,
                page_count,
                len(orders),
                total_count,
            )

            if page >= page_count:
                break
            page += 1
            if from_network:
                log.info("Rate limit: %.0f sn bekleniyor...", self.page_delay_seconds)
                time.sleep(self.page_delay_seconds)

        return all_orders

    # ------------------------------------------------------------------ #
    # Products
    # ------------------------------------------------------------------ #
    def fetch_products(self, cache_key: str) -> list[dict[str, Any]]:
        """Kategori→ürün listesini çek (rate limit 3 dk/çağrı → günde 1 kez yeter)."""
        cache_file = self.cache_dir / f"{cache_key}_products.json"
        data = self._load_cache(cache_file)
        if data is None:
            data = self._get("/Products")
            self._save_cache(cache_file, data)

        if data.get("status") != API_STATUS_OK:
            raise AdisyoApiError(
                f"Products status={data.get('status')} mesaj={data.get('message')!r}"
            )
        return data.get("data") or []

    # ------------------------------------------------------------------ #
    # Cache yardımcıları
    # ------------------------------------------------------------------ #
    def _save_cache(self, path: Path, data: dict[str, Any]) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(
            json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
        )

    def _load_cache(self, path: Path) -> dict[str, Any] | None:
        if path.exists():
            log.info("Cache kullanılıyor: %s", path.name)
            return json.loads(path.read_text(encoding="utf-8"))
        return None
