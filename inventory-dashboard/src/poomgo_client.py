from __future__ import annotations

import os
from datetime import datetime, timezone, timedelta
from pathlib import Path
from typing import Any

import pandas as pd

from browser_session import capture_storage_state
from ecount_client import DataCollectionError
from transformer import normalize_code, to_number


KST = timezone(timedelta(hours=9))


class PoomgoClient:
    URL = "https://seller.poomgo.com/"

    def __init__(
        self,
        export_dir: str | Path,
        session_dir: str | Path,
        headless: bool = False,
        use_playwright: bool = False,
    ) -> None:
        self.export_dir = Path(export_dir)
        self.session_dir = Path(session_dir)
        self.headless = headless
        self.use_playwright = use_playwright

    def collect_naver_inventory(self) -> list[dict[str, Any]]:
        export_file = self._find_export_file()
        if export_file:
            return self._read_export(export_file)

        if self.use_playwright:
            self.capture_login_session()
            raise DataCollectionError(
                "품고 로그인 세션을 저장했습니다. 네이버 재고/출고 export 파일을 data/raw/poomgo 폴더에 저장한 뒤 다시 실행하세요."
            )

        raise DataCollectionError(
            "품고/네이버 export 파일을 찾지 못했습니다. "
            "POOMGO_EXPORT_DIR에 CSV/XLSX 파일을 두거나 POOMGO_USE_PLAYWRIGHT=true로 로그인 세션을 저장하세요."
        )

    def download_naver_inventory_export(self, timeout_seconds: int = 120) -> Path:
        from playwright.sync_api import TimeoutError, sync_playwright

        state_path = self.session_dir / "poomgo_storage_state.json"
        if not state_path.exists():
            raise DataCollectionError(f"품고 세션 파일이 없습니다: {state_path}")

        self.export_dir.mkdir(parents=True, exist_ok=True)
        with sync_playwright() as playwright:
            browser = playwright.chromium.launch(headless=self.headless)
            context = browser.new_context(storage_state=str(state_path), accept_downloads=True)
            page = context.new_page()
            page.goto(
                "https://seller.poomgo.com/partner-resource-inventory",
                wait_until="domcontentloaded",
                timeout=timeout_seconds * 1000,
            )
            try:
                page.wait_for_load_state("networkidle", timeout=15000)
            except TimeoutError:
                pass
            if "login" in page.url.lower():
                browser.close()
                raise DataCollectionError("품고 저장 세션이 만료되어 로그인 화면으로 이동했습니다.")

            page.get_by_role("button", name="조회").click(timeout=30000)
            _wait_for_text(page, "총합계", timeout_seconds)
            _select_all_rows(page)
            with page.expect_download(timeout=timeout_seconds * 1000) as download_info:
                page.get_by_role("button", name="화면 엑셀 다운로드", exact=True).click(timeout=30000)
            download = download_info.value
            target = self.export_dir / f"{datetime.now(KST).strftime('%Y%m%d_%H%M%S')}_{download.suggested_filename}"
            download.save_as(str(target))
            context.storage_state(path=str(state_path))
            browser.close()
        return target

    def capture_login_session(self) -> None:
        state_path = self.session_dir / "poomgo_storage_state.json"
        capture_storage_state(self.URL, state_path, self.headless, "품고")

    def _find_export_file(self) -> Path | None:
        if not self.export_dir.exists():
            return None
        patterns = [
            "*poomgo*.xlsx",
            "*poomgo*.xls",
            "*poomgo*.csv",
            "*naver*.xlsx",
            "*naver*.xls",
            "*naver*.csv",
            "*네이버*.xlsx",
            "*네이버*.csv",
            "*SKU별재고*.xlsx",
            "*SKU별재고*.csv",
            "*.xlsx",
            "*.csv",
        ]
        candidates: list[Path] = []
        for pattern in patterns:
            candidates.extend(self.export_dir.glob(pattern))
        if not candidates:
            return None
        unique_candidates = list(dict.fromkeys(candidates))
        scored = [(_table_row_count(path), path) for path in unique_candidates]
        max_rows = max(row_count for row_count, _ in scored)
        min_rows = max(1, int(max_rows * 0.8))
        viable = [path for row_count, path in scored if row_count >= min_rows]
        return max(viable, key=lambda path: path.stat().st_mtime)

    def _read_export(self, path: Path) -> list[dict[str, Any]]:
        df = _read_table(path)
        records = []
        collected_at = datetime.now(KST).isoformat(timespec="seconds")
        for _, row in df.iterrows():
            poomgo_code = _pick(row, ["품고 상품코드", "품고코드", "SKU번호", "product_code", "상품코드"])
            seller_code = _pick(row, ["판매자 상품코드", "판매자상품코드", "seller_product_code"])
            naver_code = _pick(row, ["네이버 상품코드", "네이버상품코드", "external_product_code", "naver_code", "바코드"])
            product_code = normalize_code(poomgo_code or seller_code or naver_code)
            if not product_code:
                continue
            records.append(
                {
                    "source": "poomgo",
                    "channel": "naver",
                    "product_code": normalize_code(product_code),
                    "seller_product_code": normalize_code(seller_code),
                    "external_product_code": normalize_code(naver_code),
                    "product_name": str(_pick(row, ["상품명", "품명", "SKU명", "product_name"]) or "").strip(),
                    "stock_each": to_number(_pick(row, ["현재고", "재고수량", "총재고", "stock_each"]), default=0),
                    "available_stock_each": to_number(_pick(row, ["가용재고", "판매가능재고", "available_stock_each"]), default=None),
                    "allocated_stock_each": to_number(_pick(row, ["출고대기", "할당재고", "allocated_stock_each"]), default=0),
                    "collected_at": collected_at,
                }
            )
        return records


def _read_table(path: Path) -> pd.DataFrame:
    suffix = path.suffix.lower()
    if suffix in {".xlsx", ".xls"}:
        return pd.read_excel(path, dtype=str).fillna("")
    for encoding in ("utf-8-sig", "cp949", "euc-kr"):
        try:
            return pd.read_csv(path, dtype=str, encoding=encoding).fillna("")
        except UnicodeDecodeError:
            continue
    return pd.read_csv(path, dtype=str).fillna("")


def _table_row_count(path: Path) -> int:
    try:
        return len(_read_table(path))
    except Exception:
        return 0


def _pick(row: Any, aliases: list[str]) -> Any:
    normalized = {str(key).strip(): value for key, value in row.items()}
    lower_map = {key.lower(): key for key in normalized}
    for alias in aliases:
        if alias in normalized:
            return normalized[alias]
        key = lower_map.get(alias.lower())
        if key:
            return normalized[key]
    return ""


def _wait_for_text(page: Any, text: str, timeout_seconds: int) -> None:
    try:
        page.get_by_text(text).first.wait_for(timeout=timeout_seconds * 1000)
    except Exception:
        page.wait_for_timeout(5000)


def _select_all_rows(page: Any) -> None:
    try:
        view_count = page.locator("select#viewCount").first
        if view_count.count() > 0:
            view_count.select_option("10000")
            page.wait_for_timeout(5000)
    except Exception:
        page.wait_for_timeout(1000)


def from_env() -> PoomgoClient:
    return PoomgoClient(
        export_dir=os.getenv("POOMGO_EXPORT_DIR", "./data/raw/poomgo"),
        session_dir=os.getenv("SESSION_DIR", "./data/session"),
        headless=os.getenv("HEADLESS", "false").lower() == "true",
        use_playwright=os.getenv("POOMGO_USE_PLAYWRIGHT", "false").lower() == "true",
    )
