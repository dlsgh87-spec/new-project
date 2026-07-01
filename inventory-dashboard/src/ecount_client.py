from __future__ import annotations

import os
from datetime import datetime, timezone, timedelta
from pathlib import Path
from typing import Any
from urllib.parse import urlparse

import pandas as pd

from browser_session import capture_storage_state
from transformer import normalize_code, to_number


KST = timezone(timedelta(hours=9))
STOCK_STATUS_HASH = "#menuType=MENUTREE_000004&menuSeq=MENUTREE_000212&groupSeq=MENUTREE_000035&prgId=E040701&depth=4"
WAREHOUSE_NAMES = {
    "N001": "제품자재창고",
    "N004": "하은물류",
}


class DataCollectionError(RuntimeError):
    pass


class EcountClient:
    URL = "https://login.ecount.com/Login/"

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

    def collect_inventory(self, warehouse_code: str) -> list[dict[str, Any]]:
        export_candidate = self._find_export_file(warehouse_code)
        if export_candidate:
            export_file, is_warehouse_specific = export_candidate
            return self._read_export(
                export_file,
                warehouse_code,
                require_warehouse_column=not is_warehouse_specific,
            )

        if self.use_playwright:
            self.capture_login_session()
            raise DataCollectionError(
                "이카운트 로그인 세션을 저장했습니다. AUTO_DOWNLOAD_EXPORTS=true 상태에서 다시 실행하면 "
                f"{warehouse_code} export 자동 다운로드를 먼저 시도합니다."
            )

        raise DataCollectionError(
            f"이카운트 {warehouse_code} export 파일을 찾지 못했습니다. "
            "ECOUNT_EXPORT_DIR에 CSV/XLSX 파일을 두거나 ECOUNT_USE_PLAYWRIGHT=true로 로그인 세션을 저장하세요."
        )

    def download_inventory_export(self, warehouse_code: str, timeout_seconds: int = 120) -> Path:
        from playwright.sync_api import TimeoutError, sync_playwright

        self.export_dir.mkdir(parents=True, exist_ok=True)
        state_path = self.session_dir / "ecount_storage_state.json"
        credentials = {
            "company": os.getenv("ECOUNT_COMPANY_CODE", "").strip(),
            "user": os.getenv("ECOUNT_ID", "").strip(),
            "password": os.getenv("ECOUNT_PASSWORD", "").strip(),
        }

        with sync_playwright() as playwright:
            browser = playwright.chromium.launch(headless=self.headless)
            context_kwargs = {"accept_downloads": True}
            if state_path.exists():
                context_kwargs["storage_state"] = str(state_path)
            context = browser.new_context(**context_kwargs)
            page = context.new_page()
            page.goto(self.URL, wait_until="domcontentloaded", timeout=timeout_seconds * 1000)
            try:
                page.wait_for_load_state("networkidle", timeout=15000)
            except TimeoutError:
                pass

            if _is_login_page(page):
                if not all(credentials.values()):
                    browser.close()
                    raise DataCollectionError(
                        "이카운트 저장 세션이 만료되었고 .env의 ECOUNT_ID/ECOUNT_PASSWORD/ECOUNT_COMPANY_CODE가 비어 있습니다."
                    )
                _login(page, credentials, timeout_seconds)
                context.storage_state(path=str(state_path))

            if _is_login_page(page):
                browser.close()
                raise DataCollectionError("이카운트 로그인 후에도 내부 화면에 진입하지 못했습니다. 2FA/CAPTCHA/회사 선택 확인이 필요합니다.")

            downloaded = _try_ecount_download(page, warehouse_code, self.export_dir, timeout_seconds)
            browser.close()
            return downloaded

    def capture_login_session(self) -> None:
        state_path = self.session_dir / "ecount_storage_state.json"
        capture_storage_state(self.URL, state_path, self.headless, "이카운트")

    def _find_export_file(self, warehouse_code: str) -> tuple[Path, bool] | None:
        if not self.export_dir.exists():
            return None
        all_exports = [
            path
            for path in self.export_dir.iterdir()
            if path.is_file() and path.suffix.lower() in {".xlsx", ".xls", ".csv"}
        ]
        matching_exports = [
            path
            for path in all_exports
            if _detect_warehouse_code(path) == warehouse_code
        ]
        if matching_exports:
            return max(matching_exports, key=lambda path: path.stat().st_mtime), True

        specific_patterns = [f"*{warehouse_code}*.xlsx", f"*{warehouse_code}*.xls", f"*{warehouse_code}*.csv"]
        candidates: list[Path] = []
        for pattern in specific_patterns:
            candidates.extend(self.export_dir.glob(pattern))
        if candidates:
            return max(candidates, key=lambda path: path.stat().st_mtime), True

        generic_patterns = ["*ecount*.xlsx", "*ecount*.xls", "*ecount*.csv", "*이카운트*.xlsx", "*이카운트*.csv"]
        for pattern in generic_patterns:
            candidates.extend(self.export_dir.glob(pattern))
        if candidates:
            return max(candidates, key=lambda path: path.stat().st_mtime), False
        return None

    def _read_export(self, path: Path, warehouse_code: str, require_warehouse_column: bool) -> list[dict[str, Any]]:
        df = _read_table(path)
        warehouse_columns = {"창고구분", "창고코드", "warehouse_code"}
        has_warehouse_column = any(str(column).strip() in warehouse_columns for column in df.columns)
        if require_warehouse_column and not has_warehouse_column:
            raise DataCollectionError(
                f"{path.name} 파일명에 {warehouse_code}가 없고 창고 컬럼도 없어 이 파일을 {warehouse_code}로 사용할 수 없습니다."
            )
        records = []
        collected_at = datetime.now(KST).isoformat(timespec="seconds")
        for _, row in df.iterrows():
            product_code = _pick(row, ["상품코드", "품목코드", "제품코드", "코드", "product_code"])
            if not normalize_code(product_code):
                continue
            stock_each = _pick(row, ["현재고 수량", "현재고", "재고수량", "총량", "stock_each"])
            product_name = _pick(row, ["상품명", "품명", "품명 및 규격", "품목명[규격]", "제품명", "product_name"])
            export_warehouse = _pick(row, ["창고구분", "창고코드", "warehouse_code"])
            allowed_warehouses = {warehouse_code, "논산"} if warehouse_code == "N001" else {warehouse_code, "하은", "하은물류"}
            if normalize_code(export_warehouse) and normalize_code(export_warehouse) not in allowed_warehouses:
                continue
            records.append(
                {
                    "source": "ecount",
                    "warehouse_code": warehouse_code,
                    "warehouse_name": "논산" if warehouse_code == "N001" else "하은물류",
                    "product_code": normalize_code(product_code),
                    "product_name": str(product_name or "").strip(),
                    "stock_each": to_number(stock_each, default=0),
                    "collected_at": collected_at,
                }
            )
        return records


def _read_table(path: Path) -> pd.DataFrame:
    suffix = path.suffix.lower()
    if suffix in {".xlsx", ".xls"}:
        raw = pd.read_excel(path, dtype=str, header=None).fillna("")
        return _normalize_export_table(raw)
    for encoding in ("utf-8-sig", "cp949", "euc-kr"):
        try:
            raw = pd.read_csv(path, dtype=str, encoding=encoding, header=None).fillna("")
            return _normalize_export_table(raw)
        except UnicodeDecodeError:
            continue
    raw = pd.read_csv(path, dtype=str, header=None).fillna("")
    return _normalize_export_table(raw)


def _normalize_export_table(raw: pd.DataFrame) -> pd.DataFrame:
    header_index = _find_header_index(raw)
    if header_index is None:
        return raw.fillna("")
    headers = [str(value).strip() for value in raw.iloc[header_index].tolist()]
    data = raw.iloc[header_index + 1 :].copy()
    data.columns = headers
    data = data.loc[:, [column for column in data.columns if str(column).strip()]]
    return data.fillna("")


def _find_header_index(raw: pd.DataFrame) -> int | None:
    code_headers = {"상품코드", "품목코드", "제품코드", "코드", "product_code"}
    stock_headers = {"현재고 수량", "현재고", "재고수량", "총량", "stock_each", "가용재고"}
    for index, row in raw.iterrows():
        values = {str(value).strip() for value in row.tolist()}
        if values & code_headers and values & stock_headers:
            return int(index)
    for index, row in raw.iterrows():
        values = {str(value).strip() for value in row.tolist()}
        if values & code_headers:
            return int(index)
    return None


def _detect_warehouse_code(path: Path) -> str | None:
    try:
        suffix = path.suffix.lower()
        if suffix in {".xlsx", ".xls"}:
            first_rows = pd.read_excel(path, dtype=str, header=None, nrows=3).fillna("")
        else:
            for encoding in ("utf-8-sig", "cp949", "euc-kr"):
                try:
                    first_rows = pd.read_csv(path, dtype=str, encoding=encoding, header=None, nrows=3).fillna("")
                    break
                except UnicodeDecodeError:
                    continue
            else:
                first_rows = pd.read_csv(path, dtype=str, header=None, nrows=3).fillna("")
    except Exception:
        return None
    text = " ".join(str(value) for value in first_rows.astype(str).values.flatten())
    if "하은물류" in text or "N004" in text:
        return "N004"
    if "논산" in text or "제품자재창고" in text or "N001" in text:
        return "N001"
    return None


def _is_login_page(page: Any) -> bool:
    path = urlparse(page.url).path.lower()
    visible_login_inputs = page.locator("#com_code:visible, #id:visible, #passwd:visible").count()
    return "login" in path or visible_login_inputs >= 2


def _login(page: Any, credentials: dict[str, str], timeout_seconds: int) -> None:
    page.fill("#com_code", credentials["company"], timeout=timeout_seconds * 1000)
    page.fill("#id", credentials["user"], timeout=timeout_seconds * 1000)
    page.fill("#passwd", credentials["password"], timeout=timeout_seconds * 1000)
    try:
        page.wait_for_function("typeof excuteLogin === 'function'", timeout=15000)
        page.evaluate("excuteLogin()")
    except Exception:
        page.click("#save", timeout=timeout_seconds * 1000)
    try:
        page.wait_for_load_state("networkidle", timeout=20000)
    except Exception:
        pass
    for _ in range(4):
        page.wait_for_timeout(3000)
        _complete_new_device_notice(page)
        if "새로운 기기 로그인 알림" not in page.locator("body").inner_text(timeout=10000):
            break
    try:
        page.wait_for_load_state("networkidle", timeout=20000)
    except Exception:
        pass


def _complete_new_device_notice(page: Any) -> None:
    try:
        if page.get_by_text("새로운 기기 로그인 알림", exact=False).count() == 0:
            return
        try:
            page.locator("#toolbar_sid_toolbar_item_regist button:visible, #toolbar_sid_toolbar_item_non_regist button:visible").first.wait_for(timeout=15000)
        except Exception:
            pass
        clicked = page.evaluate(
            """
            () => {
              const nodes = [...document.querySelectorAll('#toolbar_sid_toolbar_item_regist button, #toolbar_sid_toolbar_item_non_regist button')];
              const target = nodes.find((node) => node.innerText.trim() === '등록' && (node.offsetWidth || node.offsetHeight || node.getClientRects().length)) || nodes[0];
              if (!target) return false;
              for (const type of ['mouseover', 'mousedown', 'mouseup', 'click']) {
                target.dispatchEvent(new MouseEvent(type, { bubbles: true }));
              }
              return true;
            }
            """
        )
        if clicked:
            page.wait_for_timeout(3000)
            return
        for selector in (
            "#toolbar_sid_toolbar_item_regist button",
            "#toolbar_sid_toolbar_item_non_regist button",
        ):
            try:
                button = page.locator(selector).first
                if button.count() > 0:
                    button.click(timeout=5000)
                    page.wait_for_timeout(3000)
                    return
            except Exception:
                continue
        for label in ("등록", "등록안함"):
            try:
                button = page.locator("button").filter(has_text=label).first
                if button.count() > 0:
                    button.click(timeout=5000)
                    page.wait_for_timeout(3000)
                    return
            except Exception:
                continue
    except Exception:
        return


def _try_ecount_download(page: Any, warehouse_code: str, export_dir: Path, timeout_seconds: int) -> Path:
    warehouse_name = WAREHOUSE_NAMES.get(warehouse_code)
    if not warehouse_name:
        raise DataCollectionError(f"지원하지 않는 이카운트 창고 코드입니다: {warehouse_code}")

    _open_stock_status_page(page, timeout_seconds)
    _dismiss_optional_notices(page)
    _select_warehouse(page, warehouse_code, warehouse_name, timeout_seconds)
    _run_stock_search(page, warehouse_name, timeout_seconds)

    before = set(export_dir.glob("*"))
    try:
        with page.expect_download(timeout=timeout_seconds * 1000) as download_info:
            page.locator("button#outputExcel").first.click(timeout=30000)
        download = download_info.value
        target = export_dir / f"{datetime.now(KST).strftime('%Y%m%d_%H%M%S')}_{warehouse_code}_{download.suggested_filename}"
        download.save_as(str(target))
        detected = _detect_warehouse_code(target)
        if detected and detected != warehouse_code:
            raise DataCollectionError(
                f"다운로드한 파일의 창고가 요청과 다릅니다. 요청={warehouse_code}, 파일={detected}, 파일={target}"
            )
        return target
    except Exception as exc:
        last_error = exc

    after = set(export_dir.glob("*"))
    new_files = sorted(after - before, key=lambda path: path.stat().st_mtime)
    if new_files:
        return new_files[-1]
    raise DataCollectionError(
        f"이카운트 {warehouse_code}({warehouse_name}) 자동 다운로드 버튼을 찾지 못했습니다. "
        f"메뉴 구조 확인이 필요합니다. 마지막 오류: {last_error}"
    )


def _open_stock_status_page(page: Any, timeout_seconds: int) -> None:
    base_url = page.url.split("#", 1)[0]
    page.goto(f"{base_url}{STOCK_STATUS_HASH}", wait_until="domcontentloaded", timeout=timeout_seconds * 1000)
    page.wait_for_timeout(6000)
    try:
        page.get_by_text("재고현황", exact=True).first.wait_for(timeout=timeout_seconds * 1000)
    except Exception:
        page.wait_for_timeout(3000)


def _dismiss_optional_notices(page: Any) -> None:
    for text in ("오늘하루그만보기", "확인"):
        try:
            locator = page.get_by_text(text, exact=True).last
            if locator.count() > 0:
                locator.click(timeout=2000)
                page.wait_for_timeout(500)
        except Exception:
            continue


def _select_warehouse(page: Any, warehouse_code: str, warehouse_name: str, timeout_seconds: int) -> None:
    page.locator('button[data-cid="txtSWhCd"][data-function-id="code.openpopup"]').first.click(timeout=timeout_seconds * 1000)
    try:
        page.get_by_text("창고검색", exact=True).first.wait_for(timeout=timeout_seconds * 1000)
    except Exception:
        page.wait_for_timeout(3000)

    warehouse_code_locator = page.locator(f'a[data-cid*="SALE001.WH_CD"][data-cid$="{warehouse_code}"]')
    try:
        warehouse_code_locator.first.wait_for(timeout=timeout_seconds * 1000)
    except Exception:
        try:
            page.get_by_text(warehouse_name, exact=True).first.wait_for(timeout=timeout_seconds * 1000)
        except Exception:
            page.wait_for_timeout(3000)

    selected = False
    try:
        page.locator(f'input[data-cid*="CHK_H"][data-cid$="{warehouse_code}"]').last.check(timeout=10000, force=True)
        selected = True
    except Exception:
        for locator in (
            warehouse_code_locator.last,
            page.get_by_text(warehouse_code, exact=True).last,
            page.get_by_text(warehouse_name, exact=True).last,
        ):
            try:
                locator.click(timeout=10000)
                selected = True
                break
            except Exception:
                continue
    if not selected:
        raise DataCollectionError(f"이카운트 창고 선택 팝업에서 {warehouse_code}({warehouse_name})를 찾지 못했습니다.")

    page.wait_for_timeout(1500)
    if _popup_is_visible(page, "창고검색"):
        try:
            page.locator("button#codeApply:visible").last.click(timeout=10000)
        except Exception:
            page.keyboard.press("F8")
        page.wait_for_timeout(1500)

    if _popup_is_visible(page, "창고검색"):
        raise DataCollectionError(f"이카운트 창고 선택 팝업이 닫히지 않았습니다: {warehouse_code}({warehouse_name})")


def _popup_is_visible(page: Any, title: str) -> bool:
    try:
        return bool(
            page.evaluate(
                """
                (title) => {
                  function visible(node) {
                    const rect = node.getBoundingClientRect();
                    const style = getComputedStyle(node);
                    return !!(rect.width || rect.height || node.getClientRects().length) &&
                      style.visibility !== 'hidden' &&
                      style.display !== 'none';
                  }
                  return [...document.querySelectorAll('[data-popup-id], [data-container="popup-body"], .ui-dialog')]
                    .some((node) => visible(node) && (node.innerText || '').includes(title));
                }
                """,
                title,
            )
        )
    except Exception:
        return False


def _run_stock_search(page: Any, warehouse_name: str, timeout_seconds: int) -> None:
    try:
        page.get_by_role("button", name="검색(F8)").first.click(timeout=30000)
    except Exception:
        page.keyboard.press("F8")
    try:
        page.wait_for_function(
            """
            (warehouseName) => {
              const text = document.body.innerText || '';
              return text.includes('회사명') && text.includes(warehouseName) && text.includes('품목코드');
            }
            """,
            warehouse_name,
            timeout=timeout_seconds * 1000,
        )
    except Exception:
        page.wait_for_timeout(8000)
        body = page.locator("body").inner_text(timeout=10000)
        if warehouse_name not in body or "품목코드" not in body:
            raise DataCollectionError(f"이카운트 {warehouse_name} 재고현황 검색 결과를 확인하지 못했습니다.")


def _click_text_if_visible(page: Any, text: str) -> bool:
    try:
        locator = page.get_by_text(text, exact=False).first
        if locator.count() == 0:
            return False
        locator.click(timeout=3000)
        page.wait_for_timeout(1000)
        return True
    except Exception:
        return False


def _set_text_or_select(page: Any, labels: list[str], value: str) -> bool:
    for label in labels:
        try:
            locator = page.get_by_label(label, exact=False).first
            if locator.count() > 0:
                tag = locator.evaluate("el => el.tagName.toLowerCase()")
                if tag == "select":
                    locator.select_option(label=value)
                else:
                    locator.fill(value)
                return True
        except Exception:
            pass
    try:
        option = page.get_by_text(value, exact=False).first
        if option.count() > 0:
            option.click(timeout=3000)
            return True
    except Exception:
        pass
    return False


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


def from_env() -> EcountClient:
    return EcountClient(
        export_dir=os.getenv("ECOUNT_EXPORT_DIR", "./data/raw/ecount"),
        session_dir=os.getenv("SESSION_DIR", "./data/session"),
        headless=os.getenv("HEADLESS", "false").lower() == "true",
        use_playwright=os.getenv("ECOUNT_USE_PLAYWRIGHT", "false").lower() == "true",
    )
