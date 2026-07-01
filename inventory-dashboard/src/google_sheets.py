from __future__ import annotations

import os
from dataclasses import dataclass
from datetime import datetime, timezone, timedelta
from pathlib import Path
from typing import Any

import gspread

from dashboard import build_dashboard_format_requests, build_dashboard_values
from transformer import ConversionRule, barcode_aliases, normalize_code, to_number


KST = timezone(timedelta(hours=9))


@dataclass
class MonthSheetLayout:
    spreadsheet_id: str
    sheet_id: int
    sheet_name: str
    header_row: int
    data_start_row: int
    data_end_row: int
    code_col: int
    product_name_col: int
    conversion_start_col: int
    naver_start_col: int
    timestamp_col: int
    error_col: int

    @property
    def naver_start_col_letter(self) -> str:
        return col_to_letter(self.naver_start_col)

    @property
    def timestamp_col_letter(self) -> str:
        return col_to_letter(self.timestamp_col)

    @property
    def error_col_letter(self) -> str:
        return col_to_letter(self.error_col)


class GoogleSheetsClient:
    def __init__(self, credentials_path: str | Path, spreadsheet_id: str, sheet_name: str) -> None:
        self.credentials_path = Path(credentials_path)
        self.spreadsheet_id = spreadsheet_id
        self.sheet_name = sheet_name
        if not self.credentials_path.exists():
            raise FileNotFoundError(
                f"Google Service Account 인증 파일을 찾지 못했습니다: {self.credentials_path}"
            )
        self.gc = gspread.service_account(filename=str(self.credentials_path))
        self.spreadsheet = self.gc.open_by_key(spreadsheet_id)

    @classmethod
    def from_env(cls, project_root: Path) -> "GoogleSheetsClient":
        raw_credentials = os.getenv("GOOGLE_APPLICATION_CREDENTIALS", "./credentials/google-service-account.json")
        credentials_path = Path(raw_credentials)
        if not credentials_path.is_absolute():
            credentials_path = project_root / credentials_path
        return cls(
            credentials_path=credentials_path,
            spreadsheet_id=os.getenv("GOOGLE_SHEET_ID", ""),
            sheet_name=os.getenv("GOOGLE_SHEET_NAME", "월재고현황"),
        )

    def worksheet(self, title: str) -> gspread.Worksheet:
        return self.spreadsheet.worksheet(title)

    def analyze_month_sheet(self) -> MonthSheetLayout:
        ws = self.worksheet(self.sheet_name)
        values = ws.get("A1:Z120", value_render_option="FORMATTED_VALUE")
        header_row = _find_header_row(values)
        header = _row(values, header_row)
        code_col = _find_in_row(header, "코드")
        product_name_col = _find_in_row(header, "상품명")
        data_start_row = header_row + 1
        data_end_row = _find_data_end(values, data_start_row, code_col)
        conversion_start_col = _find_conversion_start(header)

        naver_start_col = _find_group_col(values, "네이버재고")
        if not naver_start_col:
            naver_start_col = 13
        timestamp_col = _find_in_row(header, "최종수집일시") or naver_start_col + 3
        error_col = _find_in_row(header, "오류메시지") or naver_start_col + 4

        return MonthSheetLayout(
            spreadsheet_id=self.spreadsheet_id,
            sheet_id=ws.id,
            sheet_name=self.sheet_name,
            header_row=header_row,
            data_start_row=data_start_row,
            data_end_row=data_end_row,
            code_col=code_col,
            product_name_col=product_name_col,
            conversion_start_col=conversion_start_col,
            naver_start_col=naver_start_col,
            timestamp_col=timestamp_col,
            error_col=error_col,
        )

    def backup_month_sheet_if_needed(
        self,
        layout: MonthSheetLayout,
        policy: str,
        backup_before_each_sync: bool,
        manual_run: bool,
    ) -> str | None:
        policy = (policy or "daily").lower()
        if policy == "none":
            return None
        if policy == "manual_only" and not manual_run:
            return None
        if backup_before_each_sync:
            policy = "each_sync"
        today = datetime.now(KST).strftime("%Y%m%d")
        if policy == "daily":
            prefix = f"backup_{today}_"
            if any(sheet.title.startswith(prefix) for sheet in self.spreadsheet.worksheets()):
                return None
        backup_name = f"backup_{datetime.now(KST).strftime('%Y%m%d_%H%M%S')}"
        self.spreadsheet.batch_update(
            {
                "requests": [
                    {
                        "duplicateSheet": {
                            "sourceSheetId": layout.sheet_id,
                            "newSheetName": backup_name,
                        }
                    }
                ]
            }
        )
        return backup_name

    def ensure_month_sheet_layout(self, layout: MonthSheetLayout) -> MonthSheetLayout:
        ws = self.worksheet(layout.sheet_name)
        top_values = ws.get("A1:Z6", value_render_option="FORMATTED_VALUE")
        if _find_group_col(top_values, "네이버재고"):
            return self.analyze_month_sheet()

        # Current sheet has M:P empty and Q:T as the conversion table. Insert one
        # column before Q so M:Q can hold three Naver fields plus timestamp/error.
        insert_at = 17 if _area_blank(top_values, 13, 16) and layout.conversion_start_col == 17 else 13
        insert_count = 1 if insert_at == 17 else 5
        self.spreadsheet.batch_update(
            {
                "requests": [
                    {
                        "insertDimension": {
                            "range": {
                                "sheetId": layout.sheet_id,
                                "dimension": "COLUMNS",
                                "startIndex": insert_at - 1,
                                "endIndex": insert_at - 1 + insert_count,
                            },
                            "inheritFromBefore": True,
                        }
                    }
                ]
            }
        )

        start = layout.data_start_row
        end = layout.data_end_row
        ws.update(
            "M2:Q6",
            [
                ["네이버재고", "", "", "", ""],
                ["낱개", "박스", "팔렛트", "", ""],
                [f"=SUM(M{start}:M{end})", f"=SUM(N{start}:N{end})", f"=SUM(O{start}:O{end})", "", ""],
                ["현재고", "", "", "최종수집일시", "오류메시지"],
                ["낱개\n재고", "박스\n재고", "팔렛트\n재고", "최종수집일시", "오류메시지"],
            ],
            value_input_option="USER_ENTERED",
        )
        self._format_naver_columns(layout.sheet_id, start, end)
        return self.analyze_month_sheet()

    def read_products_and_conversions(
        self, layout: MonthSheetLayout
    ) -> tuple[list[dict[str, Any]], dict[str, ConversionRule]]:
        ws = self.worksheet(layout.sheet_name)
        last_col = col_to_letter(layout.conversion_start_col + 3)
        values = ws.get(
            f"A{layout.data_start_row}:{last_col}{layout.data_end_row}",
            value_render_option="UNFORMATTED_VALUE",
        )
        products: list[dict[str, Any]] = []
        conversions: dict[str, ConversionRule] = {}
        for offset, row in enumerate(values):
            sheet_row = layout.data_start_row + offset
            code = normalize_code(_cell(row, layout.code_col))
            if not code:
                continue
            product_name = _cell(row, layout.product_name_col)
            products.append({"row": sheet_row, "code": code, "product_name": product_name})
            conversion_code = normalize_code(_cell(row, layout.conversion_start_col)) or code
            conversions[conversion_code] = ConversionRule(
                pieces_per_box=to_number(_cell(row, layout.conversion_start_col + 1), default=None),
                boxes_per_pallet=to_number(_cell(row, layout.conversion_start_col + 2), default=None),
                pieces_per_pallet=to_number(_cell(row, layout.conversion_start_col + 3), default=None),
            )
            conversions.setdefault(code, conversions[conversion_code])
        return products, conversions

    def read_mapping_sheet(self) -> list[dict[str, str]]:
        try:
            ws = self.worksheet("상품매핑")
        except gspread.WorksheetNotFound:
            return []
        values = ws.get_all_values()
        if not values:
            return []
        headers = [header.strip() for header in values[0]]
        rows = []
        for value_row in values[1:]:
            if not any(value_row):
                continue
            rows.append({headers[index]: value if index < len(value_row) else "" for index, value in enumerate(value_row)})
        return rows

    def read_sku_master_mapping(self) -> list[dict[str, str]]:
        try:
            ws = self.worksheet("SKU마스터")
        except gspread.WorksheetNotFound:
            return []
        values = ws.get_all_values()
        if len(values) < 3:
            return []
        header = [normalize_code(value).replace("\n", "").replace(" ", "") for value in values[1]]
        barcode_col = _find_header_like(header, ["상품바코드(EAN13)", "상품바코드", "EAN13"])
        box_barcode_col = _find_header_like(header, ["박스바코드(EAN14)", "박스바코드", "EAN14"])
        code_col = _find_header_like(header, ["품목코드", "코드"])
        name_col = _find_header_like(header, ["제품명", "상품명", "제품,상품명"])
        if barcode_col is None and box_barcode_col is None:
            return []
        if code_col is None:
            return []

        rows: list[dict[str, str]] = []
        seen: set[tuple[str, str]] = set()
        for row in values[2:]:
            sheet_code = normalize_code(_cell(row, code_col + 1))
            product_name = normalize_code(_cell(row, name_col + 1)) if name_col is not None else ""
            if not sheet_code:
                continue
            aliases: list[str] = []
            if barcode_col is not None:
                aliases.extend(barcode_aliases(_cell(row, barcode_col + 1)))
            if box_barcode_col is not None:
                aliases.extend(barcode_aliases(_cell(row, box_barcode_col + 1)))
            for barcode in dict.fromkeys(aliases):
                key = (sheet_code, barcode)
                if key in seen:
                    continue
                seen.add(key)
                rows.append(
                    {
                        "sheet_code": sheet_code,
                        "ecount_code": sheet_code,
                        "poomgo_code": barcode,
                        "naver_code": barcode,
                        "product_name": product_name,
                    }
                )
        return rows

    def update_ecount_source_sheet(self, warehouse_code: str, rows: list[dict[str, Any]]) -> int:
        if warehouse_code == "N004":
            return self._update_source_rows(
                sheet_name="하은물류",
                rows=rows,
                header_rows=2,
                code_col=1,
                product_col=3,
                stock_cols=[4],
            )
        return self._update_source_rows(
            sheet_name="논산합본",
            rows=rows,
            header_rows=1,
            code_col=1,
            product_col=None,
            stock_cols=[2, 3],
        )

    def update_naver_columns(self, layout: MonthSheetLayout, rows: list[dict[str, Any]]) -> int:
        ws = self.worksheet(layout.sheet_name)
        row_by_code = {row["sheet_code"]: row for row in rows}
        products, _ = self.read_products_and_conversions(layout)
        start_col = col_to_letter(layout.naver_start_col)
        end_col = col_to_letter(layout.error_col)
        current_values = ws.get(
            f"{start_col}{layout.data_start_row}:{end_col}{layout.data_end_row}",
            value_render_option="UNFORMATTED_VALUE",
        )

        output: list[list[Any]] = []
        updated = 0
        for index, product in enumerate(products):
            current = _pad_row(current_values[index] if index < len(current_values) else [], 5)
            row = row_by_code.get(product["code"])
            if not row:
                output.append(current)
                continue
            if row.get("naver_stock_each") is None:
                current[4] = row.get("error_message") or "외부 시스템 미조회: naver"
                output.append(current)
                continue
            output.append(
                [
                    row.get("naver_stock_each"),
                    row.get("naver_box"),
                    row.get("naver_pallet"),
                    row.get("collected_at") or datetime.now(KST).isoformat(timespec="seconds"),
                    row.get("error_message", ""),
                ]
            )
            updated += 1

        ws.update(
            f"{start_col}{layout.data_start_row}:{end_col}{layout.data_end_row}",
            output,
            value_input_option="USER_ENTERED",
        )
        return updated

    def append_unmapped_items(self, items: list[dict[str, Any]]) -> int:
        if not items:
            return 0
        ws = self._get_or_add_sheet("미매핑상품", rows=1000, cols=8)
        header_values = ws.get("A1:G1")
        if not header_values or not any(header_values[0]):
            ws.update(
                "A1:G1",
                [["발견일시", "source", "warehouse_or_channel", "external_code", "external_product_name", "stock_each", "reason"]],
                value_input_option="USER_ENTERED",
            )
        now = datetime.now(KST).isoformat(timespec="seconds")
        ws.append_rows(
            [
                [
                    now,
                    item.get("source", ""),
                    item.get("warehouse_or_channel", ""),
                    item.get("external_code", ""),
                    item.get("external_product_name", ""),
                    item.get("stock_each", ""),
                    item.get("reason", ""),
                ]
                for item in items
            ],
            value_input_option="USER_ENTERED",
        )
        return len(items)

    def rebuild_dashboard(
        self,
        layout: MonthSheetLayout,
        last_success: str,
        result_text: str,
        unmapped_count: int,
    ) -> None:
        try:
            old = self.worksheet("Dashboard")
            self.spreadsheet.del_worksheet(old)
        except gspread.WorksheetNotFound:
            pass
        row_count = layout.data_end_row - layout.data_start_row + 1
        dashboard = self.spreadsheet.add_worksheet(title="Dashboard", rows=max(row_count + 20, 120), cols=12)
        values = build_dashboard_values(layout, last_success, result_text, unmapped_count)
        dashboard.update("A1:I" + str(len(values)), values, value_input_option="USER_ENTERED")
        self.spreadsheet.batch_update({"requests": build_dashboard_format_requests(dashboard.id, row_count)})

    def _update_source_rows(
        self,
        sheet_name: str,
        rows: list[dict[str, Any]],
        header_rows: int,
        code_col: int,
        product_col: int | None,
        stock_cols: list[int],
    ) -> int:
        ws = self.worksheet(sheet_name)
        existing_values = ws.get("A1:D1000", value_render_option="UNFORMATTED_VALUE")
        row_number_by_code: dict[str, int] = {}
        for index, value_row in enumerate(existing_values, start=1):
            if index <= header_rows:
                continue
            code = normalize_code(_cell(value_row, code_col))
            if code:
                row_number_by_code[code] = index

        updates = []
        appends = []
        for row in rows:
            code = normalize_code(row.get("sheet_code"))
            if not code:
                continue
            stock_each = row.get("stock_each")
            existing_row = row_number_by_code.get(code)
            if existing_row:
                for stock_col in stock_cols:
                    updates.append({"range": f"{col_to_letter(stock_col)}{existing_row}", "values": [[stock_each]]})
                continue
            if sheet_name == "하은물류":
                appends.append([code, "", row.get("product_name", ""), stock_each])
            else:
                appends.append([code, stock_each, stock_each])

        if updates:
            ws.batch_update(updates, value_input_option="USER_ENTERED")
        if appends:
            ws.append_rows(appends, value_input_option="USER_ENTERED")
        return len(updates) // max(len(stock_cols), 1) + len(appends)

    def _get_or_add_sheet(self, title: str, rows: int, cols: int) -> gspread.Worksheet:
        try:
            return self.worksheet(title)
        except gspread.WorksheetNotFound:
            return self.spreadsheet.add_worksheet(title=title, rows=rows, cols=cols)

    def _format_naver_columns(self, sheet_id: int, data_start: int, data_end: int) -> None:
        self.spreadsheet.batch_update(
            {
                "requests": [
                    {
                        "repeatCell": {
                            "range": {"sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": 6, "startColumnIndex": 12, "endColumnIndex": 17},
                            "cell": {
                                "userEnteredFormat": {
                                    "backgroundColor": {"red": 0.86, "green": 0.91, "blue": 1.0},
                                    "horizontalAlignment": "CENTER",
                                    "textFormat": {"bold": True},
                                }
                            },
                            "fields": "userEnteredFormat(backgroundColor,horizontalAlignment,textFormat)",
                        }
                    },
                    {
                        "repeatCell": {
                            "range": {"sheetId": sheet_id, "startRowIndex": data_start - 1, "endRowIndex": data_end, "startColumnIndex": 12, "endColumnIndex": 17},
                            "cell": {"userEnteredFormat": {"horizontalAlignment": "CENTER"}},
                            "fields": "userEnteredFormat.horizontalAlignment",
                        }
                    },
                ]
            }
        )


def col_to_letter(col: int) -> str:
    result = ""
    while col:
        col, remainder = divmod(col - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _find_header_row(values: list[list[Any]]) -> int:
    for index, row in enumerate(values, start=1):
        normalized = [normalize_code(value).replace("\n", "") for value in row]
        if "번호" in normalized and "코드" in normalized and "상품명" in normalized:
            return index
    raise ValueError("월재고현황에서 번호/코드/상품명 헤더 행을 찾지 못했습니다.")


def _find_data_end(values: list[list[Any]], data_start_row: int, code_col: int) -> int:
    last = data_start_row - 1
    for index in range(data_start_row, len(values) + 1):
        if normalize_code(_cell(_row(values, index), code_col)):
            last = index
    return last


def _find_conversion_start(header: list[Any]) -> int:
    for col_index in range(1, len(header) - 2):
        if (
            normalize_code(header[col_index - 1]) == "코드"
            and normalize_code(header[col_index]) == "낱개/박스"
            and normalize_code(header[col_index + 1]) == "박스/파렛트"
        ):
            return col_index
    raise ValueError("월재고현황 우측 기준표(코드/낱개/박스/박스/파렛트)를 찾지 못했습니다.")


def _find_group_col(values: list[list[Any]], group_name: str) -> int | None:
    for row_index in (2, 5, 6):
        row = _row(values, row_index)
        for col_index, value in enumerate(row, start=1):
            if normalize_code(value).replace("\n", "") == group_name:
                return col_index
    return None


def _find_in_row(row: list[Any], target: str) -> int | None:
    normalized_target = target.replace("\n", "")
    for col_index, value in enumerate(row, start=1):
        if normalize_code(value).replace("\n", "") == normalized_target:
            return col_index
    return None


def _find_header_like(header: list[str], candidates: list[str]) -> int | None:
    normalized_candidates = [candidate.replace("\n", "").replace(" ", "") for candidate in candidates]
    for index, value in enumerate(header):
        for candidate in normalized_candidates:
            if candidate and candidate in value:
                return index
    return None


def _row(values: list[list[Any]], row_number: int) -> list[Any]:
    if row_number - 1 >= len(values):
        return []
    return values[row_number - 1]


def _cell(row: list[Any], col_number: int) -> Any:
    if col_number - 1 >= len(row):
        return ""
    return row[col_number - 1]


def _pad_row(row: list[Any], length: int) -> list[Any]:
    padded = list(row)
    while len(padded) < length:
        padded.append("")
    return padded[:length]


def _area_blank(values: list[list[Any]], start_col: int, end_col: int) -> bool:
    for row_number in range(1, min(len(values), 6) + 1):
        row = _row(values, row_number)
        for col in range(start_col, end_col + 1):
            if normalize_code(_cell(row, col)):
                return False
    return True
