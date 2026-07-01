from __future__ import annotations

import csv
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable


NONSAN_CODES = {"N001", "논산", "nonsan"}
HAEUN_CODES = {"N004", "하은", "하은물류", "haeun"}


@dataclass(frozen=True)
class ConversionRule:
    pieces_per_box: float | None = None
    boxes_per_pallet: float | None = None
    pieces_per_pallet: float | None = None


def normalize_code(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def barcode_aliases(value: Any) -> list[str]:
    barcode = normalize_code(value)
    if not barcode:
        return []
    if barcode.endswith(".0") and barcode[:-2].isdigit():
        barcode = barcode[:-2]

    aliases = [barcode]
    if len(barcode) == 14 and barcode.isdigit():
        gtin13_base = barcode[1:13]
        aliases.append(gtin13_base + _gtin_check_digit(gtin13_base))
    return list(dict.fromkeys(aliases))


def _gtin_check_digit(base: str) -> str:
    total = 0
    for index, digit in enumerate(reversed(base)):
        total += int(digit) * (3 if index % 2 == 0 else 1)
    return str((10 - (total % 10)) % 10)


def to_number(value: Any, default: float | None = 0) -> float | None:
    if value is None:
        return default
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text:
        return default
    try:
        return float(text.replace(",", ""))
    except ValueError:
        return default


def safe_divide(value: Any, divisor: Any) -> float | None:
    numeric_value = to_number(value, default=None)
    numeric_divisor = to_number(divisor, default=None)
    if numeric_value is None or numeric_divisor in (None, 0):
        return None
    return numeric_value / numeric_divisor


def stock_units(stock_each: Any, conversion: ConversionRule | dict[str, Any] | None) -> dict[str, float | None]:
    if isinstance(conversion, dict):
        rule = ConversionRule(
            pieces_per_box=to_number(conversion.get("pieces_per_box"), default=None),
            boxes_per_pallet=to_number(conversion.get("boxes_per_pallet"), default=None),
            pieces_per_pallet=to_number(conversion.get("pieces_per_pallet"), default=None),
        )
    else:
        rule = conversion or ConversionRule()

    each = to_number(stock_each, default=None)
    return {
        "each": each,
        "box": safe_divide(each, rule.pieces_per_box),
        "pallet": safe_divide(each, rule.pieces_per_pallet),
    }


def load_product_mapping_csv(path: str | Path) -> list[dict[str, str]]:
    mapping_path = Path(path)
    if not mapping_path.exists():
        return []
    with mapping_path.open("r", encoding="utf-8-sig", newline="") as handle:
        return [
            {key.strip(): (value or "").strip() for key, value in row.items()}
            for row in csv.DictReader(handle)
        ]


def _mapping_lookup(mapping_rows: Iterable[dict[str, Any]], source: str) -> dict[str, str]:
    source_column = {
        "ecount": "ecount_code",
        "poomgo": "poomgo_code",
        "naver": "naver_code",
    }[source]
    lookup: dict[str, str] = {}
    for row in mapping_rows:
        sheet_code = normalize_code(row.get("sheet_code"))
        external_code = normalize_code(row.get(source_column))
        if sheet_code and external_code and external_code not in lookup:
            lookup[external_code] = sheet_code
    return lookup


def _resolve_code(
    candidates: Iterable[Any],
    source: str,
    sheet_codes: set[str],
    mapping_rows: list[dict[str, Any]],
) -> tuple[str | None, str | None]:
    lookups = [_mapping_lookup(mapping_rows, source)]
    if source == "poomgo":
        lookups.append(_mapping_lookup(mapping_rows, "naver"))

    normalized_candidates = [normalize_code(candidate) for candidate in candidates if normalize_code(candidate)]
    for lookup in lookups:
        for candidate in normalized_candidates:
            mapped = lookup.get(candidate)
            if mapped:
                return mapped, candidate
    for candidate in normalized_candidates:
        if candidate in sheet_codes:
            return candidate, candidate
    return None, normalized_candidates[0] if normalized_candidates else None


def _warehouse_field(record: dict[str, Any]) -> str | None:
    warehouse_code = normalize_code(record.get("warehouse_code"))
    warehouse_name = normalize_code(record.get("warehouse_name"))
    marker = warehouse_code or warehouse_name
    if marker in NONSAN_CODES or warehouse_name in NONSAN_CODES:
        return "nonsan_stock_each"
    if marker in HAEUN_CODES or warehouse_name in HAEUN_CODES:
        return "haeun_stock_each"
    return None


def _append_error(row: dict[str, Any], message: str) -> None:
    if message not in row["errors"]:
        row["errors"].append(message)


def merge_inventory(
    sheet_products: list[dict[str, Any]],
    conversions: dict[str, ConversionRule | dict[str, Any]],
    ecount_records: list[dict[str, Any]],
    poomgo_records: list[dict[str, Any]],
    mapping_rows: list[dict[str, Any]] | None = None,
    collected_at: str | None = None,
) -> dict[str, Any]:
    mapping_rows = mapping_rows or []
    sheet_codes = {normalize_code(row.get("code") or row.get("sheet_code")) for row in sheet_products}
    sheet_codes.discard("")

    merged: dict[str, dict[str, Any]] = {}
    for product in sheet_products:
        code = normalize_code(product.get("code") or product.get("sheet_code"))
        if not code:
            continue
        merged[code] = {
            "sheet_code": code,
            "product_name": product.get("product_name") or product.get("name") or "",
            "nonsan_stock_each": None,
            "haeun_stock_each": None,
            "total_stock_each": None,
            "naver_stock_each": None,
            "collected_at": collected_at,
            "errors": [],
        }

    unmapped: list[dict[str, Any]] = []
    ecount_seen: dict[str, set[str]] = {code: set() for code in merged}
    naver_seen: set[str] = set()

    for record in ecount_records:
        sheet_code, external_code = _resolve_code(
            [record.get("product_code")],
            "ecount",
            sheet_codes,
            mapping_rows,
        )
        warehouse_field = _warehouse_field(record)
        warehouse_or_channel = normalize_code(record.get("warehouse_code") or record.get("warehouse_name"))
        if not sheet_code or sheet_code not in merged or warehouse_field is None:
            unmapped.append(
                {
                    "source": "ecount",
                    "warehouse_or_channel": warehouse_or_channel,
                    "external_code": external_code or normalize_code(record.get("product_code")),
                    "external_product_name": record.get("product_name") or "",
                    "stock_each": to_number(record.get("stock_each"), default=0),
                    "reason": "상품 매핑 실패" if not sheet_code else "창고 코드 미인식",
                }
            )
            continue
        merged[sheet_code][warehouse_field] = to_number(record.get("stock_each"), default=0)
        ecount_seen[sheet_code].add("N001" if warehouse_field == "nonsan_stock_each" else "N004")

    for record in poomgo_records:
        candidates = [
            record.get("product_code"),
            record.get("seller_product_code"),
            record.get("external_product_code"),
            record.get("naver_code"),
        ]
        sheet_code, external_code = _resolve_code(candidates, "poomgo", sheet_codes, mapping_rows)
        if not sheet_code or sheet_code not in merged:
            unmapped.append(
                {
                    "source": "poomgo",
                    "warehouse_or_channel": record.get("channel") or "naver",
                    "external_code": external_code or "",
                    "external_product_name": record.get("product_name") or "",
                    "stock_each": to_number(record.get("available_stock_each", record.get("stock_each")), default=0),
                    "reason": "상품 매핑 실패",
                }
            )
            continue
        stock_value = record.get("available_stock_each")
        if stock_value in (None, ""):
            stock_value = record.get("stock_each")
        merged[sheet_code]["naver_stock_each"] = to_number(stock_value, default=0)
        naver_seen.add(sheet_code)

    for code, row in merged.items():
        nonsan = row["nonsan_stock_each"]
        haeun = row["haeun_stock_each"]
        if nonsan is not None and haeun is not None:
            row["total_stock_each"] = nonsan + haeun
        elif nonsan is not None or haeun is not None:
            row["total_stock_each"] = (nonsan or 0) + (haeun or 0)

        if "N001" not in ecount_seen[code]:
            _append_error(row, "외부 시스템 미조회: ecount N001")
        if "N004" not in ecount_seen[code]:
            _append_error(row, "외부 시스템 미조회: ecount N004")
        if code not in naver_seen:
            _append_error(row, "외부 시스템 미조회: naver")

        conversion = conversions.get(code) or ConversionRule()
        for prefix, stock_field in {
            "total": "total_stock_each",
            "nonsan": "nonsan_stock_each",
            "haeun": "haeun_stock_each",
            "naver": "naver_stock_each",
        }.items():
            units = stock_units(row[stock_field], conversion)
            row[f"{prefix}_box"] = units["box"]
            row[f"{prefix}_pallet"] = units["pallet"]
        row["error_message"] = "; ".join(row["errors"])

    return {
        "rows": list(merged.values()),
        "rows_by_code": merged,
        "unmapped": unmapped,
    }


def warehouse_update_rows(rows: Iterable[dict[str, Any]], warehouse_code: str) -> list[dict[str, Any]]:
    field = "nonsan_stock_each" if warehouse_code == "N001" else "haeun_stock_each"
    updates = []
    for row in rows:
        if row.get(field) is None:
            continue
        updates.append(
            {
                "sheet_code": row["sheet_code"],
                "product_name": row.get("product_name", ""),
                "stock_each": row[field],
            }
        )
    return updates
