from __future__ import annotations

import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from transformer import ConversionRule, barcode_aliases, merge_inventory


def base_products() -> list[dict]:
    return [
        {"code": "F00C0101", "product_name": "키튼 40g"},
        {"code": "F00C0102", "product_name": "키튼 2kg"},
    ]


def base_conversions() -> dict:
    return {
        "F00C0101": ConversionRule(pieces_per_box=200, boxes_per_pallet=40, pieces_per_pallet=8000),
        "F00C0102": ConversionRule(pieces_per_box=5, boxes_per_pallet=32, pieces_per_pallet=160),
    }


def test_merge_by_sheet_code_and_sum_n001_n004() -> None:
    result = merge_inventory(
        sheet_products=base_products(),
        conversions=base_conversions(),
        ecount_records=[
            {"warehouse_code": "N001", "product_code": "F00C0101", "stock_each": 4200},
            {"warehouse_code": "N004", "product_code": "F00C0101", "stock_each": 361},
        ],
        poomgo_records=[],
    )

    row = result["rows_by_code"]["F00C0101"]
    assert row["nonsan_stock_each"] == 4200
    assert row["haeun_stock_each"] == 361
    assert row["total_stock_each"] == 4561


def test_convert_each_to_box_and_pallet() -> None:
    result = merge_inventory(
        sheet_products=base_products(),
        conversions=base_conversions(),
        ecount_records=[
            {"warehouse_code": "N001", "product_code": "F00C0101", "stock_each": 4000},
            {"warehouse_code": "N004", "product_code": "F00C0101", "stock_each": 0},
        ],
        poomgo_records=[],
    )

    row = result["rows_by_code"]["F00C0101"]
    assert row["total_box"] == 20
    assert row["total_pallet"] == 0.5


def test_missing_conversion_values_do_not_raise() -> None:
    result = merge_inventory(
        sheet_products=[{"code": "F00C9999", "product_name": "기준값 없음"}],
        conversions={"F00C9999": ConversionRule()},
        ecount_records=[
            {"warehouse_code": "N001", "product_code": "F00C9999", "stock_each": 10},
            {"warehouse_code": "N004", "product_code": "F00C9999", "stock_each": 5},
        ],
        poomgo_records=[],
    )

    row = result["rows_by_code"]["F00C9999"]
    assert row["total_stock_each"] == 15
    assert row["total_box"] is None
    assert row["total_pallet"] is None


def test_external_system_missing_is_marked_without_zeroing() -> None:
    result = merge_inventory(
        sheet_products=base_products(),
        conversions=base_conversions(),
        ecount_records=[
            {"warehouse_code": "N001", "product_code": "F00C0101", "stock_each": 4200},
            {"warehouse_code": "N004", "product_code": "F00C0101", "stock_each": 361},
        ],
        poomgo_records=[],
    )

    missing = result["rows_by_code"]["F00C0102"]
    assert missing["nonsan_stock_each"] is None
    assert missing["haeun_stock_each"] is None
    assert "외부 시스템 미조회" in missing["error_message"]


def test_unmapped_external_product_is_separated() -> None:
    result = merge_inventory(
        sheet_products=base_products(),
        conversions=base_conversions(),
        ecount_records=[{"warehouse_code": "N001", "product_code": "UNKNOWN", "product_name": "외부상품", "stock_each": 7}],
        poomgo_records=[],
    )

    assert len(result["unmapped"]) == 1
    assert result["unmapped"][0]["external_code"] == "UNKNOWN"
    assert result["unmapped"][0]["reason"] == "상품 매핑 실패"


def test_mapping_csv_priority_maps_external_codes() -> None:
    result = merge_inventory(
        sheet_products=base_products(),
        conversions=base_conversions(),
        ecount_records=[{"warehouse_code": "N001", "product_code": "EC-A", "stock_each": 10}],
        poomgo_records=[{"product_code": "PG-A", "available_stock_each": 3}],
        mapping_rows=[
            {"sheet_code": "F00C0101", "ecount_code": "EC-A", "poomgo_code": "PG-A", "naver_code": ""},
        ],
    )

    row = result["rows_by_code"]["F00C0101"]
    assert row["nonsan_stock_each"] == 10
    assert row["naver_stock_each"] == 3
    assert result["unmapped"] == []


def test_naver_stock_is_not_added_to_total() -> None:
    result = merge_inventory(
        sheet_products=base_products(),
        conversions=base_conversions(),
        ecount_records=[
            {"warehouse_code": "N001", "product_code": "F00C0101", "stock_each": 100},
            {"warehouse_code": "N004", "product_code": "F00C0101", "stock_each": 50},
        ],
        poomgo_records=[{"product_code": "F00C0101", "available_stock_each": 999}],
    )

    row = result["rows_by_code"]["F00C0101"]
    assert row["total_stock_each"] == 150
    assert row["naver_stock_each"] == 999


def test_ean14_box_barcode_alias_maps_to_ean13_each_barcode() -> None:
    assert barcode_aliases("18809907650351") == ["18809907650351", "8809907650354"]


def test_poomgo_barcode_maps_through_sku_master_naver_code() -> None:
    result = merge_inventory(
        sheet_products=[{"code": "I01027", "product_name": "그레이비캔"}],
        conversions={"I01027": ConversionRule(pieces_per_box=24, boxes_per_pallet=10, pieces_per_pallet=240)},
        ecount_records=[],
        poomgo_records=[{"external_product_code": "8809907650354", "available_stock_each": 6048}],
        mapping_rows=[
            {
                "sheet_code": "I01027",
                "ecount_code": "I01027",
                "poomgo_code": "8809907650354",
                "naver_code": "8809907650354",
            },
        ],
    )

    assert result["rows_by_code"]["I01027"]["naver_stock_each"] == 6048
    assert result["unmapped"] == []
