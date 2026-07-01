from __future__ import annotations

import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from apps_script_sync import append_manual_products, apply_manual_mapping_rows


def test_apps_script_sync_applies_manual_naver_barcode_mapping() -> None:
    lookup: dict[str, str] = {}

    count = apply_manual_mapping_rows(
        lookup,
        [
            {
                "sheet_code": "F00D0207",
                "ecount_code": "",
                "poomgo_code": "",
                "naver_code": "8809907651535",
            }
        ],
    )

    assert count >= 1
    assert lookup["8809907651535"] == "F00D0207"


def test_apps_script_sync_adds_manual_only_products_to_dashboard_products() -> None:
    products = [{"number": 1, "code": "F00D0207", "name": "기존"}]

    added = append_manual_products(
        products,
        [
            {"sheet_code": "F00D0207", "product_name": "기존"},
            {"sheet_code": "I01041", "product_name": "데일리스틱 흰살참치"},
        ],
    )

    assert added == 1
    assert products[-1]["number"] == 2
    assert products[-1]["code"] == "I01041"
    assert products[-1]["name"] == "데일리스틱 흰살참치"
