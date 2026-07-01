from __future__ import annotations

import os
from pathlib import Path

from dotenv import load_dotenv

import ecount_client
import poomgo_client


PROJECT_ROOT = Path(__file__).resolve().parents[1]


def main() -> int:
    load_dotenv(PROJECT_ROOT / ".env")
    timeout_seconds = int(os.getenv("DOWNLOAD_TIMEOUT_SECONDS", "120"))
    failures: list[str] = []

    ecount = ecount_client.from_env()
    for warehouse_code in ("N001", "N004"):
        try:
            path = ecount.download_inventory_export(warehouse_code, timeout_seconds=timeout_seconds)
            print(f"[OK] 이카운트 {warehouse_code} export 자동 다운로드: {path}")
        except Exception as exc:
            failures.append(f"이카운트 {warehouse_code}: {exc}")
            print(f"[실패] 이카운트 {warehouse_code} export 자동 다운로드: {exc}")

    poomgo = poomgo_client.from_env()
    try:
        path = poomgo.download_naver_inventory_export(timeout_seconds=timeout_seconds)
        print(f"[OK] 품고/네이버 export 자동 다운로드: {path}")
    except Exception as exc:
        failures.append(f"품고/네이버: {exc}")
        print(f"[실패] 품고/네이버 export 자동 다운로드: {exc}")

    if failures:
        print("\n일부 다운로드가 실패했습니다.")
        print("기존 export 파일이 있으면 sync_inventory.py는 기존 최신 파일로 계속 진행합니다.")
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
