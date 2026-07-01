from __future__ import annotations

import json
import os
import traceback
from datetime import datetime, timezone, timedelta
from pathlib import Path

from dotenv import load_dotenv

import ecount_client
import poomgo_client
from ecount_client import DataCollectionError
from google_sheets import GoogleSheetsClient
from logger import setup_logger
from transformer import load_product_mapping_csv, merge_inventory, warehouse_update_rows


KST = timezone(timedelta(hours=9))
PROJECT_ROOT = Path(__file__).resolve().parents[1]


def env_bool(name: str, default: bool = False) -> bool:
    value = os.getenv(name)
    if value is None:
        return default
    return value.strip().lower() in {"1", "true", "yes", "y"}


def run_sync(manual_run: bool = True) -> int:
    load_dotenv(PROJECT_ROOT / ".env")
    if os.getenv("APPS_SCRIPT_WEBAPP_URL") or os.getenv("GOOGLE_APPS_SCRIPT_URL"):
        from apps_script_sync import run_sync as run_apps_script_sync

        return run_apps_script_sync()

    logger = setup_logger(PROJECT_ROOT)
    started_at = datetime.now(KST)
    logger.info("실행 시작 시각: %s", started_at.isoformat(timespec="seconds"))

    collected_at = started_at.isoformat(timespec="seconds")
    status_messages: list[str] = []
    ecount_records: list[dict] = []
    poomgo_records: list[dict] = []
    unmapped_count = 0

    try:
        sheets = GoogleSheetsClient.from_env(PROJECT_ROOT)
        layout = sheets.analyze_month_sheet()
        logger.info(
            "월재고현황 구조: header_row=%s data_rows=%s:%s conversion_start=%s",
            layout.header_row,
            layout.data_start_row,
            layout.data_end_row,
            layout.conversion_start_col,
        )

        backup_name = sheets.backup_month_sheet_if_needed(
            layout=layout,
            policy=os.getenv("BACKUP_POLICY", "daily"),
            backup_before_each_sync=env_bool("BACKUP_BEFORE_EACH_SYNC"),
            manual_run=manual_run,
        )
        if backup_name:
            logger.info("백업 시트 생성: %s", backup_name)
        else:
            logger.info("백업 정책에 따라 이번 실행에서는 백업을 생략했습니다.")

        layout = sheets.ensure_month_sheet_layout(layout)
        products, conversions = sheets.read_products_and_conversions(layout)
        logger.info("월재고현황 SKU 수: %s", len(products))

        mapping_rows = sheets.read_mapping_sheet()
        csv_mapping = load_product_mapping_csv(PROJECT_ROOT / "config" / "product_mapping.csv")
        sku_master_mapping = sheets.read_sku_master_mapping()
        mapping_rows = mapping_rows + csv_mapping + sku_master_mapping
        logger.info("상품 매핑 수: %s", len(mapping_rows))

        auto_download = env_bool("AUTO_DOWNLOAD_EXPORTS", True)
        download_timeout = int(os.getenv("DOWNLOAD_TIMEOUT_SECONDS", "120"))

        ecount = ecount_client.from_env()
        for warehouse_code in ("N001", "N004"):
            if auto_download:
                try:
                    downloaded = ecount.download_inventory_export(warehouse_code, timeout_seconds=download_timeout)
                    logger.info("이카운트 %s export 자동 다운로드 성공: %s", warehouse_code, downloaded)
                except Exception as exc:
                    logger.warning("이카운트 %s export 자동 다운로드 실패, 기존 파일로 진행: %s", warehouse_code, exc)
            try:
                records = ecount.collect_inventory(warehouse_code)
                ecount_records.extend(records)
                logger.info("이카운트 %s 조회 성공: SKU %s개", warehouse_code, len(records))
            except Exception as exc:
                message = f"이카운트 {warehouse_code} 조회 실패: {exc}"
                logger.exception(message)
                status_messages.append(message)

        poomgo = poomgo_client.from_env()
        if auto_download:
            try:
                downloaded = poomgo.download_naver_inventory_export(timeout_seconds=download_timeout)
                logger.info("품고/네이버 export 자동 다운로드 성공: %s", downloaded)
            except Exception as exc:
                logger.warning("품고/네이버 export 자동 다운로드 실패, 기존 파일로 진행: %s", exc)
        try:
            poomgo_records = poomgo.collect_naver_inventory()
            logger.info("품고/네이버 재고 조회 성공: SKU %s개", len(poomgo_records))
        except Exception as exc:
            message = f"품고/네이버 재고 조회 실패: {exc}"
            logger.exception(message)
            status_messages.append(message)

        result = merge_inventory(
            sheet_products=products,
            conversions=conversions,
            ecount_records=ecount_records,
            poomgo_records=poomgo_records,
            mapping_rows=mapping_rows,
            collected_at=collected_at,
        )

        haeun_updated = sheets.update_ecount_source_sheet(
            "N004", warehouse_update_rows(result["rows"], "N004")
        )
        nonsan_updated = sheets.update_ecount_source_sheet(
            "N001", warehouse_update_rows(result["rows"], "N001")
        )
        naver_updated = sheets.update_naver_columns(layout, result["rows"])
        unmapped_count = sheets.append_unmapped_items(result["unmapped"])
        logger.info("Google Sheets 업데이트 SKU 수: 하은=%s 논산=%s 네이버=%s", haeun_updated, nonsan_updated, naver_updated)
        logger.info("미매핑 SKU 수: %s", unmapped_count)

        if result["unmapped"]:
            status_messages.append(f"미매핑 SKU {len(result['unmapped'])}개")
        if not status_messages:
            status_messages.append("성공")

        result_text = " / ".join(status_messages)
        sheets.rebuild_dashboard(
            layout=sheets.analyze_month_sheet(),
            last_success=collected_at if ecount_records or poomgo_records else "",
            result_text=result_text,
            unmapped_count=unmapped_count,
        )

        processed_path = PROJECT_ROOT / "data" / "processed" / "last_result.json"
        processed_path.write_text(
            json.dumps(
                {
                    "started_at": started_at.isoformat(timespec="seconds"),
                    "finished_at": datetime.now(KST).isoformat(timespec="seconds"),
                    "ecount_sku_count": len(ecount_records),
                    "poomgo_sku_count": len(poomgo_records),
                    "unmapped_count": unmapped_count,
                    "result": result_text,
                },
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )
        logger.info("마지막 성공 업데이트 시각: %s", collected_at if ecount_records or poomgo_records else "")
        return 0
    except FileNotFoundError as exc:
        logger.error("필수 파일 누락: %s", exc)
        return 2
    except DataCollectionError as exc:
        logger.error("수집 오류: %s", exc)
        return 3
    except Exception as exc:
        logger.error("Google Sheets 업데이트 실패 또는 전체 실행 오류: %s", exc)
        logger.debug(traceback.format_exc())
        return 1
    finally:
        finished_at = datetime.now(KST)
        logger.info("실행 종료 시각: %s", finished_at.isoformat(timespec="seconds"))


if __name__ == "__main__":
    raise SystemExit(run_sync(manual_run=True))
