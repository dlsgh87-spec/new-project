from __future__ import annotations

import os
from pathlib import Path

from dotenv import load_dotenv


PROJECT_ROOT = Path(__file__).resolve().parents[1]


def exists_with_size(path: Path) -> bool:
    return path.exists() and path.is_file() and path.stat().st_size > 0


def list_data_files(path: Path) -> list[Path]:
    if not path.exists():
        return []
    return [
        item
        for item in path.iterdir()
        if item.is_file() and item.name != ".gitkeep" and item.suffix.lower() in {".csv", ".xlsx", ".xls"}
    ]


def main() -> int:
    load_dotenv(PROJECT_ROOT / ".env")
    credentials = PROJECT_ROOT / os.getenv("GOOGLE_APPLICATION_CREDENTIALS", "./credentials/google-service-account.json")
    apps_script_url = os.getenv("APPS_SCRIPT_WEBAPP_URL") or os.getenv("GOOGLE_APPS_SCRIPT_URL")
    session_dir = PROJECT_ROOT / os.getenv("SESSION_DIR", "./data/session")
    ecount_dir = PROJECT_ROOT / os.getenv("ECOUNT_EXPORT_DIR", "./data/raw/ecount")
    poomgo_dir = PROJECT_ROOT / os.getenv("POOMGO_EXPORT_DIR", "./data/raw/poomgo")

    google_write_ready = bool(apps_script_url) or exists_with_size(credentials)
    checks = [
        ("Google Sheets 쓰기 권한", Path("Apps Script Web App URL") if apps_script_url else credentials, google_write_ready),
        ("이카운트 세션", session_dir / "ecount_storage_state.json", exists_with_size(session_dir / "ecount_storage_state.json")),
        ("품고 세션", session_dir / "poomgo_storage_state.json", exists_with_size(session_dir / "poomgo_storage_state.json")),
    ]

    print("OSP 재고 자동화 준비 상태")
    print("=" * 32)
    for label, path, ok in checks:
        print(f"[{'OK' if ok else '필요'}] {label}: {path}")

    ecount_files = list_data_files(ecount_dir)
    poomgo_files = list_data_files(poomgo_dir)
    print(f"[{'OK' if ecount_files else '필요'}] 이카운트 export 파일: {len(ecount_files)}개 ({ecount_dir})")
    print(f"[{'OK' if poomgo_files else '필요'}] 품고/네이버 export 파일: {len(poomgo_files)}개 ({poomgo_dir})")

    if ecount_files:
        print("  최신 이카운트 파일:", max(ecount_files, key=lambda path: path.stat().st_mtime).name)
    if poomgo_files:
        print("  최신 품고 파일:", max(poomgo_files, key=lambda path: path.stat().st_mtime).name)

    missing = [label for label, _, ok in checks if not ok]
    if not ecount_files:
        missing.append("이카운트 export 파일")
    if not poomgo_files:
        missing.append("품고/네이버 export 파일")

    if missing:
        print("\n다음 조치:")
        if "Google Sheets 쓰기 권한" in missing:
            print("- APPS_SCRIPT_WEBAPP_URL을 설정하거나 credentials/google-service-account.json 파일을 넣으세요.")
        if "이카운트 export 파일" in missing:
            print("- python src/open_export_browser.py ecount 로 로그인된 브라우저에서 N001/N004 재고 export를 내려받으세요.")
        if "품고/네이버 export 파일" in missing:
            print("- python src/open_export_browser.py poomgo 로 로그인된 브라우저에서 네이버 재고 export를 내려받으세요.")
        return 1

    print("\n준비 완료: python src/sync_inventory.py 실행이 가능합니다.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
