from __future__ import annotations

import argparse
import os
import time
from datetime import datetime
from pathlib import Path

from dotenv import load_dotenv
from playwright.sync_api import Error, sync_playwright


PROJECT_ROOT = Path(__file__).resolve().parents[1]


TARGETS = {
    "ecount": {
        "label": "이카운트",
        "url": "https://login.ecount.com/Login/",
        "state": "ecount_storage_state.json",
        "export_env": "ECOUNT_EXPORT_DIR",
        "default_export_dir": "./data/raw/ecount",
        "hint": "재고 I > 출력물 > 재고현황에서 N001 제품자재창고, N004 하은물류 파일을 각각 내려받으세요.",
    },
    "poomgo": {
        "label": "품고/네이버",
        "url": "https://seller.poomgo.com/",
        "state": "poomgo_storage_state.json",
        "export_env": "POOMGO_EXPORT_DIR",
        "default_export_dir": "./data/raw/poomgo",
        "hint": "네이버 재고 또는 네이버 출고에서 재고 export 파일을 내려받으세요.",
    },
}


def main() -> int:
    load_dotenv(PROJECT_ROOT / ".env")
    parser = argparse.ArgumentParser(description="저장된 세션으로 export 다운로드용 브라우저를 엽니다.")
    parser.add_argument("target", choices=TARGETS.keys())
    args = parser.parse_args()
    config = TARGETS[args.target]

    session_dir = PROJECT_ROOT / os.getenv("SESSION_DIR", "./data/session")
    state_path = session_dir / config["state"]
    if not state_path.exists():
        print(f"{config['label']} 세션 파일이 없습니다: {state_path}")
        print(f"먼저 python src/capture_sessions.py {args.target} 를 실행하세요.")
        return 1

    export_dir = PROJECT_ROOT / os.getenv(config["export_env"], config["default_export_dir"])
    export_dir.mkdir(parents=True, exist_ok=True)
    headless = os.getenv("HEADLESS", "false").lower() == "true"
    saved_files: list[Path] = []

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=headless)
        context = browser.new_context(storage_state=str(state_path), accept_downloads=True)
        done = {"value": False}

        def attach_download_handler(page: object) -> None:
            page.on("download", lambda download: _save_download(download, export_dir, saved_files))

        def finish(_source: object) -> None:
            done["value"] = True

        context.on("page", attach_download_handler)
        context.expose_binding("finishExport", finish)

        page = context.new_page()
        attach_download_handler(page)
        page.goto(config["url"], wait_until="domcontentloaded")

        control = context.new_page()
        control.set_content(_control_html(config["label"], config["hint"], export_dir))
        page.bring_to_front()

        try:
            while not done["value"]:
                time.sleep(0.5)
                _show_saved_files(control, saved_files)
                if not browser.is_connected():
                    break
        finally:
            try:
                context.storage_state(path=str(state_path))
            except Error:
                pass
            try:
                browser.close()
            except Error:
                pass

    print(f"{config['label']} export 브라우저 종료")
    for file_path in saved_files:
        print(f"저장됨: {file_path}")
    return 0


def _save_download(download: object, export_dir: Path, saved_files: list[Path]) -> None:
    suggested = download.suggested_filename
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    target = export_dir / f"{timestamp}_{suggested}"
    download.save_as(str(target))
    saved_files.append(target)


def _control_html(label: str, hint: str, export_dir: Path) -> str:
    return f"""
    <html lang="ko">
      <head>
        <meta charset="utf-8" />
        <title>{label} export 도우미</title>
        <style>
          body {{
            font-family: "Malgun Gothic", Arial, sans-serif;
            margin: 32px;
            line-height: 1.55;
            color: #1f2937;
          }}
          button {{
            font-size: 18px;
            padding: 13px 18px;
            border: 0;
            border-radius: 8px;
            background: #174a7c;
            color: white;
            cursor: pointer;
          }}
          pre {{
            white-space: pre-wrap;
            background: #f3f4f6;
            padding: 14px;
            border-radius: 8px;
            border: 1px solid #e5e7eb;
          }}
        </style>
      </head>
      <body>
        <h1>{label} export 도우미</h1>
        <p>{hint}</p>
        <p>다운로드가 감지되면 자동으로 아래 폴더에 저장합니다.</p>
        <pre>{export_dir}</pre>
        <h2>저장된 파일</h2>
        <pre id="files">아직 감지된 다운로드가 없습니다.</pre>
        <button onclick="window.finishExport()">다운로드 완료, 브라우저 닫기</button>
      </body>
    </html>
    """


def _show_saved_files(page: object, saved_files: list[Path]) -> None:
    text = "\n".join(str(path) for path in saved_files) if saved_files else "아직 감지된 다운로드가 없습니다."
    page.evaluate("(text) => { document.querySelector('#files').textContent = text; }", text)


if __name__ == "__main__":
    raise SystemExit(main())
