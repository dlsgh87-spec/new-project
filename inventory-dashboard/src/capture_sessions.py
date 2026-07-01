from __future__ import annotations

import argparse
from pathlib import Path

from dotenv import load_dotenv

import ecount_client
import poomgo_client
from sync_inventory import PROJECT_ROOT


def main() -> int:
    load_dotenv(PROJECT_ROOT / ".env")
    parser = argparse.ArgumentParser(description="이카운트/품고 로그인 세션을 저장합니다.")
    parser.add_argument(
        "target",
        choices=["ecount", "poomgo", "all"],
        nargs="?",
        default="all",
        help="저장할 로그인 세션",
    )
    args = parser.parse_args()

    if args.target in {"ecount", "all"}:
        ecount_client.from_env().capture_login_session()
        print("이카운트 세션 저장 완료")
    if args.target in {"poomgo", "all"}:
        poomgo_client.from_env().capture_login_session()
        print("품고 세션 저장 완료")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
