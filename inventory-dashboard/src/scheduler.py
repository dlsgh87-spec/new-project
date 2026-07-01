from __future__ import annotations

import os
import time
from pathlib import Path

from dotenv import load_dotenv

from sync_inventory import PROJECT_ROOT, run_sync


def main() -> None:
    load_dotenv(PROJECT_ROOT / ".env")
    interval_minutes = int(os.getenv("SYNC_INTERVAL_MINUTES", "60"))
    while True:
        run_sync(manual_run=False)
        time.sleep(max(interval_minutes, 1) * 60)


if __name__ == "__main__":
    main()
