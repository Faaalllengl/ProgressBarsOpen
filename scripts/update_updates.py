"""Generate the update metadata shown by the application's update dialog."""

from __future__ import annotations

import json
import os
import subprocess
from datetime import datetime, timezone
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
UPDATES_FILE = ROOT / "updates.json"


def git(*args: str) -> str:
    return subprocess.check_output(["git", *args], cwd=ROOT, text=True).strip()


def main() -> None:
    commit_sha = os.environ.get("GITHUB_SHA") or git("rev-parse", "HEAD")
    short_sha = commit_sha[:7]
    subject = git("log", "-1", "--pretty=%s") or "Обновление программы"

    # Keep the existing title while making the actual version unique per push.
    previous = {}
    if UPDATES_FILE.exists():
        try:
            previous = json.loads(UPDATES_FILE.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            previous = {}

    payload = {
        "version": f"{datetime.now(timezone.utc):%Y.%m.%d}-{short_sha}",
        "date": datetime.now(timezone.utc).strftime("%Y-%m-%d"),
        "title": previous.get("title", "Обновление ProgressBarsOpen"),
        "changes": [subject],
    }
    UPDATES_FILE.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )


if __name__ == "__main__":
    main()
