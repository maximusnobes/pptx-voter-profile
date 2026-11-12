#!/usr/bin/env python3
"""
watch_folder.py — Phase 3 folder monitor.

Watches ./inbox for new PDI exports (.xls). Each file is one district
and becomes its own deck — drop several districts at once if you like.
The watcher waits until the inbox has been quiet for QUIET_SECONDS so
multi-file drops settle, then processes every file independently: one
file failing does not stop the others.

Rename each file to the district name first (the filename becomes the
client name on the slides).

  Success -> deck lands in ./output, the file moves to ./processed
  Failure -> the file moves to ./failed, details in output/populate_pptx.log

Start it with:  ./venv/bin/python watch_folder.py
Stop it with:   Ctrl+C
"""

import shutil
import time
from datetime import datetime
from pathlib import Path

import populate_pptx  # must sit in the same folder

HERE = Path(__file__).resolve().parent
INBOX = HERE / "inbox"
PROCESSED = HERE / "processed"
FAILED = HERE / "failed"
OUTDIR = HERE / "output"

QUIET_SECONDS = 15   # how long the inbox must be quiet before a batch runs
POLL_SECONDS = 2


def stable_files():
    """Inbox files whose size hasn't changed for one poll interval
    (i.e., they've finished copying)."""
    pdfs = sorted(list(INBOX.glob("*.xls")) + list(INBOX.glob("*.csv")))
    sizes = {p: p.stat().st_size for p in pdfs}
    time.sleep(1)
    return [p for p in pdfs if p.exists() and p.stat().st_size == sizes[p]]


def main():
    for d in (INBOX, PROCESSED, FAILED, OUTDIR):
        d.mkdir(exist_ok=True)
    print(f"Watching {INBOX} — drop PDI .xls export(s), renamed to the "
          f"district name, and wait ~{QUIET_SECONDS}s. Ctrl+C to stop.")

    last_change = None
    last_seen = set()

    while True:
        time.sleep(POLL_SECONDS)
        current = set(p.name for p in list(INBOX.glob("*.xls")) + list(INBOX.glob("*.csv")))
        if not current:
            last_change, last_seen = None, set()
            continue
        if current != last_seen:
            last_seen = current
            last_change = time.time()
            print(f"[{datetime.now():%H:%M:%S}] inbox: {sorted(current)} "
                  f"— waiting for batch to settle...")
            continue
        if last_change and time.time() - last_change >= QUIET_SECONDS:
            batch = stable_files()
            if not batch:
                continue
            ok = fail = 0
            for p in batch:
                print(f"[{datetime.now():%H:%M:%S}] processing {p.name} ...")
                argv = [str(p),
                        "--template", str(HERE / "template_placeholders.pptx"),
                        "--mapping", str(HERE / "placeholder_mapping.json"),
                        "--outdir", str(OUTDIR)]
                try:
                    rc = populate_pptx.main(argv)
                except Exception as e:
                    print(f"  ERROR: {e}")
                    rc = 1
                dest = PROCESSED if rc == 0 else FAILED
                stamped = f"{datetime.now():%Y-%m-%d_%H%M%S}_{p.name}"
                shutil.move(str(p), dest / stamped)
                if rc == 0:
                    ok += 1
                    print(f"[{datetime.now():%H:%M:%S}] {p.name}: done — "
                          f"deck is in output/")
                else:
                    fail += 1
                    print(f"[{datetime.now():%H:%M:%S}] {p.name}: FAILED — "
                          f"moved to failed/, see output/populate_pptx.log")
            print(f"[{datetime.now():%H:%M:%S}] batch finished: "
                  f"{ok} deck(s), {fail} failure(s)")
            last_change, last_seen = None, set()


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nStopped.")
