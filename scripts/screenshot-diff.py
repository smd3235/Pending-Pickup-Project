#!/usr/bin/env python3
"""Screenshot regression check using perceptual hashing (imagehash).

Compares each PNG in documentation/screenshots/ against its counterpart in
documentation/screenshots/_baseline/ (the last known-good set). If a screenshot
differs beyond the hash threshold, the step fails so a human can review whether
the change is intentional (new report layout) or a regression.

Usage:
  python scripts/screenshot-diff.py
  python scripts/screenshot-diff.py --threshold 5          # hamming distance cap
  python scripts/screenshot-diff.py --refresh-baseline     # re-save baseline from current screenshots

Exit codes:
  0  no regressions (or no baseline yet — creates one)
  1  regression detected
  2  error
"""
from __future__ import annotations

import argparse
import os
import sys
from datetime import datetime
from pathlib import Path

try:
    import imagehash
    from PIL import Image
except ImportError as exc:
    print(f"FAILED: screenshot-diff requires imagehash + Pillow: {exc}")
    sys.exit(2)

try:
    import git
except ImportError:
    git = None

ROOT = Path(__file__).resolve().parent.parent
SCREENSHOTS = ROOT / "documentation" / "screenshots"
BASELINE = SCREENSHOTS / "_baseline"

# Perceptual hash size; larger = more sensitive to small changes.
HASH_SIZE = 16


def _ts() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def log(msg: str) -> None:
    print(f"{_ts()} | {msg}")


def git_repo():
    if git is None:
        return None
    try:
        return git.Repo(ROOT, search_parent_directories=True)
    except Exception:
        return None


def working_tree_dirty() -> bool:
    """True if there are uncommitted changes anywhere in the repo."""
    repo = git_repo()
    if repo is None:
        return False
    return bool(repo.is_dirty(untracked_files=True))


def img_hash(path: Path):
    try:
        with Image.open(path) as im:
            return imagehash.average_hash(im, hash_size=HASH_SIZE)
    except Exception as exc:
        log(f"WARN: could not hash {path.name}: {exc}")
        return None


def hamming(a, b) -> int:
    if a is None or b is None:
        return 999
    return int(a - b)  # imagehash supports __sub__ -> hamming distance


def capture_current() -> dict[str, object]:
    """Return {name: hash} for every PNG in documentation/screenshots/ (excluding _baseline)."""
    result = {}
    if not SCREENSHOTS.exists():
        return result
    for p in sorted(SCREENSHOTS.glob("*.png")):
        if p.parent != SCREENSHOTS:
            continue
        h = img_hash(p)
        if h is not None:
            result[p.name] = h
    return result


def refresh_baseline(threshold: int) -> int:
    BASELINE.mkdir(parents=True, exist_ok=True)
    current = capture_current()
    written = 0
    for name, _h in current.items():
        src = SCREENSHOTS / name
        dst = BASELINE / name
        with Image.open(src) as im:
            im.save(dst)
        written += 1
    log(f"Baseline refreshed: {written} screenshots -> {BASELINE}")
    return 0


def check_regression(threshold: int) -> int:
    if not SCREENSHOTS.exists():
        log("No documentation/screenshots/ directory — nothing to compare.")
        return 0
    current = capture_current()
    if not current:
        log("No PNG screenshots found — nothing to compare.")
        return 0

    baseline_exists = BASELINE.exists() and any(BASELINE.glob("*.png"))
    if not baseline_exists:
        log("No baseline yet — creating one from current screenshots (first run).")
        return refresh_baseline(threshold)

    failures = 0
    checked = 0
    for name, chash in sorted(current.items()):
        bpath = BASELINE / name
        if not bpath.exists():
            log(f"NEW screenshot (no baseline): {name}")
            checked += 1
            continue
        bhash = img_hash(bpath)
        dist = hamming(chash, bhash)
        checked += 1
        if dist > threshold:
            failures += 1
            log(f"REGRESSION: {name} hamming distance {dist} > threshold {threshold}")
        else:
            log(f"OK: {name} (distance {dist})")

    log(f"Compared {checked} screenshot(s); {failures} regression(s).")
    return 1 if failures else 0


def main() -> int:
    ap = argparse.ArgumentParser(
        description="Perceptual-hash screenshot regression check for documentation/screenshots/"
    )
    ap.add_argument(
        "--threshold",
        type=int,
        default=5,
        help="Max hamming distance allowed before flagging a regression (default 5).",
    )
    ap.add_argument(
        "--refresh-baseline",
        action="store_true",
        help="Re-save the baseline from current screenshots and exit 0.",
    )
    args = ap.parse_args()

    # Git-awareness: warn if the working tree has uncommitted changes so the
    # baseline comparison isn't done against a stale/mixed state.
    repo = git_repo()
    if repo is not None and repo.is_dirty(untracked_files=True):
        log("WARNING: working tree has uncommitted changes; baseline comparison "
            "may be against a stale state. Commit screenshots first, or use "
            "--refresh-baseline.")

    if args.refresh_baseline:
        return refresh_baseline(args.threshold)
    return check_regression(args.threshold)


if __name__ == "__main__":
    sys.exit(main())
