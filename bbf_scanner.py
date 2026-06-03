# /// script
# dependencies = [
#     "requests",
#     "openpyxl",
# ]
# ///

"""
GC Business Benefits Finder — Program Change Tracker
=====================================================
Monitors ISED's open-government dataset for additions and removals of funding
programs. Instead of scraping the IRIS web portal (which blocks bots), this
script queries the official CKAN API on open.canada.ca to fetch the latest
published XLSX, then diffs it against the previous run.

No browser required, no need to edit path directory. 
Run command to get started :D

By @author Kev.Wo 
"""

from __future__ import annotations
import argparse
import io
import json
import logging
import sys
from datetime import datetime
from pathlib import Path
import requests
import openpyxl

# CKAN dataset ID on open.canada.ca — stable, will not change
DATASET_ID = "4e75337e-70d0-4ed7-92d1-3b85192ec6b1"
CKAN_API   = f"https://open.canada.ca/data/api/action/package_show?id={DATASET_ID}"

# Updated column names present in the current ISED release.
COL_PROGRAM_NAME = "Title  - English"
COL_CATEGORY     = "Organization - English"

# Robust fallback aliases to protect against future structural shifts
COL_NAME_ALIASES     = ["program / service name", "program name", "service name", "nom du programme", "title  - english", "title - english", "title"]
COL_CATEGORY_ALIASES = ["category", "catégorie", "sector", "secteur", "organization - english", "organization", "organisation"]

SCRIPT_DIR = Path(__file__).parent
HISTORY_FILE = SCRIPT_DIR / "bbf_program_history.json"
DELTA_LOG_FILE = SCRIPT_DIR / "bbf_program_deltas.jsonl"
LOG_FILE = SCRIPT_DIR / "bbf_tracker.log"

REQUEST_TIMEOUT = 30  # seconds

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
    ],
)
log = logging.getLogger(__name__)



def _get_latest_xlsx_url() -> tuple[str, str]:
    """
    Query the CKAN API and return the (url, name) of the most recently
    modified XLSX resource in the dataset.
    """
    log.info("Querying open.canada.ca CKAN API for latest dataset release …")
    resp = requests.get(CKAN_API, timeout=REQUEST_TIMEOUT)
    resp.raise_for_status()
    payload = resp.json()

    if not payload.get("success"):
        raise RuntimeError(f"CKAN API returned failure: {payload}")

    resources = payload["result"]["resources"]
    xlsx_resources = [r for r in resources if r.get("format", "").upper() == "XLSX"]

    if not xlsx_resources:
        raise RuntimeError("No XLSX resources found in the dataset.")

    # Pick the one with the most recent metadata_modified timestamp
    latest = max(xlsx_resources, key=lambda r: r.get("metadata_modified") or "")
    log.info("Latest release: '%s' (modified %s)", latest["name"], latest.get("metadata_modified", "unknown"))
    return latest["url"], latest["name"]


def _download_xlsx(url: str) -> openpyxl.Workbook:
    """Download the XLSX at *url* and return an openpyxl Workbook."""
    log.info("Downloading dataset: %s", url)
    resp = requests.get(url, timeout=REQUEST_TIMEOUT)
    resp.raise_for_status()
    return openpyxl.load_workbook(io.BytesIO(resp.content), read_only=True)



def _find_column(headers: list[str], exact: str, aliases: list[str]) -> int | None:
    """
    Return the 0-based column index for *exact* header name, falling back to
    case-insensitive alias matching. Returns None if not found.
    """
    for i, h in enumerate(headers):
        if h == exact:
            return i
    lower_headers = [str(h).lower() for h in headers]
    for alias in aliases:
        for i, lh in enumerate(lower_headers):
            if alias in lh:
                return i
    return None


def extract_programs(workbook: openpyxl.Workbook, category_filter: str | None) -> dict[str, dict]:
    """
    Parse the workbook and return a dict keyed by program name.
    Each value is a metadata dict (category, etc.) for richer reporting.
    """
    ws = workbook.active
    rows = ws.iter_rows(values_only=True)
    raw_headers = [str(h) if h is not None else "" for h in next(rows)]

    name_col = _find_column(raw_headers, COL_PROGRAM_NAME, COL_NAME_ALIASES)
    cat_col  = _find_column(raw_headers, COL_CATEGORY,     COL_CATEGORY_ALIASES)

    if name_col is None:
        raise RuntimeError(
            f"Could not locate a program-name column. "
            f"Headers found: {raw_headers}"
        )

    norm_filter = category_filter.lower().replace("-", " ").strip() if category_filter else None

    programs: dict[str, dict] = {}
    for row in rows:
        name = str(row[name_col]).strip() if row[name_col] else None
        if not name or name.lower() in ("none", "nan", ""):
            continue

        category = str(row[cat_col]).strip() if (cat_col is not None and row[cat_col]) else "Unknown"

        if norm_filter and norm_filter not in category.lower():
            continue

        programs[name] = {"category": category}

    return programs



def load_history() -> dict[str, dict]:
    if not HISTORY_FILE.exists():
        return {}
    try:
        data = json.loads(HISTORY_FILE.read_text(encoding="utf-8"))
        if isinstance(data, list):
            return {name: {} for name in data}
        return data
    except (json.JSONDecodeError, TypeError) as exc:
        log.warning("History file is malformed (%s) — starting fresh.", exc)
        return {}


def save_history(programs: dict[str, dict]) -> None:
    HISTORY_FILE.write_text(
        json.dumps(programs, ensure_ascii=False, indent=4, sort_keys=True),
        encoding="utf-8",
    )
    log.info("History saved to '%s'.", HISTORY_FILE)


def append_delta_log(added: set[str], removed: set[str]) -> None:
    entry = {
        "timestamp": datetime.now().isoformat(),
        "added":   sorted(added),
        "removed": sorted(removed),
    }
    with DELTA_LOG_FILE.open("a", encoding="utf-8") as f:
        f.write(json.dumps(entry, ensure_ascii=False) + "\n")
    log.info("Delta record appended to '%s'.", DELTA_LOG_FILE)




def print_report(
    added: set[str],
    removed: set[str],
    current: dict[str, dict],
    release_name: str,
    first_run: bool,
    category_filter: str | None,
) -> None:
    width = 68
    scope = f"  Category filter: {category_filter}" if category_filter else "  All categories"

    print(f"\n┌{'─' * width}┐")
    print(f"│{'GC Business Benefits Finder — Change Report':^{width}}│")
    print(f"│{datetime.now().strftime('%Y-%m-%d %H:%M:%S'):^{width}}│")
    print(f"│{scope:<{width}}│")
    print(f"│{('  Source: ' + release_name):<{width}}│")
    print(f"└{'─' * width}┘")

    if first_run:
        print(f"\n  Baseline established. Tracking {len(current)} programs.\n")
        for name in sorted(current):
            cat = current[name].get("category", "")
            print(f"  • {name}  [{cat}]")
        print()
        return

    if not added and not removed:
        print(f"\n  No changes detected. {len(current)} programs currently tracked.\n")
        return

    if added:
        print(f"\n  ADDED ({len(added)} new program{'s' if len(added) != 1 else ''})\n")
        for name in sorted(added):
            cat = current.get(name, {}).get("category", "")
            print(f"  [+] {name}  [{cat}]")

    if removed:
        print(f"\n  REMOVED ({len(removed)} program{'s' if len(removed) != 1 else ''} no longer listed)\n")
        for name in sorted(removed):
            print(f"  [-] {name}")

    print()


def cmd_scan(category_filter: str | None) -> None:
    previous = load_history()
    first_run = len(previous) == 0

    try:
        xlsx_url, release_name = _get_latest_xlsx_url()
        workbook = _download_xlsx(xlsx_url)
        current = extract_programs(workbook, category_filter)
    except requests.RequestException as exc:
        log.error("Network error: %s", exc)
        sys.exit(1)
    except RuntimeError as exc:
        log.error("%s", exc)
        sys.exit(1)

    log.info("Parsed %d program entries.", len(current))

    prev_names    = set(previous.keys())
    current_names = set(current.keys())
    added         = current_names - prev_names
    removed       = prev_names - current_names

    print_report(added, removed, current, release_name, first_run, category_filter)
    save_history(current)

    if not first_run and (added or removed):
        append_delta_log(added, removed)


def cmd_reset() -> None:
    if HISTORY_FILE.exists():
        HISTORY_FILE.unlink()
        log.info("History cleared. Next scan will establish a new baseline.")
    else:
        log.info("No history file found — nothing to reset.")


def cmd_list(category_filter: str | None) -> None:
    programs = load_history()
    if not programs:
        print("No history on record. Run a scan first.")
        return

    if category_filter:
        norm = category_filter.lower().replace("-", " ")
        programs = {k: v for k, v in programs.items() if norm in v.get("category", "").lower()}

    print(f"\nCurrently tracked programs ({len(programs)}):\n")
    for name in sorted(programs):
        cat = programs[name].get("category", "")
        print(f"  • {name}  [{cat}]")
    print()


def cmd_list_categories() -> None:
    programs = load_history()
    if not programs:
        print("No history on record. Run a scan first.")
        return
    categories = sorted({v.get("category", "Unknown") for v in programs.values()})
    print(f"\nCategories/Organizations in current history ({len(categories)}):\n")
    for cat in categories:
        count = sum(1 for v in programs.values() if v.get("category") == cat)
        print(f"  {count:>4}  {cat}")
    print()



def main() -> None:
    parser = argparse.ArgumentParser(
        description="Track program additions and removals on the GC Business Benefits Finder.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "--category", "-c",
        metavar="CATEGORY",
        help="Filter to a specific category/organization.",
    )
    parser.add_argument(
        "--reset",
        action="store_true",
        help="Clear the history file and exit.",
    )
    parser.add_argument(
        "--list",
        action="store_true",
        help="Print currently tracked programs and exit.",
    )
    parser.add_argument(
        "--categories",
        action="store_true",
        help="List all organizations/categories present in history and exit.",
    )

    args = parser.parse_args()

    if args.reset:
        cmd_reset()
    elif args.list:
        cmd_list(args.category)
    elif args.categories:
        cmd_list_categories()
    else:
        cmd_scan(args.category)


if __name__ == "__main__":
    main()
