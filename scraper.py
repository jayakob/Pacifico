#!/usr/bin/env python3
"""
Pacifico Rental Occupancy Scraper
Fetches calendar data for each unit, counts blocked nights in the next 30 days,
and writes results into the Occupancy Log sheet of the tracker spreadsheet.

Usage:
    python scraper.py                          # print results, no file update
    python scraper.py --update tracker.xlsx    # also write results into the spreadsheet
    python scraper.py --window 45              # use 45-night window instead of 30
    python scraper.py --unit C-305             # test one unit
    python scraper.py --check                  # just check connectivity per platform

Run this from your local machine (residential IP). Cloud/datacenter IPs
are blocked by Airbnb, VRBO, and most vacation rental platforms.

Platform support:
    Airbnb        — JSON calendar API (listing_id required; some block bots)
    VRBO          — public iCal feed
    Escapia       — public iCal feed (Brokers CR, PEXS)
    Special Places — scrapes embedded calendar JSON from listing page
    Booking.com   — not supported; manual check required
"""

import argparse
import json
import re
import sys
import time
from datetime import date, datetime, timedelta
from pathlib import Path

import requests

UNITS_FILE = Path(__file__).parent / "units.json"
DEFAULT_WINDOW = 30

SESSION = requests.Session()
SESSION.headers.update({
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/125.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
    "Accept-Encoding": "gzip, deflate, br",
})


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _date_range(start: date, n: int):
    for i in range(n):
        yield start + timedelta(days=i)


def _count_in_window(blocked_dates: set[date], window: int) -> int:
    today = date.today()
    return sum(1 for d in _date_range(today, window) if d in blocked_dates)


# ---------------------------------------------------------------------------
# Platform scrapers
# ---------------------------------------------------------------------------

def fetch_airbnb(listing_id: str, window: int) -> int | None:
    """Count blocked nights via Airbnb's internal calendar JSON API."""
    today = date.today()
    # Airbnb also accepts these headers for the JSON API
    headers = {
        "X-Airbnb-API-Key": "d306zoyjsyarp7uqwhtun1d19",
        "Accept": "application/json",
    }
    blocked: set[date] = set()

    # Fetch 2 calendar months to cover any 30-day span crossing a month boundary
    months_seen = set()
    for offset in range(2):
        target = today + timedelta(days=offset * 28)
        key = (target.year, target.month)
        if key in months_seen:
            continue
        months_seen.add(key)
        url = (
            "https://www.airbnb.com/api/v2/calendar_months"
            f"?listing_id={listing_id}"
            f"&month={target.month}&year={target.year}&count=1"
            "&currency=USD"
        )
        try:
            r = SESSION.get(url, headers={**SESSION.headers, **headers}, timeout=15)
            r.raise_for_status()
            data = r.json()
        except requests.HTTPError as e:
            if e.response is not None and e.response.status_code == 403:
                print(f"    [airbnb] 403 — Airbnb is blocking this IP. "
                      "Run from a residential network, or add the iCal URL to units.json.")
            else:
                print(f"    [airbnb] HTTP error for {listing_id}: {e}")
            return None
        except Exception as e:
            print(f"    [airbnb] error for {listing_id}: {e}")
            return None

        for month_data in data.get("calendar_months", []):
            for day in month_data.get("days", []):
                if not day.get("available", True):
                    try:
                        blocked.add(date.fromisoformat(day["date"]))
                    except (KeyError, ValueError):
                        pass
        time.sleep(0.5)

    return _count_in_window(blocked, window)


def fetch_ical(ical_url: str, window: int, label: str = "ical") -> int | None:
    """Count blocked nights from any iCal (.ics) feed."""
    try:
        r = SESSION.get(ical_url, timeout=20)
        r.raise_for_status()
    except requests.HTTPError as e:
        code = e.response.status_code if e.response is not None else "?"
        print(f"    [{label}] HTTP {code} fetching {ical_url}")
        return None
    except Exception as e:
        print(f"    [{label}] error fetching {ical_url}: {e}")
        return None

    blocked: set[date] = set()
    today = date.today()

    for event_block in re.split(r"BEGIN:VEVENT", r.text)[1:]:
        # Match DTSTART with or without VALUE=DATE qualifier
        dtstart = re.search(r"DTSTART(?:;[^:]+)?:(\d{8})", event_block)
        dtend   = re.search(r"DTEND(?:;[^:]+)?:(\d{8})", event_block)
        if not dtstart or not dtend:
            continue
        try:
            start = date(int(dtstart[1][:4]), int(dtstart[1][4:6]), int(dtstart[1][6:]))
            end   = date(int(dtend[1][:4]),   int(dtend[1][4:6]),   int(dtend[1][6:]))
        except ValueError:
            continue
        # iCal DTEND is exclusive — each day from start up to (not including) end is blocked
        current = start
        while current < end:
            if current >= today:  # only care about future dates
                blocked.add(current)
            current += timedelta(days=1)

    return _count_in_window(blocked, window)


def fetch_special_places(url: str, window: int) -> int | None:
    """
    Count blocked nights from a Special Places of Costa Rica listing page.
    Tries several embedded-JSON patterns used by their WP/Lodgix stack.
    """
    try:
        r = SESSION.get(url, timeout=20)
        r.raise_for_status()
    except requests.HTTPError as e:
        code = e.response.status_code if e.response is not None else "?"
        print(f"    [special_places] HTTP {code} — {url}")
        return None
    except Exception as e:
        print(f"    [special_places] error: {e}")
        return None

    html = r.text

    # Pattern 1: JSON key "unavailable_dates" array
    m = re.search(r'"unavailable_dates"\s*:\s*(\["[^"]*"(?:\s*,\s*"[^"]*")*\])', html)
    if m:
        return _parse_date_array(m.group(1), window, "unavailable_dates")

    # Pattern 2: inline iCal link
    ical_match = re.search(r'href="([^"]+\.ics[^"]*)"', html)
    if ical_match:
        return fetch_ical(ical_match.group(1), window, "special_places_ical")

    # Pattern 3: JS variable blocked_dates = [...]
    m = re.search(r'(?:blocked_dates|booked_dates)\s*[=:]\s*(\[[^\]]*\])', html)
    if m:
        return _parse_date_array(m.group(1), window, "blocked_dates")

    # Pattern 4: Lodgix flatpickr disable array (dates in "from"/"to" objects)
    m = re.search(r'"disable"\s*:\s*(\[.*?\])', html, re.DOTALL)
    if m:
        try:
            disable = json.loads(m.group(1))
            today = date.today()
            blocked: set[date] = set()
            for entry in disable:
                if isinstance(entry, str):
                    blocked.add(date.fromisoformat(entry[:10]))
                elif isinstance(entry, dict):
                    from_d = date.fromisoformat(entry.get("from", "")[:10]) if entry.get("from") else None
                    to_d   = date.fromisoformat(entry.get("to",   "")[:10]) if entry.get("to")   else None
                    if from_d and to_d:
                        current = from_d
                        while current <= to_d:
                            blocked.add(current)
                            current += timedelta(days=1)
            return _count_in_window(blocked, window)
        except Exception:
            pass

    print(f"    [special_places] could not parse calendar from {url}")
    print(f"    Tip: check page source for date arrays or an .ics link.")
    return None


def _parse_date_array(json_str: str, window: int, label: str) -> int | None:
    try:
        dates = json.loads(json_str)
        today = date.today()
        blocked = {date.fromisoformat(ds[:10]) for ds in dates if isinstance(ds, str)}
        return _count_in_window(blocked, window)
    except Exception as e:
        print(f"    [{label}] JSON parse error: {e}")
        return None


# ---------------------------------------------------------------------------
# Per-unit dispatcher
# ---------------------------------------------------------------------------

def check_unit(unit: dict, window: int) -> int | None:
    platform = unit.get("platform", "")
    unit_id  = unit["unit_id"]

    if unit.get("skip"):
        print(f"  {unit_id}: SKIP")
        return None

    if "SOLD" in platform or "Krain" in platform:
        print(f"  {unit_id}: SKIP (sold/inactive)")
        return None

    if platform == "Airbnb":
        lid = unit.get("listing_id")
        if not lid:
            print(f"  {unit_id}: SKIP (no Airbnb listing ID — update units.json)")
            return None
        # Prefer iCal if the user has added one (more reliable)
        if unit.get("ical_url"):
            blocked = fetch_ical(unit["ical_url"], window, "airbnb_ical")
        else:
            blocked = fetch_airbnb(lid, window)

    elif platform == "VRBO":
        ical_url = unit.get("ical_url")
        if not ical_url:
            print(f"  {unit_id}: SKIP (no iCal URL)")
            return None
        blocked = fetch_ical(ical_url, window, "vrbo")

    elif platform in ("Brokers CR", "PEXS"):
        ical_url = unit.get("ical_url")
        if not ical_url:
            print(f"  {unit_id}: SKIP (no iCal URL)")
            return None
        blocked = fetch_ical(ical_url, window, "escapia")

    elif platform == "Special Places":
        url = unit.get("url")
        if not url:
            print(f"  {unit_id}: SKIP (no URL)")
            return None
        blocked = fetch_special_places(url, window)

    elif platform == "Booking.com":
        print(f"  {unit_id}: SKIP (Booking.com — not supported; check manually)")
        return None

    else:
        print(f"  {unit_id}: SKIP (no URL / unknown platform: {platform})")
        return None

    if blocked is not None:
        pct = blocked / window * 100
        print(f"  {unit_id}: {blocked}/{window} nights blocked ({pct:.0f}% occupied)")
    else:
        print(f"  {unit_id}: could not retrieve data")

    return blocked


# ---------------------------------------------------------------------------
# Connectivity check mode
# ---------------------------------------------------------------------------

def run_check():
    """Quick connectivity test — one request per platform."""
    tests = [
        ("Airbnb API",     "https://www.airbnb.com/api/v2/calendar_months?listing_id=1271942562086942388&month=6&year=2026&count=1&currency=USD"),
        ("VRBO iCal",      "https://www.vrbo.com/363286/ical.ics"),
        ("Escapia iCal",   "https://bookcostarica-brokers.escapia.com/Unit/iCal/157612"),
        ("Special Places", "https://www.specialplacesofcostarica.com/vacation-rental/pacifico-c-305/"),
    ]
    for label, url in tests:
        try:
            r = SESSION.get(url, timeout=10)
            print(f"  {label}: HTTP {r.status_code}")
        except Exception as e:
            print(f"  {label}: ERROR — {e}")


# ---------------------------------------------------------------------------
# XLSX updater
# ---------------------------------------------------------------------------

def update_xlsx(xlsx_path: Path, results: dict[str, int | None], window: int) -> None:
    import openpyxl
    from openpyxl.utils import get_column_letter

    wb = openpyxl.load_workbook(xlsx_path)
    ws = wb["Occupancy Log"]

    today_dt   = datetime.combine(date.today(), datetime.min.time())
    today_date = date.today()

    # Row 4 = date headers; dates start at column G (7)
    HEADER_ROW  = 4
    FIRST_DATE_COL = 7

    date_col = None
    for col in range(FIRST_DATE_COL, ws.max_column + 2):
        val = ws.cell(row=HEADER_ROW, column=col).value
        if val is None:
            # No date here yet — insert today
            date_col = col
            ws.cell(row=HEADER_ROW, column=col).value = today_dt
            ws.cell(row=HEADER_ROW, column=col).number_format = "YYYY-MM-DD"
            print(f"\nAdded today's date in column {get_column_letter(col)}")
            break
        if isinstance(val, datetime) and val.date() == today_date:
            date_col = col
            break

    if date_col is None:
        print("ERROR: could not find or create today's column in the Occupancy Log.")
        return

    # Read cached unit IDs (data_only=True gives formula results from last save)
    wb_ro = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws_ro = wb_ro["Occupancy Log"]
    unit_rows: dict[str, int] = {}
    for row in range(5, ws_ro.max_row + 1):
        uid = ws_ro.cell(row=row, column=1).value
        if uid:
            unit_rows[str(uid).strip()] = row

    written = 0
    for unit_id, blocked in results.items():
        if blocked is None:
            continue
        row = unit_rows.get(unit_id)
        if row is None:
            print(f"  Warning: {unit_id} not found in Occupancy Log — add it to the spreadsheet.")
            continue
        ws.cell(row=row, column=date_col).value = blocked
        written += 1

    wb.save(xlsx_path)
    print(f"Wrote {written} values → {xlsx_path.name} (column {get_column_letter(date_col)})")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Pacifico occupancy scraper")
    parser.add_argument("--update", metavar="XLSX",
                        help="Path to tracker .xlsx — writes results into Occupancy Log")
    parser.add_argument("--window", type=int, default=DEFAULT_WINDOW,
                        help=f"Forward window in nights (default {DEFAULT_WINDOW})")
    parser.add_argument("--unit", metavar="ID",
                        help="Only check this unit ID (e.g. C-305)")
    parser.add_argument("--check", action="store_true",
                        help="Test network connectivity to each platform and exit")
    args = parser.parse_args()

    if args.check:
        print("Connectivity check:")
        run_check()
        return

    units = json.loads(UNITS_FILE.read_text())
    if args.unit:
        units = [u for u in units if u["unit_id"] == args.unit]
        if not units:
            print(f"Unit '{args.unit}' not found in units.json")
            sys.exit(1)

    print(f"Checking {len(units)} units — {args.window}-night window — {date.today()}\n")

    results: dict[str, int | None] = {}
    for unit in units:
        results[unit["unit_id"]] = check_unit(unit, args.window)
        time.sleep(0.3)

    have_data = {uid: v for uid, v in results.items() if v is not None}
    print(f"\n--- Summary ({len(have_data)}/{len(results)} units retrieved) ---")
    for uid, val in results.items():
        status = f"{val} nights blocked" if val is not None else "no data"
        print(f"  {uid}: {status}")

    if args.update:
        xlsx_path = Path(args.update)
        if not xlsx_path.exists():
            print(f"\nFile not found: {xlsx_path}")
            sys.exit(1)
        update_xlsx(xlsx_path, results, args.window)
    else:
        print("\nRun with --update path/to/tracker.xlsx to write results into the spreadsheet.")


if __name__ == "__main__":
    main()
