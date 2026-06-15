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
    VRBO          — public iCal feed
    Escapia       — public iCal feed (Brokers CR, PEXS)
    Special Places — scrapes embedded calendar JSON from listing page
    Airbnb        — skipped (set skip:true in units.json)
    Booking.com   — not supported; manual check required
"""

import argparse
import json
import re
import sys
import time
from collections import defaultdict
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


def _window_dates(window: int) -> list[date]:
    today = date.today()
    return list(_date_range(today, window))


def _count_in_window(blocked: set[date], window: int) -> int:
    return sum(1 for d in _window_dates(window) if d in blocked)


def _by_month(blocked: set[date], window: int) -> dict[str, int]:
    """Return {YYYY-Mon: count} for each calendar month touched by the window."""
    counts: dict[str, int] = defaultdict(int)
    for d in _window_dates(window):
        if d in blocked:
            counts[d.strftime("%Y-%b")] += 1
    return dict(counts)


# ---------------------------------------------------------------------------
# Platform scrapers — all return set[date] | None
# ---------------------------------------------------------------------------

def fetch_airbnb(listing_id: str) -> set[date] | None:
    """Fetch blocked dates via Airbnb's internal calendar JSON API."""
    today = date.today()
    headers = {
        "X-Airbnb-API-Key": "d306zoyjsyarp7uqwhtun1d19",
        "Accept": "application/json",
    }
    blocked: set[date] = set()
    months_seen: set[tuple] = set()

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
                print(f"    [airbnb] 403 — blocked. Run from a residential network "
                      "or add an iCal URL to units.json.")
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

    return blocked


def fetch_ical(ical_url: str, label: str = "ical") -> set[date] | None:
    """Fetch blocked dates from any iCal (.ics) feed."""
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
        dtstart = re.search(r"DTSTART(?:;[^:]+)?:(\d{8})", event_block)
        dtend   = re.search(r"DTEND(?:;[^:]+)?:(\d{8})", event_block)
        if not dtstart or not dtend:
            continue
        try:
            start = date(int(dtstart[1][:4]), int(dtstart[1][4:6]), int(dtstart[1][6:]))
            end   = date(int(dtend[1][:4]),   int(dtend[1][4:6]),   int(dtend[1][6:]))
        except ValueError:
            continue
        current = start
        while current < end:
            if current >= today:
                blocked.add(current)
            current += timedelta(days=1)

    return blocked


def fetch_special_places(url: str) -> set[date] | None:
    """Fetch blocked dates from a Special Places of Costa Rica listing page."""
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
        return _parse_date_array(m.group(1), "unavailable_dates")

    # Pattern 2: inline iCal link
    ical_match = re.search(r'href="([^"]+\.ics[^"]*)"', html)
    if ical_match:
        return fetch_ical(ical_match.group(1), "special_places_ical")

    # Pattern 3: JS variable blocked_dates / booked_dates = [...]
    m = re.search(r'(?:blocked_dates|booked_dates)\s*[=:]\s*(\[[^\]]*\])', html)
    if m:
        return _parse_date_array(m.group(1), "blocked_dates")

    # Pattern 4: Lodgix flatpickr disable array
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
                    from_d = date.fromisoformat(entry["from"][:10]) if entry.get("from") else None
                    to_d   = date.fromisoformat(entry["to"][:10])   if entry.get("to")   else None
                    if from_d and to_d:
                        current = from_d
                        while current <= to_d:
                            blocked.add(current)
                            current += timedelta(days=1)
            return blocked
        except Exception:
            pass

    print(f"    [special_places] could not parse calendar from {url}")
    print(f"    Tip: check page source for date arrays or an .ics link.")
    return None


def _parse_date_array(json_str: str, label: str) -> set[date] | None:
    try:
        dates = json.loads(json_str)
        return {date.fromisoformat(ds[:10]) for ds in dates if isinstance(ds, str)}
    except Exception as e:
        print(f"    [{label}] JSON parse error: {e}")
        return None


# ---------------------------------------------------------------------------
# Per-unit dispatcher
# ---------------------------------------------------------------------------

def check_unit(unit: dict, window: int) -> int | None:
    """Fetch availability, print per-unit line + month breakdown, return window count."""
    platform = unit.get("platform", "")
    unit_id  = unit["unit_id"]

    if unit.get("skip"):
        print(f"  {unit_id}: SKIP")
        return None

    if "SOLD" in platform or "Krain" in platform:
        print(f"  {unit_id}: SKIP (sold/inactive)")
        return None

    # --- fetch ---
    blocked: set[date] | None = None

    if platform == "Airbnb":
        lid = unit.get("listing_id")
        if not lid:
            print(f"  {unit_id}: SKIP (no Airbnb listing ID — update units.json)")
            return None
        blocked = fetch_ical(unit["ical_url"], "airbnb_ical") if unit.get("ical_url") \
                  else fetch_airbnb(lid)

    elif platform == "VRBO":
        ical_url = unit.get("ical_url")
        if not ical_url:
            print(f"  {unit_id}: SKIP (no iCal URL)")
            return None
        blocked = fetch_ical(ical_url, "vrbo")

    elif platform in ("Brokers CR", "PEXS"):
        ical_url = unit.get("ical_url")
        if not ical_url:
            print(f"  {unit_id}: SKIP (no iCal URL)")
            return None
        blocked = fetch_ical(ical_url, "escapia")

    elif platform == "Special Places":
        url = unit.get("url")
        if not url:
            print(f"  {unit_id}: SKIP (no URL)")
            return None
        blocked = fetch_special_places(url)

    elif platform == "Booking.com":
        print(f"  {unit_id}: SKIP (Booking.com — not supported; check manually)")
        return None

    else:
        print(f"  {unit_id}: SKIP (unknown platform: {platform})")
        return None

    # --- report ---
    if blocked is None:
        print(f"  {unit_id}: could not retrieve data")
        return None

    total = _count_in_window(blocked, window)
    pct   = total / window * 100
    by_mo = _by_month(blocked, window)
    mo_str = "  |  ".join(f"{mo}: {n}n" for mo, n in sorted(by_mo.items()))
    print(f"  {unit_id}: {total}/{window} nights blocked ({pct:.0f}%)  [{mo_str}]")

    return total


# ---------------------------------------------------------------------------
# Connectivity check
# ---------------------------------------------------------------------------

def run_check():
    tests = [
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

    HEADER_ROW     = 4
    FIRST_DATE_COL = 7

    date_col = None
    for col in range(FIRST_DATE_COL, ws.max_column + 2):
        val = ws.cell(row=HEADER_ROW, column=col).value
        if val is None:
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
# Screen report (Android / no-xlsx mode)
# ---------------------------------------------------------------------------

def print_report(results: dict[str, int | None], units_meta: dict[str, dict], window: int) -> None:
    today = date.today()
    window_end = today + timedelta(days=window)

    print()
    print("=" * 40)
    print(f"  PACIFICO OCCUPANCY REPORT")
    print(f"  {today}  |  next {window} nights")
    print(f"  window: {today} → {window_end}")
    print("=" * 40)

    have_data = [(uid, v) for uid, v in results.items() if v is not None]
    skipped   = [uid for uid, v in results.items() if v is None]

    if not have_data:
        print("\n  No data retrieved — check connectivity.")
        return

    # Group by bedroom count
    by_beds: dict = {}
    for uid, blocked in have_data:
        beds = units_meta[uid].get("bedrooms") or "?"
        by_beds.setdefault(beds, []).append((uid, blocked))

    for beds in sorted(by_beds, key=lambda x: (x == "?", x)):
        label = f"{beds}BR" if beds != "?" else "Unknown BR"
        print(f"\n  ── {label} ──────────────────────")
        for uid, blocked in sorted(by_beds[beds]):
            pct   = blocked / window * 100
            bar   = "█" * int(pct / 5) + "░" * (20 - int(pct / 5))
            rented = " ★" if units_meta[uid].get("rented_by_you") else ""
            print(f"  {uid:<14}{rented}")
            print(f"    {bar} {blocked}/{window}n  {pct:.0f}%")

    # Month breakdown across all units with data
    print(f"\n  ── By Month ─────────────────────")
    month_totals: dict[str, list[int]] = defaultdict(list)
    # Re-derive month breakdown from window counts per unit isn't possible without
    # the blocked sets here, so print a note instead
    print(f"  (per-unit month detail shown during fetch above)")

    # Overall stats
    vals = [v for _, v in have_data]
    avg  = sum(vals) / len(vals)
    print(f"\n  ── Overall ──────────────────────")
    print(f"  Units with data : {len(have_data)}")
    print(f"  Avg blocked     : {avg:.1f}/{window} nights  ({avg/window*100:.0f}%)")
    print(f"  Highest         : {max(vals)}/30n  ({max(vals)/window*100:.0f}%)")
    print(f"  Lowest          : {min(vals)}/30n  ({min(vals)/window*100:.0f}%)")
    if skipped:
        print(f"\n  Skipped ({len(skipped)}): {', '.join(skipped)}")
    print("=" * 40)


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
    parser.add_argument("--report", action="store_true",
                        help="Print a clean screen report (great for phone/tablet — no xlsx needed)")
    args = parser.parse_args()

    if args.check:
        print("Connectivity check:")
        run_check()
        return

    units_data = json.loads(UNITS_FILE.read_text())
    if args.unit:
        units_data = [u for u in units_data if u["unit_id"] == args.unit]
        if not units_data:
            print(f"Unit '{args.unit}' not found in units.json")
            sys.exit(1)

    print(f"Checking {len(units_data)} units — {args.window}-night window — {date.today()}\n")

    # Collect full blocked-date sets so we can use them in both report and xlsx
    blocked_sets: dict[str, set[date] | None] = {}
    results: dict[str, int | None] = {}
    units_meta: dict[str, dict] = {u["unit_id"]: u for u in units_data}

    for unit in units_data:
        val = check_unit(unit, args.window)
        results[unit["unit_id"]] = val
        time.sleep(0.3)

    if args.report:
        print_report(results, units_meta, args.window)
    else:
        have_data = {uid: v for uid, v in results.items() if v is not None}
        print(f"\n--- Summary ({len(have_data)}/{len(results)} units retrieved) ---")
        for uid, val in results.items():
            status = f"{val} nights blocked" if val is not None else "no data / skipped"
            print(f"  {uid}: {status}")

    if args.update:
        xlsx_path = Path(args.update)
        if not xlsx_path.exists():
            print(f"\nFile not found: {xlsx_path}")
            sys.exit(1)
        update_xlsx(xlsx_path, results, args.window)
    elif not args.report:
        print("\nRun with --update path/to/tracker.xlsx to write results into the spreadsheet.")


if __name__ == "__main__":
    main()
