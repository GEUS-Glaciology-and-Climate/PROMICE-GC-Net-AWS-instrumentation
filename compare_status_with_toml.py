"""
Compares humidity-instrument records in stations/*.toml against the field-visit
records in resources/Station and Instrument Status.xlsx ("Station Status" sheet),
and interactively offers to correct the TOML files where the two disagree.

Read-only with respect to the Excel workbook: it is only ever loaded, never saved.
The TOML files are only written when the user explicitly confirms a specific change.

How the comparison works
-------------------------
The Excel sheet does not record "a hygrometer was replaced on this visit" directly.
Instead, for a handful of humidity-sensor columns (Rotronic Installed, Lufft Upper,
Lufft Lower, Vaisla Upper Installed, Vaisla Lower Installed) each visit row shows the
INSTALL YEAR of whatever sensor is currently mounted in that slot -- the value is
carried forward unchanged on visits where nothing was swapped, and changes only on
the visit where a new sensor went in. So a "hygrometer change event" is inferred by
walking each site's visits in date order and looking for the value in a slot
changing. Two slots that pick up the same brand on the same visit date (e.g. a Lufft
installed both "Upper" and "Lower" in one visit) are merged into a single event,
since the TOML records one instrument per station, not one per upper/lower slot.

Because a TOML station code (e.g. "NUK", "THU") often corresponds to several distinct
Excel "Site" codes (e.g. NUK_B/NUK_K/NUK_L/NUK_U), and some naming differs outright
(DY2 vs DYE2, JAR vs JR1, ...), this mapping is not guessed silently. See
station_site_mapping.csv.

station_site_mapping.csv
-------------------------
A small hand-confirmable lookup table (TomlStation, ExcelSite, Confirmed), in the
same spirit as maintenance_sheets_classification.csv elsewhere in this repo: the
script proposes candidates, the user confirms once, and the answer is cached there
for future runs. Exact-string matches (e.g. TOML "CEN" <-> Excel site "CEN") are
auto-confirmed since there is no real ambiguity. Edit this file by hand at any time
to fix a wrong mapping.

Usage
-----
    python compare_status_with_toml.py [--station CODE] [--report-only] [--verbose]
                                        [--tolerance-days N]

Run with --report-only first to see the full comparison with no prompts.
"""

from __future__ import annotations

import argparse
import csv
import difflib
import sys
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Optional

try:
    import openpyxl
except ImportError:
    sys.exit("Missing dependency 'openpyxl'. Install it with: pip install openpyxl")

try:
    import tomlkit
except ImportError:
    sys.exit("Missing dependency 'tomlkit'. Install it with: pip install tomlkit")

REPO_ROOT = Path(__file__).resolve().parent
DEFAULT_XLSX = REPO_ROOT / "resources" / "Station and Instrument Status.xlsx"
DEFAULT_STATIONS_DIR = REPO_ROOT / "stations"
DEFAULT_MAPPING_FILE = REPO_ROOT / "station_site_mapping.csv"
STATUS_SHEET_NAME = "Station Status"
HEADER_ROW = 2
DATA_START_ROW = 3

# (internal slot key, exact Excel column header, sensor brand it represents)
HUMIDITY_SLOTS = [
    ("rotronic", "Rotronic Installed", "Rotronic"),
    ("lufft_upper", "Lufft Upper", "Lufft"),
    ("lufft_lower", "Lufft Lower", "Lufft"),
    ("vaisala_upper", "Vaisla Upper Installed", "Vaisala"),
    ("vaisala_lower", "Vaisla Lower Installed", "Vaisala"),
]
NA_TOKENS = {"NA", "N/A", ""}
MAPPING_FIELDS = ["TomlStation", "ExcelSite", "Confirmed"]


# --------------------------------------------------------------------------- #
# Excel side
# --------------------------------------------------------------------------- #

@dataclass
class Visit:
    site: str
    station_name: str
    visit_date: date
    visit_type: str
    record_status: str
    slot_values: dict


@dataclass
class Event:
    id: int
    site: str
    brand: str
    slots: list
    event_date: date
    visit_type: str
    first_observed: bool


def parse_excel_date(value) -> Optional[date]:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        text = value.strip()
        try:
            return date.fromisoformat(text)
        except ValueError:
            return None
    return None


def load_visits(xlsx_path: Path) -> list:
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    if STATUS_SHEET_NAME not in wb.sheetnames:
        sys.exit(f"Sheet '{STATUS_SHEET_NAME}' not found in {xlsx_path}")
    ws = wb[STATUS_SHEET_NAME]

    header = {}
    for cell in ws[HEADER_ROW]:
        if cell.value:
            header[str(cell.value).strip()] = cell.column

    required = ["Station Name", "Site", "Visited", "Visit Type"] + [h for _, h, _ in HUMIDITY_SLOTS]
    missing = [h for h in required if h not in header]
    if missing:
        sys.exit(
            f"Expected column(s) not found in '{STATUS_SHEET_NAME}' header row {HEADER_ROW}: "
            f"{missing}. The sheet layout may have changed; update HUMIDITY_SLOTS / header names."
        )

    def cell_text(row, name):
        v = ws.cell(row=row, column=header[name]).value
        return "" if v is None else str(v).strip()

    visits = []
    for row in range(DATA_START_ROW, ws.max_row + 1):
        site = cell_text(row, "Site")
        if not site:
            continue
        visit_date = parse_excel_date(ws.cell(row=row, column=header["Visited"]).value)
        if visit_date is None:
            continue
        slot_values = {slot_key: cell_text(row, header_name) for slot_key, header_name, _ in HUMIDITY_SLOTS}
        visits.append(
            Visit(
                site=site,
                station_name=cell_text(row, "Station Name"),
                visit_date=visit_date,
                visit_type=cell_text(row, "Visit Type"),
                record_status=cell_text(row, "Record Status") if "Record Status" in header else "",
                slot_values=slot_values,
            )
        )
    return visits


def build_events(visits: list) -> list:
    by_site = defaultdict(list)
    for v in visits:
        by_site[v.site].append(v)

    # First pass: per (slot) value-change events.
    raw_events = []
    for site, site_visits in by_site.items():
        site_visits.sort(key=lambda v: v.visit_date)
        for slot_key, _header_name, brand in HUMIDITY_SLOTS:
            prev_value = None
            for v in site_visits:
                value = v.slot_values.get(slot_key, "")
                if not value:
                    continue
                if value != prev_value:
                    raw_events.append(
                        {
                            "site": site,
                            "slot": slot_key,
                            "brand": brand,
                            "date": v.visit_date,
                            "visit_type": v.visit_type,
                            "first_observed": prev_value is None,
                        }
                    )
                prev_value = value

    # Second pass: merge same-site/date/brand slot events (e.g. Lufft Upper + Lufft
    # Lower installed on the same visit) into one event, since a TOML sensor record
    # doesn't distinguish upper/lower.
    merged = {}
    for e in raw_events:
        key = (e["site"], e["date"], e["brand"])
        if key not in merged:
            merged[key] = Event(
                id=-1,
                site=e["site"],
                brand=e["brand"],
                slots=[e["slot"]],
                event_date=e["date"],
                visit_type=e["visit_type"],
                first_observed=e["first_observed"],
            )
        else:
            merged[key].slots.append(e["slot"])
            merged[key].first_observed = merged[key].first_observed and e["first_observed"]

    events = sorted(merged.values(), key=lambda e: (e.site, e.event_date))
    for i, e in enumerate(events):
        e.id = i
    return events


# --------------------------------------------------------------------------- #
# Station <-> Site mapping (interactive, cached to CSV)
# --------------------------------------------------------------------------- #

def load_mapping(path: Path) -> list:
    if not path.exists():
        return []
    with open(path, newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


def save_mapping(path: Path, rows: list) -> None:
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=MAPPING_FIELDS)
        writer.writeheader()
        writer.writerows(rows)


def resolve_sites_for_station(
    station_code: str,
    xlsx_sites: list,
    mapping_rows: list,
    mapping_path: Path,
    interactive: bool,
) -> list:
    existing = [r for r in mapping_rows if r["TomlStation"] == station_code]
    confirmed = [r["ExcelSite"] for r in existing if r["Confirmed"] == "Y"]
    if confirmed:
        return confirmed

    declined = {r["ExcelSite"] for r in existing if r["Confirmed"] == "N"}

    if station_code in xlsx_sites and station_code not in declined:
        mapping_rows.append({"TomlStation": station_code, "ExcelSite": station_code, "Confirmed": "Y"})
        save_mapping(mapping_path, mapping_rows)
        return [station_code]

    if not interactive:
        return []

    def similarity(a: str, b: str) -> float:
        return difflib.SequenceMatcher(None, a.upper(), b.upper()).ratio()

    candidates = [s for s in xlsx_sites if s not in declined]
    candidates.sort(key=lambda s: similarity(station_code, s), reverse=True)
    candidates = candidates[:8]

    if not candidates:
        print(f"\nNo remaining candidate Excel 'Site' codes for TOML station '{station_code}'. Skipping.")
        return []

    print(f"\nNo confirmed Excel 'Site' mapping for TOML station '{station_code}'.")
    print("Closest Excel Site codes (pick one or more, e.g. for stations with an upper/lower logger):")
    for i, c in enumerate(candidates, 1):
        print(f"  {i}. {c}")
    print("  0. none of these")
    raw = input("Select site number(s), comma-separated, or 0: ").strip()

    chosen = []
    if raw and raw != "0":
        for token in raw.split(","):
            token = token.strip()
            if token.isdigit() and 1 <= int(token) <= len(candidates):
                chosen.append(candidates[int(token) - 1])

    for c in candidates:
        mapping_rows.append(
            {"TomlStation": station_code, "ExcelSite": c, "Confirmed": "Y" if c in chosen else "N"}
        )
    save_mapping(mapping_path, mapping_rows)
    return chosen


# --------------------------------------------------------------------------- #
# TOML side
# --------------------------------------------------------------------------- #

def parse_toml_date(value: str) -> Optional[date]:
    text = (value or "").strip()
    if text.upper() in NA_TOKENS:
        return None
    try:
        return date.fromisoformat(text)
    except ValueError:
        return None


def confidence_key(sensor) -> str:
    if "date_confidence" in sensor:
        return "date_confidence"
    return "parsing_confidence"


def find_candidates(sensor_date: date, events: list, used_ids: set, tolerance_days: int) -> list:
    candidates = []
    for e in events:
        if e.id in used_ids:
            continue
        delta = abs((e.event_date - sensor_date).days)
        if delta <= tolerance_days or e.event_date.year == sensor_date.year:
            candidates.append((delta, e))
    candidates.sort(key=lambda x: x[0])
    return candidates


def set_with_note(table, key: str, value: str, note: str) -> None:
    table[key] = value
    try:
        table[key].comment(note)
    except Exception:
        pass  # tomlkit trivia API can vary by version; the value update itself is what matters.


def prompt_yes_no(question: str) -> bool:
    answer = input(f"{question} [y/N]: ").strip().lower()
    return answer == "y"


def process_station(
    toml_path: Path,
    station_events: list,
    used_event_ids: set,
    tolerance_days: int,
    interactive: bool,
    verbose: bool,
) -> None:
    text = toml_path.read_text(encoding="utf-8")
    doc = tomlkit.parse(text)
    sensors = doc.get("sensor", [])

    print(f"\n=== {toml_path.name} ===")
    if not station_events:
        print("  (no Excel visit events matched to this station -- mapping unresolved or no humidity data)")

    dirty = False
    today_note = date.today().isoformat()

    for sensor in sensors:
        name = sensor.get("name", "?")
        inst_date = parse_toml_date(sensor.get("installation_date", ""))
        toml_brand = str(sensor.get("instrument_type", "NA")).strip()

        if inst_date is None:
            print(f"  {name}: installation_date is NA/unparseable -- skipped")
            continue

        candidates = find_candidates(inst_date, station_events, used_event_ids, tolerance_days)
        if not candidates:
            print(f"  {name}: installation_date={inst_date} type={toml_brand} -> no matching Excel event found")
            continue

        delta, event = candidates[0]
        used_event_ids.add(event.id)
        brand_matches = event.brand.upper() == toml_brand.upper()
        date_matches = delta <= tolerance_days
        note_suffix = " (first Excel record for this site -- may predate the visit)" if event.first_observed else ""

        if brand_matches and date_matches:
            if verbose:
                print(
                    f"  {name}: OK -- matches Excel visit {event.event_date} at {event.site} "
                    f"({event.brand}, {'/'.join(event.slots)}), {delta} day(s) off{note_suffix}"
                )
            continue

        print(f"  {name}: DISCREPANCY")
        print(f"    TOML:  installation_date={inst_date}  instrument_type={toml_brand}")
        print(
            f"    Excel: visit {event.event_date} at site {event.site}  "
            f"instrument={event.brand} ({'/'.join(event.slots)})  visit_type={event.visit_type}{note_suffix}"
        )

        if not interactive:
            continue

        if not date_matches:
            if prompt_yes_no(f"    Update installation_date {inst_date} -> {event.event_date}?"):
                set_with_note(
                    sensor,
                    "installation_date",
                    event.event_date.isoformat(),
                    f"date corrected from {inst_date} per Station and Instrument Status.xlsx "
                    f"({event.site} visit, {today_note})",
                )
                sensor[confidence_key(sensor)] = (
                    f"corrected from Station and Instrument Status.xlsx ({event.site} visit "
                    f"{event.event_date}) on {today_note}"
                )
                dirty = True

        if not brand_matches:
            if prompt_yes_no(f"    Update instrument_type {toml_brand!r} -> {event.brand!r}?"):
                set_with_note(
                    sensor,
                    "instrument_type",
                    event.brand,
                    f"type corrected from '{toml_brand}' per Station and Instrument Status.xlsx "
                    f"({event.site} visit, {today_note})",
                )
                dirty = True

    # Events for this station with no matching TOML sensor at all.
    unmatched = [e for e in station_events if e.id not in used_event_ids]
    if unmatched:
        print("  Excel events with no corresponding TOML sensor record:")
        for event in sorted(unmatched, key=lambda e: e.event_date):
            print(
                f"    {event.event_date} at {event.site}: {event.brand} ({'/'.join(event.slots)}), "
                f"visit_type={event.visit_type}"
            )
            if not interactive:
                continue
            if not prompt_yes_no("    Append a new [[sensor]] entry for this event?"):
                continue

            new_name = f"instrument_{len(sensors) + 1}"
            new_sensor = tomlkit.table()
            new_sensor["name"] = new_name
            new_sensor["installation_date"] = event.event_date.isoformat()
            new_sensor["decommission_date"] = "NA"
            new_sensor["instrument_type"] = event.brand
            new_sensor["serial_number"] = "NA"
            new_sensor["source_file"] = "Station and Instrument Status.xlsx"
            new_sensor["parsing_confidence"] = (
                f"added from Station and Instrument Status.xlsx ({event.site} visit "
                f"{event.event_date}, {event.visit_type}) via compare_status_with_toml.py on "
                f"{today_note}; please verify"
            )

            if sensors:
                last_sensor = sensors[-1]
                if str(last_sensor.get("decommission_date", "")).strip().upper() in NA_TOKENS:
                    if prompt_yes_no(
                        f"    Also set decommission_date of '{last_sensor.get('name')}' to {event.event_date}?"
                    ):
                        set_with_note(
                            last_sensor,
                            "decommission_date",
                            event.event_date.isoformat(),
                            f"closed out per new {new_name} added {today_note}",
                        )

            doc["sensor"].append(new_sensor)
            sensors = doc["sensor"]
            dirty = True

    if dirty:
        toml_path.write_text(tomlkit.dumps(doc), encoding="utf-8")
        print(f"  -> wrote changes to {toml_path}")


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #

def parse_args():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--xlsx", type=Path, default=DEFAULT_XLSX, help="Path to the status workbook")
    p.add_argument("--stations-dir", type=Path, default=DEFAULT_STATIONS_DIR, help="Folder of station TOML files")
    p.add_argument("--mapping-file", type=Path, default=DEFAULT_MAPPING_FILE, help="Station<->Site mapping cache")
    p.add_argument("--station", help="Only process one station (TOML filename stem, e.g. CEN)")
    p.add_argument("--tolerance-days", type=int, default=60, help="Date tolerance for matching a visit to a TOML sensor record")
    p.add_argument("--report-only", action="store_true", help="Never prompt or write; just print the comparison")
    p.add_argument("--verbose", action="store_true", help="Also print sensors that match cleanly")
    return p.parse_args()


def main():
    args = parse_args()

    if not args.xlsx.exists():
        sys.exit(f"Excel file not found: {args.xlsx}")
    if not args.stations_dir.exists():
        sys.exit(f"Stations folder not found: {args.stations_dir}")

    visits = load_visits(args.xlsx)
    events = build_events(visits)
    xlsx_sites = sorted({v.site for v in visits})
    mapping_rows = load_mapping(args.mapping_file)

    toml_files = sorted(args.stations_dir.glob("*.toml"))
    if args.station:
        toml_files = [f for f in toml_files if f.stem.upper() == args.station.upper()]
        if not toml_files:
            sys.exit(f"No TOML file found for station '{args.station}' in {args.stations_dir}")

    interactive = not args.report_only
    used_event_ids: set = set()

    for toml_path in toml_files:
        station_code = toml_path.stem
        sites = resolve_sites_for_station(station_code, xlsx_sites, mapping_rows, args.mapping_file, interactive)
        station_events = [e for e in events if e.site in sites]
        process_station(toml_path, station_events, used_event_ids, args.tolerance_days, interactive, args.verbose)

    print("\nDone.")


if __name__ == "__main__":
    main()
