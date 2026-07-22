"""
Copies maintenance sheet files identified (by manual inspection) in \\geodata\\Ice\\FIELDWORK
into this repo's "maintenance sheets" folder, renamed to a standardized STATION_YEAR.ext
scheme, and writes a manifest of original path -> copy name.

Read-only with respect to FIELDWORK: only reads (shutil.copy2) are performed there; nothing
in FIELDWORK is written, moved, or deleted.

Input:  maintenance_sheets_classification.csv (Station/Year/Status/Reason decided by hand).
Output: "maintenance sheets/" (the copies) and "maintenance_sheets_manifest.csv" (the record).
"""

import csv
import shutil
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent
CLASSIFICATION_CSV = REPO_ROOT / "maintenance_sheets_classification.csv"
DEST_DIR = REPO_ROOT / "maintenance sheets"
MANIFEST_CSV = REPO_ROOT / "maintenance_sheets_manifest.csv"

FIELDNAMES = ["OriginalPath", "Station", "Year", "CopyName", "Status", "Reason"]


def main():
    DEST_DIR.mkdir(parents=True, exist_ok=True)

    with open(CLASSIFICATION_CSV, newline="", encoding="utf-8") as f:
        rows = list(csv.DictReader(f))

    manifest = []
    counts = {}

    for row in rows:
        print(row["SourcePath"])
        source_path = Path(row["SourcePath"])
        station = row["Station"]
        year = row["Year"]
        status = row["Status"]
        reason = row["Reason"]

        if status != "Copy":
            manifest.append({
                "OriginalPath": row["SourcePath"],
                "Station": station,
                "Year": year,
                "CopyName": "",
                "Status": status,
                "Reason": reason,
            })
            continue

        if not source_path.is_file():
            print(f"WARNING: source file not found, skipping: {source_path}")
            manifest.append({
                "OriginalPath": row["SourcePath"],
                "Station": station,
                "Year": year,
                "CopyName": "",
                "Status": "Skip-NotFound",
                "Reason": "File not found at copy time",
            })
            continue

        ext = source_path.suffix
        key = f"{station}_{year}"
        counts[key] = counts.get(key, 0) + 1
        n = counts[key]

        copy_name = f"{station}_{year}{ext}" if n == 1 else f"{station}_{year}_{n}{ext}"
        dest_path = DEST_DIR / copy_name

        shutil.copy2(source_path, dest_path)

        manifest.append({
            "OriginalPath": row["SourcePath"],
            "Station": station,
            "Year": year,
            "CopyName": copy_name,
            "Status": "Copied",
            "Reason": reason,
        })

    with open(MANIFEST_CSV, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=FIELDNAMES)
        writer.writeheader()
        writer.writerows(manifest)

    copied = sum(1 for m in manifest if m["Status"] == "Copied")
    skipped = sum(1 for m in manifest if m["Status"] != "Copied")
    print(f"Copied: {copied} file(s) into '{DEST_DIR}'")
    print(f"Skipped/flagged: {skipped} row(s) - see {MANIFEST_CSV} for reasons")


if __name__ == "__main__":
    main()
