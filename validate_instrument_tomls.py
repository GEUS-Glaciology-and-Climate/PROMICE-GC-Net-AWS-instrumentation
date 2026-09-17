from pathlib import Path
from datetime import datetime
import sys

try:
    import tomllib
except ModuleNotFoundError:
    import tomli as tomllib

TOML_FOLDER = Path("../PROMICE-GC-Net-AWS-instrumentation/stations")


def parse_date(value):
    if value is None or str(value).strip().upper() == "NA":
        return None
    try:
        return datetime.strptime(str(value).strip(), "%Y-%m-%d")
    except ValueError:
        raise ValueError(f"Invalid date: {value}")


def validate_toml(path):
    print("\n" + "=" * 70)
    print(path.name)
    print("=" * 70)

    with open(path, "rb") as f:
        config = tomllib.load(f)

    sensors = config.get("sensor", [])
    if not sensors:
        print("ERROR: no [[sensor]] entries")
        return False

    valid, previous_end = True, None

    for i, sensor in enumerate(sensors, 1):
        name = sensor.get("name", f"instrument_{i}")
        station_id = sensor.get("station_id", path.stem)
        installation_raw = sensor.get("installation_date", "NA")
        decommission_raw = sensor.get("decommission_date", "NA")
        instrument_type = sensor.get("instrument_type", "NA")
        serial_number = sensor.get("serial_number", "NA")
        source_file = sensor.get("source_file", "NA")
        confidence = str(sensor.get("parsing_confidence", "NA")).lower()

        print(f"\n{name}")
        print(f"  station_id:        {station_id}")
        print(f"  installation:      {installation_raw}")
        print(f"  decommission:      {decommission_raw}")
        print(f"  instrument_type:   {instrument_type}")
        print(f"  serial_number:     {serial_number}")

        # parse and validate dates
        try:
            installation = parse_date(installation_raw)
            decommission = parse_date(decommission_raw)
        except ValueError as error:
            print(f"  ERROR: {error}")
            valid = False
            continue

        # check installation/decommission period
        if installation is None:
            print("  WARNING: installation date unresolved")
        if installation is not None and decommission is not None and decommission <= installation:
            print("  ERROR: decommission date is not after installation date")
            valid = False
        if previous_end is not None and installation is not None and installation < previous_end:
            print("  WARNING: period overlaps previous instrument")
        if decommission is not None:
            previous_end = decommission

        # flag records needing manual review or without source information
        if "review_required" in confidence:
            print("  REVIEW REQUIRED")
        if source_file is None or str(source_file).strip() in ("", "NA"):
            print("  WARNING: source_file missing")

    return valid


def main():
    if not TOML_FOLDER.exists():
        sys.exit(f"Folder not found: {TOML_FOLDER}")

    toml_files = sorted(TOML_FOLDER.glob("*.toml"))
    if not toml_files:
        sys.exit("No TOML files found.")

    print(f"\nTOML files found: {len(toml_files)}")
    failed = []

    for path in toml_files:
        try:
            if not validate_toml(path):
                failed.append(path.name)
        except Exception as error:
            print(f"\nERROR reading {path.name}:\n{error}")
            failed.append(path.name)

    print("\n" + "=" * 70)
    print("VALIDATION SUMMARY")
    print("=" * 70)
    print(f"Total TOMLs: {len(toml_files)}")
    print(f"Failed: {len(failed)}")

    if failed:
        print("\nFiles requiring attention:")
        for filename in failed:
            print(" -", filename)
    else:
        print("\nAll TOML files parsed successfully.")


if __name__ == "__main__":
    main()
