#!/usr/bin/env python3
"""Fill only the Endpoint column of an NPI spreadsheet.

The value written is the "Endpoint" cell from the NPPES provider page
(https://npiregistry.cms.hhs.gov/provider-view/<npi>), inside the Health
Information Exchange table whose class is "table table-bordered table-sm".
That cell is the public NPPES field endpoints[].endpoint, read here from
the NPPES Read API (version 2.1), which is the supported way to look up
the same record. Name, NPI, phone, address, and every other column are
left unchanged.
"""

from __future__ import annotations

import argparse
import csv
import json
import sys
import threading
import time
import urllib.error
import urllib.request
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

from openpyxl import Workbook, load_workbook

API_URL = "https://npiregistry.cms.hhs.gov/api/?version=2.1&number={npi}"
USER_AGENT = "chalet-npi-endpoint-filler/1.0"

NPI_HEADERS = {
    "npi",
    "npi number",
    "npi no",
    "npi #",
    "national provider identifier",
    "provider npi",
}
ENDPOINT_HEADERS = {
    "endpoint",
    "endpoints",
    "endpoint address",
    "hie endpoint",
}


class LookupError(Exception):
    """The registry could not be read for this NPI."""


def normalize_header(value) -> str:
    text = "" if value is None else str(value)
    return " ".join(text.strip().lower().replace("#", " # ").split())


def normalize_npi(value) -> str | None:
    """Return a 10-digit NPI, or None when the cell is not an NPI."""
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, int):
        text = str(value)
    elif isinstance(value, float):
        if value != value:  # NaN
            return None
        if not value.is_integer():
            return None
        text = str(int(value))
    else:
        text = str(value).strip()
        if not text:
            return None
        # Excel often stores NPIs as floats and writes them in scientific notation.
        if any(ch in text for ch in ".eE") and not any(ch.isalpha() and ch not in "eE" for ch in text):
            try:
                number = float(text)
            except ValueError:
                number = None
            if number is not None and number.is_integer():
                text = str(int(number))
    digits = "".join(ch for ch in text if ch.isdigit())
    if len(digits) == 10:
        return digits
    if digits.isdigit() and 0 < len(digits) < 10 and text.replace(".", "", 1).isdigit():
        padded = f"{int(float(text)):010d}" if "." in text or "e" in text.lower() else digits.zfill(10)
        if len(padded) == 10:
            return padded
    return None


def find_columns(headers: list) -> tuple[int, int]:
    npi_index = None
    endpoint_index = None
    for index, header in enumerate(headers):
        name = normalize_header(header)
        if npi_index is None and name in NPI_HEADERS:
            npi_index = index
        if endpoint_index is None and name in ENDPOINT_HEADERS:
            endpoint_index = index
    if npi_index is None:
        raise SystemExit("Could not find an NPI column. Expected a header such as 'NPI'.")
    if endpoint_index is None:
        raise SystemExit(
            "Could not find an Endpoint column. Expected a header such as 'Endpoint' or 'endpoints'."
        )
    return npi_index, endpoint_index


def endpoints_from_payload(payload: dict, npi: str) -> list[str]:
    """Pull the Endpoint column values for an exact NPI match."""
    found: list[str] = []
    for result in payload.get("results") or []:
        number = str(result.get("number") or "")
        if number and number != npi:
            continue
        for item in result.get("endpoints") or []:
            value = str(item.get("endpoint") or "").strip()
            if value and value not in found:
                found.append(value)
    return found


def fetch_endpoints(npi: str, timeout: float = 30.0, attempts: int = 5) -> list[str]:
    """Return Endpoint addresses for one NPI. Raises LookupError on failure."""
    url = API_URL.format(npi=npi)
    request = urllib.request.Request(
        url,
        headers={"Accept": "application/json", "User-Agent": USER_AGENT},
    )
    last_error: Exception | None = None
    for attempt in range(attempts):
        try:
            with urllib.request.urlopen(request, timeout=timeout) as response:
                payload = json.load(response)
            if not isinstance(payload, dict):
                raise LookupError(f"{npi}: registry returned an unexpected response")
            return endpoints_from_payload(payload, npi)
        except LookupError:
            raise
        except urllib.error.HTTPError as error:
            last_error = error
            if error.code not in {429, 500, 502, 503, 504} or attempt == attempts - 1:
                break
        except (urllib.error.URLError, TimeoutError, json.JSONDecodeError, OSError) as error:
            last_error = error
            if attempt == attempts - 1:
                break
        time.sleep(min(30.0, 1.5 * (attempt + 1)))
    raise LookupError(f"{npi}: {last_error}")


def format_endpoints(values: list[str], separator: str) -> str:
    return separator.join(values)


def read_table(path: Path, sheet_name: str | None) -> tuple[str | None, list[list]]:
    if path.suffix.lower() == ".csv":
        with path.open(newline="", encoding="utf-8-sig") as handle:
            rows = [list(row) for row in csv.reader(handle)]
        if not rows:
            raise SystemExit(f"{path} is empty.")
        return None, rows

    if path.suffix.lower() not in {".xlsx", ".xlsm"}:
        raise SystemExit("Input must be a .csv, .xlsx, or .xlsm file.")

    workbook = load_workbook(path, read_only=True, data_only=False)
    try:
        if sheet_name:
            if sheet_name not in workbook.sheetnames:
                raise SystemExit(
                    f"Sheet {sheet_name!r} was not found. Sheets: {', '.join(workbook.sheetnames)}"
                )
            worksheet = workbook[sheet_name]
            chosen = sheet_name
        else:
            chosen = None
            worksheet = None
            for name in workbook.sheetnames:
                candidate = workbook[name]
                header = next(candidate.iter_rows(min_row=1, max_row=1, values_only=True), None)
                if header and any(normalize_header(cell) in ENDPOINT_HEADERS for cell in header):
                    worksheet = candidate
                    chosen = name
                    break
            if worksheet is None:
                worksheet = workbook[workbook.sheetnames[0]]
                chosen = workbook.sheetnames[0]
        rows = [list(row) for row in worksheet.iter_rows(values_only=True)]
    finally:
        workbook.close()
    if not rows:
        raise SystemExit(f"{path} has no rows on sheet {chosen}.")
    return chosen, rows


def lookup_unique(
    npis: list[str],
    workers: int,
    cache_path: Path | None,
    pause: float,
) -> tuple[dict[str, list[str]], dict[str, str]]:
    cache: dict[str, list[str]] = {}
    if cache_path and cache_path.exists():
        stored = json.loads(cache_path.read_text(encoding="utf-8"))
        cache = {key: list(value) for key, value in stored.items()}

    pending = [npi for npi in dict.fromkeys(npis) if npi not in cache]
    errors: dict[str, str] = {}
    lock = threading.Lock()
    done = 0
    total = len(pending)

    def save() -> None:
        if cache_path is None:
            return
        cache_path.parent.mkdir(parents=True, exist_ok=True)
        temporary = cache_path.with_suffix(cache_path.suffix + ".tmp")
        temporary.write_text(json.dumps(cache, indent=2, sort_keys=True), encoding="utf-8")
        temporary.replace(cache_path)

    def one(npi: str) -> tuple[str, list[str] | None, str | None]:
        if pause:
            time.sleep(pause)
        try:
            return npi, fetch_endpoints(npi), None
        except LookupError as error:
            return npi, None, str(error)

    if not pending:
        return cache, errors

    with ThreadPoolExecutor(max_workers=max(1, workers)) as pool:
        futures = [pool.submit(one, npi) for npi in pending]
        for future in as_completed(futures):
            npi, values, error = future.result()
            with lock:
                done += 1
                if error:
                    errors[npi] = error
                    print(f"[{done}/{total}] {npi} FAILED {error}", file=sys.stderr)
                else:
                    cache[npi] = values or []
                    if done % 25 == 0 or done == total:
                        save()
                        print(f"[{done}/{total}] looked up {npi}", file=sys.stderr)
    save()
    return cache, errors


def fill_rows(
    rows: list[list],
    npi_index: int,
    endpoint_index: int,
    looked_up: dict[str, list[str]],
    errors: dict[str, str],
    separator: str,
) -> dict[str, int]:
    """Write Endpoint values. Returns counts. Other cells are not replaced."""
    stats = {
        "rows": 0,
        "filled": 0,
        "blank": 0,
        "skipped_no_npi": 0,
        "skipped_error": 0,
        "unchanged": 0,
    }
    width = max(len(row) for row in rows)
    for row in rows:
        if len(row) < width:
            row.extend([None] * (width - len(row)))

    for row in rows[1:]:
        stats["rows"] += 1
        npi = normalize_npi(row[npi_index] if npi_index < len(row) else None)
        if npi is None:
            stats["skipped_no_npi"] += 1
            continue
        if npi in errors or npi not in looked_up:
            stats["skipped_error"] += 1
            continue
        new_value = format_endpoints(looked_up[npi], separator)
        current = row[endpoint_index]
        current_text = "" if current is None else str(current)
        if current_text == new_value:
            stats["unchanged"] += 1
        row[endpoint_index] = new_value
        if new_value:
            stats["filled"] += 1
        else:
            stats["blank"] += 1
    return stats


def write_table(path: Path, rows: list[list], sheet_name: str, npi_index: int) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.suffix.lower() == ".csv":
        with path.open("w", newline="", encoding="utf-8") as handle:
            writer = csv.writer(handle)
            for row in rows:
                writer.writerow(["" if cell is None else cell for cell in row])
        return

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = sheet_name[:31] or "Sheet1"
    text_columns = {npi_index}
    # Keep phone columns as text too when the header says phone.
    for index, header in enumerate(rows[0]):
        if "phone" in normalize_header(header):
            text_columns.add(index)
    for row_number, row in enumerate(rows, start=1):
        for column, value in enumerate(row, start=1):
            cell = worksheet.cell(row=row_number, column=column, value=None if value is None else value)
            if column - 1 in text_columns and row_number > 1 and value not in (None, ""):
                cell.number_format = "@"
                cell.value = str(value)
    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions
    workbook.save(path)


def run(
    input_path: Path,
    output_path: Path,
    sheet_name: str | None,
    workers: int,
    separator: str,
    cache_path: Path | None,
    pause: float,
) -> dict[str, int]:
    chosen_sheet, rows = read_table(input_path, sheet_name)
    npi_index, endpoint_index = find_columns(list(rows[0]))
    npis = []
    for row in rows[1:]:
        npi = normalize_npi(row[npi_index] if npi_index < len(row) else None)
        if npi:
            npis.append(npi)
    print(
        f"Sheet {chosen_sheet or 'csv'}: {len(rows) - 1} data rows, {len(set(npis))} unique NPIs",
        file=sys.stderr,
    )
    looked_up, errors = lookup_unique(npis, workers=workers, cache_path=cache_path, pause=pause)
    stats = fill_rows(rows, npi_index, endpoint_index, looked_up, errors, separator)
    stats["errors"] = len(errors)
    write_table(output_path, rows, chosen_sheet or "Clinics", npi_index)
    print(
        "Wrote {output}: {filled} rows with an endpoint, {blank} with none, "
        "{skipped_error} lookup errors, {skipped_no_npi} rows without an NPI".format(
            output=output_path, **stats
        ),
        file=sys.stderr,
    )
    return stats


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path, help="Spreadsheet (.xlsx, .xlsm, or .csv) that already has an NPI column")
    parser.add_argument("-o", "--output", type=Path, help="Where to write the filled workbook. Defaults to <input stem>_with_endpoints.xlsx")
    parser.add_argument("--sheet", help="Worksheet name. Defaults to the sheet that has an Endpoint column.")
    parser.add_argument("--workers", type=int, default=4, help="Parallel registry lookups (default: 4)")
    parser.add_argument("--separator", default="; ", help="How to join multiple Endpoint values (default: '; ')")
    parser.add_argument("--cache", type=Path, help="JSON cache of NPI -> endpoints, so a rerun skips finished lookups")
    parser.add_argument("--pause", type=float, default=0.0, help="Seconds to wait before each lookup")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    output = args.output
    if output is None:
        output = args.input.with_name(f"{args.input.stem}_with_endpoints.xlsx")
    run(
        input_path=args.input,
        output_path=output,
        sheet_name=args.sheet,
        workers=args.workers,
        separator=args.separator,
        cache_path=args.cache,
        pause=args.pause,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
