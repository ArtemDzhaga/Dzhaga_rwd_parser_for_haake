#!/usr/bin/env python3
"""
CLI для локальной пакетной обработки Thermo HAAKE RheoWin .rwd файлов.

Вход:
- один .rwd файл, пример ниже:
cd /Users/artem/Desktop/Output/iam_ras/haake_parser
python3 haake_cli.py "32th_pein_15_steps_60C_1hz_max_ v2.rwd" -o result.xlsx

- несколько .rwd файлов, пример ниже:
python3 haake_cli.py "file_1.rwd" "file_2.rwd" "file_3.rwd" -o result.xlsx

- папка с .rwd файлами, пример ниже:
python3 haake_cli.py "/путь/к/папке/с/rwd" --recursive -o result.xlsx

Выход:
- одна Excel-книга .xlsx со сводкой по всем экспериментам и отдельным листом
  с метаданными прибора.
"""

from __future__ import annotations

import argparse
import math
import re
import struct
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


RWD_MAGIC = b"\x88\x88\x00\x00"
DEFAULT_OUTPUT_NAME = "haake_rwd_report.xlsx"
TEXT_RE = re.compile(rb"[\x09\x0a\x0d\x20-\x7e]{4,}")


@dataclass
class RwdRecord:
    path: Path
    size_bytes: int
    format_magic: str
    rheowin_version: str
    date: str
    time: str
    device: str
    instrument_model: str
    geometry: str
    temperature_controller: str
    gap: str
    job_file: str
    method_element: str
    mode: str
    serial_numbers: str
    driver_versions: str
    firmware_versions: str
    strings: list[tuple[int, str]]
    measurements: list[dict[str, Any]]
    numeric_blocks: list[dict[str, Any]]


NUMERIC_RECORD_MARKER = bytes.fromhex("da00da00feff")

CHANNEL_MAP = {
    (7, 4): "time_s",
    (7, 5): "time_start_s",
    (7, 8): "time_duplicate_s",
    (2, 0): "tau_pa",
    (16, 0): "gamma",
    (12, 0): "gamma_aux",
    (14, 0): "signal_aux",
    (4, 200): "temperature_k",
    (4, 201): "temperature_k_duplicate",
    (15, 0): "frequency_hz",
}


def extract_ascii_strings(blob: bytes) -> list[tuple[int, str]]:
    rows: list[tuple[int, str]] = []
    for match in TEXT_RE.finditer(blob):
        text = match.group().decode("latin1", errors="replace").strip()
        if text:
            rows.append((match.start(), text))
    return rows


def extract_numeric_blocks(blob: bytes) -> list[dict[str, Any]]:
    blocks: list[dict[str, Any]] = []
    start = 0
    while True:
        offset = blob.find(NUMERIC_RECORD_MARKER, start)
        if offset < 0:
            break
        start = offset + 1

        if offset + 60 > len(blob):
            continue

        code_1 = struct.unpack_from("<H", blob, offset + 6)[0]
        code_2 = struct.unpack_from("<H", blob, offset + 8)[0]
        dtype = struct.unpack_from("<I", blob, offset + 14)[0]
        count = struct.unpack_from("<I", blob, offset + 18)[0]

        if dtype != 4 or count <= 0 or count > 10000:
            continue

        data_offset = offset + 60
        data_end = data_offset + count * 4
        if data_end > len(blob):
            continue

        values = [struct.unpack_from("<f", blob, data_offset + index * 4)[0] for index in range(count)]
        if not all(math.isfinite(value) for value in values):
            continue

        blocks.append(
            {
                "offset": offset,
                "code_1": code_1,
                "code_2": code_2,
                "count": count,
                "channel": CHANNEL_MAP.get((code_1, code_2), f"unknown_{code_1}_{code_2}"),
                "values": values,
            }
        )

    return blocks


def looks_like_measurement_block(block: dict[str, Any]) -> bool:
    values = block["values"]
    if not values:
        return False
    if max(abs(value) for value in values) >= 1e20:
        return False
    return (block["code_1"], block["code_2"]) in CHANNEL_MAP


def extract_measurements(blocks: list[dict[str, Any]]) -> list[dict[str, Any]]:
    useful_blocks = [block for block in blocks if looks_like_measurement_block(block)]
    if not useful_blocks:
        return []

    row_count = max(block["count"] for block in useful_blocks)
    rows: list[dict[str, Any]] = [{"point": index + 1} for index in range(row_count)]

    for block in useful_blocks:
        channel = block["channel"]
        for index, value in enumerate(block["values"]):
            rows[index][channel] = clean_number(value)

    for row in rows:
        temperature_k = row.get("temperature_k")
        if isinstance(temperature_k, (int, float)):
            row["temperature_c"] = clean_number(temperature_k - 273.15)

        tau = row.get("tau_pa")
        gamma = row.get("gamma")
        if isinstance(tau, (int, float)) and isinstance(gamma, (int, float)) and gamma != 0:
            row["complex_modulus_pa_est"] = clean_number(tau / gamma)

    return rows


def clean_number(value: float) -> float | None:
    if not math.isfinite(value) or abs(value) >= 1e20:
        return None
    if abs(value) < 1e-12:
        return 0.0
    return float(value)


def unique_join(values: list[str]) -> str:
    seen: set[str] = set()
    out: list[str] = []
    for value in values:
        cleaned = value.strip()
        if cleaned and cleaned not in seen:
            seen.add(cleaned)
            out.append(cleaned)
    return " | ".join(out)


def first_match(pattern: str, text: str, flags: int = 0) -> str:
    match = re.search(pattern, text, flags)
    return match.group(1).strip() if match else ""


def all_matches(pattern: str, text: str, flags: int = 0) -> list[str]:
    return [match.strip() for match in re.findall(pattern, text, flags) if match.strip()]


def parse_pipe_metadata(text: str, key: str) -> list[str]:
    escaped = re.escape(key)
    return all_matches(rf"{escaped}:([^|$\n\r]+)", text)


def detect_device(lines: list[str], text: str) -> str:
    candidates: list[str] = []
    for line in lines:
        if "\\" in line or ".pdf" in line.lower() or ".rwj" in line.lower():
            continue
        if re.search(r"\bRheoStress\b|\bRheoWin\b|\bRS1500?\b", line, re.I):
            if not line.startswith("***") and len(line) < 120:
                candidates.append(line)
    candidates.extend(all_matches(r"(RheoStress\s+RS\d+)", text, re.I))
    candidates.extend(all_matches(r"\b(RS\d{2,4})\b", text))
    return unique_join(candidates[:8])


def detect_geometry(lines: list[str], text: str) -> str:
    candidates: list[str] = []
    for line in lines:
        if re.search(r"\b(PP|CP|DG|Z|C)\d+\b|plate|cone|sensor", line, re.I):
            if len(line) < 120:
                candidates.append(line)
    candidates.extend(all_matches(r"\b(PP\d+\s*[A-Z0-9 ]*)", text))
    return unique_join(candidates[:8])


def extract_rwd_record(path: Path) -> RwdRecord:
    blob = path.read_bytes()
    numeric_blocks = extract_numeric_blocks(blob)
    strings = extract_ascii_strings(blob)
    lines = [text for _, text in strings]
    full_text = "\n".join(lines)

    version = first_match(r"HAAKE RheoWin\s+([0-9.]+)", full_text)
    date_time = re.search(r"Date/Time:\s*([0-9.]+)\s*/\s*([0-9:]+)", full_text)
    if date_time:
        date = date_time.group(1)
        time = date_time.group(2)
    else:
        date = first_match(r"\b([0-3]\d\.[01]\d\.\d{4})\b", full_text)
        time = first_match(r"\b([0-2]\d:[0-5]\d:[0-5]\d)\b", full_text)

    job_files = all_matches(r"([A-Z]:\\[^\n\r]+?\.rwj)", full_text, re.I)
    mode = first_match(r"ElmStros::Execute:\s*Mode:\s*([^\n\r]+)", full_text)
    method_element = unique_join(all_matches(r"\b(ELM[A-Z0-9_]+)\b", full_text))

    serials = parse_pipe_metadata(full_text, "Serial number")
    driver_versions = parse_pipe_metadata(full_text, "Driver version")
    firmware_versions = parse_pipe_metadata(full_text, "Firmware version 1")
    firmware_versions.extend(parse_pipe_metadata(full_text, "Firmware version 3"))

    return RwdRecord(
        path=path,
        size_bytes=len(blob),
        format_magic="OOP_HADES" if b"OOP_HADES" in blob[:128] or b"OOP_HADES" in blob else blob[:16].hex(" "),
        rheowin_version=version,
        date=date,
        time=time,
        device=detect_device(lines, full_text),
        instrument_model=unique_join(all_matches(r"\b(RheoStress\s+RS\d+|RS1500?|DC50)\b", full_text, re.I)),
        geometry=detect_geometry(lines, full_text),
        temperature_controller=unique_join(all_matches(r"\b(DC50\b[^\n\r]*)", full_text)),
        gap=first_match(r"Gap:\s*([^\n\r]+)", full_text),
        job_file=unique_join(job_files),
        method_element=method_element,
        mode=mode,
        serial_numbers=unique_join(serials),
        driver_versions=unique_join(driver_versions),
        firmware_versions=unique_join(firmware_versions),
        strings=strings,
        measurements=extract_measurements(numeric_blocks),
        numeric_blocks=numeric_blocks,
    )


def collect_rwd_files(paths: list[Path], recursive: bool) -> list[Path]:
    files: list[Path] = []
    for path in paths:
        expanded = path.expanduser()
        if expanded.is_dir():
            pattern = "**/*.rwd" if recursive else "*.rwd"
            files.extend(expanded.glob(pattern))
        elif expanded.is_file() and expanded.suffix.lower() == ".rwd":
            files.append(expanded)
    return sorted(set(file.resolve() for file in files))


def style_header(ws) -> None:
    fill = PatternFill("solid", fgColor="1F4E78")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def append_table(ws, headers: list[str], rows: list[list[object]]) -> None:
    ws.append(headers)
    for row in rows:
        ws.append(row)
    style_header(ws)
    autosize(ws)


def autosize(ws) -> None:
    for col_idx in range(1, ws.max_column + 1):
        letter = get_column_letter(col_idx)
        width = 10
        for row_idx in range(1, min(ws.max_row, 500) + 1):
            value = ws.cell(row_idx, col_idx).value
            if value is not None:
                width = max(width, min(70, len(str(value)) + 2))
        ws.column_dimensions[letter].width = width


def write_workbook(records: list[RwdRecord], output_path: Path, include_diagnostics: bool) -> None:
    records = sorted(records, key=lambda record: record.path.name.lower())
    wb = Workbook()
    ws = wb.active
    ws.title = "data"

    data_headers = [
        "file_name",
        "point",
        "time_s",
        "time_start_s",
        "tau_pa",
        "gamma",
        "complex_modulus_pa_est",
        "frequency_hz",
        "temperature_c",
        "temperature_k",
        "gamma_aux",
        "signal_aux",
    ]
    data_rows: list[list[object]] = []
    for record in records:
        for measurement in record.measurements:
            data_rows.append([record.path.name] + [measurement.get(header) for header in data_headers[1:]])
    append_table(ws, data_headers, data_rows)

    ws = wb.create_sheet("instrument_metadata")
    metadata_rows = build_transposed_metadata(records)
    append_table(ws, ["field"] + [record.path.name for record in records], metadata_rows)

    ws = wb.create_sheet("logs")
    log_rows: list[list[object]] = []
    for record in records:
        for offset, line in iter_text_lines(record.strings):
            if is_log_line(line):
                log_rows.append([record.path.name, offset, line])
    append_table(ws, ["file_name", "offset_dec", "log_line"], log_rows)

    if include_diagnostics:
        ws = wb.create_sheet("numeric_blocks")
        rows: list[list[object]] = []
        for record in records:
            for block in record.numeric_blocks:
                values = ", ".join("" if clean_number(value) is None else f"{value:.8g}" for value in block["values"][:20])
                rows.append(
                    [
                        record.path.name,
                        hex(block["offset"]),
                        block["code_1"],
                        block["code_2"],
                        block["channel"],
                        block["count"],
                        values,
                    ]
                )
        append_table(
            ws,
            ["file_name", "offset_hex", "code_1", "code_2", "channel", "count", "first_values"],
            rows,
        )

    wb.save(output_path)


def build_transposed_metadata(records: list[RwdRecord]) -> list[list[object]]:
    fields = [
        ("file_path", lambda record: str(record.path)),
        ("size_bytes", lambda record: record.size_bytes),
        ("format", lambda record: record.format_magic),
        ("rheowin_version", lambda record: record.rheowin_version),
        ("date", lambda record: record.date),
        ("time", lambda record: record.time),
        ("instrument_model", lambda record: record.instrument_model),
        ("device", lambda record: record.device),
        ("geometry", lambda record: record.geometry),
        ("temperature_controller", lambda record: record.temperature_controller),
        ("gap", lambda record: record.gap),
        ("serial_numbers", lambda record: record.serial_numbers),
        ("driver_versions", lambda record: record.driver_versions),
        ("firmware_versions", lambda record: record.firmware_versions),
        ("job_file", lambda record: record.job_file),
        ("mode", lambda record: record.mode),
        ("measurement_points_extracted", lambda record: len(record.measurements)),
    ]
    return [[field_name] + [getter(record) for record in records] for field_name, getter in fields]


def iter_text_lines(strings: list[tuple[int, str]]) -> list[tuple[int, str]]:
    rows: list[tuple[int, str]] = []
    for offset, text in strings:
        running_offset = offset
        for line in text.splitlines():
            cleaned = line.strip()
            if cleaned:
                rows.append((running_offset, cleaned))
            running_offset += len(line) + 1
    return rows


def is_log_line(text: str) -> bool:
    return bool(
        re.search(r"^\d{2}:\d{2}:\d{2}", text)
        or text.startswith("***")
        or "AUTOSAVE" in text
        or "Execute" in text
        or "Begin" in text
        or "End" in text
    )


def choose_folder_with_gui() -> Path:
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()
    selected = filedialog.askdirectory(title="Choose folder with HAAKE .rwd files")
    root.destroy()
    if not selected:
        raise SystemExit("No folder selected.")
    return Path(selected)


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Batch-extract metadata from HAAKE RheoWin .rwd files to one Excel workbook."
    )
    parser.add_argument(
        "paths",
        nargs="*",
        type=Path,
        help="Folder(s) or .rwd file(s). You can drag files/folders into the terminal here.",
    )
    parser.add_argument("-o", "--output", type=Path, help=f"Output .xlsx path. Default: {DEFAULT_OUTPUT_NAME}")
    parser.add_argument("-r", "--recursive", action="store_true", help="Scan folders recursively.")
    parser.add_argument("--gui", action="store_true", help="Open a local folder picker instead of typing a path.")
    parser.add_argument(
        "--include-diagnostics",
        action="store_true",
        help="Include heuristic binary float32 runs for reverse-engineering diagnostics.",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv or sys.argv[1:])
    paths: list[Path] = args.paths
    if args.gui:
        paths = [choose_folder_with_gui()]
    if not paths:
        print("Error: pass a folder/.rwd file or use --gui.", file=sys.stderr)
        return 2

    files = collect_rwd_files(paths, recursive=args.recursive)
    if not files:
        print("Error: no .rwd files found.", file=sys.stderr)
        return 1

    records = [extract_rwd_record(path) for path in files]
    if args.output:
        output_path = args.output.expanduser().resolve()
    else:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        base_dir = paths[0].expanduser().resolve() if paths[0].expanduser().is_dir() else Path.cwd()
        output_path = base_dir / f"haake_rwd_report_{timestamp}.xlsx"

    write_workbook(records, output_path, include_diagnostics=args.include_diagnostics)
    print(f"Processed {len(records)} .rwd file(s)")
    print(f"Wrote {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
