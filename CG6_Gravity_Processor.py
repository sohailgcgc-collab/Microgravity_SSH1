#!/usr/bin/env python3
"""
CG-6 Gravity Data Processor
============================
Standalone GUI tool (Python standard library only) for processing Scintrex
CG-6 microgravity survey data.

Workflow
--------
1. Select one or more CG-6 .dat / ASCII files (tab or comma delimited,
   auto-detects the header row).
2. Select a coordinate file (.xlsx or .csv) with columns:
       Station ID, Easting, Northing, Elevation, Latitude, Longitude
3. Enter the base station ID and a screening range (mGal). Repeat
   readings at the same station that fall outside the range around the
   station's median are excluded.
4. The tool builds loops between successive base-station occupations,
   applies linear drift correction, then (relative-to-base) latitude
   correction, free-air correction, Bouguer slab correction, and computes
   the Bouguer gravity + relative Bouguer anomaly for every station.
5. Results are written as a .csv table plus a .json audit file recording
   every parameter used, in your chosen output folder.

No third-party packages are required — .xlsx files are read directly via
the zipfile + xml modules in the standard library.
"""

from __future__ import annotations

import csv
import json
import math
import queue
import re
import threading
import traceback
import zipfile
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree as ET

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# ---------------------------------------------------------------------------
# Physical constants
# ---------------------------------------------------------------------------
FREE_AIR_COEFF = 0.3086      # mGal / m   (free-air gradient)
BOUGUER_COEFF = 0.04193      # mGal / (m * g/cm3)
DEFAULT_DENSITY = 2.67       # g/cm3 (typical crustal density)

DATE_TIME_FORMATS = [
    "%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M:%S", "%m/%d/%Y %H:%M:%S",
    "%d/%m/%Y %H:%M:%S", "%Y-%m-%d %H:%M", "%m/%d/%Y %H:%M",
    "%d-%m-%Y %H:%M:%S", "%d/%m/%Y %H:%M",
]


class ProcessingError(Exception):
    """Raised for any recoverable problem with the input data."""


# ---------------------------------------------------------------------------
# Small helpers
# ---------------------------------------------------------------------------

def norm(value: str) -> str:
    """Normalise a header name for loose matching (case/space/punct-insensitive)."""
    return re.sub(r"[^a-z0-9]", "", str(value).strip().lower())


def pick_key(headers, choices, required=True):
    normalized = {norm(h): h for h in headers}
    for choice in choices:
        key = normalized.get(norm(choice))
        if key:
            return key
    if required:
        raise ProcessingError(f"Could not find any of {choices} among columns {headers}")
    return None


def as_float(value, field_name="value"):
    try:
        return float(str(value).strip())
    except (TypeError, ValueError):
        raise ProcessingError(f"Could not parse a number for {field_name}: {value!r}")


def parse_datetime(date_str, time_str):
    combined = f"{(date_str or '').strip()} {(time_str or '').strip()}".strip()
    for fmt in DATE_TIME_FORMATS:
        try:
            return datetime.strptime(combined, fmt)
        except ValueError:
            continue
    raise ProcessingError(f"Could not parse date/time: {date_str!r} {time_str!r}")


def normal_gravity(latitude_deg: float) -> float:
    """International Gravity Formula 1980 / Somigliana equation (mGal)."""
    phi = math.radians(latitude_deg)
    s2 = math.sin(phi) ** 2
    s2b = math.sin(2 * phi) ** 2
    return 978032.67715 * (1 + 0.0053024 * s2 - 0.0000058 * s2b)


# ---------------------------------------------------------------------------
# .xlsx reader (no external dependency — parses the OOXML directly)
# ---------------------------------------------------------------------------

_XLSX_NS = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def _excel_col_to_index(ref: str) -> int:
    col = re.match(r"[A-Z]+", ref).group(0)
    idx = 0
    for ch in col:
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx - 1


def read_xlsx_rows(path: Path):
    with zipfile.ZipFile(path) as zf:
        shared = []
        if "xl/sharedStrings.xml" in zf.namelist():
            tree = ET.parse(zf.open("xl/sharedStrings.xml"))
            for si in tree.getroot().findall("a:si", _XLSX_NS):
                text = "".join(t.text or "" for t in si.iter("{%s}t" % _XLSX_NS["a"]))
                shared.append(text)

        sheet_path = "xl/worksheets/sheet1.xml"
        if sheet_path not in zf.namelist():
            candidates = [n for n in zf.namelist() if n.startswith("xl/worksheets/sheet")]
            if not candidates:
                raise ProcessingError(f"No worksheet found in {path.name}")
            sheet_path = sorted(candidates)[0]

        tree = ET.parse(zf.open(sheet_path))
        sheet_data = tree.getroot().find("a:sheetData", _XLSX_NS)
        rows = []
        if sheet_data is None:
            return rows
        for row in sheet_data:
            cells = {}
            for c in row.findall("a:c", _XLSX_NS):
                ref = c.get("r")
                col_idx = _excel_col_to_index(ref) if ref else len(cells)
                v = c.find("a:v", _XLSX_NS)
                text = v.text if v is not None else ""
                if c.get("t") == "s" and text != "":
                    text = shared[int(text)]
                cells[col_idx] = text
            if cells:
                width = max(cells) + 1
                rows.append([cells.get(i, "") for i in range(width)])
        return rows


def read_table(path_text: str):
    """Read a .xlsx or delimited text (.csv/.txt) file into a list of rows."""
    path = Path(path_text)
    if path.suffix.lower() == ".xlsx":
        return read_xlsx_rows(path)
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        sample = f.read(4096)
        f.seek(0)
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=",;\t")
        except csv.Error:
            dialect = csv.excel
        return [row for row in csv.reader(f, dialect) if any(c.strip() for c in row)]


# ---------------------------------------------------------------------------
# Coordinate file
# ---------------------------------------------------------------------------

def read_coords(path_text: str, require_latitude: bool = True):
    rows = read_table(path_text)
    if not rows:
        raise ProcessingError("Coordinate file is empty.")
    headers = rows[0]
    station_key = pick_key(headers, ["Station ID", "Station", "StationID", "/Station"])
    easting_key = pick_key(headers, ["Easting", "East", "X"])
    northing_key = pick_key(headers, ["Northing", "North", "Y"])
    elev_key = pick_key(headers, ["Elevation", "Elev", "Z", "Height"])
    lat_key = pick_key(headers, ["Latitude", "Lat"], required=require_latitude)
    lon_key = pick_key(headers, ["Longitude", "Lon", "Long"], required=False)
    idx = {h: i for i, h in enumerate(headers)}

    coords = {}
    for r in rows[1:]:
        if len(r) <= idx[station_key]:
            continue
        sid = str(r[idx[station_key]]).strip()
        if not sid:
            continue
        coords[sid] = {
            "easting": as_float(r[idx[easting_key]], "Easting"),
            "northing": as_float(r[idx[northing_key]], "Northing"),
            "elevation": as_float(r[idx[elev_key]], "Elevation"),
            "latitude": (as_float(r[idx[lat_key]], "Latitude")
                         if lat_key and len(r) > idx[lat_key] and str(r[idx[lat_key]]).strip() else None),
            "longitude": (as_float(r[idx[lon_key]], "Longitude")
                          if lon_key and len(r) > idx[lon_key] and str(r[idx[lon_key]]).strip() else None),
        }
    if not coords:
        raise ProcessingError("No station coordinates could be read from the coordinate file.")
    return coords


# ---------------------------------------------------------------------------
# CG-6 .dat / ASCII reader
# ---------------------------------------------------------------------------

def read_cg6(path_text: str, header_scan: int = 200):
    path = Path(path_text)
    with open(path, "r", encoding="utf-8-sig", errors="replace") as f:
        lines = f.readlines()

    header_idx = None
    headers = None
    delim = "\t"
    for i, line in enumerate(lines[:header_scan]):
        stripped = line.strip()
        if not stripped or stripped.startswith(("//", "/*", "#")):
            continue
        for d in ("\t", ","):
            candidate = [c.strip() for c in stripped.split(d)]
            if len(candidate) > 3:
                nc = [norm(c) for c in candidate]
                has_station = any(x in ("station", "stationid") for x in nc)
                has_grav = any("corrgrav" in x or x == "grav" for x in nc)
                if has_station and has_grav:
                    header_idx, headers, delim = i, candidate, d
                    break
        if header_idx is not None:
            break

    if header_idx is None:
        raise ProcessingError(
            f"Could not find a CG-6 header (Station + CorrGrav columns) in the "
            f"first {header_scan} rows of {path.name}"
        )

    station_key = pick_key(headers, ["Station ID", "Station", "/Station"])
    grav_key = pick_key(headers, ["CorrGrav", "Corrected Gravity", "Corr.Grav.", "Grav"])
    date_key = pick_key(headers, ["Date"], required=False)
    time_key = pick_key(headers, ["Time"], required=False)
    idx = {h: i for i, h in enumerate(headers)}

    records = []
    for line in lines[header_idx + 1:]:
        stripped = line.strip()
        if not stripped or stripped.startswith(("//", "/*", "#")):
            continue
        parts = [c.strip() for c in stripped.split(delim)]
        if len(parts) < len(headers):
            continue
        row = {h: (parts[i] if i < len(parts) else "") for h, i in idx.items()}
        station = row.get(station_key, "").strip()
        if not station:
            continue
        try:
            gravity = as_float(row.get(grav_key), grav_key)
        except ProcessingError:
            continue
        when = parse_datetime(row.get(date_key), row.get(time_key)) if (date_key and time_key) else None
        records.append({"station": station, "gravity": gravity, "time": when, "source": path.name})

    if not records:
        raise ProcessingError(f"No usable data rows found in {path.name}")
    return records, headers


# ---------------------------------------------------------------------------
# Occupations + screening (screens the repeat readings taken together in the
# field at one station visit — NOT across separate visits/days, since a
# station's true value is expected to shift between separate occupations
# due to drift and tides)
# ---------------------------------------------------------------------------

def build_occupations(records, threshold_mgal: float):
    """Collapse consecutive readings at the same station (in time order) into
    a single 'occupation'. Within each occupation, readings further than
    threshold_mgal from that occupation's median are excluded before
    averaging the rest."""
    ordered = sorted(records, key=lambda r: (r["time"] is None, r["time"]))
    raw_occupations = []
    for r in ordered:
        if raw_occupations and raw_occupations[-1]["station"] == r["station"]:
            raw_occupations[-1]["readings"].append(r)
        else:
            raw_occupations.append({"station": r["station"], "readings": [r]})

    occupations = []
    excluded = []
    for occ in raw_occupations:
        values = [x["gravity"] for x in occ["readings"]]
        median = sorted(values)[len(values) // 2]
        kept_readings = []
        for x in occ["readings"]:
            deviation = x["gravity"] - median
            if len(values) > 1 and abs(deviation) > threshold_mgal:
                excluded.append({
                    "station": occ["station"],
                    "gravity_mgal": x["gravity"],
                    "occupation_median_mgal": median,
                    "deviation_mgal": deviation,
                    "source": x["source"],
                })
            else:
                kept_readings.append(x)
        if not kept_readings:
            # everything in this occupation was screened out — fall back to
            # keeping the reading closest to the median rather than losing
            # the whole occupation (and the loop it belongs to)
            kept_readings = [min(occ["readings"], key=lambda x: abs(x["gravity"] - median))]

        occ["readings"] = kept_readings
        occ["gravity"] = sum(x["gravity"] for x in kept_readings) / len(kept_readings)
        times = [x["time"] for x in kept_readings if x["time"] is not None]
        occ["time"] = times[len(times) // 2] if times else None
        occ["source"] = kept_readings[0]["source"]
        occupations.append(occ)

    return occupations, excluded


def build_loops(occupations, base_id: str):
    base_id = base_id.strip()
    base_idx = [i for i, occ in enumerate(occupations) if occ["station"] == base_id]
    if len(base_idx) < 2:
        raise ProcessingError(
            f"Fewer than two occupations of base station '{base_id}' were found. "
            "Drift correction needs an opening and a closing base reading."
        )
    loops = []
    for n in range(len(base_idx) - 1):
        loops.append(occupations[base_idx[n]: base_idx[n + 1] + 1])
    return loops


def apply_drift_correction(loops):
    for loop_no, loop_occs in enumerate(loops, start=1):
        open_occ, close_occ = loop_occs[0], loop_occs[-1]
        if open_occ["time"] is None or close_occ["time"] is None:
            raise ProcessingError(f"Loop {loop_no} is missing timestamps needed for drift correction.")
        elapsed_hours = (close_occ["time"] - open_occ["time"]).total_seconds() / 3600.0
        if elapsed_hours <= 0:
            raise ProcessingError(f"Loop {loop_no} has an invalid (zero or negative) duration.")
        drift_total = close_occ["gravity"] - open_occ["gravity"]
        drift_rate = drift_total / elapsed_hours

        for occ in loop_occs:
            dt_hours = (occ["time"] - open_occ["time"]).total_seconds() / 3600.0
            drift_calc = drift_rate * dt_hours
            occ["loop"] = loop_no
            occ["elapsed_hours"] = dt_hours
            occ["drift_rate_mgal_per_hour"] = drift_rate
            occ["drift_calculated_mgal"] = drift_calc
            occ["drift_correction_mgal"] = -drift_calc
            occ["gravity_drift_corrected_mgal"] = occ["gravity"] - drift_calc


# ---------------------------------------------------------------------------
# Latitude / free-air / Bouguer corrections
# ---------------------------------------------------------------------------

def apply_gravity_corrections(loops, coords, density, apply_latitude):
    missing = set()
    for loop_occs in loops:
        base_occ = loop_occs[0]
        base_coord = coords.get(base_occ["station"])
        if apply_latitude and base_coord is None:
            missing.add(base_occ["station"])
        base_theoretical_g = normal_gravity(base_coord["latitude"]) if (apply_latitude and base_coord) else 0.0

        for occ in loop_occs:
            coord = coords.get(occ["station"])
            if coord is None:
                missing.add(occ["station"])
                for key in ("easting", "northing", "elevation", "latitude", "longitude",
                            "latitude_correction_mgal", "gravity_after_latitude_correction_mgal",
                            "free_air_correction_mgal", "free_air_corrected_gravity_mgal",
                            "bouguer_slab_correction_mgal", "bouguer_corrected_gravity_mgal"):
                    occ[key] = None
                continue

            occ["easting"] = coord["easting"]
            occ["northing"] = coord["northing"]
            occ["elevation"] = coord["elevation"]
            occ["latitude"] = coord["latitude"]
            occ["longitude"] = coord["longitude"]

            gravity = occ["gravity_drift_corrected_mgal"]

            if apply_latitude:
                lat_corr = base_theoretical_g - normal_gravity(coord["latitude"])
            else:
                lat_corr = 0.0
            occ["latitude_correction_mgal"] = lat_corr
            gravity = gravity + lat_corr
            occ["gravity_after_latitude_correction_mgal"] = gravity

            fac = FREE_AIR_COEFF * coord["elevation"]
            occ["free_air_correction_mgal"] = fac
            gravity_fa = gravity + fac
            occ["free_air_corrected_gravity_mgal"] = gravity_fa

            bc = BOUGUER_COEFF * density * coord["elevation"]
            occ["bouguer_slab_correction_mgal"] = bc
            occ["bouguer_corrected_gravity_mgal"] = gravity_fa - bc

        valid = [o for o in loop_occs if o.get("bouguer_corrected_gravity_mgal") is not None]
        loop_base_bouguer = valid[0]["bouguer_corrected_gravity_mgal"] if valid else None
        for occ in loop_occs:
            occ["loop_base_bouguer_gravity_mgal"] = loop_base_bouguer
            if occ.get("bouguer_corrected_gravity_mgal") is not None and loop_base_bouguer is not None:
                occ["relative_bouguer_anomaly_mgal"] = occ["bouguer_corrected_gravity_mgal"] - loop_base_bouguer
            else:
                occ["relative_bouguer_anomaly_mgal"] = None

    return missing


# ---------------------------------------------------------------------------
# Full pipeline
# ---------------------------------------------------------------------------

OUTPUT_COLUMNS = [
    "Loop", "Station", "Date_Time", "N_Readings",
    "RawGravity_mGal", "ElapsedHours",
    "DriftRate_mGalPerHour", "DriftCalculated_mGal", "DriftCorrection_mGal",
    "CorrGrav_DriftCorrected_mGal",
    "Easting", "Northing", "Elevation", "Latitude", "Longitude",
    "LatitudeCorrection_mGal", "Gravity_AfterLatitudeCorrection_mGal",
    "FreeAirCorrection_mGal", "FreeAirCorrectedGravity_mGal",
    "BouguerSlabCorrection_mGal", "BouguerCorrectedGravity_mGal",
    "LoopBaseBouguerGravity_mGal", "RelativeBouguerAnomaly_mGal",
]


def run_process(files, coords_file, output_dir, output_name,
                 base_id, threshold_mgal, density, apply_latitude, log=print):
    if not files:
        raise ProcessingError("Select one or more CG-6 data files.")
    if not coords_file:
        raise ProcessingError("Select a coordinate file.")
    if not base_id.strip():
        raise ProcessingError("Enter the base station ID.")

    log(f"Reading {len(files)} CG-6 file(s)...")
    all_records = []
    reference_headers = None
    for f in files:
        records, headers = read_cg6(f)
        norm_headers = [norm(h) for h in headers]
        if reference_headers is None:
            reference_headers = norm_headers
        elif norm_headers != reference_headers:
            log(f"  Warning: headers in {Path(f).name} differ from the first file "
                "— continuing, but double-check the source files.")
        all_records.extend(records)
        log(f"  {Path(f).name}: {len(records)} readings")

    log(f"Total raw readings: {len(all_records)}")

    log("Reading coordinate file...")
    coords = read_coords(coords_file, require_latitude=apply_latitude)
    log(f"  {len(coords)} station coordinates loaded")

    log(f"Building occupations and screening with a +/-{threshold_mgal} mGal range "
        "around each occupation's median...")
    occupations, excluded = build_occupations(all_records, threshold_mgal)
    log(f"  {len(occupations)} occupation(s) built, {len(excluded)} outlier reading(s) excluded")

    log("Building loops from base station repeats...")
    loops = build_loops(occupations, base_id)
    log(f"  {len(loops)} loop(s) found")

    log("Applying drift correction...")
    apply_drift_correction(loops)

    log("Applying latitude / free-air / Bouguer corrections..." if apply_latitude
        else "Applying free-air / Bouguer corrections (latitude correction disabled)...")
    missing_coords = apply_gravity_corrections(loops, coords, density, apply_latitude)
    if missing_coords:
        log(f"  Warning: no coordinates found for: {', '.join(sorted(missing_coords))}")

    out_dir = Path(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    out_csv = out_dir / f"{output_name}.csv"
    out_json = out_dir / f"{output_name}_audit.json"

    log(f"Writing results to {out_csv.name}...")
    with open(out_csv, "w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(OUTPUT_COLUMNS)
        for loop_occs in loops:
            for occ in loop_occs:
                writer.writerow([
                    occ["loop"], occ["station"],
                    occ["time"].strftime("%Y-%m-%d %H:%M:%S") if occ["time"] else "",
                    len(occ["readings"]),
                    f'{occ["gravity"]:.4f}',
                    f'{occ["elapsed_hours"]:.4f}',
                    f'{occ["drift_rate_mgal_per_hour"]:.6f}',
                    f'{occ["drift_calculated_mgal"]:.4f}',
                    f'{occ["drift_correction_mgal"]:.4f}',
                    f'{occ["gravity_drift_corrected_mgal"]:.4f}',
                    occ.get("easting", ""), occ.get("northing", ""),
                    occ.get("elevation", ""), occ.get("latitude", ""), occ.get("longitude", ""),
                    _fmt(occ.get("latitude_correction_mgal")),
                    _fmt(occ.get("gravity_after_latitude_correction_mgal")),
                    _fmt(occ.get("free_air_correction_mgal")),
                    _fmt(occ.get("free_air_corrected_gravity_mgal")),
                    _fmt(occ.get("bouguer_slab_correction_mgal")),
                    _fmt(occ.get("bouguer_corrected_gravity_mgal")),
                    _fmt(occ.get("loop_base_bouguer_gravity_mgal")),
                    _fmt(occ.get("relative_bouguer_anomaly_mgal")),
                ])

    audit = {
        "input_files": [str(f) for f in files],
        "coordinate_file": str(coords_file),
        "base_station_id": base_id,
        "screening_threshold_mgal": threshold_mgal,
        "screening_excluded_readings": excluded,
        "density_g_cm3": density,
        "free_air_coefficient_mgal_per_m": FREE_AIR_COEFF,
        "bouguer_slab_coefficient": BOUGUER_COEFF,
        "latitude_correction_enabled": apply_latitude,
        "loops_found": len(loops),
        "stations_missing_coordinates": sorted(missing_coords),
        "output_csv": str(out_csv),
    }
    with open(out_json, "w", encoding="utf-8") as f:
        json.dump(audit, f, indent=2, default=str)

    log("Done.")
    return out_csv, out_json, len(loops), len(excluded), missing_coords


def _fmt(value):
    return "" if value is None else f"{value:.4f}"


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------

class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("CG-6 Gravity Processor")
        self.root.geometry("820x620")

        self.files_var = tk.StringVar(value="")
        self.coords_var = tk.StringVar(value="")
        self.out_dir_var = tk.StringVar(value="")
        self.out_name_var = tk.StringVar(value="CG6_Final")
        self.base_id_var = tk.StringVar(value="9999")
        self.threshold_var = tk.StringVar(value="0.05")
        self.density_var = tk.StringVar(value=str(DEFAULT_DENSITY))
        self.latitude_enabled_var = tk.BooleanVar(value=True)

        self._selected_files = []
        self._log_queue = queue.Queue()

        pad = {"padx": 8, "pady": 5}
        frame = ttk.Frame(self.root, padding=10)
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(1, weight=1)

        def row(r, label, var, button_text, command, width=60):
            ttk.Label(frame, text=label).grid(row=r, column=0, sticky="w", **pad)
            ttk.Entry(frame, textvariable=var, width=width).grid(row=r, column=1, sticky="ew", **pad)
            ttk.Button(frame, text=button_text, command=command).grid(row=r, column=2, **pad)

        row(0, "CG-6 data files (.dat/.txt/.csv)", self.files_var, "Browse...", self.choose_files)
        row(1, "Coordinate file (.xlsx/.csv)", self.coords_var, "Browse...", self.choose_coords)
        row(2, "Output folder", self.out_dir_var, "Browse...", self.choose_out)

        ttk.Label(frame, text="Output file name (no extension)").grid(row=3, column=0, sticky="w", **pad)
        ttk.Entry(frame, textvariable=self.out_name_var, width=30).grid(row=3, column=1, sticky="w", **pad)

        ttk.Label(frame, text="Base station ID").grid(row=4, column=0, sticky="w", **pad)
        ttk.Entry(frame, textvariable=self.base_id_var, width=20).grid(row=4, column=1, sticky="w", **pad)

        ttk.Label(frame, text="Screening range (+/- mGal around station median)").grid(
            row=5, column=0, sticky="w", **pad)
        ttk.Entry(frame, textvariable=self.threshold_var, width=20).grid(row=5, column=1, sticky="w", **pad)

        ttk.Label(frame, text="Bouguer density (g/cm3)").grid(row=6, column=0, sticky="w", **pad)
        ttk.Entry(frame, textvariable=self.density_var, width=20).grid(row=6, column=1, sticky="w", **pad)

        ttk.Checkbutton(
            frame, text="Apply latitude correction (relative to each loop's base station)",
            variable=self.latitude_enabled_var
        ).grid(row=7, column=0, columnspan=2, sticky="w", **pad)

        self.run_button = ttk.Button(frame, text="Run Processing", command=self.run)
        self.run_button.grid(row=8, column=0, columnspan=3, pady=12)

        self.status_var = tk.StringVar(value="Select inputs, then click Run Processing.")
        ttk.Label(frame, textvariable=self.status_var, wraplength=780).grid(
            row=9, column=0, columnspan=3, sticky="w", **pad)

        ttk.Label(frame, text="Log:").grid(row=10, column=0, sticky="w", padx=8)
        self.log_text = tk.Text(frame, height=18, wrap="word", state="disabled")
        self.log_text.grid(row=11, column=0, columnspan=3, sticky="nsew", padx=8, pady=(0, 8))
        frame.rowconfigure(11, weight=1)

        self.root.after(100, self._drain_log_queue)

    # -- file pickers --------------------------------------------------
    def choose_files(self):
        paths = filedialog.askopenfilenames(
            title="Select CG-6 data file(s)",
            filetypes=[("CG-6 / text files", "*.dat *.txt *.csv"), ("All files", "*.*")],
        )
        if paths:
            self._selected_files = list(paths)
            self.files_var.set("; ".join(Path(p).name for p in paths))

    def choose_coords(self):
        path = filedialog.askopenfilename(
            title="Select coordinate file",
            filetypes=[("Excel / CSV", "*.xlsx *.csv"), ("All files", "*.*")],
        )
        if path:
            self.coords_var.set(path)

    def choose_out(self):
        path = filedialog.askdirectory(title="Select output folder")
        if path:
            self.out_dir_var.set(path)

    # -- logging (thread-safe) -----------------------------------------
    def log(self, message):
        self._log_queue.put(message)

    def _drain_log_queue(self):
        try:
            while True:
                message = self._log_queue.get_nowait()
                self.log_text.configure(state="normal")
                self.log_text.insert("end", message + "\n")
                self.log_text.see("end")
                self.log_text.configure(state="disabled")
        except queue.Empty:
            pass
        self.root.after(100, self._drain_log_queue)

    # -- run --------------------------------------------------------------
    def run(self):
        if not self._selected_files:
            messagebox.showerror("Missing input", "Select one or more CG-6 data files.")
            return
        if not self.coords_var.get():
            messagebox.showerror("Missing input", "Select a coordinate file.")
            return
        if not self.out_dir_var.get():
            messagebox.showerror("Missing input", "Select an output folder.")
            return
        try:
            threshold = as_float(self.threshold_var.get(), "screening range")
            density = as_float(self.density_var.get(), "density")
        except ProcessingError as e:
            messagebox.showerror("Invalid input", str(e))
            return

        self.run_button.configure(state="disabled")
        self.status_var.set("Processing...")
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

        def worker():
            try:
                out_csv, out_json, n_loops, n_excluded, missing = run_process(
                    files=list(self._selected_files),
                    coords_file=self.coords_var.get(),
                    output_dir=self.out_dir_var.get(),
                    output_name=self.out_name_var.get() or "CG6_Final",
                    base_id=self.base_id_var.get(),
                    threshold_mgal=threshold,
                    density=density,
                    apply_latitude=self.latitude_enabled_var.get(),
                    log=self.log,
                )
                summary = (f"Done. {n_loops} loop(s) processed, {n_excluded} reading(s) screened out.")
                if missing:
                    summary += f" {len(missing)} station(s) missing coordinates (see audit file)."
                self.root.after(0, lambda: self.status_var.set(summary))
                self.root.after(0, lambda: messagebox.showinfo(
                    "Processing complete",
                    f"Results written to:\n{out_csv}\n\nAudit log:\n{out_json}"))
            except ProcessingError as e:
                self.root.after(0, lambda: self.status_var.set(f"Error: {e}"))
                self.root.after(0, lambda: messagebox.showerror("Processing error", str(e)))
            except Exception:
                tb = traceback.format_exc()
                self.log(tb)
                self.root.after(0, lambda: self.status_var.set("Unexpected error — see log."))
                self.root.after(0, lambda: messagebox.showerror(
                    "Unexpected error", "An unexpected error occurred. See the log panel for details."))
            finally:
                self.root.after(0, lambda: self.run_button.configure(state="normal"))

        threading.Thread(target=worker, daemon=True).start()

    def mainloop(self):
        self.root.mainloop()


def main():
    App().mainloop()


if __name__ == "__main__":
    main()
