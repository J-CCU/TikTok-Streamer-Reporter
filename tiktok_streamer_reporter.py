#!/usr/bin/env python3
"""
TikTok Streamer Reporter  — review-first assignment workflow

Goal
----
Import raw TikTok Shop LIVE CSV/XLSX exports, deduplicate by Room ID, lay every
stream out in one spreadsheet-like review table, let the user directly assign or
correct the Streamer field, and generate the final Excel report only when every
stream has a streamer.

The program intentionally avoids risky chronology/schedule guessing. It only
prefills streamer names when the title itself provides strong evidence (known
aliases or explicit W/ / WITH / HOST markers), or when the user made a correction/room assignment earlier in the current session.

Days with no streams are treated as normal schedule holes. Missing calendar dates
do not trigger warnings and do not block report generation.

GUI workflow
------------
1. Add raw TikTok CSV/XLSX exports.
2. Review the Streamer column. Double-click a Streamer cell to edit it.
3. Use bulk assignment and Rename/Merge to fix repeated names quickly.
4. Generate the Excel report. The button stays disabled until 100% assigned.

Run:
    python3 tiktok_streamer_reporter.py

CLI (uses automatic title detection only; refuses to export with unresolved rows):
    python3 tiktok_streamer_reporter.py export1.csv export2.xlsx -o TikTok_Streamer_Report.xlsx

Dependency for Excel output:
    python3 -m pip install XlsxWriter
"""

from __future__ import annotations

import argparse
import csv
import json
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple, Any

try:
    import xlsxwriter
except ImportError:  # pragma: no cover
    xlsxwriter = None


APP_TITLE = "TikTok Streamer Reporter"
STATE_VERSION = 1

# Stable aliases that have been explicitly established during development. This
# is NOT a fixed roster: future names from explicit title markers are added
# dynamically and can be renamed/merged in the UI.
KNOWN_ALIASES: Dict[str, str] = {
    "adam": "Adam",
    "austin": "Austin",
    "darby": "Darby",
    "darb": "Darby",
    "darbz": "Darby",
    "dc": "DC",
    "elly": "Elly",
    "george": "George",
    "georgie": "George",
    "georgie boi": "George",
    "ry": "Ryann",
    "ryry": "Ryann",
    "ryann": "Ryann",
    "tawna": "Tawna",
    "ccusports": "DC",
}

PREFERRED_ORDER = ["Adam", "Austin", "Darby", "DC", "Elly", "George", "Ryann", "Tawna"]

# Matches the current sales-master-sheet style.
GMV_BANDS = [
    (1000, "#FF0000"),
    (2000, "#FFF200"),
    (4000, "#93C47D"),
    (10000, "#00E600"),
    (float("inf"), "#00E5F0"),
]

DAY_WORDS = {
    "mon", "monday", "tue", "tues", "tuesday", "wed", "weds", "wednesday",
    "thu", "thur", "thurs", "thursday", "fri", "friday", "sat", "saturday",
    "sun", "sunday",
}

REQUIRED_HEADER_NAMES = {"room id", "room title", "start time"}


@dataclass
class StreamRow:
    source_file: str
    room_id: str
    room_title: str
    start_time: datetime
    end_time: Optional[datetime]
    duration_minutes: int
    gmv: float
    raw: Dict[str, str] = field(default_factory=dict)
    source_freshness: float = 0.0
    streamer: Optional[str] = None
    assignment_method: str = ""
    extracted_candidate: str = ""

    @property
    def stream_date(self) -> date:
        return self.start_time.date()


@dataclass
class ReportData:
    rows: List[StreamRow]
    duplicates_removed: int
    streamers: List[str]
    unresolved: List[StreamRow]


# ------------------------------- state ---------------------------------

def _blank_state() -> Dict[str, Any]:
    """Create clean session-only state.

    Manual room assignments, renames, and newly typed streamer names deliberately
    do NOT persist between app launches. This keeps every new instance independent
    from prior review work while still preserving edits during the current run.
    """
    return {
        "version": STATE_VERSION,
        "name_corrections": {},
        "room_assignments": {},
        "streamers": list(PREFERRED_ORDER),
    }


def load_state() -> Dict[str, Any]:
    """Return a fresh state for every application/CLI invocation."""
    return _blank_state()


def save_state(state: Dict[str, Any]) -> None:
    """Intentionally do nothing: review decisions are session-only.

    The GUI still calls this helper in a few places so the workflow code stays
    simple, but no JSON/config/sidecar file is written anywhere.
    """
    return None


# -------------------------- CSV / XLSX parsing -------------------------

def _clean_cell(value: object) -> str:
    if value is None:
        return ""
    return str(value).replace("\u00a0", " ").strip()


def _clean_name(value: object) -> str:
    return re.sub(r"\s+", " ", _clean_cell(value)).strip()


def _norm_header(value: str) -> str:
    return re.sub(r"\s+", " ", _clean_cell(value)).strip().lower()


def _read_csv_rows(path: Path) -> Tuple[List[str], List[List[str]]]:
    text = None
    for encoding in ("utf-8-sig", "utf-8", "cp1252"):
        try:
            text = path.read_text(encoding=encoding)
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        raise ValueError(f"Could not decode {path.name}")

    sample = text[:10000]
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",\t;")
    except csv.Error:
        dialect = csv.excel

    rows = list(csv.reader(text.splitlines(), dialect))
    for i, row in enumerate(rows[:100]):
        normalized = {_norm_header(c) for c in row}
        if REQUIRED_HEADER_NAMES.issubset(normalized):
            headers = [_clean_cell(c) for c in row]
            body = [r for r in rows[i + 1:] if any(_clean_cell(c) for c in r)]
            return headers, body
    raise ValueError(f"{path.name}: couldn't find TikTok headers. Expected Room ID, Room Title, and Start Time.")


def _excel_col_index(cell_ref: str) -> int:
    letters = re.match(r"[A-Za-z]+", cell_ref or "")
    if not letters:
        return 0
    n = 0
    for ch in letters.group(0).upper():
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n - 1


def _read_xlsx_rows(path: Path) -> Tuple[List[str], List[List[str]]]:
    """Read the TikTok XLSX directly with the standard library.

    This keeps input support lightweight and avoids requiring pandas/openpyxl.
    """
    NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
    PKG_REL_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"

    try:
        zf = zipfile.ZipFile(path)
    except zipfile.BadZipFile as exc:
        raise ValueError(f"{path.name}: not a valid .xlsx file") from exc

    with zf:
        shared_strings: List[str] = []
        if "xl/sharedStrings.xml" in zf.namelist():
            root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            for si in root.findall(f"{NS}si"):
                shared_strings.append("".join(t.text or "" for t in si.iter(f"{NS}t")))

        workbook = ET.fromstring(zf.read("xl/workbook.xml"))
        rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        rel_targets = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels.findall(f"{PKG_REL_NS}Relationship")}

        sheet_paths: List[str] = []
        sheets = workbook.find(f"{NS}sheets")
        if sheets is not None:
            for sheet in sheets.findall(f"{NS}sheet"):
                rid = sheet.attrib.get(f"{REL_NS}id", "")
                target = rel_targets.get(rid, "")
                if not target:
                    continue
                if target.startswith("/"):
                    sheet_paths.append(target.lstrip("/"))
                elif target.startswith("xl/"):
                    sheet_paths.append(target)
                else:
                    sheet_paths.append("xl/" + target.lstrip("/"))

        for sheet_path in sheet_paths:
            if sheet_path not in zf.namelist():
                continue
            root = ET.fromstring(zf.read(sheet_path))
            rows: List[List[str]] = []
            for row_el in root.iter(f"{NS}row"):
                values: List[str] = []
                for cell in row_el.findall(f"{NS}c"):
                    idx = _excel_col_index(cell.attrib.get("r", ""))
                    while len(values) <= idx:
                        values.append("")
                    cell_type = cell.attrib.get("t", "")
                    value = ""
                    if cell_type == "inlineStr":
                        inline = cell.find(f"{NS}is")
                        if inline is not None:
                            value = "".join(t.text or "" for t in inline.iter(f"{NS}t"))
                    else:
                        v = cell.find(f"{NS}v")
                        raw = v.text if v is not None and v.text is not None else ""
                        if cell_type == "s" and raw:
                            try:
                                value = shared_strings[int(raw)]
                            except (ValueError, IndexError):
                                value = raw
                        elif cell_type == "b":
                            value = "TRUE" if raw == "1" else "FALSE"
                        else:
                            value = raw
                    values[idx] = _clean_cell(value)
                rows.append(values)

            for i, row in enumerate(rows[:100]):
                normalized = {_norm_header(c) for c in row}
                if REQUIRED_HEADER_NAMES.issubset(normalized):
                    headers = [_clean_cell(c) for c in row]
                    body = [r for r in rows[i + 1:] if any(_clean_cell(c) for c in r)]
                    return headers, body

    raise ValueError(f"{path.name}: couldn't find TikTok headers. Expected Room ID, Room Title, and Start Time.")


def _read_input_rows(path: Path) -> Tuple[List[str], List[List[str]]]:
    suffix = path.suffix.lower()
    if suffix == ".csv":
        return _read_csv_rows(path)
    if suffix == ".xlsx":
        return _read_xlsx_rows(path)
    raise ValueError(f"{path.name}: unsupported input type. Use .csv or .xlsx TikTok exports.")


def _row_to_dict(headers: Sequence[str], row: Sequence[str]) -> Dict[str, str]:
    values = list(row) + [""] * max(0, len(headers) - len(row))
    return {headers[i]: _clean_cell(values[i]) for i in range(len(headers))}


def _get_col(row: Dict[str, str], *names: str) -> str:
    normalized = {_norm_header(k): v for k, v in row.items()}
    for name in names:
        key = _norm_header(name)
        if key in normalized:
            return normalized[key]
    return ""


def parse_datetime(value: str) -> Optional[datetime]:
    value = _clean_cell(value)
    if not value:
        return None
    formats = (
        "%Y-%m-%d %H:%M:%S",
        "%Y/%m/%d %H:%M:%S",
        "%m/%d/%Y %H:%M:%S",
        "%m/%d/%Y %I:%M:%S %p",
        "%m/%d/%Y %I:%M %p",
        "%Y-%m-%d %H:%M",
        "%m/%d/%Y %H:%M",
    )
    for fmt in formats:
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            pass
    try:
        return datetime.fromisoformat(value)
    except ValueError:
        pass

    # Excel stores datetimes as serial days in some exports.
    try:
        serial = float(value)
        if 20000 <= serial <= 100000:
            return datetime(1899, 12, 30) + timedelta(days=serial)
    except ValueError:
        pass
    return None


def parse_money(value: str) -> float:
    value = _clean_cell(value)
    if not value:
        return 0.0
    value = value.replace("$", "").replace(",", "").strip()
    if value.startswith("(") and value.endswith(")"):
        value = "-" + value[1:-1]
    try:
        return float(value)
    except ValueError:
        return 0.0


def parse_duration_minutes(value: str, start: Optional[datetime], end: Optional[datetime]) -> int:
    raw_value = _clean_cell(value)
    compact = raw_value.lower().replace(" ", "")
    if compact:
        h = re.search(r"(\d+)h", compact)
        m = re.search(r"(\d+)m", compact)
        s = re.search(r"(\d+)s", compact)
        if h or m or s:
            total_seconds = (int(h.group(1)) if h else 0) * 3600
            total_seconds += (int(m.group(1)) if m else 0) * 60
            total_seconds += int(s.group(1)) if s else 0
            return max(0, round(total_seconds / 60))
        parts = compact.split(":")
        if all(p.isdigit() for p in parts) and len(parts) in (2, 3):
            nums = [int(p) for p in parts]
            if len(nums) == 2:
                hours, minutes = nums
                return max(0, hours * 60 + minutes)
            hours, minutes, seconds = nums
            return max(0, round((hours * 3600 + minutes * 60 + seconds) / 60))
        # Excel duration as fraction of a day.
        try:
            numeric = float(compact)
            if 0 <= numeric <= 2:
                return max(0, round(numeric * 1440))
        except ValueError:
            pass
    if start and end:
        return max(0, round((end - start).total_seconds() / 60))
    return 0


def _source_freshness(path: Path) -> float:
    """Best-effort export recency for duplicate Room IDs.

    TikTok export filenames commonly include YYYYMMDDHHMMSS. When the same Room
    ID appears in older and newer exports, the newer export should win because
    historical GMV can be updated after the stream. Fall back to file mtime.
    """
    m = re.search(r"(?<!\d)(20\d{12})(?!\d)", path.name)
    if m:
        try:
            return datetime.strptime(m.group(1), "%Y%m%d%H%M%S").timestamp()
        except ValueError:
            pass
    try:
        return path.stat().st_mtime
    except OSError:
        return 0.0


def parse_file(path: Path) -> List[StreamRow]:
    headers, body = _read_input_rows(path)
    parsed: List[StreamRow] = []
    freshness = _source_freshness(path)
    for physical_row_num, raw_values in enumerate(body, start=1):
        raw = _row_to_dict(headers, raw_values)
        room_id = _get_col(raw, "Room ID")
        room_title = _get_col(raw, "Room Title")
        start = parse_datetime(_get_col(raw, "Start Time"))
        end = parse_datetime(_get_col(raw, "End Time"))

        if not room_id and not room_title and start is None:
            continue
        if start is None:
            # Skip malformed lines, but never make the entire batch fail.
            continue

        duration = parse_duration_minutes(_get_col(raw, "Duration"), start, end)
        gmv = parse_money(_get_col(raw, "Attributed GMV", "GMV"))
        parsed.append(
            StreamRow(
                source_file=path.name,
                room_id=room_id or f"{path.name}:row:{physical_row_num}",
                room_title=room_title,  # Preserve the source title exactly after basic whitespace cleanup.
                start_time=start,
                end_time=end,
                duration_minutes=duration,
                gmv=gmv,
                raw=raw,
                source_freshness=freshness,
            )
        )
    return parsed


# -------------------------- streamer prefilling ------------------------

def normalize_name(value: str) -> str:
    value = _clean_cell(value).upper().replace("&", " AND ")
    value = re.sub(r"[^A-Z0-9' -]+", " ", value)
    return re.sub(r"\s+", " ", value).strip(" -").lower()


def pretty_name(value: str) -> str:
    cleaned = re.sub(r"\s+", " ", _clean_cell(value)).strip(" !?.-_/#")
    if len(cleaned) <= 3 and cleaned.isalpha():
        return cleaned.upper()
    return " ".join(part.capitalize() for part in cleaned.split())


def _candidate_alias_prefix(candidate: str, aliases: Dict[str, str]) -> Optional[str]:
    key = normalize_name(candidate)
    for alias in sorted(aliases, key=len, reverse=True):
        if key == alias or key.startswith(alias + " "):
            return alias
    return None


def extract_streamer_candidate(title: str) -> Optional[str]:
    """Extract only explicit host markers; no chronology or schedule inference."""
    text = _clean_cell(title)
    if not text:
        return None
    patterns = [
        r"(?i)(?:^|\s)W\s*[/\.\-]\s*(.+?)\s*$",
        r"(?i)(?:^|\s)WITH\s+(.+?)\s*$",
        r"(?i)(?:^|\s)HOSTED\s+BY\s+(.+?)\s*$",
        r"(?i)(?:^|\s)HOST\s*[:\-]\s*(.+?)\s*$",
    ]
    candidate = None
    for pattern in patterns:
        m = re.search(pattern, text)
        if m:
            candidate = m.group(1)
            break
    if not candidate:
        return None
    candidate = re.sub(r"(?i)^YA\s+BOI\s+", "", candidate)
    candidate = re.sub(r"(?i)^YOUR\s+BOY\s+", "", candidate)
    candidate = re.sub(r"(?i)^THE\s+BOY\s+", "", candidate)
    candidate = re.sub(r"(?i)\s+(LIVE|STREAM)\s*$", "", candidate)
    candidate = candidate.strip(" !?.-_/#🧙‍♂️🧝")
    return candidate or None


def _candidate_looks_like_person(candidate: str) -> bool:
    key = normalize_name(candidate)
    if not key:
        return False
    words = key.split()
    if any(w in DAY_WORDS for w in words):
        return False
    if "?" in candidate or len(words) > 3:
        return False
    junk = {"order", "orders", "sale", "live", "stream", "break", "breaks", "rips", "rip"}
    return not (words and words[0] in junk)


def _title_contains_alias(title: str, alias: str) -> bool:
    hay = normalize_name(title)
    needle = normalize_name(alias)
    if not needle:
        return False
    pattern = r"(?<![a-z0-9])" + re.escape(needle) + r"(?:'s|s)?(?![a-z0-9])"
    return re.search(pattern, hay, re.I) is not None


def _known_alias_hits(title: str, aliases: Dict[str, str]) -> List[Tuple[str, str]]:
    hits: List[Tuple[str, str]] = []
    seen = set()
    for alias, display in sorted(aliases.items(), key=lambda kv: len(kv[0]), reverse=True):
        # One-character aliases are too risky for arbitrary title-wide matching.
        if len(alias) < 2:
            continue
        if _title_contains_alias(title, alias) and display not in seen:
            hits.append((alias, display))
            seen.add(display)
    return hits


def _resolve_explicit_candidate(candidate: str, aliases: Dict[str, str]) -> str:
    key = normalize_name(candidate)
    if key in aliases:
        return aliases[key]
    prefix = _candidate_alias_prefix(candidate, aliases)
    if prefix:
        return aliases[prefix]
    return pretty_name(candidate)


def _specific_direct_title_host(title: str, aliases: Dict[str, str]) -> Optional[Tuple[str, str]]:
    """Very narrow direct-title pattern: 'Dave on a Monday?!' -> Dave."""
    text = _clean_cell(title)
    m = re.match(
        r"^([A-Za-z][A-Za-z'\-]{1,24})\s+on\s+a\s+"
        r"(?:mon|tue|wed|thu|fri|sat|sun|monday|tuesday|wednesday|thursday|friday|saturday|sunday)\b",
        text,
        re.I,
    )
    if not m:
        return None
    candidate = m.group(1)
    return candidate, _resolve_explicit_candidate(candidate, aliases)


def assign_streamers(rows: List[StreamRow], state: Optional[Dict[str, Any]] = None) -> None:
    """Prefill only evidence that comes from the title or saved user decisions.

    Crucially, this function does NOT use prior/next rows, time slots, weekdays,
    source consistency, or special-event inheritance to silently assign a person.
    """
    state = state or _blank_state()
    aliases: Dict[str, str] = dict(KNOWN_ALIASES)

    # A rename made earlier in this session becomes a reusable alias for rescans in this session.
    for old, new in state.get("name_corrections", {}).items():
        old_key = normalize_name(str(old))
        new_name = _clean_name(new)
        if old_key and new_name:
            aliases[old_key] = new_name

    # Pass 1: stable aliases explicitly present in the title.
    for row in rows:
        hits = _known_alias_hits(row.room_title, aliases)
        if len(hits) == 1:
            alias, display = hits[0]
            row.streamer = display
            row.extracted_candidate = alias
            row.assignment_method = "auto: title alias"

    # Pass 2: explicit W/ / WITH / HOST markers can discover future streamers.
    for row in rows:
        if row.streamer:
            continue
        candidate = extract_streamer_candidate(row.room_title)
        if not candidate or not _candidate_looks_like_person(candidate):
            continue
        display = _resolve_explicit_candidate(candidate, aliases)
        row.streamer = display
        row.extracted_candidate = candidate
        row.assignment_method = "auto: explicit host marker"
        aliases[normalize_name(candidate)] = display

    # Pass 3: learned explicit aliases can now be reused elsewhere in the title.
    for row in rows:
        if row.streamer:
            continue
        hits = _known_alias_hits(row.room_title, aliases)
        if len(hits) == 1:
            alias, display = hits[0]
            row.streamer = display
            row.extracted_candidate = alias
            row.assignment_method = "auto: learned explicit alias"

    # Pass 4: a tiny, direct grammatical shape. No inference from surrounding rows.
    for row in rows:
        if row.streamer:
            continue
        hit = _specific_direct_title_host(row.room_title, aliases)
        if hit:
            candidate, display = hit
            row.streamer = display
            row.extracted_candidate = candidate
            row.assignment_method = "auto: direct host phrase"
            aliases[normalize_name(candidate)] = display

    # Pass 5: row-level human decisions from the current session always override automation.
    saved = state.get("room_assignments", {})
    for row in rows:
        if row.room_id not in saved:
            continue
        value = saved[row.room_id]
        if value is None or not _clean_name(value):
            row.streamer = None
            row.assignment_method = "session: user left unassigned"
            row.extracted_candidate = row.extracted_candidate or "[session room decision]"
        else:
            row.streamer = _clean_name(value)
            row.assignment_method = "session: room assignment"
            row.extracted_candidate = row.extracted_candidate or "[session room assignment]"

    for row in rows:
        if not row.streamer:
            row.assignment_method = row.assignment_method or "needs user assignment"


def _row_completeness(row: StreamRow) -> int:
    return sum(1 for v in row.raw.values() if _clean_cell(v))


def deduplicate_rows(rows: Iterable[StreamRow]) -> Tuple[List[StreamRow], int]:
    """Deduplicate by Room ID, preferring the newest export copy.

    TikTok can revise historical metrics. If multiple export batches contain the
    same Room ID, recency is more important than input-file ordering. If export
    recency is tied/unknown, keep the row with more populated source fields.
    """
    by_id: Dict[str, StreamRow] = {}
    removed = 0
    for row in rows:
        key = row.room_id.strip()
        if key not in by_id:
            by_id[key] = row
            continue
        removed += 1
        old = by_id[key]
        if row.source_freshness > old.source_freshness:
            by_id[key] = row
        elif row.source_freshness == old.source_freshness and _row_completeness(row) > _row_completeness(old):
            by_id[key] = row
    deduped = list(by_id.values())
    deduped.sort(key=lambda r: (r.start_time, r.room_id))
    return deduped, removed


def streamer_sort_key(name: str) -> Tuple[int, str]:
    try:
        return (PREFERRED_ORDER.index(name), "")
    except ValueError:
        return (len(PREFERRED_ORDER), name.lower())



def _refresh_report_state(report: ReportData) -> None:
    report.streamers = sorted({r.streamer for r in report.rows if r.streamer}, key=streamer_sort_key)
    report.unresolved = [r for r in report.rows if not r.streamer]


def build_report_data(paths: Sequence[Path], state: Optional[Dict[str, Any]] = None) -> ReportData:
    all_rows: List[StreamRow] = []
    errors: List[str] = []
    for path in paths:
        try:
            all_rows.extend(parse_file(path))
        except Exception as exc:
            errors.append(f"{path.name}: {exc}")
    if errors:
        raise ValueError("Could not import one or more files:\n\n" + "\n".join(errors))
    if not all_rows:
        raise ValueError("No valid TikTok live rows were found in the selected files.")

    rows, removed = deduplicate_rows(all_rows)
    assign_streamers(rows, state)
    report = ReportData(
        rows=rows,
        duplicates_removed=removed,
        streamers=[],
        unresolved=[],
    )
    _refresh_report_state(report)
    return report


def set_row_assignment(report: ReportData, row: StreamRow, streamer: str, method: str, state: Dict[str, Any]) -> None:
    name = _clean_name(streamer)
    row.streamer = name or None
    row.assignment_method = method if name else "manual: cleared assignment"
    if name:
        if name not in state["streamers"]:
            state["streamers"].append(name)
        state["room_assignments"][row.room_id] = name
    else:
        # Preserve an explicit user decision to clear a false automatic match.
        state["room_assignments"][row.room_id] = None
    _refresh_report_state(report)


def set_many_assignments(report: ReportData, rows: Sequence[StreamRow], streamer: str, method: str, state: Dict[str, Any]) -> None:
    name = _clean_name(streamer)
    if not name:
        raise ValueError("Streamer name cannot be blank.")
    if name not in state["streamers"]:
        state["streamers"].append(name)
    for row in rows:
        row.streamer = name
        row.assignment_method = method
        state["room_assignments"][row.room_id] = name
    _refresh_report_state(report)


def rename_streamer(report: ReportData, old: str, new: str, state: Dict[str, Any]) -> int:
    old = _clean_name(old)
    new = _clean_name(new)
    if not old or not new:
        raise ValueError("Both the existing and new streamer names are required.")
    count = 0
    for row in report.rows:
        if row.streamer == old:
            row.streamer = new
            row.assignment_method = (row.assignment_method + " + manual rename/merge").strip(" +")
            count += 1
            # If the room already has saved memory, keep that memory consistent.
            if row.room_id in state["room_assignments"]:
                state["room_assignments"][row.room_id] = new
    state["name_corrections"][normalize_name(old)] = new
    if new not in state["streamers"]:
        state["streamers"].append(new)
    if old in state["streamers"] and old not in PREFERRED_ORDER:
        state["streamers"] = [x for x in state["streamers"] if x != old]
    _refresh_report_state(report)
    return count


# ----------------------------- Excel output ----------------------------

def gmv_fill_color(value: float) -> str:
    for upper, color in GMV_BANDS:
        if value < upper:
            return color
    return GMV_BANDS[-1][1]


def minutes_to_decimal_hours(minutes: int) -> float:
    return minutes / 60.0


def _first_monday_for_month(year: int, month: int) -> date:
    first = date(year, month, 1)
    return first - timedelta(days=first.weekday())


def _last_sunday_for_month(year: int, month: int) -> date:
    if month == 12:
        next_month = date(year + 1, 1, 1)
    else:
        next_month = date(year, month + 1, 1)
    last = next_month - timedelta(days=1)
    return last + timedelta(days=(6 - last.weekday()))


def _month_keys(rows: Sequence[StreamRow]) -> List[Tuple[int, int]]:
    return sorted({(r.stream_date.year, r.stream_date.month) for r in rows})


def _aggregate(rows: Sequence[StreamRow]):
    agg = defaultdict(lambda: {"gmv": 0.0, "minutes": 0, "rooms": 0})
    for r in rows:
        if not r.streamer:
            continue
        key = (r.stream_date, r.streamer)
        agg[key]["gmv"] += r.gmv
        agg[key]["minutes"] += r.duration_minutes
        agg[key]["rooms"] += 1
    return agg


def _unique_sheet_name(existing: set, desired: str) -> str:
    base = desired[:31]
    name = base
    n = 2
    while name in existing:
        suffix = f" {n}"
        name = base[:31 - len(suffix)] + suffix
        n += 1
    existing.add(name)
    return name


def _raw_headers(rows: Sequence[StreamRow]) -> List[str]:
    seen = set()
    result = []
    for row in rows:
        for header in row.raw:
            key = header.strip()
            if key and key not in seen:
                seen.add(key)
                result.append(key)
    return result


def write_excel(report: ReportData, output_path: Path) -> None:
    if report.unresolved:
        raise ValueError(f"Cannot create final report: {len(report.unresolved)} stream(s) are still unassigned.")
    if xlsxwriter is None:
        raise RuntimeError("Missing dependency 'xlsxwriter'. Install it with: python3 -m pip install XlsxWriter")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook = xlsxwriter.Workbook(str(output_path))
    workbook.set_properties({
        "title": "TikTok Streamer Sales Report",
        "subject": "Daily and weekly attributed GMV and streaming hours",
        "author": APP_TITLE,
        "comments": "All rows were assigned before report generation.",
    })

    fmt_streamer = workbook.add_format({"bold": True, "align": "center", "valign": "vcenter", "border": 1})
    fmt_hrs_header = workbook.add_format({"bold": True, "align": "center", "valign": "vcenter", "border": 1})
    fmt_week = workbook.add_format({"bold": True, "align": "left", "valign": "vcenter", "border": 1})
    fmt_week_money = workbook.add_format({"num_format": "$#,##0.00", "border": 1})
    fmt_week_hours = workbook.add_format({"num_format": "0.00", "border": 1})
    fmt_day = workbook.add_format({"align": "center", "border": 1})
    fmt_date = workbook.add_format({"align": "center", "border": 1})
    fmt_blank = workbook.add_format({"border": 1})
    fmt_money_by_color = {
        color: workbook.add_format({"num_format": "$#,##0.00", "align": "right", "border": 1, "bg_color": color})
        for _, color in GMV_BANDS
    }
    fmt_hours = workbook.add_format({"num_format": "0.00", "align": "center", "border": 1})
    fmt_section = workbook.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1})
    fmt_text = workbook.add_format({"border": 1})
    fmt_wrap = workbook.add_format({"border": 1, "text_wrap": True, "valign": "top"})
    fmt_money = workbook.add_format({"num_format": "$#,##0.00", "border": 1})
    fmt_datetime = workbook.add_format({"num_format": "yyyy-mm-dd hh:mm:ss", "border": 1})
    fmt_integer = workbook.add_format({"num_format": "0", "border": 1})
    fmt_note = workbook.add_format({"font_color": "#666666", "italic": True})
    fmt_warn = workbook.add_format({"bg_color": "#FFF2CC", "border": 1})

    agg = _aggregate(report.rows)
    existing_sheets = set()
    years = {y for y, _ in _month_keys(report.rows)}

    # Calendar-style month sheets matching the master-sheet structure.
    for year, month in _month_keys(report.rows):
        month_name = date(year, month, 1).strftime("%B")
        desired_name = month_name if len(years) == 1 else f"{month_name} {year}"
        ws = workbook.add_worksheet(_unique_sheet_name(existing_sheets, desired_name))
        ws.hide_gridlines(2)
        ws.freeze_panes(2, 2)
        ws.set_column(0, 0, 6)
        ws.set_column(1, 1, 9)
        for i, _ in enumerate(report.streamers):
            sales_col = 2 + i * 2
            ws.set_column(sales_col, sales_col, 13)
            ws.set_column(sales_col + 1, sales_col + 1, 8)

        ws.write(0, 0, "", fmt_streamer)
        ws.write(0, 1, "", fmt_streamer)
        for i, streamer in enumerate(report.streamers):
            sales_col = 2 + i * 2
            ws.write(0, sales_col, streamer, fmt_streamer)
            ws.write(0, sales_col + 1, "Hrs", fmt_hrs_header)

        current = _first_monday_for_month(year, month)
        end = _last_sunday_for_month(year, month)
        row_idx = 1
        week_num = 1
        while current <= end:
            week_dates = [current + timedelta(days=d) for d in range(7)]
            ws.write(row_idx, 0, "", fmt_week)
            ws.write(row_idx, 1, f"Week {week_num}", fmt_week)
            for i, streamer in enumerate(report.streamers):
                sales_col = 2 + i * 2
                weekly_gmv = sum(agg[(d, streamer)]["gmv"] for d in week_dates)
                weekly_minutes = sum(agg[(d, streamer)]["minutes"] for d in week_dates)
                ws.write_number(row_idx, sales_col, weekly_gmv, fmt_week_money)
                ws.write_number(row_idx, sales_col + 1, minutes_to_decimal_hours(weekly_minutes), fmt_week_hours)

            for day_offset, d in enumerate(week_dates, start=1):
                rr = row_idx + day_offset
                ws.write(rr, 0, d.strftime("%a"), fmt_day)
                ws.write_number(rr, 1, d.day, fmt_date)
                for i, streamer in enumerate(report.streamers):
                    sales_col = 2 + i * 2
                    item = agg[(d, streamer)]
                    if item["rooms"]:
                        ws.write_number(rr, sales_col, item["gmv"], fmt_money_by_color[gmv_fill_color(item["gmv"])])
                        ws.write_number(rr, sales_col + 1, minutes_to_decimal_hours(item["minutes"]), fmt_hours)
                    else:
                        ws.write_blank(rr, sales_col, None, fmt_blank)
                        ws.write_blank(rr, sales_col + 1, None, fmt_blank)

            row_idx += 9
            week_num += 1
            current += timedelta(days=7)

        note_row = row_idx + 1
        ws.write(note_row, 0, "Generated from reviewed TikTok live-room assignments. Duplicate Room IDs removed automatically.", fmt_note)

    # Full audit sheet. Fixed useful columns come first, followed by every raw
    # TikTok export field so the source data is never hidden by the report schema.
    ws = workbook.add_worksheet(_unique_sheet_name(existing_sheets, "Combined Data"))
    ws.freeze_panes(1, 0)
    fixed_headers = [
        "Source File", "Room ID", "Room Title", "Start Time", "End Time", "Date",
        "Duration Minutes", "Duration Hours", "Attributed GMV", "Streamer",
        "Assignment Method", "Extracted Candidate",
    ]
    raw_headers = _raw_headers(report.rows)
    headers = fixed_headers + [f"Raw: {h}" for h in raw_headers]
    for c, h in enumerate(headers):
        ws.write(0, c, h, fmt_section)

    ws.set_column(0, 0, 27)
    ws.set_column(1, 1, 22)
    ws.set_column(2, 2, 48)
    ws.set_column(3, 5, 20)
    ws.set_column(6, 7, 16)
    ws.set_column(8, 8, 16)
    ws.set_column(9, 11, 27)
    if raw_headers:
        ws.set_column(len(fixed_headers), len(headers) - 1, 18)

    for r_idx, row in enumerate(report.rows, start=1):
        ws.write(r_idx, 0, row.source_file, fmt_text)
        ws.write(r_idx, 1, row.room_id, fmt_text)
        ws.write(r_idx, 2, row.room_title, fmt_wrap)
        ws.write_datetime(r_idx, 3, row.start_time, fmt_datetime)
        if row.end_time:
            ws.write_datetime(r_idx, 4, row.end_time, fmt_datetime)
        else:
            ws.write_blank(r_idx, 4, None, fmt_text)
        ws.write(r_idx, 5, row.stream_date.isoformat(), fmt_text)
        ws.write_number(r_idx, 6, row.duration_minutes, fmt_integer)
        ws.write_number(r_idx, 7, minutes_to_decimal_hours(row.duration_minutes), fmt_hours)
        ws.write_number(r_idx, 8, row.gmv, fmt_money)
        ws.write(r_idx, 9, row.streamer or "", fmt_text)
        ws.write(r_idx, 10, row.assignment_method, fmt_text)
        ws.write(r_idx, 11, row.extracted_candidate, fmt_text)
        for j, raw_header in enumerate(raw_headers, start=len(fixed_headers)):
            ws.write(r_idx, j, row.raw.get(raw_header, ""), fmt_text)
    ws.autofilter(0, 0, len(report.rows), len(headers) - 1)

    ws = workbook.add_worksheet(_unique_sheet_name(existing_sheets, "Run Info"))
    ws.set_column(0, 0, 30)
    ws.set_column(1, 1, 90)
    info = [
        ("Rows after dedupe", len(report.rows)),
        ("Duplicate Room IDs removed", report.duplicates_removed),
        ("Duplicate rule", "Newest export copy wins for repeated Room IDs; tied export timestamps prefer the more complete row"),
        ("Streamer assignments", "100% complete before workbook generation"),
        ("Streamers", ", ".join(report.streamers)),
        ("Hours format", "Decimal hours with 2 decimals (minutes / 60)"),
        ("Attribution policy", "Direct title evidence and saved user decisions only; no automatic chronology/schedule inheritance"),
        ("Inputs", "Raw TikTok CSV/XLSX live-performance exports"),
    ]
    for r, (label, value) in enumerate(info):
        ws.write(r, 0, label, fmt_section)
        ws.write(r, 1, value, fmt_wrap)

    workbook.close()


# -------------------------------- GUI ----------------------------------

def launch_gui() -> None:
    import tkinter as tk
    from tkinter import filedialog, messagebox, simpledialog, ttk

    root = tk.Tk()
    root.title(f"{APP_TITLE}")
    root.geometry("1420x820")
    root.minsize(1040, 650)

    state = load_state()
    selected_paths: List[Path] = []
    report: Optional[ReportData] = None
    item_to_row: Dict[str, StreamRow] = {}
    active_editor = {"widget": None}

    main = ttk.Frame(root, padding=12)
    main.pack(fill="both", expand=True)

    title_row = ttk.Frame(main)
    title_row.pack(fill="x")
    ttk.Label(title_row, text="TikTok Streamer Reporter", font=("TkDefaultFont", 18, "bold")).pack(side="left")
    ttk.Label(title_row, text="", font=("TkDefaultFont", 10)).pack(side="left", padx=(10, 0), pady=(6, 0))

    ttk.Label(
        main,
        text=(
            "Import raw TikTok exports, then review the Streamer column directly. "
            "Double-click a Streamer cell to edit it. The final report cannot be generated until every row is assigned."
        ),
        wraplength=1360,
    ).pack(anchor="w", pady=(3, 8))

    # Compact file picker.
    file_frame = ttk.LabelFrame(main, text="TikTok exports", padding=6)
    file_frame.pack(fill="x", pady=(0, 6))
    file_buttons = ttk.Frame(file_frame)
    file_buttons.pack(side="left", padx=(0, 8))
    file_list = tk.Listbox(file_frame, height=3, exportselection=False)
    file_list.pack(side="left", fill="x", expand=True)
    file_scroll = ttk.Scrollbar(file_frame, orient="vertical", command=file_list.yview)
    file_scroll.pack(side="right", fill="y")
    file_list.configure(yscrollcommand=file_scroll.set)

    # Summary / notices.
    summary_var = tk.StringVar(value="No exports loaded.")
    summary_label = ttk.Label(main, textvariable=summary_var, font=("TkDefaultFont", 10, "bold"))
    summary_label.pack(anchor="w", pady=(0, 2))
    notice_var = tk.StringVar(value="")
    notice_label = ttk.Label(main, textvariable=notice_var, wraplength=1360)
    notice_label.pack(anchor="w", pady=(0, 6))

    # Filters and assignment tools.
    tools_frame = ttk.LabelFrame(main, text="Assignment tools", padding=6)
    tools_frame.pack(fill="x", pady=(0, 6))

    filter_var = tk.StringVar(value="All rows")
    source_var = tk.StringVar(value="All sources")
    search_var = tk.StringVar(value="")
    bulk_var = tk.StringVar(value="")

    ttk.Label(tools_frame, text="Show:").grid(row=0, column=0, sticky="w")
    filter_combo = ttk.Combobox(tools_frame, textvariable=filter_var, values=["All rows", "Needs streamer", "Assigned"], state="readonly", width=16)
    filter_combo.grid(row=0, column=1, padx=(5, 12), sticky="w")

    ttk.Label(tools_frame, text="Source:").grid(row=0, column=2, sticky="w")
    source_combo = ttk.Combobox(tools_frame, textvariable=source_var, values=["All sources"], state="readonly", width=34)
    source_combo.grid(row=0, column=3, padx=(5, 12), sticky="w")

    ttk.Label(tools_frame, text="Search:").grid(row=0, column=4, sticky="w")
    search_entry = ttk.Entry(tools_frame, textvariable=search_var, width=28)
    search_entry.grid(row=0, column=5, padx=(5, 12), sticky="ew")
    tools_frame.columnconfigure(5, weight=1)

    ttk.Label(tools_frame, text="Assign selected to:").grid(row=1, column=0, sticky="w", pady=(7, 0))
    bulk_combo = ttk.Combobox(tools_frame, textvariable=bulk_var, state="normal", width=24)
    bulk_combo.grid(row=1, column=1, padx=(5, 6), pady=(7, 0), sticky="w")
    assign_btn = ttk.Button(tools_frame, text="Apply to selected")
    assign_btn.grid(row=1, column=2, padx=(0, 6), pady=(7, 0), sticky="w")
    clear_assign_btn = ttk.Button(tools_frame, text="Clear selected")
    clear_assign_btn.grid(row=1, column=3, padx=(0, 6), pady=(7, 0), sticky="w")
    prev_btn = ttk.Button(tools_frame, text="Use previous")
    prev_btn.grid(row=1, column=4, padx=(0, 6), pady=(7, 0), sticky="w")
    next_btn = ttk.Button(tools_frame, text="Use next")
    next_btn.grid(row=1, column=5, padx=(0, 6), pady=(7, 0), sticky="w")
    rename_btn = ttk.Button(tools_frame, text="Rename / Merge Name…")
    rename_btn.grid(row=1, column=6, padx=(6, 0), pady=(7, 0), sticky="e")

    # Main review table.
    table_frame = ttk.Frame(main)
    table_frame.pack(fill="both", expand=True)

    columns = ("status", "date", "start", "end", "hours", "gmv", "title", "streamer", "method", "room_id", "source")
    tree = ttk.Treeview(table_frame, columns=columns, show="headings", selectmode="extended")
    yscroll = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
    xscroll = ttk.Scrollbar(table_frame, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
    tree.grid(row=0, column=0, sticky="nsew")
    yscroll.grid(row=0, column=1, sticky="ns")
    xscroll.grid(row=1, column=0, sticky="ew")
    table_frame.rowconfigure(0, weight=1)
    table_frame.columnconfigure(0, weight=1)

    headings = {
        "status": "Status", "date": "Date", "start": "Start", "end": "End",
        "hours": "Hrs", "gmv": "GMV", "title": "Room Title", "streamer": "Streamer",
        "method": "Assignment", "room_id": "Room ID", "source": "Source File",
    }
    widths = {
        "status": 70, "date": 92, "start": 82, "end": 82, "hours": 65,
        "gmv": 100, "title": 350, "streamer": 135, "method": 190,
        "room_id": 160, "source": 260,
    }
    anchors = {"hours": "e", "gmv": "e", "status": "center", "date": "center", "start": "center", "end": "center"}
    for col in columns:
        tree.heading(col, text=headings[col])
        tree.column(col, width=widths[col], minwidth=55, anchor=anchors.get(col, "w"), stretch=(col in {"title", "method", "source"}))

    tree.tag_configure("missing", foreground="#A33A00")
    tree.tag_configure("manual", foreground="#245A9A")
    tree.tag_configure("memory", foreground="#5B3A8A")

    bottom = ttk.Frame(main)
    bottom.pack(fill="x", pady=(7, 0))
    status_var = tk.StringVar(value="Ready")
    ttk.Label(bottom, textvariable=status_var).pack(side="left")
    generate_btn = ttk.Button(bottom, text="Generate Excel Report", state="disabled")
    generate_btn.pack(side="right")
    show_missing_btn = ttk.Button(bottom, text="Show Needs Streamer")
    show_missing_btn.pack(side="right", padx=(0, 8))

    def current_name_choices() -> List[str]:
        names = list(PREFERRED_ORDER)
        names.extend(state.get("streamers", []))
        if report:
            names.extend(report.streamers)
        deduped = {}
        for name in names:
            cleaned = _clean_name(name)
            if cleaned:
                deduped.setdefault(normalize_name(cleaned), cleaned)
        return sorted(deduped.values(), key=streamer_sort_key)

    def update_name_widgets() -> None:
        values = current_name_choices()
        bulk_combo.configure(values=values)

    def refresh_file_list() -> None:
        file_list.delete(0, tk.END)
        for p in selected_paths:
            file_list.insert(tk.END, str(p))

    def refresh_summary() -> None:
        if not report:
            summary_var.set("No exports loaded.")
            notice_var.set("")
            generate_btn.configure(state="disabled")
            return
        assigned = len(report.rows) - len(report.unresolved)
        summary_var.set(
            f"{len(selected_paths)} file(s) • {len(report.rows)} unique rows • "
            f"{report.duplicates_removed} duplicate(s) removed • {assigned}/{len(report.rows)} assigned • "
            f"{len(report.unresolved)} need streamer"
        )
        notices = []
        if report.unresolved:
            notices.append("Report generation is locked until every row has a streamer. Filter to ‘Needs streamer’ to finish quickly.")
        else:
            notices.append("All rows are assigned. The report is ready to generate.")
        notice_var.set("\n".join(notices))
        generate_btn.configure(state="normal" if report.rows and not report.unresolved else "disabled")
        update_name_widgets()

    def filtered_rows() -> List[StreamRow]:
        if not report:
            return []
        mode = filter_var.get()
        source = source_var.get()
        query = search_var.get().strip().lower()
        rows = []
        for row in report.rows:
            if mode == "Needs streamer" and row.streamer:
                continue
            if mode == "Assigned" and not row.streamer:
                continue
            if source != "All sources" and row.source_file != source:
                continue
            if query:
                hay = " ".join([
                    row.room_title, row.streamer or "", row.room_id, row.source_file,
                    row.assignment_method, row.start_time.strftime("%Y-%m-%d %H:%M"),
                ]).lower()
                if query not in hay:
                    continue
            rows.append(row)
        # Source-local chronology makes prior/next rows visually adjacent without
        # mixing separate exports/accounts together.
        rows.sort(key=lambda r: (r.source_file.lower(), r.start_time, r.room_id))
        return rows

    def _selected_room_ids() -> set:
        ids = set()
        for item in tree.selection():
            row = item_to_row.get(item)
            if row:
                ids.add(row.room_id)
        return ids

    def refresh_tree(preserve_selection: bool = True) -> None:
        selected_ids = _selected_room_ids() if preserve_selection else set()
        if active_editor["widget"] is not None:
            try:
                active_editor["widget"].destroy()
            except tk.TclError:
                pass
            active_editor["widget"] = None
        for item in tree.get_children():
            tree.delete(item)
        item_to_row.clear()

        for idx, row in enumerate(filtered_rows()):
            iid = f"r{idx}"
            method_lower = row.assignment_method.lower()
            tag = "missing" if not row.streamer else ("manual" if method_lower.startswith("session") else ("manual" if "manual" in method_lower else ""))
            values = (
                "NEEDS" if not row.streamer else "✓",
                row.start_time.strftime("%Y-%m-%d"),
                row.start_time.strftime("%I:%M %p").lstrip("0"),
                row.end_time.strftime("%I:%M %p").lstrip("0") if row.end_time else "",
                f"{minutes_to_decimal_hours(row.duration_minutes):.2f}",
                f"${row.gmv:,.2f}",
                row.room_title or "[BLANK TITLE]",
                row.streamer or "",
                row.assignment_method,
                row.room_id,
                row.source_file,
            )
            tree.insert("", "end", iid=iid, values=values, tags=((tag,) if tag else ()))
            item_to_row[iid] = row
            if row.room_id in selected_ids:
                tree.selection_add(iid)

        if report:
            status_var.set(f"Showing {len(tree.get_children())} of {len(report.rows)} rows")
        else:
            status_var.set("Ready")

    def update_source_filter() -> None:
        sources = ["All sources"]
        if report:
            sources.extend(sorted({r.source_file for r in report.rows}))
        source_combo.configure(values=sources)
        if source_var.get() not in sources:
            source_var.set("All sources")

    def scan_selected() -> None:
        nonlocal report
        if not selected_paths:
            report = None
            update_source_filter()
            refresh_tree(False)
            refresh_summary()
            return
        try:
            root.configure(cursor="watch")
            status_var.set("Importing and deduplicating TikTok exports…")
            root.update_idletasks()
            report = build_report_data(selected_paths, state)
            update_source_filter()
            refresh_tree(False)
            refresh_summary()
            status_var.set("Import complete")
        except Exception as exc:
            report = None
            update_source_filter()
            refresh_tree(False)
            refresh_summary()
            messagebox.showerror("Could not import TikTok exports", str(exc))
            status_var.set("Import failed")
        finally:
            root.configure(cursor="")

    def add_files() -> None:
        picked = filedialog.askopenfilenames(
            title="Select TikTok CSV / XLSX exports",
            filetypes=[("TikTok exports", "*.csv *.xlsx"), ("CSV files", "*.csv"), ("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if not picked:
            return
        existing = {str(p.resolve()) for p in selected_paths}
        for raw in picked:
            p = Path(raw).expanduser()
            key = str(p.resolve())
            if key not in existing:
                selected_paths.append(p)
                existing.add(key)
        refresh_file_list()
        scan_selected()

    def remove_files() -> None:
        chosen = list(file_list.curselection())
        if not chosen:
            return
        for i in reversed(chosen):
            del selected_paths[i]
        refresh_file_list()
        scan_selected()

    def clear_files() -> None:
        nonlocal report
        selected_paths.clear()
        report = None
        refresh_file_list()
        update_source_filter()
        refresh_tree(False)
        refresh_summary()
        status_var.set("Ready")

    ttk.Button(file_buttons, text="Add Exports…", command=add_files).pack(fill="x")
    ttk.Button(file_buttons, text="Remove Selected", command=remove_files).pack(fill="x", pady=(4, 0))
    ttk.Button(file_buttons, text="Clear", command=clear_files).pack(fill="x", pady=(4, 0))

    def selected_rows() -> List[StreamRow]:
        return [item_to_row[i] for i in tree.selection() if i in item_to_row]

    def persist_and_refresh() -> None:
        save_state(state)
        refresh_tree(True)
        refresh_summary()

    def bulk_assign() -> None:
        if not report:
            return
        rows = selected_rows()
        if not rows:
            messagebox.showinfo("Select rows", "Select one or more rows in the table first.")
            return
        name = _clean_name(bulk_var.get())
        if not name:
            messagebox.showwarning("Streamer required", "Choose or type a streamer name first.")
            bulk_combo.focus_set()
            return
        set_many_assignments(report, rows, name, "manual: bulk table assignment", state)
        persist_and_refresh()
        status_var.set(f"Assigned {len(rows)} row(s) to {name}")

    def clear_selected_assignments() -> None:
        if not report:
            return
        rows = selected_rows()
        if not rows:
            messagebox.showinfo("Select rows", "Select one or more rows first.")
            return
        for row in rows:
            row.streamer = None
            row.assignment_method = "manual: cleared assignment"
            state["room_assignments"][row.room_id] = None
        _refresh_report_state(report)
        persist_and_refresh()
        status_var.set(f"Cleared {len(rows)} assignment(s)")

    def source_neighbor(row: StreamRow, direction: int) -> Optional[StreamRow]:
        if not report:
            return None
        seq = sorted([r for r in report.rows if r.source_file == row.source_file], key=lambda r: (r.start_time, r.room_id))
        try:
            idx = seq.index(row)
        except ValueError:
            return None
        j = idx + direction
        return seq[j] if 0 <= j < len(seq) else None

    def assign_from_neighbor(direction: int) -> None:
        if not report:
            return
        rows = selected_rows()
        if len(rows) != 1:
            messagebox.showinfo("Select one row", "Select exactly one row for Previous/Next assignment.")
            return
        target = rows[0]
        neighbor = source_neighbor(target, direction)
        label = "previous" if direction < 0 else "next"
        if neighbor is None:
            messagebox.showinfo("No neighboring row", f"There is no {label} row in this source export.")
            return
        if not neighbor.streamer:
            messagebox.showinfo("Neighbor is unassigned", f"The {label} row does not have a streamer yet.")
            return
        set_row_assignment(report, target, neighbor.streamer, f"manual: copied {label} row", state)
        persist_and_refresh()
        status_var.set(f"Assigned selected row to {neighbor.streamer} from {label} row")

    def rename_merge() -> None:
        if not report:
            return
        old = ""
        rows = selected_rows()
        assigned = sorted({r.streamer for r in rows if r.streamer})
        if len(assigned) == 1:
            old = assigned[0]
        if not old:
            old = simpledialog.askstring("Rename / Merge Name", "Existing streamer name:", parent=root) or ""
        old = _clean_name(old)
        if not old:
            return
        if old not in {r.streamer for r in report.rows if r.streamer}:
            messagebox.showwarning("Name not found", f"No current rows are assigned to ‘{old}’.")
            return
        new = simpledialog.askstring(
            "Rename / Merge Name",
            f"Replace every current ‘{old}’ assignment with:\n\n(Use an existing name to merge them.)",
            initialvalue=old,
            parent=root,
        )
        new = _clean_name(new)
        if not new or new == old:
            return
        count = rename_streamer(report, old, new, state)
        save_state(state)
        bulk_var.set(new)
        refresh_tree(True)
        refresh_summary()
        status_var.set(f"Renamed/merged {count} row(s): {old} → {new}")

    assign_btn.configure(command=bulk_assign)
    clear_assign_btn.configure(command=clear_selected_assignments)
    prev_btn.configure(command=lambda: assign_from_neighbor(-1))
    next_btn.configure(command=lambda: assign_from_neighbor(1))
    rename_btn.configure(command=rename_merge)

    def begin_cell_edit(event=None) -> None:
        if not report:
            return
        if event is not None:
            item = tree.identify_row(event.y)
            col_num = tree.identify_column(event.x)
            streamer_index = columns.index("streamer") + 1
            if not item or col_num != f"#{streamer_index}":
                return
        else:
            sel = tree.selection()
            if not sel:
                return
            item = sel[0]
        row = item_to_row.get(item)
        if not row:
            return
        bbox = tree.bbox(item, "streamer")
        if not bbox:
            return

        if active_editor["widget"] is not None:
            try:
                active_editor["widget"].destroy()
            except tk.TclError:
                pass
        x, y, w, h = bbox
        var = tk.StringVar(value=row.streamer or "")
        editor = ttk.Combobox(tree, textvariable=var, values=current_name_choices(), state="normal")
        editor.place(x=x, y=y, width=max(w, 150), height=h)
        active_editor["widget"] = editor
        editor.focus_set()
        editor.selection_range(0, tk.END)
        finished = {"done": False}

        def close(commit: bool) -> None:
            if finished["done"]:
                return
            finished["done"] = True
            value = _clean_name(var.get())
            try:
                editor.destroy()
            except tk.TclError:
                pass
            active_editor["widget"] = None
            if not commit:
                return
            set_row_assignment(report, row, value, "manual: table cell edit", state)
            save_state(state)
            refresh_tree(True)
            refresh_summary()
            status_var.set(f"Updated streamer for {row.start_time:%Y-%m-%d %I:%M %p}")

        def _commit_key(event=None):
            close(True)
            return "break"

        def _cancel_key(event=None):
            close(False)
            return "break"

        editor.bind("<Return>", _commit_key)
        editor.bind("<Tab>", _commit_key)
        editor.bind("<Escape>", _cancel_key)
        editor.bind("<<ComboboxSelected>>", _commit_key)
        # Focus-out commits typed edits; the done guard prevents duplicate commits
        # when ComboboxSelected and FocusOut happen close together.
        editor.bind("<FocusOut>", lambda e: root.after_idle(lambda: close(True)))

    tree.bind("<Double-1>", begin_cell_edit)
    tree.bind("<Return>", lambda e: begin_cell_edit())

    def on_filter_change(event=None) -> None:
        refresh_tree(False)

    filter_combo.bind("<<ComboboxSelected>>", on_filter_change)
    source_combo.bind("<<ComboboxSelected>>", on_filter_change)
    search_entry.bind("<KeyRelease>", on_filter_change)

    def show_missing() -> None:
        filter_var.set("Needs streamer")
        refresh_tree(False)
        children = tree.get_children()
        if children:
            tree.selection_set(children[0])
            tree.focus(children[0])
            tree.see(children[0])

    show_missing_btn.configure(command=show_missing)

    def generate() -> None:
        if not report:
            messagebox.showinfo("Add exports", "Add at least one TikTok CSV/XLSX export first.")
            return
        if report.unresolved:
            show_missing()
            messagebox.showwarning("Streamer assignments required", f"{len(report.unresolved)} stream(s) still need a streamer. Fill the Streamer column before generating the report.")
            return
        if xlsxwriter is None:
            messagebox.showerror("Missing dependency", "This app needs XlsxWriter. Run:\n\npython3 -m pip install XlsxWriter")
            return
        output = filedialog.asksaveasfilename(
            title="Save final streamer report",
            defaultextension=".xlsx",
            initialfile="TikTok_Streamer_Report.xlsx",
            filetypes=[("Excel workbook", "*.xlsx")],
        )
        if not output:
            return
        try:
            root.configure(cursor="watch")
            status_var.set("Generating final Excel report…")
            root.update_idletasks()
            save_state(state)
            write_excel(report, Path(output))
            status_var.set("Report created")
            messagebox.showinfo(
                "Report created",
                f"Saved:\n{output}\n\n{len(report.rows)} streams • 100% assigned • {len(report.streamers)} streamer(s)",
            )
        except Exception as exc:
            status_var.set("Report generation failed")
            messagebox.showerror("Could not create report", str(exc))
        finally:
            root.configure(cursor="")

    generate_btn.configure(command=generate)

    # Useful keyboard shortcuts without changing the data model.
    def _focus_search(event=None):
        search_entry.focus_set()
        search_entry.selection_range(0, tk.END)
        return "break"

    root.bind("<Command-f>", _focus_search)
    root.bind("<Control-f>", _focus_search)

    update_source_filter()
    update_name_widgets()
    refresh_summary()
    refresh_tree(False)
    root.mainloop()


# -------------------------------- CLI ----------------------------------

def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="Build a reviewed streamer GMV/hours report from TikTok CSV/XLSX exports.")
    parser.add_argument("inputs", nargs="*", help="TikTok CSV or XLSX export files")
    parser.add_argument("-o", "--output", help="Output .xlsx path")
    args = parser.parse_args(argv)

    if not args.inputs:
        launch_gui()
        return 0

    paths = [Path(p).expanduser().resolve() for p in args.inputs]
    missing = [str(p) for p in paths if not p.exists()]
    if missing:
        print("Missing file(s):", *missing, sep="\n  ", file=sys.stderr)
        return 2

    state = load_state()
    try:
        report = build_report_data(paths, state)
    except Exception as exc:
        print(f"Import failed: {exc}", file=sys.stderr)
        return 2

    if report.unresolved:
        print(
            f"Not creating a report: {len(report.unresolved)} of {len(report.rows)} stream(s) still need a streamer.\n"
            "Open the program without command-line input and complete the Streamer column first.",
            file=sys.stderr,
        )
        for row in report.unresolved[:20]:
            print(f"  {row.start_time:%Y-%m-%d %H:%M} | {row.room_title or '[BLANK TITLE]'} | {row.source_file}", file=sys.stderr)
        if len(report.unresolved) > 20:
            print(f"  ...and {len(report.unresolved) - 20} more", file=sys.stderr)
        return 3

    output = Path(args.output).expanduser().resolve() if args.output else Path.cwd() / "TikTok_Streamer_Report.xlsx"
    try:
        write_excel(report, output)
    except Exception as exc:
        print(f"Report generation failed: {exc}", file=sys.stderr)
        return 4

    print(f"Created: {output}")
    print(f"Unique live rooms: {len(report.rows)}")
    print(f"Duplicates removed: {report.duplicates_removed}")
    print(f"Streamers: {', '.join(report.streamers)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
