"""
CAPA Merge Utility
==================
Drop your year tracking files into this folder and run:
    python merge.py

Outputs a single CAPA_Master.xlsx ready to upload to the dashboard.
Supports 2023+ file formats across CAR, PTO, PAR, CAF sheets.
"""

import os
import re
import sys
import pandas as pd
from pathlib import Path
from datetime import datetime

# ── Config ────────────────────────────────────────────────────────
OUTPUT_FILE   = "CAPA_Master.xlsx"
SKIP_LOCS     = ['A&B Labs', 'VOIDED', 'Extras', 'Warehouse', 'Additives',
                 'Utah', 'Cameron', 'Specialty', 'Kenner', 'Santurce', 'Boucherville']

# ── Sheet name patterns by doc type ───────────────────────────────
SHEET_PATTERNS = {
    'CAR': [r'CAR', r'car'],
    'PTO': [r'PTO', r'pto'],
    'PAR': [r'PAR', r'par'],
    'CAF': [r'CAF', r'caf'],
}

# ── Column normalization maps ──────────────────────────────────────
# Each entry: (normalized_name, [possible raw column names])
CAR_COL_MAP = [
    ('location',      ['Location \n(drop-down)', 'Location']),
    ('location_id',   ['Location ID']),
    ('car_number',    ['CAR #']),
    ('car_type',      ['CAR Type']),
    ('area',          ['Area \n(Lab, Ops, Misc.)\n(drop-down)',
                       'Area (Lab, Ops, PT, Misc.)\n(drop-down)', 'Area']),
    ('description',   ['BRIEF DESCRIPTION OF NC\nif VOID:  "VOID, Location, type, \ndetails (initials, date)"']),
    ('init_date',     ['CAR initialized date']),
    ('close_date',    ['Complete corrective actions (Approved Date)',  # 2025 col M
                       'Corrective Action Approved Date',        # 2026
                       'Corrective Action Completion Date',      # 2025 col L fallback
                       'Effectiveness Review & date deemed effective',  # 2023/2024
                       'Effectiveness Review (Closed Date)',     # 2025 alt
                       'Effectiveness Review Date',              # 2026 alt
                       'Date Corrective Action completed']),     # 2023 fallback
    ('effectiveness_date', ['Effectiveness Review (Closed Date)',
                            'Effectiveness Review Date',
                            'Effectiveness Review & date deemed effective']),
    ('qa_initials',   ['QA team member initials']),
    ('notes',         ['Notes']),
    ('days_open',     ['Days open if not closed', 'Days to Close\n(Cells in red are not closed)']),
]

PTO_COL_MAP = [
    ('location',      ['Location \n(drop-down)', 'Location']),
    ('location_id',   ['Location ID']),
    ('pto_number',    ['PTO #']),
    ('pt_program',    ['PT Program']),
    ('parameter',     ['Parameter (s)']),
    ('z_score',       ['Z-score (s)']),
    ('description',   ['BRIEF DESCRIPTION OF NC\nif VOID:  "VOID, Location,  \ndetails (initials, date)"']),
    ('init_date',     ['PTO initialized date']),
    ('close_date',    ['Effectiveness Review & date deemed effective',  # 2024+
                       'Date Corrective Action completed']),            # 2023
    ('qa_initials',   ['QA team member initials']),
    ('notes',         ['Notes']),
    ('days_open',     ['Days open if not closed', 'Days to Close\n(Cells in red are not closed)']),
]

PAR_COL_MAP = [
    ('location',      ['Location \n(drop-down)', 'Location']),
    ('location_id',   ['Location ID']),
    ('par_number',    ['PAR #']),
    ('par_type',      ['PAR type']),
    ('description',   ['BRIEF DESCRIPTION OF PAR\nif VOID:  Location, type, details (initials, date)']),
    ('init_date',     ['PAR initialized date']),
    ('close_date',    ['Date closed']),
    ('area',          ['Area (Ops, Lab, Misc.)\n(drop-down)',
                       'Area (Ops, Lab, Misc., PT)\n(drop-down)']),
    ('qa_initials',   ['QA team member initials']),
    ('notes',         ['Notes']),
]

CAF_COL_MAP = [
    ('location',      ['Location \n(drop-down)', 'Location']),
    ('location_id',   ['Location ID']),
    ('caf_number',    ['CAF#']),
    ('area',          ['Area']),
    ('init_date',     ['Date initiated']),
    ('shared_date',   ['Date shared w/ customer']),
    ('close_date',    ['Corrective Action completed, CAR approved.',  # 2025+
                       'Date closed after effectiveness review']),    # 2023-2024
    ('description',   ['Brief description of complaint']),
    ('car_par_ref',   ['CAR / PAR # (if applicable)']),
    ('notes',         ['Notes']),
]

COL_MAPS = {'CAR': CAR_COL_MAP, 'PTO': PTO_COL_MAP,
            'PAR': PAR_COL_MAP, 'CAF': CAF_COL_MAP}

# ══════════════════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════════════════
def find_sheet(xl_sheets, doc_type):
    """Find the right sheet for a doc type regardless of year naming."""
    pattern = SHEET_PATTERNS[doc_type][0]
    for name in xl_sheets:
        if re.search(pattern, name, re.IGNORECASE):
            return name
    return None


def detect_header_row(df_raw):
    """Return 0 if headers are on row 1, else 1 if headers are on row 2."""
    # If 'Location' appears in column names, header=0 is correct
    cols = df_raw.columns.astype(str).str.lower().tolist()
    if any('location' in c for c in cols):
        return 0
    # Otherwise the first row is the real header
    return 1


def normalize_cols(df, col_map):
    """
    Rename raw columns to normalized names using the priority list.
    First match wins. Unrecognized columns are dropped.
    """
    raw_cols = {c.strip(): c for c in df.columns}
    rename   = {}
    used_raw = set()

    for norm_name, candidates in col_map:
        if norm_name in rename.values():
            continue  # already mapped
        for candidate in candidates:
            candidate_stripped = candidate.strip()
            if candidate_stripped in raw_cols and raw_cols[candidate_stripped] not in used_raw:
                rename[raw_cols[candidate_stripped]] = norm_name
                used_raw.add(raw_cols[candidate_stripped])
                break

    df = df.rename(columns=rename)
    # Keep only normalized columns that exist
    keep = [norm for norm, _ in col_map if norm in df.columns]
    return df[keep].copy()


def filter_qa_initials(df, doc_type):
    """Exclude records with specific QA initials by doc type."""
    EXCLUDE = {
        'PTO': ['JN'],
    }
    if doc_type not in EXCLUDE:
        return df
    if 'qa_initials' not in df.columns:
        return df
    before = len(df)
    df = df[~df['qa_initials'].astype(str).str.strip().str.upper().isin(EXCLUDE[doc_type])]
    removed = before - len(df)
    if removed > 0:
        print(f"      ↳ Excluded {removed} {doc_type} rows with initials {EXCLUDE[doc_type]}")
    return df


def clean_location(df):
    """Strip, filter VOIDed/skip locations, drop nulls."""
    if 'location' not in df.columns:
        return df
    df['location'] = df['location'].astype(str).str.strip()
    df = df[~df['location'].isin(['nan', 'None', '', 'NaT'])]
    df = df[df['location'].notna()]
    df = df[~df['location'].apply(lambda x: any(s in x for s in SKIP_LOCS))]
    df = df[~df['location'].str.upper().str.contains('VOID', na=False)]
    return df


def parse_dates(df, date_cols):
    """Coerce date columns to datetime."""
    for col in date_cols:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')
    return df


def infer_status(df, close_col='close_date'):
    """Add a status column: CLOSED if close_date exists, else OPEN."""
    if close_col in df.columns:
        df['status'] = df[close_col].apply(
            lambda x: 'CLOSED' if pd.notna(x) else 'OPEN')
    else:
        df['status'] = 'OPEN'
    return df


def days_to_close(df):
    """Compute days_to_close from init_date to close_date."""
    if 'init_date' in df.columns and 'close_date' in df.columns:
        df['days_to_close'] = (df['close_date'] - df['init_date']).dt.days
        df['days_to_close'] = df['days_to_close'].where(df['status'] == 'CLOSED', other=None)
    return df


def read_sheet(path, sheet_name, doc_type, year):
    """Read, detect header, normalize, clean one sheet."""
    try:
        raw = pd.read_excel(path, sheet_name=sheet_name, header=0)
        raw.columns = raw.columns.astype(str).str.strip()

        # Detect if real header is on row 2
        if detect_header_row(raw) == 1:
            raw = pd.read_excel(path, sheet_name=sheet_name, header=1)
            raw.columns = raw.columns.astype(str).str.strip()

        col_map   = COL_MAPS[doc_type]
        df        = normalize_cols(raw, col_map)
        df        = clean_location(df)
        df        = filter_qa_initials(df, doc_type)
        date_cols = ['init_date', 'close_date', 'effectiveness_date',
                     'shared_date', 'close_date']
        df        = parse_dates(df, date_cols)
        df        = infer_status(df)
        df        = days_to_close(df)
        df['year']     = year
        df['doc_type'] = doc_type
        df['source_file'] = Path(path).name
        return df, []
    except Exception as e:
        return None, [f"  ⚠  {doc_type} {year}: {e}"]


# ══════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════
def main():
    script_dir = Path(__file__).parent
    xlsx_files = sorted([
        f for f in script_dir.glob('*.xlsx')
        if f.name != OUTPUT_FILE and not f.name.startswith('~')
    ])

    if not xlsx_files:
        print("❌  No .xlsx files found in this folder. Drop your year files here and re-run.")
        sys.exit(1)

    print(f"\n{'='*55}")
    print(f"  CAPA Merge Utility")
    print(f"  {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"{'='*55}")
    print(f"\n  Found {len(xlsx_files)} file(s):")
    for f in xlsx_files:
        print(f"    • {f.name}")

    results  = {'CAR': [], 'PTO': [], 'PAR': [], 'CAF': []}
    warnings = []

    for path in xlsx_files:
        # Extract year from filename
        yr_match = re.search(r'(20\d{2})', path.name)
        year     = int(yr_match.group(1)) if yr_match else 0

        try:
            xl = pd.read_excel(path, sheet_name=None)
        except Exception as e:
            warnings.append(f"  ⚠  Could not open {path.name}: {e}")
            continue

        print(f"\n  Processing {path.name} (year={year})...")

        for doc_type in ['CAR', 'PTO', 'PAR', 'CAF']:
            sheet_name = find_sheet(list(xl.keys()), doc_type)
            if not sheet_name:
                warnings.append(f"  ⚠  No {doc_type} sheet found in {path.name}")
                continue

            df, errs = read_sheet(str(path), sheet_name, doc_type, year)
            warnings.extend(errs)

            if df is not None and len(df) > 0:
                results[doc_type].append(df)
                closed = (df['status'] == 'CLOSED').sum()
                open_  = (df['status'] == 'OPEN').sum()
                print(f"    ✓ {doc_type:4s} — {len(df):4d} rows  "
                      f"({closed} closed, {open_} open)  [{sheet_name}]")
            else:
                warnings.append(f"  ⚠  {doc_type} in {path.name} returned 0 usable rows")

    # ── Combine & deduplicate ──────────────────────────────────────
    print(f"\n  Combining and deduplicating...")
    combined = {}
    num_col  = {'CAR': 'car_number', 'PTO': 'pto_number',
                'PAR': 'par_number', 'CAF': 'caf_number'}

    for doc_type, frames in results.items():
        if not frames:
            warnings.append(f"  ⚠  No data collected for {doc_type}")
            combined[doc_type] = pd.DataFrame()
            continue

        df = pd.concat(frames, ignore_index=True)

        # Deduplicate on record number + location (keep latest year)
        id_col = num_col[doc_type]
        if id_col in df.columns and 'location' in df.columns:
            before = len(df)
            df = (df.sort_values('year', ascending=False)
                    .drop_duplicates(subset=[id_col, 'location'], keep='first')
                    .sort_values(['year', id_col], ascending=[True, True]))
            dupes = before - len(df)
            if dupes > 0:
                print(f"    ↳ {doc_type}: removed {dupes} duplicate record(s)")

        combined[doc_type] = df
        print(f"    ✓ {doc_type:4s} — {len(df):4d} total rows across all years")

    # ── Write output ──────────────────────────────────────────────
    out_path = script_dir / OUTPUT_FILE
    print(f"\n  Writing {OUTPUT_FILE}...")

    with pd.ExcelWriter(str(out_path), engine='openpyxl') as writer:
        for doc_type, df in combined.items():
            if df.empty:
                continue
            sheet = f"Data - {doc_type}s"
            df.to_excel(writer, sheet_name=sheet, index=False)

        # Summary sheet
        summary_rows = []
        for doc_type, df in combined.items():
            if df.empty:
                continue
            for yr in sorted(df['year'].unique()):
                ydf = df[df['year'] == yr]
                summary_rows.append({
                    'Doc Type': doc_type,
                    'Year':     yr,
                    'Total':    len(ydf),
                    'Closed':   (ydf['status'] == 'CLOSED').sum(),
                    'Open':     (ydf['status'] == 'OPEN').sum(),
                    'Avg Days to Close': round(
                        ydf.loc[ydf['status']=='CLOSED','days_to_close'].mean(), 1)
                        if (ydf['status']=='CLOSED').any() else 'N/A',
                })
        pd.DataFrame(summary_rows).to_excel(writer, sheet_name='Summary', index=False)

    # ── Report ────────────────────────────────────────────────────
    print(f"\n{'='*55}")
    print(f"  ✅  Done! Output: {out_path}")
    print(f"{'='*55}")

    if warnings:
        print(f"\n  Warnings ({len(warnings)}):")
        for w in warnings:
            print(w)

    print(f"\n  Summary:")
    print(f"  {'Doc Type':<10} {'Total':>7} {'Closed':>8} {'Open':>6}")
    print(f"  {'-'*35}")
    for doc_type, df in combined.items():
        if df.empty:
            continue
        closed = (df['status'] == 'CLOSED').sum()
        open_  = (df['status'] == 'OPEN').sum()
        print(f"  {doc_type:<10} {len(df):>7,} {closed:>8,} {open_:>6,}")

    print(f"\n  Upload {OUTPUT_FILE} to your CAPA dashboard.\n")


if __name__ == '__main__':
    main()
