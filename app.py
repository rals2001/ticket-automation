import re
from datetime import datetime
from pathlib import Path

import pandas as pd
import geopandas as gpd
from openpyxl import load_workbook

# =========================
# FILE PATHS (relative to this file's location)
# =========================
BASE_DIR = Path(__file__).resolve().parent

INPUT_FOLDER = BASE_DIR / "Input CSV"
SHAPEFILE = BASE_DIR / "Shapefile" / "gadm41_IDN_3.shp"
OUTPUT_DIR = BASE_DIR / "Output"

# Indonesian month names
bulan_id = {
    1: "Januari",
    2: "Februari",
    3: "Maret",
    4: "April",
    5: "Mei",
    6: "Juni",
    7: "Juli",
    8: "Agustus",
    9: "September",
    10: "Oktober",
    11: "November",
    12: "Desember"
}

today = datetime.now()
tanggal_indonesia = f"{today.day:02d} {bulan_id[today.month]}"

OUTPUT_FILE = OUTPUT_DIR / f"Tiket Indihome {tanggal_indonesia} butuh feedback.xlsx"

# Create the Output folder automatically if it does not exist
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

WAKTU_COL = 6  # Column F

# How the column should look in the final Excel file — minutes only, no seconds
OUTPUT_NUMBER_FORMAT = "dd/mm/yyyy hh:mm"

# Formats we'll try, in order, to parse the raw source text
CANDIDATE_FORMATS = [
    "%d/%m/%Y %H:%M:%S",
    "%d/%m/%Y %H:%M",
    "%Y-%m-%d %H:%M:%S",
    "%Y-%m-%d %H:%M:%S.%f",
    "%Y-%m-%d %H:%M",
    "%d-%m-%Y %H:%M:%S",
    "%m/%d/%Y %H:%M:%S",
]


def manual_parse_csv_text(raw_text):
    """
    Pure character-by-character CSV parser (no csv module, no pandas).
    Respects quoted fields so commas inside quotes don't break column
    alignment. Performs zero interpretation of values.
    """
    rows = []
    field = []
    row = []
    in_quotes = False
    i = 0
    n = len(raw_text)

    while i < n:
        ch = raw_text[i]

        if in_quotes:
            if ch == '"':
                if i + 1 < n and raw_text[i + 1] == '"':
                    field.append('"')
                    i += 2
                    continue
                in_quotes = False
                i += 1
                continue
            field.append(ch)
            i += 1
            continue

        if ch == '"':
            in_quotes = True
            i += 1
            continue
        elif ch == ",":
            row.append("".join(field))
            field = []
            i += 1
            continue
        elif ch == "\r":
            i += 1
            continue
        elif ch == "\n":
            row.append("".join(field))
            field = []
            rows.append(row)
            row = []
            i += 1
            continue
        else:
            field.append(ch)
            i += 1
            continue

    if field or row:
        row.append("".join(field))
        rows.append(row)

    return rows


def parse_waktu_value(raw_value):
    """
    Converts raw column F text into a real Python datetime object.
    Tries several known formats explicitly. Falls back to pandas'
    flexible parser (dayfirst=True) if none match. Returns the
    original raw string only if every attempt fails.
    """
    cleaned = re.sub(r"\s+", " ", str(raw_value).strip())

    if not cleaned:
        return raw_value

    for fmt in CANDIDATE_FORMATS:
        try:
            return datetime.strptime(cleaned, fmt)
        except ValueError:
            continue

    # Fallback: flexible parsing, day-first (26/09/2026 style)
    try:
        parsed = pd.to_datetime(cleaned, dayfirst=True, errors="raise")
        return parsed.to_pydatetime()
    except Exception:
        return raw_value


# =========================
# CHECK REQUIRED FOLDERS / FILES
# =========================
if not INPUT_FOLDER.is_dir():
    raise FileNotFoundError(
        f"Input folder not found: {INPUT_FOLDER}\n"
        "Create a folder named 'Input CSV' next to app.py and put the CSV files in it."
    )

if not SHAPEFILE.is_file():
    raise FileNotFoundError(
        f"Shapefile not found: {SHAPEFILE}\n"
        "Place gadm41_IDN_3.shp (plus .shx, .dbf, .prj) in a folder named 'Shapefile' next to app.py."
    )

# =========================
# FIND CSV FILES
# =========================
csv_files = sorted(
    (
        p for p in INPUT_FOLDER.iterdir()
        if p.is_file() and p.suffix.lower() == ".csv"
    ),
    key=lambda p: p.name.lower(),
)

if not csv_files:
    raise ValueError("No CSV files found in input folder.")

# =========================
# SAVE ORIGINAL WAKTU COLUMN (F) AS PARSED DATETIME
# =========================

waktu_data = []
waktu_col_name = None
all_dfs = []

for csv_file in csv_files:

    print(f"Reading: {csv_file}")

    with open(csv_file, "r", encoding="utf-8-sig", newline="") as f:
        raw_text = f.read()

    parsed_rows = manual_parse_csv_text(raw_text)

    if not parsed_rows:
        raise ValueError(f"File '{csv_file}' appears to be empty.")

    header = parsed_rows[0]

    if len(header) < WAKTU_COL:
        raise ValueError(
            f"File '{csv_file}' has fewer than {WAKTU_COL} columns."
        )

    this_waktu_col_name = header[WAKTU_COL - 1].strip()

    if waktu_col_name is None:
        waktu_col_name = this_waktu_col_name

    for data_row in parsed_rows[1:]:
        raw_value = data_row[WAKTU_COL - 1] if len(data_row) > WAKTU_COL - 1 else ""
        waktu_data.append(parse_waktu_value(raw_value))

    # Read the SAME file with pandas, EXCLUDING column F entirely
    temp_df = pd.read_csv(
        csv_file,
        dtype=str,
        keep_default_na=False,
        low_memory=False,
        encoding="utf-8-sig",
        usecols=lambda col: col.strip() != this_waktu_col_name,
    )

    temp_df.columns = temp_df.columns.astype(str).str.strip()

    all_dfs.append(temp_df)

print("DEBUG - first 3 waktu_data values:", waktu_data[:3])
print("DEBUG - types:", [type(v) for v in waktu_data[:3]])

# =========================
# MERGE CSV FILES (WITHOUT COLUMN F)
# =========================
df = pd.concat(all_dfs, ignore_index=True)
df.columns = df.columns.astype(str).str.strip()

# =========================
# ENSURE REQUIRED COLUMNS
# =========================
if "Grid GeoJSON" not in df.columns:
    raise ValueError("Column 'Grid GeoJSON' not found.")

if "Mapping RCA" not in df.columns:
    print("'Mapping RCA' not found. Creating column...")
    df["Mapping RCA"] = ""

# =========================
# KEEP ONLY COLUMNS FROM A UNTIL MAPPING RCA
# =========================
mapping_idx = df.columns.get_loc("Mapping RCA")
keep_cols = list(df.columns[:mapping_idx + 1])

if "Grid GeoJSON" not in keep_cols:
    keep_cols.append("Grid GeoJSON")

df = df[keep_cols]

# =========================
# EXTRACT LONGITUDE & LATITUDE
# =========================
df["Longitude"] = df["Grid GeoJSON"].astype(str).str.extract(
    r"\[\[\[\[\s*([-\d\.]+)"
)
df["Latitude"] = df["Grid GeoJSON"].astype(str).str.extract(
    r"\[\[\[\[\s*[-\d\.]+\s*,\s*([-\d\.]+)"
)

df["Longitude"] = pd.to_numeric(df["Longitude"], errors="coerce")
df["Latitude"] = pd.to_numeric(df["Latitude"], errors="coerce")

# =========================
# CREATE GEODATAFRAME
# =========================
gdf_points = gpd.GeoDataFrame(
    df.copy(),
    geometry=gpd.points_from_xy(df["Longitude"], df["Latitude"]),
    crs="EPSG:4326"
)

# =========================
# READ SHAPEFILE
# =========================
gdf_admin = gpd.read_file(SHAPEFILE)

if gdf_admin.crs != gdf_points.crs:
    gdf_admin = gdf_admin.to_crs(gdf_points.crs)

# =========================
# SPATIAL JOIN
# =========================
result = gpd.sjoin(
    gdf_points,
    gdf_admin,
    how="left",
    predicate="within"
)

# =========================
# FALLBACK FOR BOUNDARY POINTS
# =========================
missing = result["NAME_1"].isna()

if missing.any():
    fallback = gpd.sjoin_nearest(
        gdf_points[missing],
        gdf_admin,
        how="left"
    )
    result.loc[missing, ["NAME_1", "NAME_2"]] = fallback[["NAME_1", "NAME_2"]].values

# =========================
# ADD PROVINSI & KABUPATEN
# =========================
df["Provinsi"] = result["NAME_1"]
df["Kabupaten/Kota"] = result["NAME_2"]

# =========================
# INSERT AFTER MAPPING RCA
# =========================
mapping_idx = df.columns.get_loc("Mapping RCA")

provinsi = df.pop("Provinsi")
kabupaten = df.pop("Kabupaten/Kota")

df.insert(mapping_idx + 1, "Provinsi", provinsi)
df.insert(mapping_idx + 2, "Kabupaten/Kota", kabupaten)

# =========================
# RE-INSERT WAKTU AS BLANK PLACEHOLDER
# =========================

if len(waktu_data) != len(df):
    print(
        "WARNING: number of column F values "
        f"({len(waktu_data)}) does not match number of "
        f"output rows ({len(df)}). Check source CSVs for "
        "blank lines or ragged rows."
    )

insert_pos = min(WAKTU_COL - 1, len(df.columns))
df.insert(insert_pos, waktu_col_name, "")

# =========================
# SAVE RESULT
# =========================
df.to_excel(
    OUTPUT_FILE,
    index=False
)

# =========================
# OPEN OUTPUT FILE
# =========================
wb_output = load_workbook(OUTPUT_FILE)
ws_output = wb_output.active

# =========================
# CLEAR COLUMN F
# =========================
for row in range(2, ws_output.max_row + 1):
    ws_output.cell(row=row, column=WAKTU_COL).value = None

# =========================
# RESTORE COLUMN F AS REAL DATETIME
# =========================
converted_count = 0
fallback_count = 0

for row, value in enumerate(waktu_data, start=2):

    if row > ws_output.max_row:
        break

    tgt = ws_output.cell(row=row, column=WAKTU_COL)
    tgt.value = value

    if isinstance(value, datetime):
        tgt.number_format = OUTPUT_NUMBER_FORMAT
        converted_count += 1
    else:
        tgt.number_format = "@"
        fallback_count += 1

# =========================
# SAVE FINAL FILE
# =========================
wb_output.save(OUTPUT_FILE)

print("Waktu column converted to real datetime values.")
print("Completed!")
print(f"Rows processed : {len(df)}")
print(f"CSV files processed : {len(csv_files)}")
print(f"Converted to datetime : {converted_count}")
print(f"Left as text (unparsed) : {fallback_count}")
print(f"Output saved to : {OUTPUT_FILE}")