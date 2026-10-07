import json
import re
from pathlib import Path

import pandas as pd


# ============================================================
# CONFIG
# ============================================================

BASE_DIR = Path(__file__).resolve().parent.parent

INPUT_FILE = BASE_DIR / "data" / "lineage.xlsx"
OUTPUT_DIR = BASE_DIR / "metadata"

OUTPUT_DIR.mkdir(exist_ok=True)

HEADER_MARKERS = (
    "section",
    "sub_section",
    "sub section",
    "data entit",
    "report",
    "table",
    "base table",
    "migration",
    "status",
    "direct",
    "derived",
    "tab",
)


# ============================================================
# HELPERS
# ============================================================

def clean(value):
    """Clean Excel cell values."""
    if pd.isna(value):
        return None

    value = str(value).strip()

    if not value or value.lower() in {"nan", "none", "null", "-", "--"}:
        return None

    return value


def normalize_name(value):
    """Normalize table/column names for matching."""
    value = clean(value)

    if not value:
        return None

    # Oracle names sometimes arrive as OWNER.TABLE.
    if "." in value:
        value = value.split(".")[-1]

    value = value.upper()
    value = re.sub(r"[^A-Z0-9_]", "_", value)
    value = re.sub(r"_+", "_", value)

    return value.strip("_") or None


def find_column(df, candidates, exclude=()):
    """
    Find a column whose name contains one of the candidate words.
    """
    columns = []

    for col in df.columns:
        col_lower = str(col).strip().lower()

        if not col_lower or col_lower.startswith("unnamed"):
            continue

        if any(word in col_lower for word in exclude):
            continue

        columns.append((col_lower, col))

    normalized = {col_lower: col for col_lower, col in columns}

    for candidate in candidates:
        if candidate.lower() in normalized:
            return normalized[candidate.lower()]

    for col_lower, col in columns:
        for candidate in candidates:
            if candidate.lower() in col_lower:
                return col

    return None


def split_tables(value):
    """
    Convert a cell containing multiple table names into a list.

    Handles separators such as:
        ,
        ;
        newline
        |
    """
    value = clean(value)

    if not value:
        return []

    parts = re.split(r"[,;\n|]+", value)

    tables = []

    for part in parts:
        part = normalize_name(part)

        if part:
            tables.append(part)

    return list(dict.fromkeys(tables))


def find_header_row(raw):
    """
    Title rows sit above the real headers in this workbook.
    Pick the early row that looks most like a header.
    """
    best_row = 0
    best_score = -1

    for index in range(min(12, len(raw))):
        score = 0

        for value in raw.iloc[index].tolist():
            text = clean(value)

            if not text or len(text) > 60:
                continue

            lowered = text.lower()

            if any(marker in lowered for marker in HEADER_MARKERS):
                score += 1

        if score > best_score:
            best_row = index
            best_score = score

    return best_row


def load_sheet(sheet_name):
    raw = pd.read_excel(
        INPUT_FILE,
        sheet_name=sheet_name,
        header=None,
    )

    header_row = find_header_row(raw)

    df = pd.read_excel(
        INPUT_FILE,
        sheet_name=sheet_name,
        header=header_row,
    )

    df = df.dropna(axis=1, how="all")
    df = df.dropna(how="all")
    df.columns = [str(col).strip() for col in df.columns]

    print(f"  header row: {header_row + 1}")

    return df


def fill_down(df, columns):
    for column in columns:
        if not column or column not in df.columns:
            continue

        df[column] = df[column].map(
            lambda value: clean(value) if clean(value) else pd.NA
        )
        df[column] = df[column].ffill()

    return df


def save_json(filename, data):
    path = OUTPUT_DIR / filename

    with open(path, "w", encoding="utf-8") as f:
        json.dump(
            data,
            f,
            indent=2,
            ensure_ascii=False,
        )

    print(f"[OK] {path}")


def print_columns(df):
    print("Columns:")

    for col in df.columns:
        print(f"  - {col}")


# ============================================================
# SHEET EXTRACTORS
# ============================================================

def extract_lighthouse(df):
    """ERP lighthouse sections, with merged cells filled down."""
    section_col = find_column(
        df,
        ["lighthouse section", "section"],
        exclude=("sub",),
    )
    subsection_col = find_column(
        df,
        ["sub_section_1", "sub section_1", "sub_section 1"],
    )
    report_col = find_column(
        df,
        ["sub_section_2", "sub section_2", "sub_section 2"],
    )
    entity_col = find_column(
        df,
        ["data entit", "data entity", "entities used"],
    )

    fill_down(df, [section_col, subsection_col, report_col])

    groups = {}
    order = []

    for _, row in df.iterrows():
        entities = split_tables(row[entity_col]) if entity_col else []

        if not entities:
            continue

        key = (
            clean(row[section_col]) if section_col else None,
            clean(row[subsection_col]) if subsection_col else None,
            clean(row[report_col]) if report_col else None,
        )

        if key not in groups:
            groups[key] = []
            order.append(key)

        for entity in entities:
            if entity not in groups[key]:
                groups[key].append(entity)

    return [
        {
            "tab": tab,
            "subsection": subsection,
            "report": report,
            "tables": groups[(tab, subsection, report)],
        }
        for tab, subsection, report in order
    ]


def extract_reports(df):
    """TOC AI bot report catalog."""
    tab_col = find_column(df, ["base tab", "base tabs"])
    subsection_col = find_column(df, ["sub tab", "subsection"])
    report_col = find_column(df, ["report name", "report"])
    table_col = find_column(
        df,
        ["table name", "table"],
        exclude=("type",),
    )

    fill_down(df, [tab_col, subsection_col])

    reports = []

    for _, row in df.iterrows():
        report = clean(row[report_col]) if report_col else None
        tables = split_tables(row[table_col]) if table_col else []

        if not report and not tables:
            continue

        reports.append({
            "tab": clean(row[tab_col]) if tab_col else None,
            "subsection": clean(row[subsection_col]) if subsection_col else None,
            "report": report,
            "tables": tables,
        })

    return reports


def extract_lineage(df):
    adw_col = find_column(
        df,
        ["adw_to", "adw table", "table_list", "table list"],
    )
    type_col = find_column(
        df,
        ["type_of_table", "type"],
        exclude=("table_list", "adw_to"),
    )
    source_col = find_column(
        df,
        ["base table", "base tables"],
    )

    lineage = []

    for _, row in df.iterrows():
        target = normalize_name(row[adw_col]) if adw_col else None

        if not target:
            continue

        lineage.append({
            "adw_table": target,
            "type": clean(row[type_col]) if type_col else None,
            "source_tables": split_tables(row[source_col]) if source_col else [],
        })

    return lineage


def extract_migration(df):
    table_col = find_column(
        df,
        ["data entit", "table name"],
        exclude=("type", "status"),
    )
    type_col = find_column(df, ["type_of_table", "type"])
    status_col = find_column(df, ["migration status", "status"])

    migration_status = []

    for _, row in df.iterrows():
        table = normalize_name(row[table_col]) if table_col else None

        if not table:
            continue

        migration_status.append({
            "table": table,
            "type": clean(row[type_col]) if type_col else None,
            "status": clean(row[status_col]) if status_col else None,
        })

    return migration_status


def sheet_kind(df):
    names = " | ".join(str(col).lower() for col in df.columns)

    if "migration status" in names:
        return "migration"

    if "base table" in names or "adw_to" in names:
        return "lineage"

    if "report name" in names or "base tab" in names:
        return "reports"

    if "data entit" in names or "lighthouse section" in names:
        return "lighthouse"

    return None


# ============================================================
# LOAD EXCEL
# ============================================================

print(f"\nReading: {INPUT_FILE}")

if not INPUT_FILE.exists():
    raise FileNotFoundError(
        f"Could not find {INPUT_FILE}"
    )

excel = pd.ExcelFile(INPUT_FILE)

print("\nSheets found:")

for sheet in excel.sheet_names:
    print(f"  - {sheet}")


reports = []
lineage = []
migration_status = []

for sheet_name in excel.sheet_names:
    print(f"\nProcessing sheet: {sheet_name}")

    df = load_sheet(sheet_name)
    print_columns(df)

    kind = sheet_kind(df)
    print(f"  detected: {kind}")

    if kind == "lighthouse":
        reports.extend(extract_lighthouse(df))
    elif kind == "reports":
        reports.extend(extract_reports(df))
    elif kind == "lineage":
        lineage.extend(extract_lineage(df))
    elif kind == "migration":
        migration_status.extend(extract_migration(df))


save_json("reports.json", reports)
save_json("lineage.json", lineage)
save_json("migration_status.json", migration_status)


# ============================================================
# BUILD TABLE CATALOG
# ============================================================

tables = {}

for item in lineage:
    table = item["adw_table"]

    tables.setdefault(
        table,
        {
            "table": table,
            "type": item.get("type"),
            "source_tables": item.get("source_tables", []),
        },
    )

for report in reports:
    for table in report["tables"]:
        tables.setdefault(
            table,
            {
                "table": table,
                "type": None,
                "source_tables": [],
            },
        )

for item in migration_status:
    tables.setdefault(
        item["table"],
        {
            "table": item["table"],
            "type": item.get("type"),
            "source_tables": [],
        },
    )

save_json("tables.json", list(tables.values()))


# ============================================================
# SUMMARY
# ============================================================

print("\n" + "=" * 60)
print("EXTRACTION COMPLETE")
print("=" * 60)

for file in sorted(OUTPUT_DIR.iterdir()):
    if file.suffix != ".json":
        continue

    with open(file, encoding="utf-8") as f:
        data = json.load(f)

    print(f"{file.name:<25} {len(data):>6} records")

print("=" * 60)
