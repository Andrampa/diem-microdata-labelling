"""
DIEM microdata labelling script

This script takes a DIEM household survey microdata file (CSV or Excel) and produces
a labelled version of the dataset by automatically applying value labels from the
official DIEM codebook.

The script:
- Automatically detects the microdata infrastructure version (pre- or post-Dec 2022)
- Downloads the appropriate DIEM codebook from the DIEM Hub if not already available
- Applies coded-value labels field by field, using one dictionary sheet per variable
- Handles derived yes/no fields consistently
- Preserves the original file format (CSV, XLS, XLSX) in the output

The output file is saved in the same folder as the input file, with '_LABELLED'
added to the filename.
"""

import os
import pandas as pd

# ------------------------------------------------------------
# User inputs start
# ------------------------------------------------------------

# Set only the input microdata file. The script will automatically detect the infrastructure version,
# download the appropriate codebook from the DIEM Hub, apply labels where appropriate,
# and save a labelled version of the microdata file.

MICRODATA_FILE = r"C:\DIEM\DIEM_household_surveys_microdata.csv"  # Replace with the path, filename, and extension of the DIEM microdata file you want to label


# ------------------------------------------------------------
# User inputs end
# ------------------------------------------------------------



# Detect infrastructure by reading the header (field list)
# Detect infrastructure by reading only the header / columns (works for .csv, .xls, .xlsx)
_ext = os.path.splitext(MICRODATA_FILE)[1].lower().strip()

if _ext == ".csv":
    _micro_cols = list(pd.read_csv(MICRODATA_FILE, nrows=0).columns)
elif _ext in (".xlsx", ".xls"):
    # Read only header row (no data) -> fast
    _micro_cols = list(pd.read_excel(MICRODATA_FILE, nrows=0).columns)
else:
    raise ValueError(f"Unsupported microdata extension '{_ext}'. Use .xlsx, .xls, or .csv.")


infr = 2.0 if "qc_step0_date" in _micro_cols else 1.0 #field introduced after questinnaire revision in dec 2022

# Dictionary download URLs (ArcGIS items)
DICT_URL_1 = "https://hqfao.maps.arcgis.com/sharing/rest/content/items/e59d08ded7c1440587493bf65236cf44/data"
DICT_URL_2 = "https://hqfao.maps.arcgis.com/sharing/rest/content/items/41fa55934d2f462f86cd381ee8dc1fda/data"

# Save the dict in the same folder as the microdata CSV
_micro_folder = os.path.dirname(MICRODATA_FILE)
DICT_XLSX = os.path.join(_micro_folder, "DIEM_codebook_V10.xlsx" if infr == 1.0 else "DIEM_codebook_V20.xlsx")

# Download only if missing
if not os.path.exists(DICT_XLSX):
    import urllib.request
    dict_url = DICT_URL_1 if infr == 1.0 else DICT_URL_2
    print(f"Downloading dictionary (infr={infr}) to: {DICT_XLSX}")
    urllib.request.urlretrieve(dict_url, DICT_XLSX)
else:
    print(f"Dictionary already present (infr={infr}), skipping download: {DICT_XLSX}")


ADD_LABEL_COLUMNS = False  # if false, original coded value is REPLACED by the labelled value
LABEL_SUFFIX = "_label"
DERIVED_FIELDS_SHEET = "derived_fields"

MIN_OK_MATCH_RATE = 0.95


# ------------------------------------------------------------
# Helpers
# ------------------------------------------------------------
def _normalize_code(x):
    if pd.isna(x):
        return pd.NA
    s = str(x).strip()
    if s == "":
        return pd.NA
    try:
        f = float(s)
        if f.is_integer():
            return str(int(f))
        return str(f)
    except Exception:
        return s


def build_mapping_from_sheet(df_sheet):
    cols = {str(c).lower().strip(): c for c in df_sheet.columns}
    if "code" not in cols or "label" not in cols:
        return None

    tmp = df_sheet[[cols["code"], cols["label"]]].copy()
    tmp[cols["code"]] = tmp[cols["code"]].apply(_normalize_code)
    tmp[cols["label"]] = tmp[cols["label"]].astype(str).str.strip()
    tmp = tmp.dropna(subset=[cols["code"]])

    return dict(zip(tmp[cols["code"]], tmp[cols["label"]]))


def read_derived_fields(path_xlsx, sheet_name):
    """
    Your derived_fields format stores field names in the SECOND column
    (first column is an index).
    """
    try:
        df = pd.read_excel(path_xlsx, sheet_name=sheet_name)
    except Exception:
        return set()

    if df is None or df.empty or df.shape[1] < 2:
        return set()

    series = df.iloc[:, 1]
    return set(
        series.dropna()
        .astype(str)
        .map(str.strip)
        .loc[lambda s: s != ""]
        .tolist()
    )


def output_path_same_format(input_path: str, suffix: str) -> str:
    """
    Create output path with same extension as input.
    Example: data.xlsx -> data_LABELLED.xlsx
    """
    base, ext = os.path.splitext(input_path)
    return base + suffix + ext


def save_with_same_format(df: pd.DataFrame, input_path: str, out_path: str):
    """
    Save df using the same format as the input file extension.
    Supported: .xlsx, .xls, .csv
    """
    _, ext = os.path.splitext(input_path)
    ext = ext.lower().strip()

    if ext == ".csv":
        df.to_csv(out_path, index=False, encoding="utf-8")
    elif ext in (".xlsx", ".xls"):
        # Note: writing .xls may require additional dependencies and has row limits.
        # If .xls export fails, consider switching input/output to .xlsx.
        df.to_excel(out_path, index=False)
    else:
        raise ValueError(f"Unsupported input extension '{ext}'. Use .xlsx, .xls, or .csv.")


def read_microdata_any(input_path: str) -> pd.DataFrame:
    """
    Read microdata from .xlsx, .xls, or .csv.
    """
    _, ext = os.path.splitext(input_path)
    ext = ext.lower().strip()

    if ext == ".csv":
        return pd.read_csv(input_path, dtype=str, keep_default_na=True, na_values=["", "NA", "NaN", "nan"])
    if ext in (".xlsx", ".xls"):
        return pd.read_excel(input_path)

    raise ValueError(f"Unsupported microdata extension '{ext}'. Use .xlsx, .xls, or .csv.")


# ------------------------------------------------------------
# Main
# ------------------------------------------------------------
def main():
    print("Reading microdata:", MICRODATA_FILE)
    micro = read_microdata_any(MICRODATA_FILE)

    # Dictionary is expected to be an Excel workbook with one sheet per field
    dict_ext = os.path.splitext(DICT_XLSX)[1].lower().strip()
    if dict_ext not in (".xlsx", ".xls"):
        raise ValueError(
            f"DICT_XLSX must be an Excel file with sheets (got '{dict_ext}'). "
            f"Please point DICT_XLSX to the dictionary workbook (.xlsx/.xls)."
        )

    print("Reading coded-value dictionary:", DICT_XLSX)
    xls = pd.ExcelFile(DICT_XLSX)

    # Output path: keep the same format/extension as the input microdata
    OUT_PATH = output_path_same_format(MICRODATA_FILE, "_LABELLED")
    print("Output will be saved as:", OUT_PATH)

    # Read derived fields (0=No, 1=Yes)
    derived_fields = set()
    sheet_names_clean = [s.strip() for s in xls.sheet_names]
    if DERIVED_FIELDS_SHEET in sheet_names_clean:
        derived_fields = read_derived_fields(DICT_XLSX, DERIVED_FIELDS_SHEET)

    yes_no_mapping = {"0": "No", "1": "Yes"}

    # Collect label columns here and add them once (fast, avoids fragmentation)
    label_columns = {}

    # ------------------------------------------------------------------
    # Tracking / stats
    # ------------------------------------------------------------------
    labelled_fields = set()            # fields we actually processed (dedicated mapping or derived yes/no)
    skipped_fields = {}                # field -> reason (for candidates we expected to label but did not)
    missing_in_micro = []              # fields present in dict/derived list but not in microdata columns

    # Per-field detailed issues for labelled fields
    # field -> dict with counts and some sample unmapped codes
    per_field_details = {}

    # Global summary stats (global only)
    fields_labelled = 0
    fields_skipped_empty = 0
    fields_low_match = 0
    total_values = 0
    mapped_values = 0

    # Candidate fields from dedicated sheets
    dedicated_sheet_fields = []
    for s in xls.sheet_names:
        s_clean = s.strip()
        if s_clean == DERIVED_FIELDS_SHEET:
            continue
        dedicated_sheet_fields.append(s_clean)

    # 1) Dedicated dictionary sheets (one sheet per field)
    for sheet_name in xls.sheet_names:
        if sheet_name.strip() == DERIVED_FIELDS_SHEET:
            continue

        field = sheet_name.strip()

        if field not in micro.columns:
            missing_in_micro.append(field)
            continue

        mapping = build_mapping_from_sheet(pd.read_excel(DICT_XLSX, sheet_name=sheet_name))
        if mapping is None:
            skipped_fields[field] = "Dictionary sheet missing 'code' and/or 'label' columns"
            continue

        normalized = micro[field].apply(_normalize_code)
        non_null = normalized.dropna()

        # Skip fields with no values at all
        if non_null.empty:
            fields_skipped_empty += 1
            skipped_fields[field] = "All values are null/empty in microdata"
            continue

        labels = normalized.map(mapping)

        # Field-level match rate and unmapped info
        mapped_mask = labels.loc[non_null.index].notna()
        mapped_here = int(mapped_mask.sum())
        total_here = int(len(non_null))
        match_rate_here = mapped_here / total_here if total_here else 0.0
        unmapped_here = total_here - mapped_here

        # Collect top unmapped codes (only among non-null)
        unmapped_codes = normalized.loc[non_null.index][~mapped_mask]
        top_unmapped = (
            unmapped_codes.value_counts(dropna=True)
            .head(15)
            .to_dict()
        )

        per_field_details[field] = {
            "source": "dedicated_sheet",
            "non_null_values": total_here,
            "mapped_values": mapped_here,
            "unmapped_values": unmapped_here,
            "match_rate": match_rate_here,
            "top_unmapped_codes": top_unmapped,
        }

        if match_rate_here < MIN_OK_MATCH_RATE:
            fields_low_match += 1

        total_values += total_here
        mapped_values += mapped_here

        if ADD_LABEL_COLUMNS:
            label_columns[field + LABEL_SUFFIX] = labels
        else:
            micro[field] = labels

        labelled_fields.add(field)
        fields_labelled += 1

    # 2) Derived yes/no fields listed in derived_fields sheet
    for field in derived_fields:
        if field not in micro.columns:
            missing_in_micro.append(field)
            continue

        # Do not overwrite if already created by a dedicated mapping sheet
        if ADD_LABEL_COLUMNS:
            if (field + LABEL_SUFFIX) in label_columns:
                continue
        else:
            # If we already labelled it via dedicated mapping, skip
            if field in labelled_fields:
                continue

        normalized = micro[field].apply(_normalize_code)
        non_null = normalized.dropna()

        if non_null.empty:
            fields_skipped_empty += 1
            skipped_fields[field] = "All values are null/empty in microdata (derived yes/no)"
            continue

        labels = normalized.map(yes_no_mapping)

        mapped_mask = labels.loc[non_null.index].notna()
        mapped_here = int(mapped_mask.sum())
        total_here = int(len(non_null))
        match_rate_here = mapped_here / total_here if total_here else 0.0
        unmapped_here = total_here - mapped_here

        unmapped_codes = normalized.loc[non_null.index][~mapped_mask]
        top_unmapped = (
            unmapped_codes.value_counts(dropna=True)
            .head(15)
            .to_dict()
        )

        per_field_details[field] = {
            "source": "derived_yes_no",
            "non_null_values": total_here,
            "mapped_values": mapped_here,
            "unmapped_values": unmapped_here,
            "match_rate": match_rate_here,
            "top_unmapped_codes": top_unmapped,
        }

        if match_rate_here < MIN_OK_MATCH_RATE:
            fields_low_match += 1

        total_values += total_here
        mapped_values += mapped_here

        if ADD_LABEL_COLUMNS:
            label_columns[field + LABEL_SUFFIX] = labels
        else:
            micro[field] = labels

        labelled_fields.add(field)
        fields_labelled += 1


    # 3) Any remaining *_other fields: apply yes/no mapping (0=No, 1=Yes) if not already labelled
    other_fields = [c for c in micro.columns if c.lower().endswith("_other")]

    for field in other_fields:
        # Skip if already labelled via dedicated sheet or derived_fields
        if ADD_LABEL_COLUMNS:
            if (field + LABEL_SUFFIX) in label_columns:
                continue
        else:
            if field in labelled_fields:
                continue

        normalized = micro[field].apply(_normalize_code)
        labels = normalized.map(yes_no_mapping)

        if ADD_LABEL_COLUMNS:
            label_columns[field + LABEL_SUFFIX] = labels
        else:
            micro[field] = labels
            labelled_fields.add(field)


    # Add all label columns in one shot (fast)
    if ADD_LABEL_COLUMNS and label_columns:
        labels_df = pd.DataFrame(label_columns)
        micro_out = pd.concat([micro, labels_df], axis=1)
    else:
        micro_out = micro

    # ------------------------------------------------------------------
    # Final reporting
    # ------------------------------------------------------------------
    # Candidates we expected to label (present in microdata and present in dict sheets or derived list)
    candidate_fields = set([f for f in dedicated_sheet_fields if f in micro.columns]) | set(
        [f for f in derived_fields if f in micro.columns]
    ) | set(other_fields)

    not_labelled_candidates = sorted([f for f in candidate_fields if f not in labelled_fields])

    # Labelled fields with issues: any unmapped values > 0
    labelled_with_unmapped = []
    for f in sorted(labelled_fields):
        d = per_field_details.get(f, {})
        if d.get("unmapped_values", 0) > 0:
            labelled_with_unmapped.append(f)

    print("\n--- Labelling summary (global) ---")
    print(f"Candidate fields to label (present in micro + in dict/derived list): {len(candidate_fields)}")
    print(f"Fields labelled: {len(labelled_fields)}")
    print(f"Fields not labelled (among candidates): {len(not_labelled_candidates)}")
    print(f"Fields skipped (all null): {fields_skipped_empty}")
    print(f"Fields with low match rate (< {int(MIN_OK_MATCH_RATE*100)}%): {fields_low_match}")
    if total_values:
        print(f"Overall value match rate: {mapped_values / total_values:.1%}")

    if not_labelled_candidates:
        print("\n--- Fields not labelled (candidates) ---")
        for f in not_labelled_candidates:
            reason = skipped_fields.get(f, "Unknown reason (check dictionary sheet format and data)")
            print(f"- {f}: {reason}")

    # Optional extra: show fields present in dict/derived but missing in micro
    missing_in_micro_unique = sorted(set(missing_in_micro))

    # Flag issues where codes could not be mapped to labels
    if labelled_with_unmapped:
        print("\n--- Labelled fields with unmapped codes (issues to review) ---")
        for f in labelled_with_unmapped:
            d = per_field_details.get(f, {})
            print(
                f"- {f} | source={d.get('source')} | "
                f"non-null={d.get('non_null_values')} | "
                f"mapped={d.get('mapped_values')} | "
                f"unmapped={d.get('unmapped_values')} | "
                f"match={d.get('match_rate', 0):.1%}"
            )
            top_unmapped = d.get("top_unmapped_codes", {}) or {}
            if top_unmapped:
                # Print a compact list of most frequent unmapped codes
                preview = ", ".join([f"{k}({v})" for k, v in top_unmapped.items()])
                print(f"  top unmapped codes: {preview}")

    # Save using same format as input
    print("\nSaving output to:", OUT_PATH)
    save_with_same_format(micro_out, MICRODATA_FILE, OUT_PATH)

    print("\nDone.")


if __name__ == "__main__":
    main()
