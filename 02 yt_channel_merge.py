import argparse
from pathlib import Path
import pandas as pd


# =========================
# CONFIGURATION
# =========================
BASE_DIR = Path(__file__).resolve().parent
INPUT_DIR = BASE_DIR / "outputs" / "youtube"
OUTPUT_FILE = BASE_DIR / "outputs" / "youtube_merged.xlsx"


REQUIRED_COLUMNS = [
    "id",
    "title",
    "channel",
    "upload_date",
    "duration",
    "duration_seconds",
    "view_count",
    "comment_count",
    "like_count",
    "categories",
    "tags",
    "description",
    "availability",
    "webpage_url",
    "url",
]


FINAL_COLUMNS = [
    "title",
    "channel",
    "upload_date",
    "duration_seconds",
    "view_count",
    "comment_count",
    "like_count",
    "categories",
    "tags",
    "description",
    "availability",
    "webpage_url",
    "source_file",
    "section",
]


TEXT_COLUMNS = [
    "id",
    "title",
    "channel",
    "upload_date",
    "categories",
    "tags",
    "description",
    "availability",
    "webpage_url",
    "url",
]


NUMERIC_COLUMNS = [
    "duration",
    "duration_seconds",
    "view_count",
    "comment_count",
    "like_count",
]


SECTION_SUFFIXES = {
    "__videos.xlsx": "videos",
    "__streams.xlsx": "streams",
    "__shorts.xlsx": "shorts",
    "__videos.xls": "videos",
    "__streams.xls": "streams",
    "__shorts.xls": "shorts",
}


# =========================
# HELPERS
# =========================
def is_empty_series(series: pd.Series) -> pd.Series:
    """
    Detects empty values robustly:
    - NaN
    - None
    - <NA>
    - empty strings
    - string versions of null values
    """

    return (
        series.isna()
        | series.astype(str).str.strip().isin(["", "nan", "None", "<NA>", "NaT"])
    )


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Cleans column names.
    """

    df.columns = [str(col).strip() for col in df.columns]
    return df


def ensure_required_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Adds missing required columns.
    Important: use pd.NA instead of None.
    """

    for col in REQUIRED_COLUMNS:
        if col not in df.columns:
            df[col] = pd.NA

    return df


def force_text_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Forces text-like columns to object dtype.

    This prevents errors like:
    Invalid value '<StringArray> ...' for dtype 'float64'
    """

    for col in TEXT_COLUMNS:
        if col in df.columns:
            df[col] = df[col].astype("object")

    return df


def repair_duration(df: pd.DataFrame) -> pd.DataFrame:
    """
    Repairs duration_seconds using duration when needed.
    Converts duration fields to numeric values safely.
    """

    if "duration_seconds" not in df.columns and "duration" in df.columns:
        df["duration_seconds"] = df["duration"]

    if "duration_seconds" in df.columns and "duration" not in df.columns:
        df["duration"] = df["duration_seconds"]

    df["duration_seconds"] = pd.to_numeric(
        df["duration_seconds"],
        errors="coerce"
    )

    df["duration"] = pd.to_numeric(
        df["duration"],
        errors="coerce"
    )

    df["duration_seconds"] = df["duration_seconds"].fillna(df["duration"])

    return df


def repair_webpage_url(df: pd.DataFrame) -> pd.DataFrame:
    """
    Repairs webpage_url.

    Priority:
    1. Existing webpage_url
    2. url
    3. Constructed URL from video id
    """

    for col in ["webpage_url", "url", "id"]:
        if col in df.columns:
            df[col] = df[col].astype("object")

    missing_webpage_url = is_empty_series(df["webpage_url"])
    has_url = ~is_empty_series(df["url"])

    df.loc[missing_webpage_url & has_url, "webpage_url"] = (
        df.loc[missing_webpage_url & has_url, "url"]
        .astype(str)
        .str.strip()
    )

    missing_webpage_url = is_empty_series(df["webpage_url"])
    has_id = ~is_empty_series(df["id"])

    df.loc[missing_webpage_url & has_id, "webpage_url"] = (
        "https://www.youtube.com/watch?v="
        + df.loc[missing_webpage_url & has_id, "id"]
        .astype(str)
        .str.strip()
    )

    return df


def clean_text_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Cleans text columns after import.
    Keeps real missing values as empty strings only at the final stage.
    """

    for col in TEXT_COLUMNS:
        if col in df.columns:
            df[col] = df[col].astype("object")

            mask_empty = is_empty_series(df[col])
            df.loc[mask_empty, col] = pd.NA

            df.loc[~mask_empty, col] = (
                df.loc[~mask_empty, col]
                .astype(str)
                .str.strip()
            )

    return df


def clean_numeric_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Converts numeric columns safely.
    """

    for col in NUMERIC_COLUMNS:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    return df


def extract_section_from_filename(filename: str) -> str:
    """
    Extracts section from filename:
    - channel__videos.xlsx  -> videos
    - channel__streams.xlsx -> streams
    - channel__shorts.xlsx  -> shorts

    Returns empty string if no known section suffix is found.
    """

    name = str(filename).strip()

    for suffix, section in SECTION_SUFFIXES.items():
        if name.endswith(suffix):
            return section

    return ""


def clean_source_file_name(filename: str) -> str:
    """
    Removes section suffixes from source file names:
    - __videos.xlsx
    - __streams.xlsx
    - __shorts.xlsx
    - __videos.xls
    - __streams.xls
    - __shorts.xls

    If no section suffix is found, removes only the Excel extension.
    """

    name = str(filename).strip()

    for suffix in SECTION_SUFFIXES:
        if name.endswith(suffix):
            return name[: -len(suffix)]

    if name.endswith(".xlsx"):
        return name[:-5]

    if name.endswith(".xls"):
        return name[:-4]

    return name


def load_excel_file(path: Path) -> pd.DataFrame:
    """
    Loads and repairs one Excel file.
    """

    df = pd.read_excel(
        path,
        dtype={
            "id": "object",
            "url": "object",
            "webpage_url": "object",
            "title": "object",
            "channel": "object",
            "upload_date": "object",
            "categories": "object",
            "tags": "object",
            "description": "object",
            "availability": "object",
        }
    )

    df = normalize_columns(df)
    # API export field names differ from legacy yt-dlp names.
    aliases = {"video_id": "id", "video_title": "title", "video_url": "webpage_url", "published_at": "upload_date"}
    df = df.rename(columns={old: new for old, new in aliases.items() if new not in df.columns})
    df = ensure_required_columns(df)

    df = force_text_columns(df)
    df = clean_text_columns(df)

    df = repair_duration(df)
    df = repair_webpage_url(df)

    df = clean_numeric_columns(df)

    df = df[REQUIRED_COLUMNS].copy()

    df = df.drop(columns=["duration", "url", "id"])

    df["source_file"] = clean_source_file_name(path.name)
    df["section"] = extract_section_from_filename(path.name)

    df = df[FINAL_COLUMNS]

    return df


def get_excel_files(input_dir: Path) -> list[Path]:
    """
    Finds Excel files and ignores temporary Excel lock files.
    Also ignores the final merged output file, if it already exists.
    """

    files = sorted(
        list(input_dir.glob("*.xlsx")) +
        list(input_dir.glob("*.xls"))
    )

    files = [
        file for file in files
        if any(file.name.endswith(suffix) for suffix in SECTION_SUFFIXES)
        and not file.name.startswith("~$")
        and file.resolve() != OUTPUT_FILE.resolve()
        and file.name != "youtube_merged_failed_files.xlsx"
    ]

    return files


def save_result(result: pd.DataFrame, output_file: Path) -> None:
    """
    Saves merged dataframe to Excel.
    Disables automatic URL conversion to avoid Excel hyperlink issues.
    """

    output_file.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(
        output_file,
        engine="xlsxwriter",
        engine_kwargs={"options": {"strings_to_urls": False, "strings_to_formulas": False}},
    ) as writer:
        result.to_excel(writer, index=False, sheet_name="youtube_data")

        worksheet = writer.sheets["youtube_data"]

        # Optional: make columns wider
        worksheet.set_column("A:A", 50)   # title
        worksheet.set_column("B:B", 30)   # channel
        worksheet.set_column("C:C", 15)   # upload_date
        worksheet.set_column("D:G", 15)   # numeric fields
        worksheet.set_column("H:I", 35)   # categories, tags
        worksheet.set_column("J:J", 70)   # description
        worksheet.set_column("K:K", 18)   # availability
        worksheet.set_column("L:L", 55)   # webpage_url
        worksheet.set_column("M:M", 35)   # source_file
        worksheet.set_column("N:N", 15)   # section


# =========================
# MAIN
# =========================
def main():
    global INPUT_DIR, OUTPUT_FILE
    parser = argparse.ArgumentParser(description="Merge video/stream/short exports into a flat legacy table. Use build_research_workbook.py for the AI notebooks.")
    parser.add_argument("--input-dir", type=Path, default=INPUT_DIR)
    parser.add_argument("--output", type=Path, default=OUTPUT_FILE)
    args = parser.parse_args()
    INPUT_DIR, OUTPUT_FILE = args.input_dir, args.output
    if not INPUT_DIR.exists():
        print(f"Input folder does not exist: {INPUT_DIR}")
        return 2

    files = get_excel_files(INPUT_DIR)

    if not files:
        print("No Excel files found.")
        return 2

    frames = []
    failed_files = []

    print(f"Found Excel files: {len(files)}")
    print(f"Input folder: {INPUT_DIR}")
    print("-" * 60)

    for file in files:
        try:
            print(f"Importing: {file.name}")

            df = load_excel_file(file)

            if df.empty:
                print(f"Skipped empty file: {file.name}")
                continue

            frames.append(df)

            print(f"Imported rows: {len(df)}")

        except Exception as e:
            print(f"Error in {file.name}: {e}")

            failed_files.append({
                "source_file": clean_source_file_name(file.name),
                "section": extract_section_from_filename(file.name),
                "original_file": file.name,
                "error": str(e)
            })

    print("-" * 60)

    if not frames:
        print("No data imported.")
        return 2

    result = pd.concat(frames, ignore_index=True)

    before_dedup = len(result)
    result = result.drop_duplicates()
    after_dedup = len(result)

    save_result(result, OUTPUT_FILE)

    print("Done.")
    print(f"Successfully imported files: {len(frames)}")
    print(f"Rows before deduplication: {before_dedup}")
    print(f"Rows after deduplication: {after_dedup}")
    print(f"Saved to: {OUTPUT_FILE}")

    if failed_files:
        failed_df = pd.DataFrame(failed_files)
        failed_output = OUTPUT_FILE.with_name("youtube_merged_failed_files.xlsx")

        with pd.ExcelWriter(
            failed_output,
            engine="xlsxwriter",
            engine_kwargs={"options": {"strings_to_urls": False, "strings_to_formulas": False}},
        ) as writer:
            failed_df.to_excel(writer, index=False, sheet_name="failed_files")

        print(f"Failed files: {len(failed_files)}")
        print(f"Failure log saved to: {failed_output}")


if __name__ == "__main__":
    main()