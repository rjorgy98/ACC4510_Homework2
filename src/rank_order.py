from __future__ import annotations

import logging
import os
import re
from pathlib import Path
from typing import Iterable, Optional

import matplotlib.pyplot as plt
import pandas as pd


logging.basicConfig(level=logging.INFO, format="[%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)

REPO_ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = REPO_ROOT / "data"
OUTPUTS_DIR = REPO_ROOT / "outputs"

FORCED_SCALE_LABELS = {
    "most": (1, 3),
    "neutral": (4, 5),
    "least": (6, 8),
}
LIKERT_LABELS = {
    "most": (4, 5),
    "neutral": (3, 3),
    "least": (1, 2),
}

TEXT_BUCKET_MAP = {
    "most beneficial": "most",
    "most": "most",
    "beneficial": "most",
    "neutral": "neutral",
    "neither": "neutral",
    "least beneficial": "least",
    "least": "least",
}

BLOCK_COLUMN_HINTS = (
    "rank",
    "beneficial",
    "course",
    "program",
    "acc",
    "q84",
    "preparation",
)


def snake_case(value: str) -> str:
    value = str(value)
    value = value.strip().lower()
    value = re.sub(r"[^a-z0-9]+", "_", value)
    value = re.sub(r"_+", "_", value).strip("_")
    return value or "unnamed"


def infer_year_from_name(path: Path) -> Optional[str]:
    match = re.search(r"(20\d{2})", path.name)
    return match.group(1) if match else None


def pick_data_file(data_dir: Path) -> Path:
    env_file = os.getenv("DATA_FILE", "").strip()
    if env_file:
        candidate = Path(env_file)
        if not candidate.is_absolute():
            candidate = data_dir / candidate
        if candidate.exists():
            logger.info("Using DATA_FILE override: %s", candidate)
            return candidate
        raise FileNotFoundError(f"DATA_FILE was set but not found: {candidate}")

    files = sorted(data_dir.glob("*.xlsx"))
    if not files:
        raise FileNotFoundError(
            f"No .xlsx files found in {data_dir}. Add one year of data in /data."
        )
    logger.info("No DATA_FILE override provided. Using first file: %s", files[0])
    return files[0]


def categorize_numeric(value: float, scale_max: int) -> Optional[str]:
    if pd.isna(value):
        return None
    if scale_max >= 8:
        bins = FORCED_SCALE_LABELS
    elif scale_max == 5:
        bins = LIKERT_LABELS
    else:
        return None

    for label, (low, high) in bins.items():
        if low <= value <= high:
            return label
    return None


def categorize_text(value: object) -> Optional[str]:
    if pd.isna(value):
        return None
    text = str(value).strip().lower()
    if not text:
        return None
    if "did not take" in text or "n/a" == text:
        return None
    for key, bucket in TEXT_BUCKET_MAP.items():
        if key in text:
            return bucket
    return None


def looks_like_course_col(name: str) -> bool:
    lowered = name.lower()
    return any(hint in lowered for hint in BLOCK_COLUMN_HINTS)


def standardize_df(df: pd.DataFrame) -> pd.DataFrame:
    renamed = {}
    used = set()
    for c in df.columns:
        base = snake_case(c)
        name = base
        i = 2
        while name in used:
            name = f"{base}_{i}"
            i += 1
        used.add(name)
        renamed[c] = name
    return df.rename(columns=renamed)


def find_rank_records(df: pd.DataFrame) -> pd.DataFrame:
    records = []

    # Case 1: wide forced-rank/rating columns where each course is its own column
    numeric_candidates = []
    for col in df.columns:
        series = pd.to_numeric(df[col], errors="coerce")
        non_na = series.dropna()
        if non_na.empty:
            continue
        # rank/rating data should have many small integers
        if non_na.between(1, 8).mean() >= 0.7 and non_na.nunique() >= 3:
            numeric_candidates.append((col, series, int(non_na.max())))

    filtered_candidates = [
        (col, series, mx)
        for col, series, mx in numeric_candidates
        if looks_like_course_col(col)
    ]
    if filtered_candidates:
        numeric_candidates = filtered_candidates

    if numeric_candidates:
        logger.info("Detected %s numeric ranking columns.", len(numeric_candidates))
        for col, series, scale_max in numeric_candidates:
            for value in series:
                bucket = categorize_numeric(value, scale_max)
                if bucket is None:
                    continue
                records.append({"course": col, "bucket": bucket})
        if records:
            return pd.DataFrame(records)

    # Case 2: long form with a course column and rank/value column
    course_cols = [c for c in df.columns if "course" in c or "program" in c]
    rank_cols = [c for c in df.columns if "rank" in c or "rating" in c or c.startswith("q")]
    for ccol in course_cols:
        course_series = df[ccol].astype("string")
        for rcol in rank_cols:
            rank_series = pd.to_numeric(df[rcol], errors="coerce")
            valid = rank_series.dropna()
            if valid.empty:
                continue
            if valid.between(1, 8).mean() < 0.7:
                continue
            scale_max = int(valid.max())
            logger.info("Detected long-form ranking columns: %s + %s", ccol, rcol)
            tmp = pd.DataFrame({"course": course_series, "value": rank_series})
            tmp = tmp.dropna(subset=["course", "value"])
            tmp["course"] = tmp["course"].str.strip()
            tmp = tmp[tmp["course"] != ""]
            tmp["bucket"] = tmp["value"].apply(lambda v: categorize_numeric(v, scale_max))
            tmp = tmp.dropna(subset=["bucket"])
            if not tmp.empty:
                return tmp[["course", "bucket"]]

    # Case 3: bucketed text responses
    text_bucket_cols = [
        c
        for c in df.columns
        if any(k in c for k in ("beneficial", "preference", "rank", "rating", "course"))
    ]
    for col in text_bucket_cols:
        mapped = df[col].map(categorize_text)
        if mapped.notna().sum() < 5:
            continue
        if "course" in df.columns:
            tmp = pd.DataFrame({"course": df["course"], "bucket": mapped})
            tmp = tmp.dropna(subset=["course", "bucket"])
            tmp["course"] = tmp["course"].astype(str).str.strip()
            tmp = tmp[tmp["course"] != ""]
            if not tmp.empty:
                logger.info("Detected bucketed text with explicit course column.")
                return tmp[["course", "bucket"]]

    raise ValueError(
        "Could not detect ranking fields. Expected either: (1) per-course numeric ranking "
        "columns, (2) a course column plus numeric rank/rating column, or (3) bucketed "
        "Most/Neutral/Least responses."
    )


def summarize_nas(rank_records: pd.DataFrame) -> pd.DataFrame:
    grouped = rank_records.groupby("course", dropna=True)
    summary = grouped["bucket"].value_counts().unstack(fill_value=0)
    for col in ("most", "neutral", "least"):
        if col not in summary.columns:
            summary[col] = 0
    summary = summary[["most", "neutral", "least"]]
    summary = summary.rename(
        columns={"most": "n_most", "neutral": "n_neutral", "least": "n_least"}
    )
    summary["n_total"] = summary[["n_most", "n_neutral", "n_least"]].sum(axis=1)
    summary = summary[summary["n_total"] > 0]
    summary["pct_most"] = summary["n_most"] / summary["n_total"] * 100
    summary["pct_least"] = summary["n_least"] / summary["n_total"] * 100
    summary["nas"] = summary["pct_most"] - summary["pct_least"]

    summary = summary.sort_values(
        by=["nas", "pct_most", "n_total"], ascending=[False, False, False]
    )
    summary = summary.reset_index()
    summary.insert(0, "rank", range(1, len(summary) + 1))
    return summary[
        [
            "rank",
            "course",
            "n_total",
            "n_most",
            "n_neutral",
            "n_least",
            "pct_most",
            "pct_least",
            "nas",
        ]
    ]


def plot_rank_order(summary: pd.DataFrame, output_path: Path, year: Optional[str]) -> None:
    ordered = summary.sort_values("nas", ascending=True)

    plt.figure(figsize=(10, max(4, len(ordered) * 0.45)))
    plt.barh(ordered["course"], ordered["nas"], color="#1f77b4")
    plt.axvline(0, color="black", linestyle="--", linewidth=1)
    title_year = f" ({year})" if year else ""
    plt.title(f"Course Rank Order by Net Approval Score{title_year}")
    plt.xlabel("Net Approval Score (pct_most - pct_least)")
    plt.ylabel("Course / Program")
    plt.tight_layout()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    plt.savefig(output_path, dpi=200)
    plt.close()


def load_all_sheets(path: Path) -> pd.DataFrame:
    xls = pd.ExcelFile(path)
    frames = []
    for sheet in xls.sheet_names:
        df = pd.read_excel(path, sheet_name=sheet)
        if df.empty:
            continue
        df = standardize_df(df)
        frames.append(df)
    if not frames:
        raise ValueError(f"No non-empty sheets found in workbook: {path}")
    return pd.concat(frames, ignore_index=True, sort=False)


def main() -> None:
    data_file = pick_data_file(DATA_DIR)
    OUTPUTS_DIR.mkdir(parents=True, exist_ok=True)

    df = load_all_sheets(data_file)
    rank_records = find_rank_records(df)
    summary = summarize_nas(rank_records)

    csv_path = OUTPUTS_DIR / "rank_order.csv"
    png_path = OUTPUTS_DIR / "rank_order.png"

    summary.to_csv(csv_path, index=False)
    year = infer_year_from_name(data_file)
    plot_rank_order(summary, png_path, year)

    logger.info("Wrote ranking table: %s", csv_path)
    logger.info("Wrote figure: %s", png_path)


if __name__ == "__main__":
    main()
