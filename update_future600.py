import argparse
import re
from pathlib import Path

import pandas as pd


def _normalize_col(name: str) -> str:
    return re.sub(r"[^a-z0-9]", "", str(name).strip().lower())


def _find_column(columns, aliases, required=True):
    normalized = {_normalize_col(c): c for c in columns}
    for alias in aliases:
        key = _normalize_col(alias)
        if key in normalized:
            return normalized[key]
    if required:
        raise ValueError(f"Missing required column. Expected one of: {aliases}")
    return None


def _extract_numeric_part(value):
    if pd.isna(value):
        return pd.NA
    match = re.search(r"(\d+)", str(value))
    return match.group(1) if match else pd.NA


def _clean_value(value):
    if pd.isna(value):
        return pd.NA
    text = str(value).strip()
    return pd.NA if text == "" else value


def _first_non_null(series: pd.Series):
    cleaned = series.map(_clean_value).dropna()
    return cleaned.iloc[0] if not cleaned.empty else pd.NA


def _build_position_key(df: pd.DataFrame, requisition_col: str, position_col: str) -> pd.Series:
    req = df[requisition_col].astype(str).str.strip()
    pos_num = df[position_col].map(_extract_numeric_part)
    return req + "|" + pos_num.astype(str)


def update_future600(
    master_file: str,
    future650_file: str,
    future654_file: str,
    future663_file: str,
    output_file: str = "Future600_updated.xlsx",
    min_completed_date: str | None = None,
):
    master_path = Path(master_file)
    output_path = Path(output_file)
    if not output_path.is_absolute():
        output_path = master_path.parent / output_path

    master_df = pd.read_excel(master_path)
    master_req_col = _find_column(master_df.columns, ["Job Requisition ID", "Job_Requisition_ID", "job_requisition_id"])
    master_pos_col = _find_column(master_df.columns, ["Position", "All Positions", "All Position"])
    master_df["_position_key"] = _build_position_key(master_df, master_req_col, master_pos_col)

    if min_completed_date:
        status_col = _find_column(master_df.columns, ["Job Requisition Status", "job_requisition_status"], required=False)
        completed_col = _find_column(
            master_df.columns,
            ["Job Requisition Completed", "Job Requisition Complete", "job_requisition_completed"],
            required=False,
        )

        if status_col and completed_col:
            cutoff = pd.to_datetime(min_completed_date, errors="coerce")
            if pd.isna(cutoff):
                raise ValueError(
                    f"Invalid min_completed_date '{min_completed_date}'. Use YYYY-MM-DD format."
                )

            status_text = master_df[status_col].astype(str).str.lower()
            is_closed_filled = status_text.str.contains("closed", na=False) | status_text.str.contains(
                "filled", na=False
            )
            completed_dates = pd.to_datetime(master_df[completed_col], errors="coerce")
            should_remove = is_closed_filled & completed_dates.notna() & (completed_dates < cutoff)
            master_df = master_df.loc[~should_remove].copy()

    candidate_sources = []
    source_files = [future654_file, future650_file]
    for file_path in source_files:
        source_df = pd.read_excel(file_path)
        req_col = _find_column(source_df.columns, ["Job Requisition ID", "Job_Requisition_ID", "job_requisition_id"])
        pos_col = _find_column(source_df.columns, ["All Positions", "All Position", "Position"])
        name_col = _find_column(source_df.columns, ["Candidate Name", "Full Name", "Candidate_Name"], required=False)
        offer_col = _find_column(source_df.columns, ["Offer Date", "offer_date", "Offer_Date"], required=False)
        start_col = _find_column(
            source_df.columns,
            ["Candidate Start Date", "Candidate_Start_Date", "Start Date", "start_date"],
            required=False,
        )

        source_df["_position_key"] = _build_position_key(source_df, req_col, pos_col)
        selected = pd.DataFrame({"_position_key": source_df["_position_key"]})
        selected["Candidate Name"] = source_df[name_col] if name_col else pd.NA
        selected["WD Offer Created Date"] = source_df[offer_col] if offer_col else pd.NA
        selected["Candidate Start Date"] = source_df[start_col] if start_col else pd.NA
        candidate_sources.append(selected)

    candidate_union = pd.concat(candidate_sources, ignore_index=True)
    candidate_by_key = (
        candidate_union.groupby("_position_key", dropna=False, as_index=False)
        .agg(
            {
                "Candidate Name": _first_non_null,
                "WD Offer Created Date": _first_non_null,
                "Candidate Start Date": _first_non_null,
            }
        )
    )

    merged = master_df.merge(candidate_by_key, on="_position_key", how="left")

    f663_df = pd.read_excel(future663_file)
    f663_req_col = _find_column(f663_df.columns, ["Job_Requisition_ID", "Job Requisition ID", "job_requisition_id"])
    sent_col = _find_column(
        f663_df.columns,
        ["Offer_letter_sent_date", "Offer Letter Sent Date", "offer_letter_sent_date"],
    )
    signed_col = _find_column(
        f663_df.columns,
        [
            "Offer Letter Signed / Declined Date",
            "Offer_Letter_Signed_Declined_Date",
            "offer_letter_signed_declined_date",
        ],
    )

    req_level_663 = (
        f663_df[[f663_req_col, sent_col, signed_col]]
        .groupby(f663_req_col, as_index=False)
        .agg({sent_col: _first_non_null, signed_col: _first_non_null})
        .rename(
            columns={
                f663_req_col: "_req_join",
                sent_col: "_offer_letter_sent_date_src",
                signed_col: "_offer_letter_signed_declined_src",
            }
        )
    )

    merged["_req_join"] = merged[master_req_col].astype(str).str.strip()
    merged = merged.merge(req_level_663, on="_req_join", how="left")

    merged["Offer_letter_sent_date"] = merged["_offer_letter_sent_date_src"]
    merged["Offer Letter Signed / Declined Date"] = merged["_offer_letter_signed_declined_src"]

    drop_cols = [
        "_position_key",
        "_req_join",
        "_offer_letter_sent_date_src",
        "_offer_letter_signed_declined_src",
    ]
    merged = merged.drop(columns=[c for c in drop_cols if c in merged.columns])

    merged.to_excel(output_path, index=False)
    return output_path


def main():
    parser = argparse.ArgumentParser(
        description="Update Future600 with candidate and offer details from Future650/Future654/Future663."
    )
    parser.add_argument("--master", required=True, help="Path to Future600.xlsx")
    parser.add_argument("--future650", required=True, help="Path to Future650.xlsx")
    parser.add_argument("--future654", required=True, help="Path to Future654.xlsx")
    parser.add_argument("--future663", required=True, help="Path to Future663.xlsx")
    parser.add_argument(
        "--output",
        default="Future600_updated.xlsx",
        help="Output file name/path (default: Future600_updated.xlsx in master file folder)",
    )
    parser.add_argument(
        "--min-completed-date",
        default=None,
        help=(
            "Optional filter cutoff (YYYY-MM-DD). Removes records where Job Requisition Status is closed+filled "
            "and Job Requisition Completed is older than this date."
        ),
    )
    args = parser.parse_args()

    output = update_future600(
        master_file=args.master,
        future650_file=args.future650,
        future654_file=args.future654,
        future663_file=args.future663,
        output_file=args.output,
        min_completed_date=args.min_completed_date,
    )
    print(f"Updated file written to: {output}")


if __name__ == "__main__":
    main()
