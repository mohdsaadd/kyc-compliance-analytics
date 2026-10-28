"""Data cleaning: turn the raw export into analysis-ready data plus an issues log."""
import pandas as pd

STATUS_MAP = {
    "pending": "Pending", "in progress": "Pending",
    "approved": "Approved", "done": "Approved",
    "rejected": "Rejected", "declined": "Rejected",
}
REQUIRED_COLUMNS = ["case_id", "customer_name", "document_type", "region",
                    "risk_level", "status", "analyst", "submitted_date", "completed_date"]


def clean_cases(raw: pd.DataFrame):
    """Return (clean_df, issues_df). issues_df lists every row removed and why."""
    missing = set(REQUIRED_COLUMNS) - set(raw.columns)
    if missing:
        raise ValueError(f"Raw data is missing columns: {sorted(missing)}")

    df = raw[REQUIRED_COLUMNS].copy()
    issues = []

    def drop(mask, reason):
        for case_id in df.loc[mask, "case_id"]:
            issues.append({"case_id": case_id, "reason": reason})
        return df[~mask]

    # 1. Standardise text: trim whitespace, consistent casing
    for col in ["case_id", "customer_name", "document_type", "region",
                "risk_level", "status", "analyst"]:
        df[col] = df[col].astype("string").str.strip()
    df["customer_name"] = df["customer_name"].str.title()
    df["region"] = df["region"].str.title()
    df["risk_level"] = df["risk_level"].str.title()

    # 2. Normalise status wording (Done -> Approved, In Progress -> Pending, ...)
    df["status"] = df["status"].str.lower().map(STATUS_MAP)
    df = drop(df["status"].isna(), "unrecognised status")

    # 3. Remove duplicate case IDs (keep the first occurrence)
    df = drop(df.duplicated(subset="case_id", keep="first"), "duplicate case_id")

    # 4. Parse dates; invalid values become NaT
    df["submitted_date"] = pd.to_datetime(df["submitted_date"], errors="coerce", format="mixed")
    df["completed_date"] = pd.to_datetime(df["completed_date"], errors="coerce", format="mixed")
    df = drop(df["submitted_date"].isna(), "invalid or missing submitted_date")

    # 5. Logical checks
    df = drop(df["completed_date"] < df["submitted_date"],
              "completed_date earlier than submitted_date")
    df = drop((df["status"] == "Pending") & df["completed_date"].notna(),
              "pending case has a completed_date")
    df = drop((df["status"] != "Pending") & df["completed_date"].isna(),
              "closed case has no completed_date")

    df = df.reset_index(drop=True)
    for col in df.select_dtypes("string").columns:
        df[col] = df[col].astype(object)
    issues_df = pd.DataFrame(issues, columns=["case_id", "reason"])
    return df, issues_df
