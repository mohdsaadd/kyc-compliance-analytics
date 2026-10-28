"""SQLite loading and named-query execution."""
import re
import sqlite3
from pathlib import Path

import pandas as pd

QUERY_FILE = Path(__file__).resolve().parent.parent / "sql" / "queries.sql"

CREATE_TABLE = """
CREATE TABLE cases (
    case_id        TEXT PRIMARY KEY,
    customer_name  TEXT NOT NULL,
    document_type  TEXT NOT NULL,
    region         TEXT NOT NULL,
    risk_level     TEXT NOT NULL CHECK (risk_level IN ('Low', 'Medium', 'High')),
    status         TEXT NOT NULL CHECK (status IN ('Pending', 'Approved', 'Rejected')),
    analyst        TEXT NOT NULL,
    submitted_date TEXT NOT NULL,
    completed_date TEXT
);
CREATE INDEX idx_cases_status ON cases(status);
"""


def load_queries(path=QUERY_FILE):
    """Parse queries.sql into {name: sql}."""
    text = Path(path).read_text()
    parts = re.split(r"^-- name:\s*(\w+)\s*$", text, flags=re.MULTILINE)
    # parts = [preamble, name1, sql1, name2, sql2, ...]
    return {parts[i]: parts[i + 1].strip() for i in range(1, len(parts), 2)}


def load_cases(df, db_path=":memory:"):
    """Create the schema and insert cleaned cases. Returns an open connection."""
    conn = sqlite3.connect(db_path)
    conn.executescript("DROP TABLE IF EXISTS cases;" + CREATE_TABLE)
    out = df.copy()
    for col in ("submitted_date", "completed_date"):
        out[col] = out[col].dt.strftime("%Y-%m-%d")
    out = out.astype(object).where(out.notna(), None)
    out.to_sql("cases", conn, if_exists="append", index=False)
    conn.commit()
    return conn


def run_query(conn, name, **params):
    queries = load_queries()
    if name not in queries:
        raise KeyError(f"Unknown query '{name}'. Available: {sorted(queries)}")
    return pd.read_sql_query(queries[name], conn, params=params)
