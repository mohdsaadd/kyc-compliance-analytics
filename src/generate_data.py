"""Generate a realistic, deliberately messy KYC case export (as an Excel file).

Real exports are never clean, so the data contains the problems the cleaning
step has to handle: duplicates, inconsistent status wording, stray whitespace,
mixed casing, invalid dates and completion dates earlier than submission.
"""
import random
from datetime import date, timedelta

import pandas as pd

FIRST = ["Aarav", "Vivaan", "Aditya", "Sana", "Ayesha", "Riya", "Kabir", "Ishaan",
         "Meera", "Zara", "Rohan", "Fatima", "Neha", "Arjun", "Tanvi", "Imran"]
LAST = ["Sharma", "Khan", "Verma", "Gupta", "Singh", "Ansari", "Mehta", "Reddy",
        "Iyer", "Malik", "Joshi", "Nair"]
DOC_TYPES = ["Passport", "Aadhaar", "PAN Card", "Driving Licence", "Address Proof"]
REGIONS = ["North", "South", "East", "West"]
RISK = ["Low", "Medium", "High"]
ANALYSTS = ["Analyst A", "Analyst B", "Analyst C", "Analyst D"]

STATUS_VARIANTS = {
    "Approved": ["Approved", "approved", "APPROVED", "Done", " Approved "],
    "Rejected": ["Rejected", "rejected", "Declined", "REJECTED"],
    "Pending": ["Pending", "pending", "In Progress", "PENDING ", "pending"],
}


def generate_cases(n_rows=500, as_of=None, seed=42):
    rng = random.Random(seed)
    as_of = as_of or date.today()
    rows = []
    for i in range(1, n_rows + 1):
        submitted = as_of - timedelta(days=rng.randint(0, 120))
        tat = rng.randint(1, 12)
        completed = None
        if submitted + timedelta(days=tat) <= as_of and rng.random() < 0.9:
            status = "Approved" if rng.random() < 0.8 else "Rejected"
            completed = submitted + timedelta(days=tat)
        else:
            status = "Pending"
        name = f"{rng.choice(FIRST)} {rng.choice(LAST)}"
        rows.append({
            "case_id": f"KYC-{i:05d}",
            "customer_name": rng.choice([name, f"  {name}", name.upper(), name.lower()]),
            "document_type": rng.choice(DOC_TYPES),
            "region": rng.choice(REGIONS + [r.lower() for r in REGIONS] + ["North "]),
            "risk_level": rng.choices(RISK, weights=[60, 30, 10])[0],
            "status": rng.choice(STATUS_VARIANTS[status]),
            "analyst": rng.choice(ANALYSTS),
            "submitted_date": submitted,
            "completed_date": completed,
        })

    df = pd.DataFrame(rows)
    # --- inject data-quality problems ---
    dupes = df.sample(frac=0.03, random_state=seed)
    df = pd.concat([df, dupes], ignore_index=True)
    bad = df.sample(frac=0.02, random_state=seed + 1).index
    df["submitted_date"] = df["submitted_date"].astype(object)
    df.loc[bad, "submitted_date"] = "N/A"
    early = df[df["completed_date"].notna()].sample(frac=0.02, random_state=seed + 2).index
    df.loc[early, "completed_date"] = pd.to_datetime("2000-01-01").date()
    return df.sample(frac=1, random_state=seed).reset_index(drop=True)


def save_raw_excel(df, path):
    df.to_excel(path, index=False, sheet_name="Raw Export")
    return path
