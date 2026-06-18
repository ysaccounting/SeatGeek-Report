"""
OPTIONAL: bulk-lock categories from an already-categorised report.

Usage:
    python seed_lookup.py "April_2026_-_SG2_Report.xlsx" [more_reports...]

By default the app lets its rules categorise everything and only "locks" a category
when you explicitly change one in the review gate. If instead you want to force a batch
of names to specific categories from a historical report (overriding the rules for those
names), run this. It reads each file's Invoice Details sheet, takes the Performer/Team +
Category columns, and writes them into overrides.json (which wins over the rules).

Most users will not need this — it's here for the case where you want to pin a large set
of past decisions verbatim.
"""
import sys, pandas as pd, app

def load_invoice_sheet(path):
    if path.lower().endswith(".csv"):
        return pd.read_csv(path)
    sheets = pd.read_excel(path, sheet_name=None)
    for name in ("Invoice Details", "Invoice Detail", "Invoices"):
        if name in sheets:
            return sheets[name]
    return max(sheets.values(), key=lambda d: d.shape[0])

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__); sys.exit(1)
    for p in sys.argv[1:]:
        df = load_invoice_sheet(p)
        n = app.seed_overrides_from_df(df)
        print(f"  {p}: {n} performer(s) locked")
    print(f"overrides.json now has {len(app.load_overrides())} locked entries.")
