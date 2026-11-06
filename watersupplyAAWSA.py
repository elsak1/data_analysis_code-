
from decimal import Decimal, ROUND_HALF_UP

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

INPUT = "supply.xlsx"
OUTPUT = "branch_by_subcity.xlsx"

# Display order of branches (any branch not listed is added afterwards in file order)
BRANCH_ORDER = ["Arada", "Gulele", "Addis Ketema", "Gurde Shola",
                "Megenaga", "Akaki", "Nifas Silk Lafto"]


# ---------------------------------------------------------------- load + clean
def load_data(path):
    # data_only=True reads the saved values of the Days/Hours formulas
    ws = load_workbook(path, data_only=True).active
    rows = list(ws.iter_rows(min_row=2, values_only=True))
    df = pd.DataFrame([r[:7] for r in rows],
                      columns=["Branch", "ID", "Sub", "Woreda", "Text", "Days", "Hours"])
    df["Branch"] = df["Branch"].ffill().str.strip().str.title()
    df["Sub"] = df["Sub"].str.strip().replace({"Gullele": "Gulele"})
    df["Days"] = pd.to_numeric(df["Days"], errors="coerce")
    df["Hours"] = pd.to_numeric(df["Hours"], errors="coerce")
    return df


def classify(days, hours):
    """Continuous = 7 d x 24 h; Daily <24 h = 7 d with fewer (or unknown) hours;
    Non-daily = fewer than 7 days/week."""
    if days < 7:
        return "Non-daily"
    if pd.notna(hours) and hours >= 24:
        return "Continuous"
    return "Daily <24 h"


# ---------------------------------------------------------------- formatting
def woreda_ranges(values):
    """[6,7,8,9,10,12,13] -> '6-10, 12, 13' (ranges only for 3+ consecutive)."""
    nums = sorted({int(w) for w in values if isinstance(w, (int, float))})
    text = sorted({str(w) for w in values if not isinstance(w, (int, float))})
    parts, i = [], 0
    while i < len(nums):
        j = i
        while j + 1 < len(nums) and nums[j + 1] == nums[j] + 1:
            j += 1
        parts += [f"{nums[i]}-{nums[j]}"] if j - i >= 2 else [str(n) for n in nums[i:j + 1]]
        i = j + 1
    return ", ".join(parts + text)


def n_pct(n, total):
    if not total:
        return "0 (0.0%)"
    # round half up (like Excel), not Python's default round-half-even
    pct = (Decimal(int(n)) * 100 / Decimal(total)).quantize(Decimal("0.1"), rounding=ROUND_HALF_UP)
    return f"{n} ({pct}%)"


# ---------------------------------------------------------------- build table
def build_rows(df):
    df = df.copy()
    df["Class"] = [classify(d, h) for d, h in zip(df["Days"], df["Hours"])]
    branches = BRANCH_ORDER + [b for b in df["Branch"].unique() if b not in BRANCH_ORDER]
    rows = []
    for b in branches:
        g = df[df["Branch"] == b]
        if g.empty:
            continue
        # the sub-city sharing the branch's name goes first, then alphabetical
        subs = sorted(g["Sub"].unique(), key=lambda s: (s.split()[0] != b.split()[0], s))
        for k, s in enumerate(subs):
            gs = g[g["Sub"] == s]
            rows.append({
                "branch": b if k == 0 else None,
                "sub": s,
                "woredas": woreda_ranges(gs["Woreda"]),
                "n": len(gs),
                "cont": (gs["Class"] == "Continuous").sum(),
                "daily": (gs["Class"] == "Daily <24 h").sum(),
                "non": (gs["Class"] == "Non-daily").sum(),
            })
    return rows, df["Class"].value_counts()


# ---------------------------------------------------------------- write sheet
def write_sheet(rows, overall, path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Branch by sub-city"
    font = "Arial"
    thin = Side(style="thin", color="BFBFBF")
    heavy = Side(style="medium", color="1F4E78")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)

    heads = ["Branch", "Sub cities served", "Woredas served", "Entries",
             "Continuous (7 d x 24 h), n (%)", "Daily, under 24 h, n (%)",
             "Non-daily (<7 d/week), n (%)"]
    for j, h in enumerate(heads, 1):
        c = ws.cell(1, j, h)
        c.font = Font(name=font, bold=True, color="FFFFFF", size=10)
        c.fill = PatternFill("solid", fgColor="1F4E78")
        c.border = box
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 42

    r = 2
    for row in rows:
        vals = [row["branch"], row["sub"], row["woredas"], row["n"],
                n_pct(row["cont"], row["n"]), n_pct(row["daily"], row["n"]),
                n_pct(row["non"], row["n"])]
        for j, v in enumerate(vals, 1):
            c = ws.cell(r, j, v)
            c.font = Font(name=font, size=10, bold=(j == 1))
            top = heavy if (row["branch"] and r > 2) else thin  # line between branches
            c.border = Border(left=thin, right=thin, bottom=thin, top=top)
            c.alignment = Alignment(vertical="top", horizontal="left" if j <= 3 else "center")
        r += 1

    total = sum(row["n"] for row in rows)
    vals = ["Total", None, None, total,
            n_pct(overall.get("Continuous", 0), total),
            n_pct(overall.get("Daily <24 h", 0), total),
            n_pct(overall.get("Non-daily", 0), total)]
    for j, v in enumerate(vals, 1):
        c = ws.cell(r, j, v)
        c.font = Font(name=font, size=10, bold=True)
        c.fill = PatternFill("solid", fgColor="DDEBF7")
        c.border = box
        c.alignment = Alignment(horizontal="left" if j <= 3 else "center")

    note = ("Continuous = 7 days/week for 24 h. Daily, under 24 h = 7 days/week for fewer (or unrecorded) "
            "hours. Non-daily = fewer than 7 days/week. Percentages are shares of each row's entries. "
            "Ranges are used where three or more woredas are consecutive.")
    ws.cell(r + 2, 1, note).font = Font(name=font, size=9, italic=True)

    for j, w in enumerate([20, 20, 26, 10, 18, 18, 18], 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = "A2"
    wb.save(path)


if __name__ == "__main__":
    data = load_data(INPUT)
    table_rows, overall_counts = build_rows(data)
    write_sheet(table_rows, overall_counts, OUTPUT)
    print(f"Saved {OUTPUT} with {len(table_rows)} rows from {len(data)} entries")