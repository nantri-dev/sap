"""
universal_ingest.py - read ANY finance spreadsheet and write output/briefing.csv

Works with messy real-world files:
  * title rows / blank rows above the table (it finds the header row itself)
  * any column names ("Principal Amt", "Notional", "Balance", ...)
  * several sheets with different layouts
  * amounts stored as text ("$1,200.50", "(500)", "INR 12,34,567")
  * dates stored as real dates or as text
  * total / subtotal rows (dropped automatically)

How it decides what a column is: header keywords + what the values look like.
If it guesses wrong, fix it in memory/column_map.json (see bottom of this file).

Everything is plain rules + arithmetic. No AI, no internet.
The local AI is only used later, to EXPLAIN the flagged items.

Command line:   python universal_ingest.py data/myfile.xlsx
From app.py:    universal_ingest.run(path)  ->  list of report lines
"""
import csv
import json
import math
import numbers
import os
import re
import sys
import warnings
from datetime import date, datetime

import pandas as pd

BASE = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE, "data")
OUT_DIR = os.path.join(BASE, "output")
BRIEFING_CSV = os.path.join(OUT_DIR, "briefing.csv")
SUMMARY_TXT = os.path.join(OUT_DIR, "data_summary.txt")
MAP_FILE = os.path.join(BASE, "memory", "column_map.json")

# 1 = day first (India / UK / EU, 03/04/2025 = 3 April). Set LEDGERMIND_DAYFIRST=0 for US style.
DAYFIRST = os.environ.get("LEDGERMIND_DAYFIRST", "1") != "0"

# --------------------------------------------------------------------------
# what we look for
# --------------------------------------------------------------------------
KEYWORDS = {
    "due_date": ["due date", "due", "maturity", "matures", "expiry", "expiration", "end date",
                 "settlement date", "payment date", "value date", "repay", "faellig", "vencimiento", "echeance"],
    "amount": ["amount", "amt", "principal", "balance", "notional", "outstanding", "exposure",
               "value", "total", "cash", "nominal", "sum", "debit", "credit", "betrag", "importe", "montant", "monto"],
    "currency": ["currency", "ccy", "curr"],
    "status": ["status", "state", "stage"],
    "counterparty": ["counterparty", "customer", "client", "vendor", "supplier", "borrower",
                     "payee", "issuer", "party", "bank", "name", "entity", "company", "kunde", "cliente", "proveedor"],
    "type": ["type", "category", "product", "instrument", "class", "segment", "kind",
             "description"],
    "date": ["date", "booked", "start", "posting", "trade", "created", "transaction", "opened"],
    "id": ["id", "ref", "reference", "number", "contract", "deal", "invoice", "account", "no"],
}
# order matters: more specific fields claim their column first
ORDER = ["due_date", "amount", "currency", "status", "counterparty", "type", "date", "id"]

CLOSED_RE = re.compile(r"\b(closed|paid|settled|repaid|complete\w*|cancel\w*|matured|redeemed|written off)\b", re.I)
RISK_RE = re.compile(r"\b(default\w*|delinquen\w*|overdue|late|breach\w*|fail\w*|reject\w*|disput\w*|"
                     r"blocked|on hold|npa|watchlist|escalat\w*)\b", re.I)
TOTAL_RE = re.compile(r"^\s*(grand\s+total|sub\s*-?\s*total|totals?|sum)\s*:?\s*$", re.I)

NUM_RE = re.compile(r"^\(?-?[\d,]*\.?\d+\)?$")
CUR_RE = re.compile(r"(?i)\b(usd|eur|gbp|inr|rs)\b\.?|[$\u20ac\u00a3\u20b9\u00a5]")


# --------------------------------------------------------------------------
# small parsers
# --------------------------------------------------------------------------
def _blank(v):
    if isinstance(v, (list, tuple, dict)):
        return False
    try:
        if pd.isna(v):
            return True
    except (TypeError, ValueError):
        pass
    return isinstance(v, str) and not v.strip()


def parse_number(v):
    """'$1,200.50' -> 1200.5, '(500)' -> -500, 'Loan-12' -> nan."""
    if _blank(v) or isinstance(v, bool):
        return math.nan
    if isinstance(v, numbers.Number):
        return float(v)
    if isinstance(v, (datetime, date)):
        return math.nan
    s = CUR_RE.sub("", str(v)).replace(" ", "").strip()
    if not NUM_RE.match(s):
        return math.nan
    neg = s.startswith("(") and s.endswith(")")
    s = s.strip("()").replace(",", "")
    try:
        x = float(s)
    except ValueError:
        return math.nan
    return -abs(x) if neg else x


def parse_date(v):
    if _blank(v):
        return pd.NaT
    if isinstance(v, (datetime, date)):
        return pd.Timestamp(v)
    if isinstance(v, str):
        s = v.strip()
        if NUM_RE.match(CUR_RE.sub("", s).replace(" ", "")):
            return pd.NaT  # a bare number is not a date
        if len(s) < 6 or not re.search(r"\d", s):
            return pd.NaT
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            try:
                return pd.to_datetime(s, dayfirst=DAYFIRST, errors="coerce")
            except (ValueError, TypeError, OverflowError):
                return pd.NaT
    return pd.NaT


def clean_text(v):
    if _blank(v):
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return re.sub(r"\s+", " ", str(v)).strip()


def norm(name):
    return re.sub(r"[^a-z0-9]+", " ", str(name).lower()).strip()


# --------------------------------------------------------------------------
# step 1: read the file, find the real table inside each sheet
# --------------------------------------------------------------------------
def read_raw_sheets(path):
    ext = os.path.splitext(path)[1].lower()
    if ext in (".csv", ".txt"):
        with open(path, newline="", encoding="utf-8-sig", errors="replace") as f:
            sample = f.read(4096)
            f.seek(0)
            try:
                dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
            except csv.Error:
                dialect = csv.excel
            rows = list(csv.reader(f, dialect))
        if not rows:
            raise ValueError("The file is empty.")
        width = max(len(r) for r in rows)
        rows = [r + [None] * (width - len(r)) for r in rows]
        return {os.path.splitext(os.path.basename(path))[0]: pd.DataFrame(rows, dtype=object)}
    if ext in (".xlsx", ".xlsm"):
        return pd.read_excel(path, sheet_name=None, header=None, dtype=object)
    raise ValueError("Unsupported file type. Please use .xlsx, .xlsm or .csv "
                     "(for old .xls files: open in Excel and Save As .xlsx).")


def find_header_row(raw):
    """The header is the early row with the most text cells (title rows have only one)."""
    best_i, best_score = None, -10 ** 9
    for i in range(min(len(raw), 15)):
        cells = [v for v in raw.iloc[i] if not _blank(v)]
        texts = [v for v in cells if isinstance(v, str) and math.isnan(parse_number(v))]
        if len(texts) < 2:
            continue
        score = 2 * len(texts) - (len(cells) - len(texts))
        if score > best_score:
            best_i, best_score = i, score
    return best_i


def extract_table(raw):
    """Raw sheet -> tidy DataFrame with real column names, or (None, reason)."""
    raw = raw.dropna(how="all").dropna(axis=1, how="all").reset_index(drop=True)
    raw = raw.loc[~raw.apply(lambda r: all(_blank(v) for v in r), axis=1)].reset_index(drop=True)
    if len(raw) < 3:
        return None, "too few rows", None
    hi = find_header_row(raw)
    if hi is None:
        return None, "could not find a header row", None
    names, seen = [], {}
    for j, v in enumerate(raw.iloc[hi]):
        name = clean_text(v) or f"col_{j + 1}"
        seen[name] = seen.get(name, 0) + 1
        names.append(name if seen[name] == 1 else f"{name}_{seen[name]}")
    df = raw.iloc[hi + 1:].copy()
    df.columns = names
    df = df.reset_index(drop=True)

    # drop repeated header rows and total / subtotal rows
    lowered = {n.lower() for n in names}

    def bad_row(r):
        vals = [v for v in r if isinstance(v, str)]
        if sum(1 for v in vals if v.strip().lower() in lowered) >= 2:
            return True
        return any(TOTAL_RE.match(v) for v in vals)

    dropped_mask = df.apply(bad_row, axis=1)
    df = df.loc[~dropped_mask].reset_index(drop=True)
    return df, None, {"header_row": hi + 1, "dropped": int(dropped_mask.sum())}


# --------------------------------------------------------------------------
# step 2: work out which column is which
# --------------------------------------------------------------------------
def header_score(header, field):
    h = f" {header} "
    best = 0
    for k in KEYWORDS[field]:
        if header == k:
            best = max(best, 4)
        elif f" {k} " in h:
            best = max(best, 3)
        elif len(k) > 3 and k in header:
            best = max(best, 2)
    return best


def _vals(col, limit=200):
    return [v for v in col if not _blank(v)][:limit]


def num_ratio(vs):
    return sum(not math.isnan(parse_number(v)) for v in vs) / len(vs) if vs else 0.0


def date_ratio(vs):
    return sum(not pd.isna(parse_date(v)) for v in vs) / len(vs) if vs else 0.0


def valid_for(col, field):
    vs = _vals(col)
    if not vs:
        return False
    if field == "amount":
        return num_ratio(vs) >= 0.7
    if field in ("date", "due_date"):
        return date_ratio(vs) >= 0.6
    if field in ("counterparty", "type", "status", "currency"):
        if num_ratio(vs) >= 0.5 or date_ratio(vs) >= 0.5:
            return False
        distinct = len({str(v).strip().lower() for v in vs})
        if field == "status":
            return distinct <= 15
        if field == "currency":
            return distinct <= 12 and max(len(str(v)) for v in vs) <= 6
        return True
    return True  # id


def load_overrides(sheet):
    try:
        with open(MAP_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
    except (OSError, json.JSONDecodeError):
        return {}
    merged = dict(data.get("*", {}))
    for key, val in data.items():
        if key.lower() == sheet.lower():
            merged.update(val)
    return merged


def _median_date(col):
    ds = sorted(d for d in (parse_date(v) for v in _vals(col)) if not pd.isna(d))
    return ds[len(ds) // 2] if ds else None


def _fallbacks(df, cols, mapping, guessed, used):
    """Columns with unhelpful headers ('Who', 'Ends', 'Col A'): guess from the values."""
    def take(field, col):
        mapping[field] = col
        guessed.add(field)
        used.add(col)

    # dates
    dcols = [c for c in cols if c not in used and valid_for(df[c], "date") and _median_date(df[c]) is not None]
    dcols.sort(key=lambda c: _median_date(df[c]))
    if dcols:
        if "date" not in mapping and "due_date" not in mapping:
            if len(dcols) == 1:
                recent = _median_date(df[dcols[0]]) >= pd.Timestamp.today() - pd.Timedelta(days=30)
                take("due_date" if recent else "date", dcols[0])
            else:
                first, last = dcols[0], dcols[-1]
                take("date", first)
                take("due_date", last)
        elif "date" in mapping and "due_date" not in mapping:
            if _median_date(df[dcols[-1]]) > _median_date(df[mapping["date"]]):
                take("due_date", dcols[-1])
        elif "due_date" in mapping and "date" not in mapping:
            if _median_date(df[dcols[0]]) < _median_date(df[mapping["due_date"]]):
                take("date", dcols[0])

    # counterparty: a text column whose values repeat (a few parties, many rows)
    if "counterparty" not in mapping:
        best, best_d = None, 0
        for c in cols:
            if c in used or not valid_for(df[c], "counterparty"):
                continue
            vs = _vals(df[c])
            d = len({str(v).strip().lower() for v in vs})
            if 2 <= d <= 0.8 * len(vs) and d > best_d:
                best, best_d = c, d
        if best is not None:
            take("counterparty", best)

    # id: a text column where every value is different
    if "id" not in mapping:
        for c in cols:
            if c in used or not valid_for(df[c], "counterparty"):
                continue
            vs = _vals(df[c])
            if len(vs) > 2 and len({str(v).strip().lower() for v in vs}) == len(vs):
                take("id", c)
                break


def detect_columns(df, sheet):
    cols = list(df.columns)
    lower = {c.lower(): c for c in cols}
    mapping, guessed, used = {}, set(), set()

    for field, name in load_overrides(sheet).items():
        c = lower.get(str(name).lower())
        if c is not None and field in KEYWORDS:
            mapping[field] = c
            used.add(c)

    for field in ORDER:
        if field in mapping:
            continue
        best, best_s = None, 0
        for c in cols:
            if c in used:
                continue
            s = header_score(norm(c), field)
            if s > best_s and valid_for(df[c], field):
                best, best_s = c, s
        if best is not None:
            mapping[field] = best
            used.add(best)

    if "amount" not in mapping:  # last resort: biggest-looking numeric column
        best, best_med = None, -1
        for c in cols:
            if c in used or header_score(norm(c), "id") or header_score(norm(c), "date"):
                continue
            vs = _vals(df[c])
            if vs and num_ratio(vs) >= 0.8:
                nums = sorted(abs(parse_number(v)) for v in vs)
                med = nums[len(nums) // 2]
                if med > best_med:
                    best, best_med = c, med
        if best is not None:
            mapping["amount"] = best
            guessed.add("amount")
    _fallbacks(df, cols, mapping, guessed, used)
    return mapping, guessed


def to_canonical(df, mapping, sheet):
    out = pd.DataFrame(index=df.index)
    out["sheet"] = sheet
    out["amount"] = df[mapping["amount"]].map(parse_number)
    for f in ("counterparty", "type", "status", "id", "currency"):
        out[f] = df[mapping[f]].map(clean_text) if f in mapping else ""
    for f in ("date", "due_date"):
        if f in mapping:
            out[f] = pd.to_datetime(pd.Series([parse_date(v) for v in df[mapping[f]]], index=df.index),
                                    errors="coerce")
        else:
            out[f] = pd.Series(pd.NaT, index=df.index, dtype="datetime64[ns]")
    out["sheet_has_due"] = "due_date" in mapping
    out["category"] = [t.lower() if t else sheet.lower() for t in out["type"]]
    out = out[out["amount"].notna()].reset_index(drop=True)
    return out


# --------------------------------------------------------------------------
# step 3: the rules that create briefing items
# --------------------------------------------------------------------------
def money(x, cur=""):
    return f"{cur}{'-' if x < 0 else ''}{abs(x):,.0f}"


def cnt(n, word):
    """1 -> '1 item', 3 -> '3 items'"""
    return f"{n} {word}" + ("" if n == 1 else "s")


def dfmt(d):
    return d.strftime("%d %b %Y")


def row_label(r):
    cp, rid = r["counterparty"], r["id"]
    if cp and rid and cp != rid:
        return f"{cp} ({rid})"
    return cp or rid or r["category"].title()


def pick_as_of(data):
    today = pd.Timestamp.today().normalize()
    dates = pd.concat([data["date"], data["due_date"]]).dropna()
    if dates.empty:
        return today, False
    latest = dates.max()
    if latest < today - pd.Timedelta(days=400):  # old file: judge it against its own newest date
        return latest, True
    return today, False


def currency_info(data):
    cs = sorted({c.upper() for c in data["currency"] if c})
    if len(cs) == 1:
        return cs[0] + " ", False
    return "", len(cs) > 1


def build_items(data, has, as_of, stale, cur):
    items = []
    n = len(data)
    absamt = data["amount"].abs()
    tot_abs = float(absamt.sum()) or 1.0
    p95 = absamt.quantile(0.95)
    p90 = absamt.quantile(0.90)

    active = data
    if has["status"]:
        active = data[~data["status"].map(lambda s: bool(CLOSED_RE.search(s)))]

    # 1. biggest positions (priority follows the size, so the biggest is never 'low')
    k = min(5, max(1, n // 10))
    for _, r in data.loc[absamt.nlargest(k).index].iterrows():
        kind = "outflow" if r["amount"] < 0 else "position"
        pr = "high" if abs(r["amount"]) >= p95 else "medium"
        share = abs(r["amount"]) / tot_abs * 100
        items.append((r["category"], f"Large {kind}: {row_label(r)} {money(r['amount'], cur)}",
                      f"{money(r['amount'], cur)} in {r['category']}, about {share:.0f}% of the total "
                      f"exposure in the file. " + ("The biggest entry" if k == 1 else f"One of the {k} biggest entries")
                      + f" out of {n} rows.", pr))

    # 2. possible duplicates
    others = [c for c in ("counterparty", "date", "due_date") if has[c]]
    if others:
        def key(r):
            parts = [str(r["sheet"]), f"{r['amount']:.2f}"]
            got = 0
            for c in others:
                v = r[c]
                if isinstance(v, str):
                    v = v or None
                elif pd.isna(v):
                    v = None
                if v is not None:
                    parts.append(str(v))
                    got += 1
            return "|".join(parts) if got else None

        keys = data.apply(key, axis=1)
        mask = keys.notna() & keys.duplicated(keep=False)
        groups = [g for _, g in data[mask].groupby(keys[mask])]
        groups.sort(key=lambda g: -abs(g["amount"].iloc[0]))
        same = ", ".join(["amount"] + [o.replace("_", " ") for o in others])
        for g in groups[:5]:
            r = g.iloc[0]
            pr = "high" if abs(r["amount"]) >= p90 else "medium"
            items.append((r["category"], f"Possible duplicate: {row_label(r)} {money(r['amount'], cur)}",
                          f"{len(g)} rows share the same {same}. Check whether one was entered twice.", pr))

    # 3. dates
    if has["due_date"]:
        due = active.dropna(subset=["due_date"])
        note = f" Measured as of {dfmt(as_of)}, the newest date in the file." if stale else ""
        over = due[due["due_date"] < as_of].sort_values("amount", key=lambda s: -s.abs())
        if len(over):
            top = "; ".join(f"{row_label(r)} {money(r['amount'], cur)} (due {dfmt(r['due_date'])})"
                            for _, r in over.head(3).iterrows())
            items.append(("dates", f"{cnt(len(over), 'item')} {'is' if len(over) == 1 else 'are'} past {'its' if len(over) == 1 else 'their'} due date",
                          f"Together {money(over['amount'].sum(), cur)}. Largest: {top}.{note}", "high"))
        soon = due[(due["due_date"] >= as_of) & (due["due_date"] <= as_of + pd.Timedelta(days=7))]
        if len(soon):
            top = "; ".join(f"{row_label(r)} {money(r['amount'], cur)} (due {dfmt(r['due_date'])})"
                            for _, r in soon.sort_values("amount", key=lambda s: -s.abs()).head(3).iterrows())
            items.append(("dates", f"{cnt(len(soon), 'item')} due in the next 7 days",
                          f"Together {money(soon['amount'].sum(), cur)}. Largest: {top}.", "medium"))
        month = due[(due["due_date"] > as_of + pd.Timedelta(days=7)) &
                    (due["due_date"] <= as_of + pd.Timedelta(days=30))]
        if len(month):
            items.append(("dates", f"{cnt(len(month), 'item')} due in the next 8-30 days",
                          f"Together {money(month['amount'].sum(), cur)}.", "low"))
        missing = active[active["due_date"].isna() & active["sheet_has_due"]]
        if len(missing):
            pr = "medium" if abs(missing["amount"]).sum() / tot_abs >= 0.05 else "low"
            items.append(("data quality", f"{cnt(len(missing), 'item')} {'has' if len(missing) == 1 else 'have'} no due date on file",
                          f"Together {money(missing['amount'].sum(), cur)}. Their timing cannot be checked.", pr))

    # 4. negative amounts
    neg = data[data["amount"] < 0]
    if len(neg):
        items.append(("cash flow", f"{cnt(len(neg), 'negative entry').replace('entrys', 'entries')} (outflows or overdrafts)",
                      f"Together {money(neg['amount'].sum(), cur)}. Largest: "
                      f"{row_label(neg.loc[neg['amount'].idxmin()])} {money(neg['amount'].min(), cur)}.",
                      "medium"))

    # 5. concentration
    if has["counterparty"]:
        named = data[data["counterparty"] != ""]
        by = named.groupby("counterparty")["amount"].apply(lambda s: s.abs().sum())
        if len(by) >= 3 and by.sum() > 0:
            share = by.max() / by.sum()
            if share >= 0.30:
                items.append(("concentration", f"Concentration: {by.idxmax()} holds {share * 100:.0f}% of the total",
                              f"{money(by.max(), cur)} of {money(by.sum(), cur)} across {len(by)} counterparties.",
                              "high" if share >= 0.5 else "medium"))

    # 6. risky statuses
    if has["status"]:
        risky = data[data["status"].map(lambda s: bool(RISK_RE.search(s)))]
        if len(risky):
            breakdown = ", ".join(f"{s}: {c}" for s, c in risky["status"].value_counts().head(4).items())
            items.append(("risk status", f"{cnt(len(risky), 'item')} {'carries' if len(risky) == 1 else 'carry'} a risk status",
                          f"Together {money(risky['amount'].sum(), cur)} ({breakdown}).", "high"))

    if not items:
        items.append(("info", f"No exceptions found across {n} rows",
                      "No large positions, duplicates, date problems or risk statuses were detected.", "low"))
    return items


# --------------------------------------------------------------------------
# step 4: a compact summary so the ask bar can answer questions about ALL rows
# --------------------------------------------------------------------------
def build_summary(data, has, as_of, stale, cur, mixed):
    lines = [f"DATA SUMMARY (whole file: {len(data)} rows; sheets: {', '.join(sorted(set(data['sheet'])))}).",
             f"Total of all amounts: {money(data['amount'].sum(), cur)}. "
             f"Total ignoring signs: {money(data['amount'].abs().sum(), cur)}."]
    if mixed:
        lines.append("WARNING: several currencies are present and were NOT converted, so totals mix currencies.")
    lines.append("By category: " + "; ".join(
        f"{c}: {len(g)} rows, {money(g['amount'].sum(), cur)}"
        for c, g in data.groupby("category")) + ".")
    if has["counterparty"]:
        by = data[data["counterparty"] != ""].groupby("counterparty")["amount"].apply(lambda s: s.abs().sum())
        lines.append("Biggest counterparties: " + "; ".join(
            f"{n} {money(v, cur)}" for n, v in by.sort_values(ascending=False).head(8).items()) + ".")
    dates = pd.concat([data["date"], data["due_date"]]).dropna()
    if len(dates):
        lines.append(f"Dates in file run from {dfmt(dates.min())} to {dfmt(dates.max())}. "
                     f"'As of' date used for overdue checks: {dfmt(as_of)}.")
    lines.append("Rows (largest first): label | category | amount | due date | status")
    for _, r in data.loc[data["amount"].abs().sort_values(ascending=False).index].head(40).iterrows():
        due = dfmt(r["due_date"]) if pd.notna(r["due_date"]) else "no due date"
        lines.append(f"- {row_label(r)} | {r['category']} | {money(r['amount'], cur)} | {due} | {r['status'] or '-'}")
    return "\n".join(lines)[:4500]


# --------------------------------------------------------------------------
# main entry point
# --------------------------------------------------------------------------
def run(path):
    report = []
    frames = []
    for sheet, raw in read_raw_sheets(path).items():
        sheet = str(sheet)
        df, why, info = extract_table(raw)
        if df is None:
            report.append(f"Sheet '{sheet}': skipped ({why}).")
            continue
        mapping, guessed = detect_columns(df, sheet)
        if "amount" not in mapping:
            report.append(f"Sheet '{sheet}': skipped (no amount column found).")
            continue
        frame = to_canonical(df, mapping, sheet)
        if frame.empty:
            report.append(f"Sheet '{sheet}': skipped (no rows with a usable amount).")
            continue
        frames.append(frame)
        found = ", ".join(f"{f} = '{mapping[f]}'" + (" (guess)" if f in guessed else "")
                          for f in ORDER if f in mapping)
        extra = f", {info['dropped']} total rows ignored" if info["dropped"] else ""
        report.append(f"Sheet '{sheet}': {len(frame)} rows, header on row {info['header_row']}{extra}. "
                      f"Read as: {found}.")
        missing = [f for f in ("counterparty", "due_date") if f not in mapping]
        if missing:
            report.append(f"   Not found in '{sheet}': {', '.join(m.replace('_', ' ') for m in missing)} "
                          f"(some checks are skipped). Fix in memory/column_map.json if it exists under another name.")
    if not frames:
        raise ValueError("Could not find a usable table with an amount column. "
                         "Add a mapping in memory/column_map.json (see universal_ingest.py).")

    data = pd.concat(frames, ignore_index=True)
    for c in ("date", "due_date"):
        data[c] = pd.to_datetime(data[c], errors="coerce")
    has = {f: bool((data[f].notna() & (data[f] != "")).any()) if f in ("counterparty", "status", "id")
           else bool(data[f].notna().any()) for f in ("counterparty", "status", "id", "date", "due_date")}
    as_of, stale = pick_as_of(data)
    cur, mixed = currency_info(data)
    if mixed:
        report.append("Warning: several currencies found. Amounts are not converted, so totals mix currencies.")
    if stale:
        report.append(f"Note: dates in this file are old, so overdue checks use {dfmt(as_of)} (newest date in the file).")

    items = build_items(data, has, as_of, stale, cur)
    os.makedirs(OUT_DIR, exist_ok=True)
    with open(BRIEFING_CSV, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["category", "title", "detail", "priority"])
        for category, title, detail, priority in items:
            w.writerow([category, title, detail, priority])
    with open(SUMMARY_TXT, "w", encoding="utf-8") as f:
        f.write(build_summary(data, has, as_of, stale, cur, mixed))
    report.append(f"Created {cnt(len(items), 'briefing item')} from {len(data)} rows.")
    return report


# --------------------------------------------------------------------------
# Optional: fix a wrong guess. Create memory/column_map.json like this:
#   {
#     "*":     {"amount": "Principal Amt"},                # applies to every sheet
#     "Loans": {"due_date": "Maturity", "counterparty": "Borrower"}
#   }
# Field names: amount, due_date, date, counterparty, type, status, currency, id
# Column names must match the header in your file (upper/lower case does not matter).
# --------------------------------------------------------------------------
if __name__ == "__main__":
    target = sys.argv[1] if len(sys.argv) > 1 else None
    if target is None:
        files = [os.path.join(DATA_DIR, n) for n in os.listdir(DATA_DIR)
                 if n.lower().endswith((".xlsx", ".xlsm", ".csv")) and not n.startswith("~$")]
        if not files:
            sys.exit("No spreadsheet found. Usage: python universal_ingest.py path/to/file.xlsx")
        target = max(files, key=os.path.getmtime)
    for line in run(target):
        print(line)
    print(f"\nWrote {BRIEFING_CSV}")
