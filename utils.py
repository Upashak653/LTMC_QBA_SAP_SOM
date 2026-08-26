import html
import pandas as pd


def parse_date_safe(val):
    try:
        return pd.to_datetime(str(val).strip(), errors="coerce")
    except Exception:
        return pd.NaT


def find_col_dynamic(df, keywords):
    """Locate a source column by header text using priority keyword matching."""
    for kw in keywords:
        for i, c in enumerate(df.columns):
            first_line = str(c).split("\n")[0].strip().rstrip("*").strip()
            if kw.lower() == first_line.lower():
                return i
    for kw in keywords:
        for i, c in enumerate(df.columns):
            first_line = str(c).split("\n")[0].strip().rstrip("*").strip()
            if kw.lower() in first_line.lower():
                return i
    return None


def format_date_ddmmyyyy(val):
    d = parse_date_safe(val)
    return d.strftime("%d-%m-%Y") if pd.notna(d) else ""


def negate_num(val):
    """Ensure a numeric string is negative (e.g. depreciation amounts)."""
    s = str(val).strip()
    if not s or s.lower() == "nan":
        return s
    try:
        f = float(s)
    except ValueError:
        return s
    if f == 0:
        return s
    if s.startswith("-"):
        return s
    return "-" + s


def validate_rows(rows, headers, required_labels, numeric_labels, key_label, area_label=None):
    idx = {h: i for i, h in enumerate(headers)}

    blank_counts = {}
    for label in required_labels:
        i = idx.get(label)
        if i is None:
            continue
        cnt = sum(1 for r in rows if not str(r[i]).strip())
        if cnt:
            blank_counts[label] = cnt

    bad_numeric = {}
    for label in numeric_labels:
        i = idx.get(label)
        if i is None:
            continue
        bad = 0
        for r in rows:
            v = str(r[i]).strip()
            if v == "":
                continue
            try:
                float(v)
            except ValueError:
                bad += 1
        if bad:
            bad_numeric[label] = bad

    dup_count = 0
    key_i = idx.get(key_label)
    area_i = idx.get(area_label) if area_label else None
    if key_i is not None:
        seen = {}
        for r in rows:
            k = (r[key_i], r[area_i]) if area_i is not None else r[key_i]
            seen[k] = seen.get(k, 0) + 1
        dup_count = sum(1 for v in seen.values() if v > 1)

    return {
        "total_rows": len(rows),
        "blank_required": blank_counts,
        "invalid_numeric": bad_numeric,
        "duplicate_keys": dup_count,
    }


def letter_to_idx(letter):
    idx = 0
    for c in letter:
        idx = idx * 26 + (ord(c.upper()) - 64)
    return idx - 1


def clean_xml(val):
    return html.escape(str(val).strip())