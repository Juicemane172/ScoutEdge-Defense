"""
DefensiveIQ-style opponent tendency analyzer.

Takes a Hudl playlist export (.xlsx/.csv) and produces a multi-tab Excel
tendency report: field-zone tendencies, run/pass/hash/formation breakdowns,
down & distance tendencies, concept efficiency, situational summary,
practice scripts, a call sheet builder, and an auto-filled game day call sheet.

Built by reverse-engineering the INPUT (Hudl export columns) and OUTPUT
(DefensiveIQ workbook structure) of one real report — not from any source
code, which was never available. Some internal formulas (exact "success
rate" definition, field-zone boundaries) are inferred from the data and are
documented inline; tweak the constants below if your own eye test disagrees.
"""

import pandas as pd
import numpy as np
from collections import Counter

# ----------------------------------------------------------------------
# Tunable definitions (inferred from sample data - adjust to taste)
# ----------------------------------------------------------------------

ZONE_ORDER = ["BZ", "OF", "MF", "FZ", "RZ", "GL"]
ZONE_LABEL = {
    "BZ": "Backed Up · Own 1-20",
    "OF": "Open Field · Own 21-49",
    "MF": "Midfield · 50-Opp 40",
    "FZ": "Fringe · Opp 39-21",
    "RZ": "Red Zone · Opp 20-11",
    "GL": "Goal Line · Opp 10 and in",
}

EXPLOSIVE_RUN_YDS = 10
EXPLOSIVE_PASS_YDS = 15
MIN_SAMPLE_CONCEPT = 3     # concepts need >=3 calls to appear in concept tables
MIN_SAMPLE_FORMATION = 3  # formations need >=3 snaps to appear in formation table
SMALL_SAMPLE_FLAG = 5     # below this, flag with "*" as small sample


def classify_zone(yard_ln):
    """Hudl YARD LN convention: negative = own side (abs value = own yard
    line), positive = opponent side (value = distance from opp goal)."""
    if pd.isna(yard_ln):
        return None
    y = yard_ln
    if y < 0:
        v = abs(y)
        return "BZ" if v <= 20 else "OF"
    else:
        if y >= 40:
            return "MF"
        elif y >= 21:
            return "FZ"
        elif y >= 11:
            return "RZ"
        else:
            return "GL"


def dd_bucket(dn, dist):
    """Down & distance situation bucket."""
    if pd.isna(dn):
        return "OTHER"
    dn = int(dn)
    if dn == 0:
        # Hudl tags the first snap of a new possession as DN=0. In practice
        # this always comes with DIST=10 in real exports - it's "P & 10"
        # (Possession and 10), the first-down of a new series, not a random
        # zero. Treat it as its own bucket rather than folding it into OTHER.
        return "P & 10"
    if dn == 1:
        return "1ST & 10" if (pd.isna(dist) or dist >= 10) else "1ST & SHORT"
    if dn == 2:
        # 2nd down: Short 1-3, Medium 4-6, Long 7+
        if pd.isna(dist):
            return "2ND & MEDIUM"
        if dist >= 7:
            return "2ND & LONG"
        elif dist >= 4:
            return "2ND & MEDIUM"
        else:
            return "2ND & SHORT"
    if dn == 3:
        # 3rd down uses wider bands than 2nd: Short 1-3, Medium 4-8, Long 9+
        if pd.isna(dist):
            return "3RD & MEDIUM"
        if dist >= 9:
            return "3RD & LONG"
        elif dist >= 4:
            return "3RD & MEDIUM"
        else:
            return "3RD & SHORT"
    if dn == 4:
        return "4TH DOWN"
    return "OTHER"


def is_success(dn, dist, gain):
    """Standard down-based success-rate heuristic:
    1st down: >=40% of distance. 2nd down: >=60% of distance.
    3rd/4th down: full conversion (>= distance)."""
    if pd.isna(dn) or pd.isna(dist) or pd.isna(gain):
        return None
    dn = int(dn)
    if dn == 1:
        return gain >= 0.4 * dist
    if dn == 2:
        return gain >= 0.6 * dist
    if dn in (3, 4):
        return gain >= dist
    return None


def load_playlist(path):
    """Read a Hudl playlist export (.xlsx or .csv) and clean it up."""
    if str(path).lower().endswith(".csv"):
        df = pd.read_csv(path)
    else:
        df = pd.read_excel(path)

    df.columns = [c.strip().upper() for c in df.columns]

    # Drop no-play / penalty rows and rows missing the essentials
    if "RESULT" in df.columns:
        df = df[~df["RESULT"].astype(str).str.contains("Penalty", case=False, na=False)]
    df = df[df["PLAY TYPE"].isin(["Run", "Pass"])].copy()

    df["ZONE"] = df["YARD LN"].apply(classify_zone)
    df["DD_BUCKET"] = df.apply(lambda r: dd_bucket(r.get("DN"), r.get("DIST")), axis=1)
    df["GAIN"] = pd.to_numeric(df.get("GN/LS"), errors="coerce")
    df["SUCCESS"] = df.apply(lambda r: is_success(r.get("DN"), r.get("DIST"), r["GAIN"]), axis=1)
    df["EXPLOSIVE"] = df.apply(
        lambda r: (r["PLAY TYPE"] == "Run" and r["GAIN"] >= EXPLOSIVE_RUN_YDS)
        or (r["PLAY TYPE"] == "Pass" and r["GAIN"] >= EXPLOSIVE_PASS_YDS),
        axis=1,
    )
    df["CONCEPT"] = df.get("PLAY")
    df["FORMATION"] = df.get("FORMATION")
    df["HASH"] = df.get("HASH")
    df["BACKFIELD"] = df.get("BACKFIELD")

    return df.reset_index(drop=True)


def top_n_counts(series, n=3):
    """Return list of '<value> (<count>)' strings for the top-n values."""
    vc = series.dropna()
    vc = vc[vc.astype(str).str.strip() != ""]
    counts = Counter(vc)
    top = counts.most_common(n)
    out = [f"{val} ({cnt})" for val, cnt in top]
    while len(out) < n:
        out.append("—")
    return out


def pct(numer, denom):
    if not denom:
        return None
    return numer / denom


def run_pass_split(sub):
    runs = (sub["PLAY TYPE"] == "Run").sum()
    passes = (sub["PLAY TYPE"] == "Pass").sum()
    total = runs + passes
    return runs, passes, total, pct(runs, total), pct(passes, total)


def call_idea(run_pct, pass_pct):
    if run_pct is None:
        return ""
    if run_pct >= 0.65:
        return "Sell run first, rob crossers on play-action"
    if pass_pct is not None and pass_pct >= 0.65:
        return "Bring pressure, sit zone underneath"
    return "Balanced look — stay multiple, disguise front"
