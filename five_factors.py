"""
Five-factor snapshot metrics for one offense (the "Quick Facts" report).

The five factors are Efficiency, Explosiveness, Field Position, Finishing
Drives and Turnovers - a widely used way to summarize how an offense wins.
Everything here is computed from a single team's offensive film (one Hudl
playlist export), so it describes what THIS offense does; it does not
compare against opponents or league averages.

Everything tunable lives in the constants below, with comments - edit them
if your eye test disagrees with a number.
"""

import numpy as np
import pandas as pd

from analyzer import EXPLOSIVE_RUN_YDS, EXPLOSIVE_PASS_YDS

# ----------------------------------------------------------------------
# Expected-points curve (my own approximation, not anyone's proprietary
# model): points an offense can expect to score from a given spot, by
# yards-to-goal. A play's value = expected points AFTER the play minus
# expected points BEFORE it, so a 10-yard gain near the goal line is worth
# more than a 10-yard gain on your own 20. Edit the two lists to reshape it.
# ----------------------------------------------------------------------
EP_YDS_TO_GOAL = [0, 1, 5, 10, 20, 30, 40, 50, 60, 70, 80, 90, 99]
EP_VALUE = [7.0, 6.4, 5.6, 4.8, 3.8, 3.0, 2.3, 1.7, 1.2, 0.7, 0.3, -0.1, -0.5]
TD_EP = 7.0                 # value of the ball after a touchdown
TD_POINTS = 6               # points credited to a scoring drive (no PAT/FG data)
SCORING_OPP_LINE = 40       # a drive that reaches the opponent's 40 is a scoring opportunity
FUMBLE_LOST_RATE = 0.5      # fumbles have no recovery tag, so count each as half a turnover
DRIVE_FORWARD_JUMP_YDS = 20  # forward field-position jump (bigger than any penalty) that signals a new drive
DRIVE_BACK_TOLERANCE = 1     # a 1st down that starts this far BEHIND the last spot is a new possession

# ----------------------------------------------------------------------
# Grade bands (my own defaults - tune them against the offenses you see).
# Each list is checked best-to-worst; first match wins, else POOR.
# ----------------------------------------------------------------------
GRADE_BANDS = {
    "success":   ("high", [(0.46, "ELITE"), (0.40, "STRONG"), (0.34, "AVERAGE"), (0.28, "BELOW AVG")]),
    "xpp":       ("high", [(0.45, "ELITE"), (0.30, "STRONG"), (0.15, "AVERAGE"), (0.05, "BELOW AVG")]),
    "start_d":   ("low",  [(55, "ELITE"), (65, "STRONG"), (75, "AVERAGE"), (85, "BELOW AVG")]),
    "ppso":      ("high", [(4.5, "ELITE"), (3.5, "STRONG"), (2.5, "AVERAGE"), (1.5, "BELOW AVG")]),
    "to_rate":   ("low",  [(0.010, "ELITE"), (0.025, "STRONG"), (0.040, "AVERAGE"), (0.060, "BELOW AVG")]),
    "stuff":     ("low",  [(0.12, "ELITE"), (0.18, "STRONG"), (0.24, "AVERAGE"), (0.30, "BELOW AVG")]),
    "sack":      ("low",  [(0.03, "ELITE"), (0.05, "STRONG"), (0.07, "AVERAGE"), (0.10, "BELOW AVG")]),
}

# Radar-chart anchors: (value that plots at the center, value at the rim).
# Chosen so a typical offense lands near the middle ring.
RADAR_ANCHORS = {
    "efficiency":  (0.22, 0.52),
    "explosive":   (-0.10, 0.55),
    "field_pos":   (90.0, 50.0),   # yards-to-goal at drive start (lower is better)
    "finishing":   (0.0, 6.0),
    "turnovers":   (0.08, 0.0),    # turnovers per play (lower is better)
}


def grade(metric, value):
    """Return a grade label for value, or None when there is no value."""
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return None
    direction, bands = GRADE_BANDS[metric]
    for threshold, label in bands:
        if (direction == "high" and value >= threshold) or (direction == "low" and value <= threshold):
            return label
    return "POOR"


# ----------------------------------------------------------------------
# Play-level values
# ----------------------------------------------------------------------
def yards_to_goal(yard_ln):
    """Hudl YARD LN: negative = own side (abs = own yard line), positive =
    opponent side (value = yards from their goal line)."""
    if pd.isna(yard_ln):
        return np.nan
    if yard_ln < 0:
        return 100.0 - abs(float(yard_ln))
    if yard_ln == 0:
        return 50.0
    return float(yard_ln)


def _ep(d):
    return np.interp(d, EP_YDS_TO_GOAL, EP_VALUE)


def prepare(df):
    """Copy of the analyzed playlist with play values and event flags added."""
    out = df.copy()
    res = out["RESULT"].astype(str).str.lower() if "RESULT" in out.columns else pd.Series("", index=out.index)
    out["IS_TD"] = res.str.contains(r"\btd\b", regex=True, na=False)
    out["IS_INT"] = res.str.contains("interception", na=False)
    out["IS_FUMBLE"] = res.str.contains("fumble", na=False)
    out["IS_SACK"] = res.str.contains("sack", na=False)
    out["IS_SCRAMBLE"] = res.str.contains("scramble", na=False)
    out["IS_COMPLETE"] = res.str.startswith("complete")

    d0 = out["YARD LN"].map(yards_to_goal)
    gain = pd.to_numeric(out["GAIN"], errors="coerce").fillna(0)
    d1 = (d0 - gain).clip(lower=0, upper=99)

    ep0 = _ep(d0)
    ep1 = np.where(out["IS_TD"], TD_EP, _ep(d1))
    normal = ep1 - ep0
    # A turnover hands the opponent the ball roughly where the play ended.
    opp_ep = _ep(np.clip(100 - d1, 1, 99))
    turnover = -opp_ep - ep0

    value = np.where(out["IS_INT"], turnover, normal)
    value = np.where(out["IS_FUMBLE"], (1 - FUMBLE_LOST_RATE) * normal + FUMBLE_LOST_RATE * turnover, value)
    out["D0"], out["D1"] = d0, d1
    out["VALUE"] = pd.Series(value, index=out.index).where(d0.notna())
    out["SUCC"] = pd.to_numeric(out["SUCCESS"].map(lambda v: np.nan if v is None or pd.isna(v) else float(v)),
                                errors="coerce")
    out["TO_EQ"] = out["IS_INT"].astype(float) + FUMBLE_LOST_RATE * out["IS_FUMBLE"].astype(float)
    return out


# ----------------------------------------------------------------------
# Drives
# ----------------------------------------------------------------------
def assign_drives(d, breaks=None):
    """Label every play with a drive number.

    Uses your P&10 tags (DN = 0 starts a possession) when the export has
    them. If it doesn't, drive breaks are inferred: a 1st-down snap that
    follows a score or turnover, or that follows a play that did not earn
    a first down AND starts behind the last spot (or jumps far ahead of
    it), begins a new drive. Returns (ids, method).
    """
    dn = d["DN"]
    breaks = pd.Series(False, index=d.index) if breaks is None else breaks
    if (dn == 0).any():
        return ((dn == 0) | breaks).cumsum(), "tagged"

    gain = pd.to_numeric(d["GAIN"], errors="coerce").fillna(0)
    converted = (gain >= d["DIST"]) | d["IS_TD"]
    prev_converted = converted.shift(1, fill_value=True)
    prev_end = (d["IS_TD"] | d["IS_INT"] | d["IS_FUMBLE"]).shift(1, fill_value=False)
    prev_end_spot = d["D1"].shift(1)
    moved_back = d["D0"] > prev_end_spot + DRIVE_BACK_TOLERANCE      # a penalty can't push you backward on a new 1st down
    jumped_ahead = (prev_end_spot - d["D0"]) > DRIVE_FORWARD_JUMP_YDS  # farther than any penalty gains
    new = (dn == 1) & (prev_end | ((~prev_converted) & (moved_back | jumped_ahead)))
    new = new | breaks
    new.iloc[0] = True
    return new.cumsum(), "inferred"


def summarize_drives(d, ids, method):
    rows = []
    for drive_id, g in d.groupby(ids, sort=True):
        reach = np.nanmin(np.minimum(g["D0"].to_numpy(), g["D1"].to_numpy())) if len(g) else np.nan
        rows.append({
            "id": drive_id,
            "plays": len(g),
            "start_d": g["D0"].iloc[0],
            # In tagged mode, plays before the first P&10 are a partial drive with no known start.
            "has_start": not (method == "tagged" and drive_id == 0),
            "reach_d": reach,
            "td": bool(g["IS_TD"].any()),
        })
    return pd.DataFrame(rows)


# ----------------------------------------------------------------------
# Aggregations
# ----------------------------------------------------------------------
def _mean(series):
    series = pd.to_numeric(series, errors="coerce").dropna()
    return float(series.mean()) if len(series) else None


def agg(sub, kind):
    """Core metrics for a slice of run plays or pass plays."""
    n = len(sub)
    if n == 0:
        return {"plays": 0}
    gain = pd.to_numeric(sub["GAIN"], errors="coerce")
    succ = sub["SUCC"]
    out = {
        "plays": n,
        "success": _mean(succ),
        "xpp": _mean(sub["VALUE"]),
        "iso": _mean(sub.loc[succ == 1, "VALUE"]),
    }
    if kind == "run":
        out["stuff"] = float((gain <= 0).mean())
        out["opp5"] = float((gain >= 5).mean())
        out["explosive"] = float((gain >= EXPLOSIVE_RUN_YDS).mean())
        out["ypc"] = _mean(gain)
    else:
        sacks = int(sub["IS_SACK"].sum())
        attempts = sub[~sub["IS_SACK"] & ~sub["IS_SCRAMBLE"]]
        comps = int(attempts["IS_COMPLETE"].sum())
        out["sack"] = sacks / n
        out["sacks"] = sacks
        out["attempts"] = len(attempts)
        out["comp_pct"] = comps / len(attempts) if len(attempts) else None
        out["ypa"] = _mean(pd.to_numeric(attempts["GAIN"], errors="coerce"))
        out["explosive"] = float((gain >= EXPLOSIVE_PASS_YDS).mean())
        out["int_rate"] = float(sub["IS_INT"].sum() / len(attempts)) if len(attempts) else None
        out["disruption"] = float((sub["IS_SACK"] | sub["IS_INT"] | sub["IS_FUMBLE"]).mean())
    return out


def split_rows(d, kind, by):
    """Rows for a by-down or by-quarter table (first row is always Overall)."""
    play_type = "Run" if kind == "run" else "Pass"
    base = d[d["PLAY TYPE"] == play_type]
    rows = [("Overall", agg(base, kind))]
    if by == "down":
        labels = [("P & 10", 0)] if (d["DN"] == 0).any() else []
        labels += [("1st", 1), ("2nd", 2), ("3rd", 3), ("4th", 4)]
        for label, dn in labels:
            rows.append((label, agg(base[base["DN"] == dn], kind)))
    else:
        qtr = pd.to_numeric(base["QTR"], errors="coerce") if "QTR" in base.columns else None
        if qtr is None or qtr.notna().sum() == 0:
            return None
        for q in sorted(qtr.dropna().unique()):
            label = f"Q{int(q)}" if q <= 4 else "OT"
            rows.append((label, agg(base[qtr == q], kind)))
    return rows


MAX_BREAK_SHARE = 0.05   # more section breaks than this share of plays = not play-by-play order


def section_breaks(d):
    """Rows where the play numbering restarts (a new game or playlist
    section starts there). Falls back to QTR if there is no PLAY #."""
    for col in ("PLAY #", "QTR"):
        if col in d.columns:
            v = pd.to_numeric(d[col], errors="coerce")
            if v.notna().sum() > 1:
                filled = v.ffill()
                return (filled < filled.shift(1)).fillna(False)
    return pd.Series(False, index=d.index)


def is_chronological(d, breaks=None):
    """True when the export is in game order, allowing a few section
    restarts (several games in one playlist). A clip playlist grouped by
    play type restarts constantly, so drive-based factors can't be built."""
    breaks = section_breaks(d) if breaks is None else breaks
    return bool(breaks.sum() <= MAX_BREAK_SHARE * max(len(d), 1))


# ----------------------------------------------------------------------
# The full five-factor summary
# ----------------------------------------------------------------------
def _scale(value, anchors):
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return None
    lo, hi = anchors
    return float(np.clip((value - lo) / (hi - lo), 0, 1))


def compute_factors(df):
    d = prepare(df)
    runs, passes = d[d["PLAY TYPE"] == "Run"], d[d["PLAY TYPE"] == "Pass"]
    n = len(d)
    run_stats, pass_stats = agg(runs, "run"), agg(passes, "pass")

    success = _mean(d["SUCC"])
    xpp = _mean(d["VALUE"])
    iso = _mean(d.loc[d["SUCC"] == 1, "VALUE"])
    explosive_all = float(((runs["GAIN"] >= EXPLOSIVE_RUN_YDS).sum() + (passes["GAIN"] >= EXPLOSIVE_PASS_YDS).sum()) / n) if n else None

    has_tags = bool((d["DN"] == 0).any())
    breaks = section_breaks(d)
    chrono = is_chronological(d, breaks)
    drives, method, asfp, opps, ppso, td_drives = None, "unavailable", None, None, None, None
    if has_tags:
        # P&10 tags give true drive-start field position even if rows are shuffled.
        asfp = float(d.loc[d["DN"] == 0, "D0"].mean())
        method = "tagged"
    if chrono:
        ids, method = assign_drives(d, breaks)
        drives = summarize_drives(d, ids, method)
        if not has_tags:
            asfp = float(drives["start_d"].mean()) if len(drives) else None
        opps = drives[drives["reach_d"] <= SCORING_OPP_LINE]
        td_drives = int(drives["td"].sum())
        ppso = (TD_POINTS * int(opps["td"].sum()) / len(opps)) if len(opps) else 0.0

    turnovers_eq = float(d["TO_EQ"].sum())
    to_rate = turnovers_eq / n if n else None

    ff = {
        "plays": n, "runs": len(runs), "passes": len(passes),
        "run_pct": len(runs) / n if n else None,
        "success": success, "run_success": run_stats.get("success"), "pass_success": pass_stats.get("success"),
        "xpp": xpp, "iso": iso, "explosive_rate": explosive_all,
        "drive_method": method, "chronological": chrono, "sections": int(breaks.sum()) + 1,
        "drives": len(drives) if drives is not None else None, "td_drives": td_drives,
        "asfp_d": asfp, "asfp_own": (100 - asfp) if asfp is not None else None,
        "scoring_opps": len(opps) if opps is not None else None,
        "scoring_opp_tds": int(opps["td"].sum()) if opps is not None else None, "ppso": ppso,
        "ints": int(d["IS_INT"].sum()), "fumbles": int(d["IS_FUMBLE"].sum()),
        "turnovers_eq": turnovers_eq, "to_rate": to_rate,
        "run": run_stats, "pass": pass_stats,
    }
    ff["grades"] = {
        "success": grade("success", success), "xpp": grade("xpp", xpp),
        "start_d": grade("start_d", asfp), "ppso": grade("ppso", ppso),
        "to_rate": grade("to_rate", to_rate),
        "run_success": grade("success", run_stats.get("success")),
        "pass_success": grade("success", pass_stats.get("success")),
        "run_xpp": grade("xpp", run_stats.get("xpp")), "pass_xpp": grade("xpp", pass_stats.get("xpp")),
        "stuff": grade("stuff", run_stats.get("stuff")), "sack": grade("sack", pass_stats.get("sack")),
    }
    ff["radar"] = {
        "Efficiency": _scale(success, RADAR_ANCHORS["efficiency"]),
        "Explosiveness": _scale(xpp, RADAR_ANCHORS["explosive"]),
        "Field Position": _scale(asfp, RADAR_ANCHORS["field_pos"]),
        "Finishing Drives": _scale(ppso, RADAR_ANCHORS["finishing"]),
        "Turnovers": _scale(to_rate, RADAR_ANCHORS["turnovers"]),
    }
    ff["_prepared"] = d
    return ff


# ----------------------------------------------------------------------
# Plain-English takeaways (written from the defense's point of view)
# ----------------------------------------------------------------------
def _pct(x, digits=0):
    return "n/a" if x is None or (isinstance(x, float) and np.isnan(x)) else f"{x * 100:.{digits}f}%"


def key_takeaways(ff):
    out = []
    rp = ff["run_pct"] or 0
    identity = "Run-first" if rp >= 0.58 else ("Pass-first" if rp <= 0.42 else "Balanced")
    out.append(f"<b>{identity} identity.</b> {_pct(ff['run_pct'])} of {ff['plays']} snaps are runs.")

    rs, ps = ff["run_success"], ff["pass_success"]
    if rs is not None and ps is not None:
        better = "running" if rs > ps else "throwing"
        out.append(f"<b>Efficiency: {ff['grades']['success'] or 'n/a'}.</b> {_pct(ff['success'])} of plays stay on schedule "
                   f"(run {_pct(rs)}, pass {_pct(ps)}) - they are more efficient {better} the ball.")

    out.append(f"<b>Explosiveness: {ff['grades']['xpp'] or 'n/a'}.</b> {ff['xpp']:.2f} expected points per play; "
               f"{_pct(ff['run'].get('explosive'))} of runs gain {EXPLOSIVE_RUN_YDS}+ and "
               f"{_pct(ff['pass'].get('explosive'))} of dropbacks gain {EXPLOSIVE_PASS_YDS}+.")

    soft = []
    if ff["run"].get("stuff") is not None:
        soft.append(f"{_pct(ff['run']['stuff'])} of runs are stuffed (0 yards or less)")
    if ff["pass"].get("sack") is not None:
        soft.append(f"{_pct(ff['pass']['sack'])} of dropbacks end in a sack")
    if soft:
        out.append("<b>Where to attack.</b> " + " and ".join(soft) + ".")

    if ff["asfp_own"] is not None and ff["drives"] is not None:
        est = " (drive starts estimated from the play sequence)" if ff["drive_method"] == "inferred" else ""
        out.append(f"<b>Field position and finishing.</b> Drives start near their own {ff['asfp_own']:.0f} on average{est}; "
                   f"{ff['scoring_opps']} of {ff['drives']} drives reached our {SCORING_OPP_LINE} or closer"
                   + (f", producing {ff['ppso']:.1f} points per opportunity." if ff["ppso"] is not None else "."))

    ints, fums = ff["ints"], ff["fumbles"]
    out.append(f"<b>Ball security.</b> {ints} interception{'s' if ints != 1 else ''} and "
               f"{fums} fumble{'s' if fums != 1 else ''} in {ff['plays']} plays.")
    return out
