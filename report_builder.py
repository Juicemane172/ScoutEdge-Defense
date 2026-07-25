"""Builds the DefensiveIQ-style Excel workbook from an analyzed dataframe."""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import pandas as pd
import numpy as np
from collections import Counter

from analyzer import (
    ZONE_ORDER, ZONE_LABEL, MIN_SAMPLE_CONCEPT, MIN_SAMPLE_FORMATION,
    SMALL_SAMPLE_FLAG, top_n_counts, run_pass_split, call_idea,
)

TITLE_FONT = Font(bold=True, size=13, color="FFFFFF", name="Arial")
TITLE_FILL = PatternFill("solid", fgColor="1F2937")
HEADER_FONT = Font(bold=True, color="FFFFFF", name="Arial", size=10)
HEADER_FILL = PatternFill("solid", fgColor="374151")
RUN_FILL = PatternFill("solid", fgColor="FBE1E1")
PASS_FILL = PatternFill("solid", fgColor="DCE9F9")
BASE_FONT = Font(name="Arial", size=10)


def _title(ws, text, span=10):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=span)
    c = ws.cell(row=1, column=1, value=text)
    c.font = TITLE_FONT
    c.fill = TITLE_FILL
    ws.row_dimensions[1].height = 22


def _header_row(ws, row, headers):
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=row, column=i, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")


def _autosize(ws, widths=None):
    if widths:
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w
        return
    for col in ws.columns:
        length = max((len(str(c.value)) if c.value is not None else 0) for c in col)
        letter = get_column_letter(col[0].column)
        ws.column_dimensions[letter].width = min(max(length + 2, 10), 32)


def _write_row(ws, r, values, fills=None):
    for i, v in enumerate(values, start=1):
        cell = ws.cell(row=r, column=i, value=v)
        cell.font = BASE_FONT
        if fills and i - 1 < len(fills) and fills[i - 1]:
            cell.fill = fills[i - 1]


# ------------------------------------------------------------------
# Sheet builders
# ------------------------------------------------------------------

def build_film_log(wb, df):
    ws = wb.create_sheet("1. Film Log")
    _title(ws, "FILM LOG  —  CONCEPT = play concept for ALL plays", span=13)
    _header_row(ws, 2, ["QTR", "DN", "DIST", "HASH", "YARD LN", "ZONE", "OFF FORM",
                        "OFF STR", "BACKFIELD", "PLAY DIR", "PLAY TYPE", "CONCEPT", "GN/LS"])
    r = 3
    for _, row in df.iterrows():
        _write_row(ws, r, [
            None, row.get("DN"), row.get("DIST"), row.get("HASH"), None,
            row.get("ZONE"), row.get("FORMATION"), None, row.get("BACKFIELD"),
            None, row.get("PLAY TYPE"), row.get("CONCEPT"), row.get("GAIN"),
        ])
        r += 1
    _autosize(ws)


def build_field_zone_tendencies(wb, df):
    ws = wb.create_sheet("2. Field Zone Tendencies")
    _title(ws, "FIELD ZONE TENDENCIES", span=10)
    ws.cell(row=2, column=1, value="  Gray=plays  Red%=Run  Blue%=Pass  Yellow=Call Idea").font = Font(italic=True, size=9)
    dd_cols = ["1st Down", "2nd & 7+", "2nd & 4-6", "2nd & 1-3", "3rd & 7+", "3rd & 4-6", "3rd & 1-3", "4th Down"]
    dd_map = {
        "1st Down": lambda s: s["DN"] == 1,
        "2nd & 7+": lambda s: (s["DN"] == 2) & (s["DIST"] >= 7),
        "2nd & 4-6": lambda s: (s["DN"] == 2) & (s["DIST"].between(4, 6)),
        "2nd & 1-3": lambda s: (s["DN"] == 2) & (s["DIST"].between(1, 3)),
        "3rd & 7+": lambda s: (s["DN"] == 3) & (s["DIST"] >= 7),
        "3rd & 4-6": lambda s: (s["DN"] == 3) & (s["DIST"].between(4, 6)),
        "3rd & 1-3": lambda s: (s["DN"] == 3) & (s["DIST"].between(1, 3)),
        "4th Down": lambda s: s["DN"] == 4,
    }
    _header_row(ws, 3, ["FIELD ZONE", "METRIC"] + dd_cols)
    r = 4
    for zone in ZONE_ORDER:
        zdf = df[df["ZONE"] == zone]
        ws.cell(row=r, column=1, value=f"  {zone}  ·  {ZONE_LABEL[zone]}  ({len(zdf)} plays)").font = Font(bold=True, size=10)
        r += 1
        plays_row, run_row, pass_row, runs_row, passes_row = [], [], [], [], []
        for col in dd_cols:
            mask = dd_map[col](zdf)
            sub = zdf[mask]
            runs, passes, total, run_pct, pass_pct = run_pass_split(sub)
            plays_row.append(total if total else 0)
            run_row.append(round(run_pct, 2) if run_pct is not None else None)
            pass_row.append(round(pass_pct, 2) if pass_pct is not None else None)
            runs_row.append(runs)
            passes_row.append(passes)
        _write_row(ws, r, [None, "Plays"] + plays_row); r += 1
        _write_row(ws, r, [None, "Run %"] + run_row, fills=[None, None] + [RUN_FILL] * len(dd_cols)); r += 1
        _write_row(ws, r, [None, "Pass %"] + pass_row, fills=[None, None] + [PASS_FILL] * len(dd_cols)); r += 1
        _write_row(ws, r, [None, "Runs"] + runs_row); r += 1
        _write_row(ws, r, [None, "Passes"] + passes_row); r += 1
        runs_z, passes_z, total_z, run_pct_z, pass_pct_z = run_pass_split(zdf)
        _write_row(ws, r, [None, "▶ Call Idea", call_idea(run_pct_z, pass_pct_z)])
        r += 2
    for col in ws.iter_cols(min_row=4, min_col=3, max_col=10):
        for cell in col:
            if isinstance(cell.value, (int, float)) and cell.row not in (4,):
                pass
    _autosize(ws, widths=[32, 10] + [11] * 8)
    # percent formatting
    for row in ws.iter_rows(min_row=4, min_col=3, max_col=10):
        for cell in row:
            if isinstance(cell.value, float):
                cell.number_format = "0%"


def build_run_pass_tendencies(wb, df, play_type, sheet_name, expl_yards, top_col3_name, top_col3_short, top_col3_source):
    ws = wb.create_sheet(sheet_name)
    _title(ws, f"{play_type.upper()} TENDENCIES  —  Auto-calculated from Hudl data", span=13)
    _header_row(ws, 2, ["ZONE", "COUNTS", "", "", "TOP FORMATIONS", "", "", f"TOP {play_type.upper()} CONCEPTS", "", "", top_col3_name, "", ""])
    _header_row(ws, 3, ["Zone", f"{play_type}\nCount", f"{play_type} %", "Explosive\n" + play_type + "s",
                        "#1 Formation", "#2 Formation", "#3 Formation",
                        "#1 Concept", "#2 Concept", "#3 Concept",
                        f"#1 {top_col3_short}", f"#2 {top_col3_short}", f"#3 {top_col3_short}"])
    r = 4
    for zone in ZONE_ORDER:
        zdf = df[(df["ZONE"] == zone) & (df["PLAY TYPE"] == play_type)]
        all_zdf = df[df["ZONE"] == zone]
        count = len(zdf)
        pct_ = count / len(all_zdf) if len(all_zdf) else 0
        expl = zdf["EXPLOSIVE"].sum()
        forms = top_n_counts(zdf["FORMATION"], 3)
        concepts = top_n_counts(zdf["CONCEPT"], 3)
        col3 = top_n_counts(zdf[top_col3_source], 3) if count else ["—", "—", "—"]
        _write_row(ws, r, [zone, count, round(pct_, 2), int(expl)] + forms + concepts + col3)
        r += 1
    for row in ws.iter_rows(min_row=4, min_col=3, max_col=3):
        for cell in row:
            cell.number_format = "0%"
    _autosize(ws)


def build_hash_tendencies(wb, df):
    ws = wb.create_sheet("5. Hash Tendencies")
    _title(ws, "HASH TENDENCIES  —  Left Hash · Middle · Right Hash", span=13)
    _header_row(ws, 2, ["FIELD ZONE", "L Plays", "L Run%", "L Pass%", "M Plays", "M Run%", "M Pass%",
                        "R Plays", "R Run%", "R Pass%", "Top L Concept", "Top M Concept", "Top R Concept"])

    def hash_block(sub):
        out = []
        for h in ["L", "M", "R"]:
            hsub = sub[sub["HASH"] == h]
            runs, passes, total, run_pct, pass_pct = run_pass_split(hsub)
            out += [total, round(run_pct, 2) if run_pct is not None else None,
                    round(pass_pct, 2) if pass_pct is not None else None]
        concepts = []
        for h in ["L", "M", "R"]:
            hsub = sub[sub["HASH"] == h]
            top = top_n_counts(hsub["CONCEPT"], 1)[0] if len(hsub) else "—"
            concepts.append(top)
        return out, concepts

    r = 3
    out, concepts = hash_block(df)
    _write_row(ws, r, ["OVERALL"] + out + concepts); r += 1
    for zone in ZONE_ORDER:
        zdf = df[df["ZONE"] == zone]
        out, concepts = hash_block(zdf)
        _write_row(ws, r, [f"{zone} — {ZONE_LABEL[zone]}"] + out + concepts)
        r += 1
    for row in ws.iter_rows(min_row=3, min_col=3, max_col=10):
        for cell in row:
            if isinstance(cell.value, (int, float)) and cell.column in (3, 4, 6, 7, 9, 10):
                cell.number_format = "0%"
    _autosize(ws, widths=[26] + [9] * 9 + [18, 18, 18])


DD_SITUATIONS = ["1ST & 10", "1ST & SHORT", "2ND & LONG", "2ND & MEDIUM", "2ND & SHORT",
                 "3RD & LONG", "3RD & MEDIUM", "3RD & SHORT", "4TH DOWN"]
ZONE_SITUATIONS = [("RED ZONE", "RZ"), ("GOAL LINE", "GL"), ("BACKED UP", "BZ")]


def build_down_distance(wb, df):
    ws = wb.create_sheet("6. Down & Distance")
    _title(ws, "DOWN & DISTANCE TENDENCIES  —  Favorite Runs, Passes & Formations by Situation", span=13)
    _header_row(ws, 2, ["SITUATION", "Plays", "Run%", "Pass%", "#1 Run Concept", "#2 Run Concept", "#3 Run Concept",
                        "#1 Pass Concept", "#2 Pass Concept", "#3 Pass Concept", "#1 Formation", "#2 Formation", "#3 Formation"])
    r = 3

    def write_situation(label, sub):
        nonlocal r
        runs, passes, total, run_pct, pass_pct = run_pass_split(sub)
        run_concepts = top_n_counts(sub[sub["PLAY TYPE"] == "Run"]["CONCEPT"], 3)
        pass_concepts = top_n_counts(sub[sub["PLAY TYPE"] == "Pass"]["CONCEPT"], 3)
        forms = top_n_counts(sub["FORMATION"], 3)
        _write_row(ws, r, [label, total, round(run_pct, 2) if run_pct is not None else None,
                           round(pass_pct, 2) if pass_pct is not None else None] + run_concepts + pass_concepts + forms)
        r += 1

    for sit in DD_SITUATIONS:
        write_situation(sit, df[df["DD_BUCKET"] == sit])
    for label, zone in ZONE_SITUATIONS:
        write_situation(label, df[df["ZONE"] == zone])
    for row in ws.iter_rows(min_row=3, min_col=3, max_col=4):
        for cell in row:
            cell.number_format = "0%"
    _autosize(ws)


def build_concepts(wb, df, play_type, sheet_name, expl_yards):
    ws = wb.create_sheet(sheet_name)
    _title(ws, f"{play_type.upper()} CONCEPTS  —  Explosive ({expl_yards}+) & Success Rate by Concept ({MIN_SAMPLE_CONCEPT}+ calls)", span=13)
    _header_row(ws, 2, [f"{play_type.upper()} CONCEPT", "Called", "Avg Yd", "Expl%", "Succ%", "Dir R%", "Dir L%",
                        "1st Dn Succ", "2nd Dn Succ", "3rd Dn Succ", "RedZone Succ", "Top Form", "Top Hash"])
    sub = df[df["PLAY TYPE"] == play_type]
    counts = sub["CONCEPT"].value_counts()
    r = 3
    for concept, called in counts.items():
        if called < MIN_SAMPLE_CONCEPT or pd.isna(concept):
            continue
        cdf = sub[sub["CONCEPT"] == concept]
        avg_yd = cdf["GAIN"].mean()
        expl_pct = cdf["EXPLOSIVE"].mean()
        succ_series = cdf["SUCCESS"].dropna()
        succ_pct = succ_series.mean() if len(succ_series) else None

        def succ_by_down(dn):
            d = cdf[cdf["DN"] == dn]["SUCCESS"].dropna()
            return round(d.mean(), 2) if len(d) else "—"

        rz = cdf[cdf["ZONE"] == "RZ"]["SUCCESS"].dropna()
        rz_succ = round(rz.mean(), 2) if len(rz) else "—"
        top_form = top_n_counts(cdf["FORMATION"], 1)[0].split(" (")[0]
        top_hash = cdf["HASH"].mode().iloc[0] if not cdf["HASH"].mode().empty else "—"
        label = f"{concept} *" if called < SMALL_SAMPLE_FLAG else concept
        _write_row(ws, r, [
            label, int(called), round(avg_yd, 1) if pd.notna(avg_yd) else None,
            round(expl_pct, 2), round(succ_pct, 2) if succ_pct is not None else None,
            None, None, succ_by_down(1), succ_by_down(2), succ_by_down(3), rz_succ, top_form, top_hash,
        ])
        r += 1
    ws.cell(row=r + 1, column=1, value=f"* = small sample (under {SMALL_SAMPLE_FLAG} calls) — read with caution").font = Font(italic=True, size=9)
    for row in ws.iter_rows(min_row=3, min_col=4, max_col=5):
        for cell in row:
            if isinstance(cell.value, float):
                cell.number_format = "0%"
    for row in ws.iter_rows(min_row=3, min_col=8, max_col=11):
        for cell in row:
            if isinstance(cell.value, float):
                cell.number_format = "0%"
    _autosize(ws)


def build_formation_tendencies(wb, df):
    ws = wb.create_sheet("9. Formation Tendencies")
    _title(ws, f"FORMATION TENDENCIES  —  Favorite Runs & Passes by Formation ({MIN_SAMPLE_FORMATION}+ snaps)", span=12)
    _header_row(ws, 2, ["FORMATION", "Snaps", "Run%", "Pass%", "Run R%", "Run L%",
                        "#1 Run Concept", "#2 Run Concept", "#1 Pass Concept", "#2 Pass Concept",
                        "Top Down/Dist", "Top Zone"])
    counts = df["FORMATION"].value_counts()
    r = 3
    for form, snaps in counts.items():
        if snaps < MIN_SAMPLE_FORMATION or pd.isna(form):
            continue
        fdf = df[df["FORMATION"] == form]
        runs, passes, total, run_pct, pass_pct = run_pass_split(fdf)
        run_concepts = top_n_counts(fdf[fdf["PLAY TYPE"] == "Run"]["CONCEPT"], 2)
        pass_concepts = top_n_counts(fdf[fdf["PLAY TYPE"] == "Pass"]["CONCEPT"], 2)
        top_dd = top_n_counts(fdf["DD_BUCKET"], 1)[0]
        top_zone_code = fdf["ZONE"].mode().iloc[0] if not fdf["ZONE"].mode().empty else None
        top_zone_count = (fdf["ZONE"] == top_zone_code).sum() if top_zone_code else 0
        top_zone = f"{ZONE_LABEL.get(top_zone_code,'—').split('·')[0].strip()} ({top_zone_count})" if top_zone_code else "—"
        _write_row(ws, r, [form, int(snaps), round(run_pct, 2), round(pass_pct, 2), None, None] +
                   run_concepts + pass_concepts + [top_dd, top_zone])
        r += 1
    for row in ws.iter_rows(min_row=3, min_col=3, max_col=4):
        for cell in row:
            if isinstance(cell.value, float):
                cell.number_format = "0%"
    _autosize(ws)


SITUATIONAL_ROWS = [
    ("P & 10 (Drive Start)", lambda df: df[df["DN"] == 0]),
    ("1ST DOWN", lambda df: df[df["DN"] == 1]),
    ("2ND & LONG", lambda df: df[df["DD_BUCKET"] == "2ND & LONG"]),
    ("2ND & MEDIUM", lambda df: df[df["DD_BUCKET"] == "2ND & MEDIUM"]),
    ("2ND & SHORT", lambda df: df[df["DD_BUCKET"] == "2ND & SHORT"]),
    ("3RD & LONG", lambda df: df[df["DD_BUCKET"] == "3RD & LONG"]),
    ("3RD & MEDIUM", lambda df: df[df["DD_BUCKET"] == "3RD & MEDIUM"]),
    ("3RD & SHORT", lambda df: df[df["DD_BUCKET"] == "3RD & SHORT"]),
    ("4TH DOWN", lambda df: df[df["DN"] == 4]),
    ("RED ZONE", lambda df: df[df["ZONE"] == "RZ"]),
    ("GOAL LINE", lambda df: df[df["ZONE"] == "GL"]),
    ("BACKED UP", lambda df: df[df["ZONE"] == "BZ"]),
    ("COMING OUT", lambda df: df[df["ZONE"].isin(["BZ", "OF"])]),
]


def build_situational_summary(wb, df):
    ws = wb.create_sheet("10. Situational Summary")
    _title(ws, "SITUATIONAL SUMMARY", span=10)
    _header_row(ws, 2, ["Situation", "Run\nCount", "Pass\nCount", "L Hash\nRun%", "M Hash\nRun%",
                        "R Hash\nRun%", "Top Run Concept", "Top Pass Concept", "Best Call", "Notes"])
    r = 3
    for label, fn in SITUATIONAL_ROWS:
        sub = fn(df)
        runs = (sub["PLAY TYPE"] == "Run").sum()
        passes = (sub["PLAY TYPE"] == "Pass").sum()

        def hash_run_pct(h):
            hsub = sub[sub["HASH"] == h]
            r_, p_, t_, rp, pp = run_pass_split(hsub)
            return round(rp, 2) if rp is not None else None

        top_run = top_n_counts(sub[sub["PLAY TYPE"] == "Run"]["CONCEPT"], 1)[0]
        top_pass = top_n_counts(sub[sub["PLAY TYPE"] == "Pass"]["CONCEPT"], 1)[0]
        _write_row(ws, r, [label, int(runs), int(passes), hash_run_pct("L"), hash_run_pct("M"), hash_run_pct("R"),
                           top_run, top_pass, None, None])
        r += 1
    # No live-clock data in a Hudl playlist export, so two-minute / must-have
    # situations can't be isolated - mirror overall totals as the original tool does.
    all_runs = (df["PLAY TYPE"] == "Run").sum()
    all_passes = (df["PLAY TYPE"] == "Pass").sum()
    top_run = top_n_counts(df[df["PLAY TYPE"] == "Run"]["CONCEPT"], 1)[0]
    top_pass = top_n_counts(df[df["PLAY TYPE"] == "Pass"]["CONCEPT"], 1)[0]
    for label in ["TWO-MINUTE", "MUST-HAVE"]:
        _write_row(ws, r, [label, int(all_runs), int(all_passes), None, None, None, top_run, top_pass, None,
                           "No clock data in export - shown as overall totals"])
        r += 1
    for row in ws.iter_rows(min_row=3, min_col=4, max_col=6):
        for cell in row:
            if isinstance(cell.value, float):
                cell.number_format = "0%"
    _autosize(ws)


def build_practice_scripts(wb, df):
    ws = wb.create_sheet("11. Practice Scripts")
    _title(ws, "PRACTICE SCRIPTS   ·   auto-pulled from film", span=8)
    ws.cell(row=3, column=1, value="MONDAY  ·  FAVORITES (ALL DOWNS)").font = Font(bold=True)
    ws.cell(row=4, column=1, value="FORMATION ALIGNMENT  (top 12)").font = Font(bold=True, italic=True)
    _header_row(ws, 5, ["#", "HASH", "FORMATION", "PERS", "FRONT", "COVERAGE"])
    counts = df["FORMATION"].value_counts().head(12)
    r = 6
    for i, (form, cnt) in enumerate(counts.items(), start=1):
        fdf = df[df["FORMATION"] == form]
        top_hash = fdf["HASH"].mode().iloc[0] if not fdf["HASH"].mode().empty else "—"
        _write_row(ws, r, [i, top_hash, form, None, None, None])
        r += 1
    r += 1
    ws.cell(row=r, column=1, value="INSIDE SCRIPT  (top run concepts, coach fills front/coverage)").font = Font(bold=True, italic=True)
    r += 1
    _header_row(ws, r, ["#", "HASH", "RUN CONCEPT", "FORMATION", "FRONT", "COVERAGE"])
    r += 1
    run_counts = df[df["PLAY TYPE"] == "Run"]["CONCEPT"].value_counts().head(12)
    for i, (concept, cnt) in enumerate(run_counts.items(), start=1):
        cdf = df[(df["PLAY TYPE"] == "Run") & (df["CONCEPT"] == concept)]
        top_hash = cdf["HASH"].mode().iloc[0] if not cdf["HASH"].mode().empty else "—"
        top_form = top_n_counts(cdf["FORMATION"], 1)[0].split(" (")[0]
        _write_row(ws, r, [i, top_hash, concept, top_form, None, None])
        r += 1
    r += 1
    ws.cell(row=r, column=1, value="SKELETON / PASS SCRIPT  (top pass concepts)").font = Font(bold=True, italic=True)
    r += 1
    _header_row(ws, r, ["#", "HASH", "PASS CONCEPT", "FORMATION", "FRONT", "COVERAGE"])
    r += 1
    pass_counts = df[df["PLAY TYPE"] == "Pass"]["CONCEPT"].value_counts().head(12)
    for i, (concept, cnt) in enumerate(pass_counts.items(), start=1):
        cdf = df[(df["PLAY TYPE"] == "Pass") & (df["CONCEPT"] == concept)]
        top_hash = cdf["HASH"].mode().iloc[0] if not cdf["HASH"].mode().empty else "—"
        top_form = top_n_counts(cdf["FORMATION"], 1)[0].split(" (")[0]
        _write_row(ws, r, [i, top_hash, concept, top_form, None, None])
        r += 1
    _autosize(ws)


CALL_SHEET_SITUATIONS = [
    "1st & 10", "1st & 10 (Own Half)", "1st & 10 (Opp Half)", "2nd & Long (8+)",
    "2nd & Medium (4-7)", "2nd & Short (1-3)", "3rd & Long (7+)", "3rd & Medium (4-6)",
    "3rd & Short (1-3)", "4th Down", "Red Zone — 1st", "Red Zone — 2nd", "Red Zone — 3rd",
    "Goal Line", "Backed Up", "Coming Out", "Two-Minute (Lead)", "Two-Minute (Trail)",
    "Must-Have Plays", "Two-Point Play", "Overtime",
]


def build_call_sheet_builder(wb):
    ws = wb.create_sheet("12. Call Sheet Builder")
    _title(ws, "CALL SHEET BUILDER  —  Fill in your calls", span=11)
    _header_row(ws, 2, ["Situation", "Field\nZone", "Down /\nDistance", "Opponent Tendency", "Formation / Alert",
                        "Best Pressure", "Best Coverage", "Best Front", "Adjustment", "Priority", "Notes"])
    r = 3
    for sit in CALL_SHEET_SITUATIONS:
        _write_row(ws, r, [sit] + [None] * 10)
        r += 1
    _autosize(ws)


def build_game_day_call_sheet(wb, df, opponent, week):
    ws = wb.create_sheet("13. Game Day Call Sheet")
    _title(ws, f"GAME DAY CALL SHEET   ·   {opponent}   ·   WK {week}", span=11)
    _header_row(ws, 3, ["Situation", "Run%", "Pass%", "Top Run", "Top Pass", "Form", "", "BIGGEST TENDENCIES", "", "", ""])

    dd_labels = [("1st & 10", "1ST & 10"), ("2nd & Short", "2ND & SHORT"), ("2nd & Med", "2ND & MEDIUM"),
                 ("2nd & Long", "2ND & LONG"), ("3rd & Short", "3RD & SHORT"), ("3rd & Med", "3RD & MEDIUM"),
                 ("3rd & Long", "3RD & LONG"), ("4th Down", "4TH DOWN")]
    r = 4
    for label, bucket in dd_labels:
        sub = df[df["DD_BUCKET"] == bucket]
        runs, passes, total, run_pct, pass_pct = run_pass_split(sub)
        top_run = top_n_counts(sub[sub["PLAY TYPE"] == "Run"]["CONCEPT"], 1)[0].split(" (")[0]
        top_pass = top_n_counts(sub[sub["PLAY TYPE"] == "Pass"]["CONCEPT"], 1)[0].split(" (")[0]
        top_form = top_n_counts(sub["FORMATION"], 1)[0].split(" (")[0]
        _write_row(ws, r, [label, round(run_pct, 2) if run_pct is not None else None,
                           round(pass_pct, 2) if pass_pct is not None else None, top_run, top_pass, top_form])
        r += 1

    r += 1
    ws.cell(row=r, column=1, value="FIELD ZONE").font = Font(bold=True)
    r += 1
    _header_row(ws, r, ["Zone", "Run%", "Pass%", "Top Run", "Top Pass", "Form"])
    r += 1
    zone_labels = [("Coming Out", ["BZ"]), ("Open Field", ["OF"]), ("Midfield", ["MF"]),
                   ("Fringe", ["FZ"]), ("Red Zone", ["RZ"]), ("Goal Line", ["GL"])]
    for label, zones in zone_labels:
        sub = df[df["ZONE"].isin(zones)]
        runs, passes, total, run_pct, pass_pct = run_pass_split(sub)
        top_run = top_n_counts(sub[sub["PLAY TYPE"] == "Run"]["CONCEPT"], 1)[0].split(" (")[0]
        top_pass = top_n_counts(sub[sub["PLAY TYPE"] == "Pass"]["CONCEPT"], 1)[0].split(" (")[0]
        top_form = top_n_counts(sub["FORMATION"], 1)[0].split(" (")[0]
        _write_row(ws, r, [label, round(run_pct, 2) if run_pct is not None else None,
                           round(pass_pct, 2) if pass_pct is not None else None, top_run, top_pass, top_form])
        r += 1

    # Right column: biggest tendencies (formations with most lopsided run/pass split, min sample)
    tendencies = []
    for form, cnt in df["FORMATION"].value_counts().items():
        if cnt < MIN_SAMPLE_FORMATION or pd.isna(form):
            continue
        fdf = df[df["FORMATION"] == form]
        runs, passes, total, run_pct, pass_pct = run_pass_split(fdf)
        skew = max(run_pct, pass_pct)
        tag = "Run" if run_pct >= pass_pct else "Pass"
        flag = " *" if cnt < SMALL_SAMPLE_FLAG else ""
        tendencies.append((skew, f"{form} = {round(skew*100)}% {tag}{flag}   [{cnt}]"))
    tendencies.sort(reverse=True)
    _write_row(ws, 4, [None] * 7 + ["#1", tendencies[0][1] if tendencies else "—"])
    for i in range(1, min(5, len(tendencies))):
        _write_row(ws, 4 + i, [None] * 7 + [f"#{i+1}", tendencies[i][1]])

    # Heavy pass situations
    heavy_row = 4 + 8
    ws.cell(row=heavy_row, column=8, value="HEAVY PASS SITUATIONS").font = Font(bold=True)
    heavy = []
    for label, fn in [("1st & 10", lambda d: d[d["DD_BUCKET"] == "1ST & 10"]),
                       ("2nd & Long", lambda d: d[d["DD_BUCKET"] == "2ND & LONG"]),
                       ("3rd & Long", lambda d: d[d["DD_BUCKET"] == "3RD & LONG"]),
                       ("3rd & Medium", lambda d: d[d["DD_BUCKET"] == "3RD & MEDIUM"]),
                       ("3rd & Short", lambda d: d[d["DD_BUCKET"] == "3RD & SHORT"]),
                       ("Backed Up", lambda d: d[d["ZONE"] == "BZ"])]:
        sub = fn(df)
        runs, passes, total, run_pct, pass_pct = run_pass_split(sub)
        if pass_pct is not None and total:
            heavy.append((pass_pct, f"{label}: {round(pass_pct*100)}% Pass" + (" *" if total < SMALL_SAMPLE_FLAG else "")))
    heavy.sort(reverse=True)
    for i, (pp, text) in enumerate(heavy[:5], start=1):
        _write_row(ws, heavy_row + i, [None] * 7 + [f"#{i}", text])

    # Red zone cheat sheet
    rz_row = heavy_row + 7
    ws.cell(row=rz_row, column=8, value="RED ZONE CHEAT SHEET").font = Font(bold=True)
    rzdf = df[df["ZONE"] == "RZ"]
    runs, passes, total, run_pct, pass_pct = run_pass_split(rzdf)
    top_run = top_n_counts(rzdf[rzdf["PLAY TYPE"] == "Run"]["CONCEPT"], 1)[0].split(" (")[0]
    top_pass = top_n_counts(rzdf[rzdf["PLAY TYPE"] == "Pass"]["CONCEPT"], 1)[0].split(" (")[0]
    top_form = top_n_counts(rzdf["FORMATION"], 1)[0].split(" (")[0]
    gldf = df[df["ZONE"] == "GL"]
    gl_top = top_n_counts(gldf["CONCEPT"], 1)[0] if len(gldf) else "—"
    rows = [
        ("Inside 20", f"Run {round((run_pct or 0)*100)}% / Pass {round((pass_pct or 0)*100)}%  ({total} plays)"),
        ("Top Run", top_run),
        ("Top Pass", top_pass),
        ("Top Formation", top_form),
        ("Goal Line Favorite", gl_top),
    ]
    for i, (label, val) in enumerate(rows, start=1):
        _write_row(ws, rz_row + i, [None] * 7 + [label, val])

    for row in ws.iter_rows(min_row=4, min_col=2, max_col=3):
        for cell in row:
            if isinstance(cell.value, float):
                cell.number_format = "0%"

    footer_row = rz_row + 8
    ws.cell(row=footer_row, column=1,
            value="* = small sample (under 5 plays). Auto-generated from uploaded film.").font = Font(italic=True, size=9)
    _autosize(ws, widths=[18, 9, 9, 16, 16, 16, 3, 16, 30, 6, 6])


def build_workbook(df, opponent="Opponent", week="1"):
    wb = Workbook()
    wb.remove(wb.active)
    build_film_log(wb, df)
    build_field_zone_tendencies(wb, df)
    build_run_pass_tendencies(wb, df, "Run", "3. Run Tendencies", 10, "TOP HASH (no true direction data in export)", "Hash", "HASH")
    build_run_pass_tendencies(wb, df, "Pass", "4. Pass Tendencies", 15, "TOP BACKFIELD", "Backfield", "BACKFIELD")
    build_hash_tendencies(wb, df)
    build_down_distance(wb, df)
    build_concepts(wb, df, "Run", "7. Run Concepts", 10)
    build_concepts(wb, df, "Pass", "8. Pass Concepts", 15)
    build_formation_tendencies(wb, df)
    build_situational_summary(wb, df)
    build_practice_scripts(wb, df)
    build_call_sheet_builder(wb)
    build_game_day_call_sheet(wb, df, opponent, week)
    return wb
