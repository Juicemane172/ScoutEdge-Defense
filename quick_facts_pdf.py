"""
Quick Facts PDF - a short, printable scouting summary built around the five
factors (efficiency, explosiveness, field position, finishing drives,
turnovers), with run-game and pass-game pages split by down and quarter.

Grade chips are colored by THREAT TO OUR DEFENSE: red = strong offensive
trait, amber = average, green = a soft spot we can attack.
"""

import datetime
import math

from reportlab.graphics.shapes import Circle, Drawing, Line, Polygon, String
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import (BaseDocTemplate, Frame, KeepTogether, PageBreak, PageTemplate,
                                Paragraph, Spacer, Table, TableStyle)

import five_factors as F
from analyzer import EXPLOSIVE_PASS_YDS, EXPLOSIVE_RUN_YDS, SMALL_SAMPLE_FLAG, SUCCESS_1ST, SUCCESS_2ND

NAVY = colors.HexColor("#16213E")
RED = colors.HexColor("#C0392B")
BLUE = colors.HexColor("#1A5276")
TEAL = colors.HexColor("#0E7060")
LIGHT = colors.HexColor("#F5F5F5")
RUN_TINT = colors.HexColor("#FDE8E8")
PASS_TINT = colors.HexColor("#E8F0FE")
INK = colors.HexColor("#1F2937")
MUTED = colors.HexColor("#6B7280")
RULE = colors.HexColor("#D1D5DB")

# (background, text) per grade - threat-to-defense coloring.
CHIP = {
    "ELITE": (colors.HexColor("#C0392B"), colors.white),
    "STRONG": (colors.HexColor("#E57368"), INK),
    "AVERAGE": (colors.HexColor("#F4D03F"), INK),
    "BELOW AVG": (colors.HexColor("#8FCDB0"), INK),
    "POOR": (colors.HexColor("#0E7060"), colors.white),
    None: (colors.HexColor("#D1D5DB"), INK),
}

PAGE_W, PAGE_H = letter
MARGIN = 0.6 * inch
CONTENT_W = PAGE_W - 2 * MARGIN


def _styles():
    base = dict(fontName="Helvetica", textColor=INK, leading=11.5, fontSize=8.8)
    return {
        "title": ParagraphStyle("title", fontName="Helvetica-Bold", fontSize=25, leading=28, textColor=colors.white),
        "meta": ParagraphStyle("meta", fontName="Helvetica", fontSize=10, leading=13, textColor=colors.HexColor("#CADCFC")),
        "h1": ParagraphStyle("h1", fontName="Helvetica-Bold", fontSize=17, leading=20, textColor=INK),
        "h2": ParagraphStyle("h2", fontName="Helvetica-Bold", fontSize=10, leading=13, textColor=NAVY, spaceBefore=8, spaceAfter=4),
        "body": ParagraphStyle("body", **base),
        "bullet": ParagraphStyle("bullet", **{**base, "leftIndent": 10, "bulletIndent": 0, "spaceAfter": 4}),
        "small": ParagraphStyle("small", **{**base, "fontSize": 7.4, "leading": 9.5, "textColor": MUTED}),
        "card_label": ParagraphStyle("card_label", fontName="Helvetica-Bold", fontSize=7.2, leading=9, textColor=MUTED),
        "card_value": ParagraphStyle("card_value", fontName="Helvetica-Bold", fontSize=21, leading=25, textColor=INK),
        "card_sub": ParagraphStyle("card_sub", fontName="Helvetica", fontSize=7.2, leading=9, textColor=MUTED),
        "chip": ParagraphStyle("chip", fontName="Helvetica-Bold", fontSize=7.2, leading=9, alignment=TA_CENTER),
        "th": ParagraphStyle("th", fontName="Helvetica-Bold", fontSize=8, leading=10, textColor=colors.white),
        "td": ParagraphStyle("td", fontName="Helvetica", fontSize=8.4, leading=10.5, textColor=INK),
        "tdb": ParagraphStyle("tdb", fontName="Helvetica-Bold", fontSize=8.4, leading=10.5, textColor=INK),
        "gloss_term": ParagraphStyle("gloss_term", fontName="Helvetica-Bold", fontSize=8.8, leading=11.5, textColor=NAVY),
    }


# ----------------------------------------------------------------------
# formatting helpers
# ----------------------------------------------------------------------
def _none(v):
    return v is None or (isinstance(v, float) and math.isnan(v))


def pct(v, d=1):
    return "—" if _none(v) else f"{v * 100:.{d}f}%"


def num(v, d=2):
    return "—" if _none(v) else f"{v:.{d}f}"


def _chip(label, st, width=0.95 * inch):
    bg, fg = CHIP.get(label, CHIP[None])
    p = Paragraph(label or "N/A", ParagraphStyle("chip_i", parent=st["chip"], textColor=fg))
    t = Table([[p]], colWidths=[width])
    t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), bg), ("TOPPADDING", (0, 0), (-1, -1), 2),
                           ("BOTTOMPADDING", (0, 0), (-1, -1), 3), ("LEFTPADDING", (0, 0), (-1, -1), 2),
                           ("RIGHTPADDING", (0, 0), (-1, -1), 2)]))
    t.hAlign = "LEFT"
    return t


# ----------------------------------------------------------------------
# page chrome
# ----------------------------------------------------------------------
def _page_chrome(opponent, week):
    def draw(canvas, doc):
        canvas.saveState()
        canvas.setFillColor(NAVY)
        canvas.rect(0, PAGE_H - 0.34 * inch, PAGE_W, 0.34 * inch, stroke=0, fill=1)
        canvas.setFillColor(colors.white)
        canvas.setFont("Helvetica-Bold", 8.5)
        canvas.drawString(MARGIN, PAGE_H - 0.22 * inch, "SCOUTEDGE DEFENSE")
        canvas.setFont("Helvetica", 8.5)
        canvas.drawRightString(PAGE_W - MARGIN, PAGE_H - 0.22 * inch, "QUICK FACTS")
        canvas.setStrokeColor(RULE)
        canvas.setLineWidth(0.5)
        canvas.line(MARGIN, 0.5 * inch, PAGE_W - MARGIN, 0.5 * inch)
        canvas.setFillColor(MUTED)
        canvas.setFont("Helvetica", 7.5)
        canvas.drawString(MARGIN, 0.35 * inch, f"{opponent}  ·  Week {week}  ·  built from one offense's film - no opponent or league comparison")
        canvas.drawRightString(PAGE_W - MARGIN, 0.35 * inch, f"Page {doc.page}")
        canvas.restoreState()
    return draw


# ----------------------------------------------------------------------
# radar chart
# ----------------------------------------------------------------------
def radar_drawing(radar, labels, st):
    items = [(k, v) for k, v in radar.items() if v is not None]
    if len(items) < 3:
        return Paragraph("Not enough measurable factors in this export to draw the radar chart.", st["small"])

    w, h = 285, 250
    cx, cy, R = w / 2, h / 2 - 2, 72
    n = len(items)
    angles = [math.pi / 2 - 2 * math.pi * i / n for i in range(n)]

    def pt(frac, a):
        return cx + R * frac * math.cos(a), cy + R * frac * math.sin(a)

    d = Drawing(w, h)
    for ring in (0.25, 0.5, 0.75, 1.0):
        pts = [c for a in angles for c in pt(ring, a)]
        d.add(Polygon(pts, strokeColor=colors.HexColor("#CBD5E1"), strokeWidth=0.6, fillColor=None))
    for a in angles:
        x, y = pt(1.0, a)
        d.add(Line(cx, cy, x, y, strokeColor=colors.HexColor("#CBD5E1"), strokeWidth=0.5))

    # benchmark: a typical offense sits on the middle ring
    ref = [c for a in angles for c in pt(0.5, a)]
    d.add(Polygon(ref, strokeColor=colors.HexColor("#6B7280"), strokeWidth=1.1, strokeDashArray=[3, 2], fillColor=None))

    team = [c for (k, v), a in zip(items, angles) for c in pt(max(v, 0.03), a)]
    d.add(Polygon(team, strokeColor=BLUE, strokeWidth=1.8, fillColor=colors.Color(0.10, 0.32, 0.46, alpha=0.28)))
    for (k, v), a in zip(items, angles):
        x, y = pt(max(v, 0.03), a)
        d.add(Circle(x, y, 2.6, fillColor=BLUE, strokeColor=colors.white, strokeWidth=0.8))

    for (k, v), a in zip(items, angles):
        lx, ly = pt(1.0, a)
        lx, ly = lx + 14 * math.cos(a), ly + 14 * math.sin(a)
        cosv, sinv = math.cos(a), math.sin(a)
        anchor = "middle" if abs(cosv) < 0.25 else ("start" if cosv > 0 else "end")
        dy = 6 if sinv > 0.6 else (-14 if sinv < -0.6 else -4)
        d.add(String(lx, ly + dy + 9, k, fontName="Helvetica-Bold", fontSize=8, fillColor=INK, textAnchor=anchor))
        d.add(String(lx, ly + dy, labels.get(k, ""), fontName="Helvetica", fontSize=7.4, fillColor=MUTED, textAnchor=anchor))
    return d


# ----------------------------------------------------------------------
# tables
# ----------------------------------------------------------------------
def _table(header, rows, header_color, col_widths, st, chip_cols=None, first_col_bold=True):
    """rows: list of lists of strings. chip_cols: {col_index: [grade per row]} colors that cell."""
    data = [[Paragraph(h, st["th"]) for h in header]]
    for r in rows:
        data.append([Paragraph(str(c), st["tdb"] if (i == 0 and first_col_bold) else st["td"]) for i, c in enumerate(r)])
    t = Table(data, colWidths=col_widths, repeatRows=1)
    style = [
        ("BACKGROUND", (0, 0), (-1, 0), header_color),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING", (0, 0), (-1, -1), 6), ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("LINEBELOW", (0, 1), (-1, -1), 0.4, RULE),
    ]
    for i in range(1, len(data)):
        if i % 2 == 0:
            style.append(("BACKGROUND", (0, i), (-1, i), LIGHT))
    if chip_cols:
        for col, grades in chip_cols.items():
            for i, g in enumerate(grades, start=1):
                if g:
                    bg, fg = CHIP[g]
                    style.append(("BACKGROUND", (col, i), (col, i), bg))
                    cell = data[i][col]
                    data[i][col] = Paragraph(f"<b>{cell.text}</b>", ParagraphStyle("c", parent=st["td"], textColor=fg, alignment=TA_CENTER))
    t.setStyle(TableStyle(style))
    return t


def _section_title(text, accent, st):
    t = Table([[Paragraph(text, st["h1"])]], colWidths=[CONTENT_W])
    t.setStyle(TableStyle([("LINEBELOW", (0, 0), (-1, -1), 2.2, accent), ("LEFTPADDING", (0, 0), (-1, -1), 0),
                           ("BOTTOMPADDING", (0, 0), (-1, -1), 5), ("TOPPADDING", (0, 0), (-1, -1), 0)]))
    return t


# ----------------------------------------------------------------------
# pages
# ----------------------------------------------------------------------
def _factor_cards(ff, st):
    g = ff["grades"]
    drives_ok = ff["drives"] is not None
    asfp_own = ff["asfp_own"]
    cards = [
        ("EFFICIENCY", pct(ff["success"]), "Success rate", g["success"]),
        ("EXPLOSIVENESS", num(ff["xpp"]), f"xPP  ·  IsoPPP {num(ff['iso'])}", g["xpp"]),
        ("FIELD POSITION", "—" if _none(asfp_own) else f"Own {asfp_own:.0f}",
         ("Avg drive start" + (" (est.)" if ff["drive_method"] == "inferred" else "")) if not _none(asfp_own)
         else "Needs P&10 tags or game-order film", g["start_d"]),
        ("FINISHING DRIVES", "—" if ff["ppso"] is None else num(ff["ppso"], 1),
         f"pts / scoring opp  ·  {ff['scoring_opps']} of {ff['drives']} drives" if drives_ok
         else "Needs P&10 tags or game-order film", g["ppso"]),
        ("TURNOVERS", str(ff["ints"] + ff["fumbles"]), f"{ff['ints']} INT  ·  {ff['fumbles']} FUM", g["to_rate"]),
    ]
    cells = []
    for label, value, sub, grade in cards:
        cells.append([Paragraph(label, st["card_label"]), Paragraph(value, st["card_value"]),
                      _chip(grade, st), Spacer(1, 4), Paragraph(sub, st["card_sub"])])
    w = CONTENT_W / 5
    t = Table([cells], colWidths=[w] * 5)
    t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), LIGHT), ("VALIGN", (0, 0), (-1, -1), "TOP"),
                           ("LINEAFTER", (0, 0), (-2, -1), 6, colors.white),
                           ("TOPPADDING", (0, 0), (-1, -1), 8), ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
                           ("LEFTPADDING", (0, 0), (-1, -1), 8), ("RIGHTPADDING", (0, 0), (-1, -1), 4)]))
    return t


def _snapshot_page(ff, opponent, week, game_date, tendencies, st):
    story = []
    date_str = game_date.strftime("%B %d, %Y") if hasattr(game_date, "strftime") else str(game_date or "")
    meta = f"Week {week}  ·  {date_str}  ·  {ff['plays']} plays analyzed"
    if ff["drives"] is not None:
        meta += f"  ·  {ff['drives']} drives" + (" (est.)" if ff["drive_method"] == "inferred" else "")
    if ff.get("sections", 1) > 1 and ff["chronological"]:
        meta += f"  ·  {ff['sections']} games/sections"
    banner = Table([[[Paragraph(opponent.upper(), st["title"]), Spacer(1, 3), Paragraph(meta, st["meta"])]]], colWidths=[CONTENT_W])
    banner.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), NAVY), ("LEFTPADDING", (0, 0), (-1, -1), 16),
                                ("TOPPADDING", (0, 0), (-1, -1), 14), ("BOTTOMPADDING", (0, 0), (-1, -1), 14)]))
    story += [banner, Spacer(1, 12), Paragraph("THE FIVE FACTORS", st["h2"]), _factor_cards(ff, st), Spacer(1, 6)]

    labels = {
        "Efficiency": pct(ff["success"]), "Explosiveness": f"{num(ff['xpp'])} xPP",
        "Field Position": "—" if _none(ff["asfp_own"]) else f"Own {ff['asfp_own']:.0f}",
        "Finishing Drives": "—" if ff["ppso"] is None else f"{ff['ppso']:.1f} pts/opp",
        "Turnovers": f"{ff['ints'] + ff['fumbles']} total",
    }
    radar = radar_drawing(ff["radar"], labels, st)
    takeaways = [Paragraph("WHAT THIS MEANS FOR OUR DEFENSE", st["h2"])]
    takeaways += [Paragraph(t, st["bullet"], bulletText="•") for t in F.key_takeaways(ff)]
    body = Table([[radar, takeaways]], colWidths=[3.95 * inch, CONTENT_W - 3.95 * inch])
    body.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0),
                              ("RIGHTPADDING", (0, 0), (-1, -1), 0)]))
    story += [Paragraph("PROFILE", st["h2"]), body, Spacer(1, 4)]

    if tendencies:
        story += [Paragraph("BIGGEST SITUATIONAL TENDENCIES", st["h2"])]
        rows = [[t] for t in tendencies]
        tbl = _table(["Tendency  (plays in sample)"], rows, NAVY, [CONTENT_W], st, first_col_bold=False)
        story.append(tbl)
    story += [Spacer(1, 8), Paragraph(
        "<b>How to read this page.</b> Chip color shows the threat to our defense: "
        "<font color='#C0392B'><b>red</b></font> = strong offensive trait, "
        "<font color='#B7950B'><b>amber</b></font> = average, "
        "<font color='#0E7060'><b>green</b></font> = a soft spot to attack. "
        "On the radar, the shaded shape is this offense and the dashed ring is a typical-offense benchmark "
        "(an adjustable assumption, not opponent data). Grades are rule-of-thumb bands - see the last page.", st["small"])]
    return story


def _run_rows(ff, kind):
    r, g = ff[kind], ff["grades"]
    if kind == "run":
        return [
            ["Success rate", pct(r.get("success")), g["run_success"], "Runs that gain their share of yards-to-go"],
            ["Expected points per play (xPP)", num(r.get("xpp")), g["run_xpp"], "How much each run moves the scoreboard"],
            ["IsoPPP (successful runs only)", num(r.get("iso")), None, "How big the wins are when they win"],
            ["Stuff rate", pct(r.get("stuff")), g["stuff"], "Runs for 0 or fewer yards (lower is better for them)"],
            ["Opportunity rate (5+ yds)", pct(r.get("opp5")), None, "Runs that reach a workable 5 yards"],
            [f"Explosive runs ({EXPLOSIVE_RUN_YDS}+ yds)", pct(r.get("explosive")), None, "Big-play share of carries"],
            ["Yards per carry", num(r.get("ypc")), None, "Average gain"],
        ]
    return [
        ["Success rate", pct(r.get("success")), g["pass_success"], "Dropbacks that gain their share of yards-to-go"],
        ["Expected points per play (xPP)", num(r.get("xpp")), g["pass_xpp"], "How much each dropback moves the scoreboard"],
        ["IsoPPP (successful passes only)", num(r.get("iso")), None, "How big the wins are when they win"],
        ["Sack rate", pct(r.get("sack")), g["sack"], "Sacks per dropback (lower is better for them)"],
        ["Disruption rate", pct(r.get("disruption")), None, "Sacks + interceptions + fumbles per dropback"],
        ["Completion %", pct(r.get("comp_pct")), None, f"{r.get('attempts', 0)} attempts (sacks and scrambles excluded)"],
        ["Yards per attempt", num(r.get("ypa")), None, "Average gain on attempts"],
        [f"Explosive passes ({EXPLOSIVE_PASS_YDS}+ yds)", pct(r.get("explosive")), None, "Big-play share of dropbacks"],
        ["Interception rate", pct(r.get("int_rate")), None, "Interceptions per attempt"],
    ]


def _game_page(ff, d, kind, st):
    is_run = kind == "run"
    accent, tint = (RED, RUN_TINT) if is_run else (BLUE, PASS_TINT)
    r = ff[kind]
    story = [_section_title("Run Game Summary" if is_run else "Passing Game Summary", accent, st), Spacer(1, 8)]
    if not r.get("plays"):
        return story + [Paragraph("No plays of this type in the export.", st["body"])]

    share = (ff["runs"] if is_run else ff["passes"]) / ff["plays"]
    noun, unit = ("runs", "carry") if is_run else ("dropbacks", "attempt")
    story.append(Paragraph(
        f"{r['plays']} {noun} ({share * 100:.0f}% of snaps). Success rate {pct(r.get('success'))} "
        f"and {num(r.get('xpp'))} expected points per play. Chip color shows threat to our defense.", st["body"]))
    story += [Spacer(1, 8), Paragraph(f"{'RUN' if is_run else 'PASS'} GAME SNAPSHOT", st["h2"])]

    rows = _run_rows(ff, kind)
    story.append(_table(["Metric", "Value", "Grade", "What it tells us"],
                        [[a, b, (c or ""), e] for a, b, c, e in rows], accent,
                        [2.45 * inch, 0.8 * inch, 0.95 * inch, CONTENT_W - 4.2 * inch], st,
                        chip_cols={2: [c for _, _, c, _ in rows]}))

    if is_run:
        header = ["Down", "Plays", "Success %", "xPP", "IsoPPP", "Stuff %", "5+ yd %"]
    else:
        header = ["Down", "Plays", "Success %", "xPP", "IsoPPP", "Sack %", f"{EXPLOSIVE_PASS_YDS}+ yd %"]

    def split_table(rows_):
        out = []
        for label, a in rows_:
            if not a.get("plays"):
                out.append([label, "0", "—", "—", "—", "—", "—"])
                continue
            n_txt = f"{a['plays']}*" if a["plays"] < SMALL_SAMPLE_FLAG else str(a["plays"])
            if is_run:
                out.append([label, n_txt, pct(a["success"]), num(a["xpp"]), num(a["iso"]), pct(a["stuff"]), pct(a["opp5"])])
            else:
                out.append([label, n_txt, pct(a["success"]), num(a["xpp"]), num(a["iso"]), pct(a["sack"]), pct(a["explosive"])])
        return out

    widths = [1.25 * inch] + [(CONTENT_W - 1.25 * inch) / 6] * 6
    down_rows = F.split_rows(d, kind, "down")
    story += [Spacer(1, 6), Paragraph("BY DOWN", st["h2"]),
              _table(["Down"] + header[1:], split_table(down_rows), accent, widths, st)]

    q_rows = F.split_rows(d, kind, "qtr")
    story += [Spacer(1, 4), Paragraph("BY QUARTER", st["h2"])]
    if q_rows:
        story.append(_table(["Quarter"] + header[1:], split_table(q_rows), accent, widths, st))
    else:
        story.append(Paragraph("Quarter splits need a QTR column in the Hudl export - this file doesn't have one.", st["small"]))
    story += [Spacer(1, 6), Paragraph(f"* = fewer than {SMALL_SAMPLE_FLAG} plays - read with caution.", st["small"])]
    return story


def _glossary_page(st):
    story = [_section_title("Glossary and Method", NAVY, st), Spacer(1, 8)]
    terms = [
        ("Success rate (efficiency)",
         f"A play is a success if it gains {SUCCESS_1ST * 100:.0f}% of the yards needed on 1st down (P&10 counts as 1st), "
         f"{SUCCESS_2ND * 100:.0f}% on 2nd down, and all of them on 3rd or 4th down. The share of plays that succeed."),
        ("Expected points per play, xPP (explosiveness)",
         "Every spot on the field has an expected-points value. A play's value is how much it changed that number, so a "
         "10-yard gain inside the 20 is worth more than 10 yards at midfield. xPP is the average over all plays."),
        ("IsoPPP", "The same play value, averaged over successful plays only: how big their wins are when they win."),
        ("Stuff rate", "Runs for zero or negative yards divided by all runs."),
        ("Opportunity rate", "Runs gaining 5 or more yards divided by all runs."),
        ("Sack rate", "Sacks divided by dropbacks (every pass play, including scrambles)."),
        ("Disruption rate", "Sacks plus interceptions plus fumbles, divided by dropbacks."),
        ("Average starting field position",
         "Where drives begin, shown as the offense's own yard line. Needs P&10 tags (DN 0) or plays in game order."),
        ("Scoring opportunity and points per opportunity",
         "A drive that reaches the opponent's 40 or closer. Points per opportunity = touchdowns x 6 divided by opportunities."),
        ("Turnovers", "Interceptions plus fumbles. Fumbles count as half a turnover each when judging ball security, "
                      "because the export doesn't say who recovered."),
    ]
    for term, text in terms:
        story += [Paragraph(term, st["gloss_term"]), Paragraph(text, st["body"]), Spacer(1, 4)]

    story += [Spacer(1, 4), Paragraph("METHOD NOTES AND ASSUMPTIONS", st["h2"])]
    notes = [
        "The expected-points curve is an approximation set in five_factors.py (EP_VALUE). Edit it there to reshape every xPP number.",
        "Grades are rule-of-thumb bands (GRADE_BANDS in five_factors.py), not league percentiles. Treat them as a quick read, then check the film.",
        "The radar's dashed ring is a typical-offense benchmark built from those same assumptions. It is not data from other teams.",
        "Touchdowns count as 6 points. The export has no PAT or field-goal data, so those points aren't in finishing-drive numbers.",
        "Drive starts come from your P&10 tags when present. With no tags and plays in game order, drive breaks are estimated "
        "(in testing against a tagged file, 15 of 17 drive starts were recovered with no false ones). Clip playlists that are out of "
        "game order show Field Position and Finishing Drives as n/a.",
        "Penalty plays and non-run, non-pass plays are excluded. Quarter splits appear only when the export has a QTR column.",
        "Everything here describes one offense's film. There is no opponent comparison, expected margin, or league average.",
    ]
    story += [Paragraph(n, st["bullet"], bulletText="•") for n in notes]
    return story


# ----------------------------------------------------------------------
# entry point
# ----------------------------------------------------------------------
def build_quick_facts_pdf(df, opponent="Opponent", week="1", game_date=None, out="quick_facts.pdf"):
    """Build the PDF. `out` can be a file path or a file-like object (e.g. BytesIO)."""
    st = _styles()
    ff = F.compute_factors(df)
    d = ff["_prepared"]

    try:  # reuse the same tendency finder the PowerPoint uses
        from pptx_builder import biggest_tendencies
        tendencies = biggest_tendencies(df)
    except Exception:
        tendencies = []

    doc = BaseDocTemplate(out, pagesize=letter, leftMargin=MARGIN, rightMargin=MARGIN,
                          topMargin=0.62 * inch, bottomMargin=0.65 * inch,
                          title=f"ScoutEdge Defense Quick Facts - {opponent}", author="ScoutEdge Defense")
    frame = Frame(MARGIN, 0.65 * inch, CONTENT_W, PAGE_H - 0.62 * inch - 0.65 * inch, id="body",
                  leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)
    doc.addPageTemplates([PageTemplate(id="main", frames=[frame], onPage=_page_chrome(opponent, week))])

    story = _snapshot_page(ff, opponent, week, game_date, tendencies, st)
    story += [PageBreak()] + _game_page(ff, d, "run", st)
    story += [PageBreak()] + _game_page(ff, d, "pass", st)
    story += [PageBreak()] + _glossary_page(st)
    doc.build(story)
    return ff
