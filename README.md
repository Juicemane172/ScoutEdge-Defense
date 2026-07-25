# DefensiveIQ (Personal Build)

A personal opponent-tendency analyzer for Hudl playlist exports, rebuilt from
scratch to match the input/output of the DefensiveIQ web tool — for your own
use, on your own machine.

## What it does

Upload a Hudl playlist export (`.xlsx` or `.csv`) and it generates a 13-tab
Excel tendency report:

1. Film Log
2. Field Zone Tendencies
3. Run Tendencies
4. Pass Tendencies
5. Hash Tendencies
6. Down & Distance
7. Run Concepts (efficiency, explosive %, success %)
8. Pass Concepts
9. Formation Tendencies
10. Situational Summary
11. Practice Scripts (auto-built Monday inside/pass script)
12. Call Sheet Builder (blank template for your calls)
13. Game Day Call Sheet (auto-filled cheat sheet)

It also generates a 7-slide **"Player Presentation" scouting deck** (`.pptx`) —
title slide, offensive overview with metric cards, down & distance table,
favorite runs/passes, formation tendencies, and red zone keys — styled to
match the original DefensiveIQ deck (dark title/red-zone slides, colored
data tables). Swap in your own team/opponent logos and it's ready to present.

Plus **Formation Hit Chart slides** — the classic 4-panel-per-page format:
formation name + snap count, run/pass split, backfield tendency, a quick
alignment diagram, and the top 3 run/pass concepts out of that look. Follows
your rule of thumb: top 6 formations, skipping any run fewer than 7 times.

**The alignment diagrams read your own formation-naming tags:**
- Base names like DEUCE, TRIPS, SLOT, BUNCH, CLUSTER are treated as open/
  spread sets — even an 11-personnel "TE" is drawn flexed out at receiver
  depth, not attached to the line.
- Everything else (PRO, TREY, DOUBLES, etc.) is treated as TE-based: the TE
  is drawn tight to the tackle, on the line — unless the name carries a
  **YO** tag (e.g. "TREY YO"), in which case it's drawn tight to the tackle
  horizontally but stepped off the ball, matching a Y-off alignment.
- A **FIB** tag (Formation In Boundary) flips the strong side of the
  diagram to the left/boundary instead of the default right/field side —
  so "TWIN SPLIT FIB" correctly draws 2 receivers to the boundary and 1 to
  the field.

These three tags (base name, YO, FIB) are read directly off your FORMATION
column text, so as you keep tagging film with the same vocabulary, new
opponents' hit charts will follow the same rules automatically. If you
start using a new base name that should count as "open" (not TE-based),
add it to `OPEN_BASE_NAMES` at the top of `hit_chart_builder.py`.

## Setup

```bash
pip install streamlit pandas openpyxl numpy python-pptx
```

## Run it as a local web app (matches the DefensiveIQ interface)

```bash
streamlit run app.py
```

This opens a local page in your browser — upload your playlist, fill in
opponent/week, click "Run Analysis," download the Excel workbook and/or the
PowerPoint scouting deck.

## Or run it from the command line, no browser needed

```bash
python run_report.py your_playlist.xlsx --opponent "Christ The King" --week 1 --game-date 2026-08-21
```

This generates both `..._DefensiveIQ.xlsx` and `..._DefensiveIQ_Scouting.pptx`.
Add `--out custom_name.xlsx` to control the output filename, or `--no-pptx` to
skip the slide deck.

## Files

- `analyzer.py` — reads the Hudl export, classifies every play by field zone,
  down & distance bucket, explosive/success, etc. All the tunable definitions
  (zone yard-line boundaries, explosive-play thresholds, success-rate formula)
  live at the top of this file with comments — edit them if your tendencies
  don't match what you see on tape.
- `report_builder.py` — builds all 13 Excel tabs from the analyzed data.
- `pptx_builder.py` — builds the 7-slide scouting-deck PowerPoint.
- `hit_chart_builder.py` — builds the Formation Hit Chart slides (appended
  to the end of the same deck). **Honest caveat on the alignment diagrams:**
  the receiver/TE/back position letters (X, Z, A, B, H, T, etc.) follow a
  generic, common scouting convention inferred from your PERSONNEL and
  FORMATION FAMILY columns — not a verified match to this specific
  opponent's own numbering system. Relabel in PowerPoint if your staff uses
  different letters; the formation name, snap count, run/pass split, and
  concept lists next to it are fully data-driven and accurate.
- `run_report.py` — command-line entry point (Excel + PowerPoint).
- `app.py` — Streamlit UI (Excel + PowerPoint download buttons).

## Honest caveats

This was rebuilt by reverse-engineering your own uploaded input file
(`PlaylistData...xlsx`) against your own output file
(`Christ_The_King_Week1_DefensiveIQ.xlsx`) — not from the original tool's
source code, which I never had access to. A few things are inferred, not
exact:

- **Field zone boundaries** (BZ/OF/MF/FZ/RZ/GL) — matched to your sample data,
  double-check them against a few known plays.
- **"Success rate"** uses the common analytics convention (1st down ≥40% of
  distance, 2nd down ≥60%, 3rd/4th down = full conversion). The original tool
  may define it differently.
- **Play direction** (left/right) isn't in a standard Hudl playlist export, so
  that column shows hash data instead — same limitation the original tool
  had (its own "Top Direction" columns were also blank on your sample).
- On your sample file, this produces 106 analyzed plays vs. the original
  tool's 111 — a few edge-case rows (no run/pass tag, certain penalties)
  are handled slightly differently. Close, not pixel-perfect.

Test it against a few weeks of your own tape and tune `analyzer.py`'s
constants until the tendencies match what you already know about an
opponent — then trust it going forward.
