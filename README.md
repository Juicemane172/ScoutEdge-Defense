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
