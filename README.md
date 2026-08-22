# ScoutEdge Defense (Personal Build)

A personal opponent-tendency analyzer for Hudl playlist exports, originally
rebuilt from scratch to match the input/output of a web tool called
DefensiveIQ, then rebranded as ScoutEdge Defense — for your own use, on
your own machine.

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
11. Practice Scripts (Monday/Tuesday/Wednesday inside, perimeter, and team scripts)
12. Call Sheet Builder (blank template for your calls)
13. Game Day Call Sheet (auto-filled cheat sheet)

It also generates a 7-slide **"Player Presentation" scouting deck** (`.pptx`) —
title slide, offensive overview with metric cards, down & distance table,
favorite runs/passes, formation tendencies, and red zone keys — styled to
match the original DefensiveIQ deck this was rebuilt from (dark title/
red-zone slides, colored data tables). Swap in your own team/opponent logos
and it's ready to present.

Plus **Formation Hit Chart slides** — the classic 4-panel-per-page format:
formation name + snap count, run/pass split, backfield tendency, a quick
alignment diagram, and the top 3 run/pass concepts out of that look. Follows
your rule of thumb: top 6 formations, skipping any run fewer than 7 times.

Plus **Down & Distance Hit Chart slides** — same 4-panel-per-page format,
one panel per situation (1st & 10, 1st & Short, 2nd & Short/Medium/Long,
3rd & Short/Medium/Long, 4th Down). Since a down/distance situation doesn't
have one alignment to draw, each panel shows a run/pass proportion bar and
top 3 formations used in that situation instead, plus the same top run/pass
concepts box. Situation labels follow the same down-progression color scale
as the Excel workbook: teal (1st) → blue (2nd) → red (3rd) → maroon (4th).

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

## Run it as a local web app (matches the ScoutEdge Defense interface)

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

This generates both `..._ScoutEdge.xlsx` and `..._ScoutEdge_Scouting.pptx`.
Add `--out custom_name.xlsx` to control the output filename, or `--no-pptx` to
skip the slide deck.

## Files

- `analyzer.py` — reads the Hudl export, classifies every play by field zone,
  down & distance bucket, explosive/success, etc. All the tunable definitions
  (zone yard-line boundaries, explosive-play thresholds, success-rate formula)
  live at the top of this file with comments — edit them if your tendencies
  don't match what you see on tape.
- `report_builder.py` — builds all 13 Excel tabs from the analyzed data,
  including tab colors, section colors, and cell shading matched directly
  to a real DefensiveIQ workbook's own formatting (see below).
- `pptx_builder.py` — builds the 7-slide scouting-deck PowerPoint.
- `hit_chart_builder.py` — builds the Formation Hit Chart slides (appended
  to the end of the same deck). **Honest caveat on the alignment diagrams:**
  the receiver/TE/back position letters (X, Z, A, B, H, T, etc.) follow a
  generic, common scouting convention inferred from your PERSONNEL and
  FORMATION FAMILY columns — not a verified match to this specific
  opponent's own numbering system. Relabel in PowerPoint if your staff uses
  different letters; the formation name, snap count, run/pass split, and
  concept lists next to it are fully data-driven and accurate.
- `situation_hit_chart_builder.py` — builds the Down & Distance Hit Chart
  slides (also appended to the deck). Fully data-driven — no generic
  guessing involved, since it's just run/pass split, top formations, and
  top concepts per situation.
- `run_report.py` — command-line entry point (Excel + PowerPoint).
- `app.py` — Streamlit UI (Excel + PowerPoint download buttons).

## Color coding

Every tab matches the color system from a real DefensiveIQ workbook — the
tool this project was reverse-engineered from before being rebranded
ScoutEdge Defense
(pulled directly from its cell fills, not guessed):

- **Run columns** are always tinted light red, **pass columns** light blue,
  everywhere in the workbook — Run Concepts, Pass Concepts, Formation
  Tendencies, Down & Distance, Situational Summary, Game Day Call Sheet.
- **Field zones** each get their own color that's reused for that zone's
  label cell and row tint everywhere it appears: Backed Up = red, Open
  Field = blue, Midfield = teal, Fringe = olive, Red Zone = red (distinct
  tint from Backed Up), Goal Line = purple.
- **Down & distance situations** follow an urgency scale: 1st down = teal,
  2nd down = blue, 3rd down = red, 4th down = dark maroon (most critical),
  drive-start = yellow.
- **Formation names** are sky blue wherever they appear as a row label.
- Every sheet tab is colored to match its theme (e.g. Run Concepts tab is
  dark red, Pass Concepts tab is navy).

All of this is driven by constants at the top of `report_builder.py`
(`ZONE_STYLE`, `situation_color()`, `TAB_COLORS`, `TITLE_FILLS`) — tweak
those in one place if you want to shift the palette.

## Down & distance thresholds

- **2nd down:** Short = 1-3, Medium = 4-6, Long = 7+
- **3rd down:** Short = 1-3, Medium = 4-8, Long = 9+ (wider bands than 2nd,
  per DeAirus's own convention)

These live in one place, `dd_bucket()` in `analyzer.py`, and drive every
sheet and slide that breaks plays out by situation. If you ever want to
shift a cutoff, that function is the only place it needs to change.

## Down & distance edge cases

Hudl tags the first snap of a new possession as **DN=0**. In every real
export this comes paired with **DIST=10**, so the tool now reads it as its
own proper situation — **P & 10** (Possession and 10) — instead of silently
dropping it into a catch-all bucket. It shows up everywhere down/distance
situations do: the Down & Distance and Situational Summary Excel tabs, the
Game Day Call Sheet, and as its own panel (first in line) on the Down &
Distance Hit Chart slides in the PowerPoint deck.

## Practice Scripts (multi-day)

The Practice Scripts tab now covers three full practice days, matching the
original workbook's structure — not just Monday:

- **Monday** — Favorites (all downs)
- **Tuesday** — 2nd down situations only
- **Wednesday** — 3rd down situations only

Each day has the same five sections: Formation Alignment (top 12, padded
with blank numbered slots if fewer formations qualify), Inside Script (top
5 run concepts, 2 real reps each), Perimeter Script (top 8 pass concepts, 2
real reps each), and two Team Scripts that interleave the run/pass reps
into a team-period rep sheet.

Each "rep" is a real play instance pulled from your film — real hash, real
formation, real down & distance for that specific snap — not a fabricated
placeholder, and reps for the same concept prefer different hashes where
your film has them, for actual practice variety. The run/pass interleave
pattern in the Team Scripts (run, run, pass, pass...) is a clean, documented
convention I built rather than an exact replica of the original's internal
selection logic, which wasn't fully recoverable from the data alone.

## Different film-tagging conventions (pass concepts in another column)

By default, both run and pass concepts are read from the **PLAY** column,
same as this whole tool was built around. If someone else's film is tagged
differently — pass concepts live in a separate column like **PASS FAMILY**
instead of PLAY — that's supported without changing anything for you:

- **Streamlit app:** open "Advanced settings" above the Run Analysis button
  and enter the column name (e.g. `PASS FAMILY`).
- **Command line:** add `--pass-concept-col "PASS FAMILY"`.

Leave it blank/unset and everything works exactly as it always has. When
set, it only affects **pass** plays — run concepts always come from PLAY,
untouched — and falls back to PLAY for any pass row where that column
happens to be blank. A typo'd column name raises a clear error listing the
real column names in that file, rather than a cryptic crash.

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
