#!/usr/bin/env python3
"""
Usage:
    python run_report.py PLAYLIST.xlsx --opponent "Christ The King" --week 1 --game-date 2026-08-21 --out report.xlsx

PLAYLIST.xlsx is a Hudl playlist export (.xlsx or .csv).
Generates both the Excel workbook and a PowerPoint scouting-report deck.
"""
import argparse
import datetime
from analyzer import load_playlist
from report_builder import build_workbook
from pptx_builder import build_presentation


def main():
    ap = argparse.ArgumentParser(description="Generate an opponent tendency report from a Hudl playlist export.")
    ap.add_argument("playlist", help="Path to Hudl playlist export (.xlsx or .csv)")
    ap.add_argument("--opponent", default="Opponent", help="Opponent name")
    ap.add_argument("--week", default="1", help="Week number")
    ap.add_argument("--game-date", default=None, help="Game date, e.g. 2026-08-21 (defaults to today)")
    ap.add_argument("--out", default=None, help="Output .xlsx path (default: <Opponent>_<Week>_ScoutEdge.xlsx)")
    ap.add_argument("--no-pptx", action="store_true", help="Skip generating the PowerPoint deck")
    ap.add_argument("--pass-concept-col", default=None,
                     help="Optional: pull PASS concepts from this column instead of PLAY "
                          "(e.g. 'PASS FAMILY'). Run concepts always come from PLAY, unaffected.")
    args = ap.parse_args()

    df = load_playlist(args.playlist, pass_concept_col=args.pass_concept_col)
    print(f"Loaded {len(df)} run/pass plays from {args.playlist}")

    base = (args.out or f"{args.opponent.replace(' ', '_')}_Week{args.week}_ScoutEdge.xlsx")
    if base.lower().endswith(".xlsx"):
        base = base[:-5]

    wb = build_workbook(df, opponent=args.opponent, week=args.week)
    xlsx_path = base + ".xlsx"
    wb.save(xlsx_path)
    print(f"Saved Excel workbook to {xlsx_path}")

    if not args.no_pptx:
        game_date = datetime.datetime.strptime(args.game_date, "%Y-%m-%d").date() if args.game_date else datetime.date.today()
        prs = build_presentation(df, opponent=args.opponent, week=args.week, game_date=game_date)
        pptx_path = base + "_Scouting.pptx"
        prs.save(pptx_path)
        print(f"Saved PowerPoint deck to {pptx_path}")


if __name__ == "__main__":
    main()
