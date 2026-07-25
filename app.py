"""
DefensiveIQ (personal build) — Streamlit front-end.

Run locally with:
    pip install streamlit pandas openpyxl
    streamlit run app.py

Then open the local URL it prints (usually http://localhost:8501).
"""
import io
import datetime
import streamlit as st
from analyzer import load_playlist
from report_builder import build_workbook
from pptx_builder import build_presentation

st.set_page_config(page_title="DefensiveIQ (Personal)", page_icon="🏈", layout="centered")

st.markdown(
    "<h1 style='margin-bottom:0'>DEFENSIVE<span style='color:#e04b3c'>IQ</span></h1>"
    "<p style='color:#9ca3af;margin-top:0'>Upload your Hudl playlist export and get your full "
    "opponent tendency report in seconds.</p>",
    unsafe_allow_html=True,
)

col1, col2, col3 = st.columns(3)
opponent = col1.text_input("Opponent Name", value="Opponent")
week = col2.text_input("Week", value="1")
game_date = col3.date_input("Game Date")

st.subheader("Upload Hudl Playlist Export (.xlsx or .csv)")
playlist_file = st.file_uploader("Playlist export", type=["xlsx", "csv"], label_visibility="collapsed")

run = st.button("⚡ RUN ANALYSIS", type="primary", use_container_width=True)

if run:
    if not playlist_file:
        st.error("Upload a Hudl playlist export first.")
    else:
        with st.spinner("Analyzing plays..."):
            df = load_playlist(playlist_file)
            wb = build_workbook(df, opponent=opponent, week=week)
            buf = io.BytesIO()
            wb.save(buf)
            buf.seek(0)

            prs = build_presentation(df, opponent=opponent, week=week, game_date=game_date)
            pptx_buf = io.BytesIO()
            prs.save(pptx_buf)
            pptx_buf.seek(0)

        st.success(f"Analysis complete — {len(df)} plays analyzed")

        runs = (df["PLAY TYPE"] == "Run").sum()
        passes = (df["PLAY TYPE"] == "Pass").sum()
        total = runs + passes
        rz = df[df["ZONE"] == "RZ"]
        gl = df[df["ZONE"] == "GL"]

        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Total Plays", total)
        c2.metric("Run %", f"{round(runs/total*100) if total else 0}%")
        c3.metric("Pass %", f"{round(passes/total*100) if total else 0}%")
        c4.metric("RZ Run %", f"{round((rz['PLAY TYPE']=='Run').mean()*100) if len(rz) else 0}%")
        c5.metric("GL Run %", f"{round((gl['PLAY TYPE']=='Run').mean()*100) if len(gl) else 0}%")

        st.subheader("Download Your Reports")
        col_a, col_b = st.columns(2)
        col_a.download_button(
            "📊 EXCEL WORKBOOK",
            data=buf,
            file_name=f"{opponent.replace(' ', '_')}_Week{week}_DefensiveIQ.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        col_b.download_button(
            "🎬 PLAYER PRESENTATION",
            data=pptx_buf,
            file_name=f"{opponent.replace(' ', '_')}_Week{week}_DefensiveIQ_Scouting.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
        )

st.caption(
    "Personal rebuild — analysis logic (zone boundaries, success-rate formula, explosive-play "
    "thresholds) is inferred from a sample report and documented in analyzer.py. Adjust the "
    "constants there if your own tendencies don't match your eye test."
)
