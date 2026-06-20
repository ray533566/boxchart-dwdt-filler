"""
app.py  –  Streamlit UI for BoxChart dW/dT Filler
==================================================
Upload RawData_for_BoxPlot.xlsx  +  Format_Mode_hopping_BoxChart620.xlsx
→  Download the filled output with box-plot charts intact.
"""

import io
import zipfile
import re

import pandas as pd
import streamlit as st

from fill_boxchart import (
    agg_dwdt,
    parse_shared_strings,
    build_shared_strings_xml,
    build_data_rows_xml,
    replace_sheet_data,
    OP_CHANNELS,
    MX_CHANNELS,
)

# ---------------------------------------------------------------------------
# Page config
# ---------------------------------------------------------------------------

st.set_page_config(
    page_title="BoxChart dW/dT Filler",
    page_icon="📊",
    layout="centered",
)

# ---------------------------------------------------------------------------
# Custom CSS  (dark navy, IBM Plex Mono accent)
# ---------------------------------------------------------------------------

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600&display=swap');

html, body, [class*="css"] {
    font-family: 'IBM Plex Sans', sans-serif;
}

/* ---- header ---- */
.hero {
    background: linear-gradient(135deg, #0d1b2a 0%, #1a2e45 100%);
    border-left: 4px solid #00c2ff;
    border-radius: 6px;
    padding: 1.4rem 1.6rem 1.2rem;
    margin-bottom: 1.4rem;
}
.hero h1 {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 1.45rem;
    font-weight: 600;
    color: #e0f4ff;
    margin: 0 0 0.25rem;
}
.hero p {
    font-size: 0.88rem;
    color: #8ab4cc;
    margin: 0;
}

/* ---- step badges ---- */
.step-badge {
    display: inline-block;
    background: #00c2ff;
    color: #0d1b2a;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.7rem;
    font-weight: 600;
    padding: 2px 8px;
    border-radius: 3px;
    margin-bottom: 6px;
}

/* ---- info / stat boxes ---- */
.stat-row {
    display: flex;
    gap: 12px;
    margin: 0.8rem 0 1rem;
    flex-wrap: wrap;
}
.stat-box {
    background: #1a2e45;
    border: 1px solid #2a4560;
    border-radius: 6px;
    padding: 0.6rem 1rem;
    flex: 1;
    min-width: 130px;
}
.stat-box .val {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 1.35rem;
    font-weight: 600;
    color: #00c2ff;
}
.stat-box .lbl {
    font-size: 0.75rem;
    color: #8ab4cc;
}

/* ---- channel table ---- */
.ch-table { font-size: 0.82rem; }

/* ---- success banner ---- */
.success-banner {
    background: #0a2e1a;
    border: 1px solid #1db954;
    border-radius: 6px;
    padding: 0.9rem 1.2rem;
    color: #a8ffca;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.82rem;
    margin: 1rem 0;
}
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Hero
# ---------------------------------------------------------------------------

st.markdown("""
<div class="hero">
  <h1>📊 BoxChart dW/dT Filler</h1>
  <p>Upload your raw data + format template → get a filled Excel with box-plot charts intact.</p>
</div>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Step 1 – Upload files
# ---------------------------------------------------------------------------

st.markdown('<span class="step-badge">STEP 1</span>', unsafe_allow_html=True)
st.markdown("**Upload both files**")

col1, col2 = st.columns(2)
with col1:
    raw_file = st.file_uploader(
        "RawData_for_BoxPlot.xlsx",
        type=["xlsx"],
        key="raw",
        help="Must contain columns: TESTSN, CHNumber, dW/dT",
    )
with col2:
    fmt_file = st.file_uploader(
        "Format_Mode_hopping_BoxChart620.xlsx",
        type=["xlsx"],
        key="fmt",
        help="The format template with box-plot charts",
    )

# ---------------------------------------------------------------------------
# Step 2 – Preview raw data
# ---------------------------------------------------------------------------

if raw_file:
    st.markdown("---")
    st.markdown('<span class="step-badge">STEP 2</span>', unsafe_allow_html=True)
    st.markdown("**Raw data preview**")

    df = pd.read_excel(raw_file)
    raw_file.seek(0)  # reset for later use

    required = {"TESTSN", "CHNumber", "dW/dT"}
    if not required.issubset(df.columns):
        missing = required - set(df.columns)
        st.error(f"❌ Missing columns in RawData: {missing}")
        st.stop()

    raw_sns = sorted(df["TESTSN"].unique().tolist())
    n_op_rows = len(df[df["CHNumber"].str.endswith("Operational", na=False)])
    n_mx_rows = len(df[df["CHNumber"].str.endswith("Maximum",     na=False)])

    st.markdown(f"""
    <div class="stat-row">
      <div class="stat-box"><div class="val">{len(raw_sns)}</div><div class="lbl">Unique TESTSN</div></div>
      <div class="stat-box"><div class="val">{len(df):,}</div><div class="lbl">Total rows</div></div>
      <div class="stat-box"><div class="val">{n_op_rows:,}</div><div class="lbl">Operational rows</div></div>
      <div class="stat-box"><div class="val">{n_mx_rows:,}</div><div class="lbl">Maximum rows</div></div>
    </div>
    """, unsafe_allow_html=True)

    # Channel distribution
    ch_counts = df["CHNumber"].value_counts().reindex(OP_CHANNELS + MX_CHANNELS, fill_value=0)
    ch_df = ch_counts.reset_index()
    ch_df.columns = ["CHNumber", "Row count"]
    ch_df["Sheet"] = ch_df["CHNumber"].apply(
        lambda x: "Normal_Operational Current" if "Operational" in x else "Bias400_Maximum Current"
    )
    with st.expander("Channel distribution", expanded=False):
        st.dataframe(ch_df, use_container_width=True, hide_index=True)

    # Sample dW/dT for first SN
    with st.expander(f"dW/dT preview — {raw_sns[0]}", expanded=False):
        sn0 = raw_sns[0]
        preview_rows = []
        for ch in OP_CHANNELS + MX_CHANNELS:
            sub = df.loc[(df["TESTSN"] == sn0) & (df["CHNumber"] == ch), "dW/dT"]
            if not sub.empty:
                preview_rows.append({
                    "CHNumber": ch,
                    "avg": round(sub.mean(), 4),
                    "max": round(sub.max(),  4),
                    "min": round(sub.min(),  4),
                    "n":   len(sub),
                })
        st.dataframe(pd.DataFrame(preview_rows), use_container_width=True, hide_index=True)

# ---------------------------------------------------------------------------
# Step 3 – Process & Download
# ---------------------------------------------------------------------------

if raw_file and fmt_file:
    st.markdown("---")
    st.markdown('<span class="step-badge">STEP 3</span>', unsafe_allow_html=True)
    st.markdown("**Generate filled Excel**")

    if st.button("▶  Run — Fill BoxChart", type="primary", use_container_width=True):
        with st.spinner("Aggregating dW/dT and patching Excel…"):
            try:
                # Re-read raw (already loaded df above)
                df = pd.read_excel(raw_file)
                raw_sns = sorted(df["TESTSN"].unique().tolist())

                # Aggregate
                op_rows = [[sn] + agg_dwdt(df, sn, OP_CHANNELS) for sn in raw_sns]
                mx_rows = [[sn] + agg_dwdt(df, sn, MX_CHANNELS) for sn in raw_sns]

                # Read format ZIP
                fmt_bytes = fmt_file.read()
                orig_contents = {}
                with zipfile.ZipFile(io.BytesIO(fmt_bytes)) as zin:
                    for name in zin.namelist():
                        orig_contents[name] = zin.read(name)

                shared_strings = parse_shared_strings(orig_contents["xl/sharedStrings.xml"])

                # Build row XMLs
                op_row_xmls, shared_after_op = build_data_rows_xml(op_rows, shared_strings, start_row=2)
                mx_row_xmls, shared_final    = build_data_rows_xml(mx_rows, shared_after_op, start_row=2)

                new_sheet2 = replace_sheet_data(
                    orig_contents["xl/worksheets/sheet2.xml"].decode("utf-8"), op_row_xmls
                )
                new_sheet3 = replace_sheet_data(
                    orig_contents["xl/worksheets/sheet3.xml"].decode("utf-8"), mx_row_xmls
                )
                new_ss_xml = build_shared_strings_xml(shared_final)

                replacements = {
                    "xl/worksheets/sheet2.xml":  new_sheet2.encode("utf-8"),
                    "xl/worksheets/sheet3.xml":  new_sheet3.encode("utf-8"),
                    "xl/sharedStrings.xml":      new_ss_xml,
                }

                # Write output ZIP in memory
                out_buf = io.BytesIO()
                with zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as zout:
                    for name, content in orig_contents.items():
                        zout.writestr(name, replacements.get(name, content))
                out_bytes = out_buf.getvalue()

                # Count preserved charts
                with zipfile.ZipFile(io.BytesIO(out_bytes)) as zcheck:
                    chart_count = sum(1 for f in zcheck.namelist() if "chart" in f)

                st.markdown(f"""
                <div class="success-banner">
                ✅  Done!&nbsp;&nbsp;
                {len(raw_sns)} SNs written to both sheets &nbsp;|&nbsp;
                {chart_count} chart files preserved
                </div>
                """, unsafe_allow_html=True)

                st.download_button(
                    label="⬇  Download Filled Excel",
                    data=out_bytes,
                    file_name="Format_Mode_hopping_BoxChart620_filled.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

            except Exception as e:
                st.error(f"❌ Error: {e}")
                st.exception(e)

# ---------------------------------------------------------------------------
# Footer
# ---------------------------------------------------------------------------

st.markdown("---")
st.markdown(
    "<small style='color:#8ab4cc;'>BoxChart dW/dT Filler · "
    "Optical Transceiver QE Tools · "
    "<a href='https://github.com/ray533566' style='color:#00c2ff;'>GitHub</a></small>",
    unsafe_allow_html=True,
)
