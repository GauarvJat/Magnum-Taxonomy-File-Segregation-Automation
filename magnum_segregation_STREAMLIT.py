"""
DG Magnum Taxonomy— Data Segregation Tool  (Streamlit version)
Deploy: push to GitHub → connect repo to streamlit.app
"""

import io
import logging
import zipfile
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# =========================
# CONFIGURATION
# =========================
PLATFORM_CONFIG = {
    "AdServer":     ["CAMPAIGN", "PLACEMENT", "CREATIVE"],
    "Paid Social":  ["CAMPAIGN", "PLACEMENT", "CREATIVE"],
    "Programmatic": ["CAMPAIGN", "PLACEMENT", "CREATIVE"],   # CREATIVE not PLACEMENTGROUP
    "Search":       ["CAMPAIGN", "PLACEMENT", "CREATIVE"]
}

# Columns to remove from output (case-insensitive match)
COLS_TO_DROP = ["clicks", "currency", "spend"]

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(message)s")
logger = logging.getLogger()


# =========================
# CORE LOGIC
# =========================

def get_nc_indices(df):
    """Finds the first and last column index containing 'NC'."""
    nc_cols = [i for i, col in enumerate(df.columns) if "NC" in str(col)]
    return (nc_cols[0], nc_cols[-1] + 1) if nc_cols else (None, None)


def process_dataframe(df):
    """
    1. Fill NaN with 'Is missing' in NC range — skipping 'Free Text' columns.
    2. Rename duplicate Market column.
    3. Drop all NC columns.
    4. Drop Clicks, Currency, Spend columns.
    """

    # Step 1: Fill NaN with 'Is missing' across NC range
    # Free Text columns are excluded — empty cells there are intentional and
    # should not be flagged as missing or highlighted in red.
    start, end_plus_one = get_nc_indices(df)
    if start is not None:
        for col_idx in range(start, end_plus_one + 1):
            col_name = df.columns[col_idx]
            if "free text" not in str(col_name).lower():  # Skip Free Text columns
                df.iloc[:, col_idx] = df.iloc[:, col_idx].fillna("Is missing")

    # Step 2: Rename duplicate Market column to avoid confusion
    market_cols = [c for c in df.columns if str(c).lower() == "market"]
    if len(market_cols) > 1:
        df.rename(columns={market_cols[1]: "Market_MK"}, inplace=True)

    # Step 3: Drop all columns whose name contains 'NC'
    nc_col_names = [c for c in df.columns if "NC" in str(c)]
    df.drop(columns=nc_col_names, inplace=True, errors="ignore")

    # Step 4: Drop Clicks, Currency, Spend (case-insensitive)
    lower_map = {str(c).lower(): c for c in df.columns}
    for drop_col in COLS_TO_DROP:
        if drop_col in lower_map:
            df.drop(columns=[lower_map[drop_col]], inplace=True, errors="ignore")

    return df


def apply_formatting(file_bytes: bytes) -> bytes:
    """
    Loads workbook from bytes, applies red highlight to every cell
    containing 'Is missing' or 'Invalid' — skipping 'Free Text' columns.
    Returns the updated workbook as bytes.
    """
    wb = load_workbook(io.BytesIO(file_bytes))
    red_fill = PatternFill(start_color="FFFF0000", fill_type="solid")
    bad_values = {"is missing", "invalid", "invalid value"}

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]

        # Read header row once to know each column's name
        headers = [
            str(c.value).lower() if c.value else ""
            for c in ws[1]
        ]

        for row in ws.iter_rows(min_row=2):
            for cell in row:
                # Skip Free Text columns — empty cells there are not errors
                col_header = headers[cell.column - 1]
                if "free text" in col_header:
                    continue
                if cell.value and str(cell.value).strip().lower() in bad_values:
                    cell.fill = red_fill

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def parse_filename(filename: str):
    """
    Extracts platform, level, and date from a structured filename.
    Expected format: 'Platform LEVEL - Date.xlsx'
    Example: 'Programmatic CREATIVE - Jan 2025.xlsx'
    """
    try:
        name_part, date_part = filename.rsplit(" - ", 1)
        date = date_part.replace(".xlsx", "").strip()
        for level in ["CAMPAIGN", "PLACEMENT", "PLACEMENTGROUP", "CREATIVE"]:
            if name_part.upper().endswith(level):
                platform = name_part[: -len(level)].strip()
                return platform, level, date
    except ValueError:
        return None, None, None
    return None, None, None


def run_automation(uploaded_files, log_fn):
    """
    Core processing — works entirely in-memory (no disk I/O).
    Returns a dict of {output_filename: bytes} and a stats summary.
    """
    stats = {"files_created": 0, "tabs_created": 0, "platforms_processed": 0}
    inputs = {}

    # ── Parse filenames and group by platform / date / level ─────────────────
    for uf in uploaded_files:
        platform_raw, level, date = parse_filename(uf.name)
        if not platform_raw:
            log_fn(f"⚠️  Skipped (unrecognised name): {uf.name}")
            continue
        platform_map = {k.lower(): k for k in PLATFORM_CONFIG.keys()}
        platform = platform_map.get(platform_raw.lower())
        if platform and level in PLATFORM_CONFIG[platform]:
            inputs.setdefault(platform, {}).setdefault(date, {})[level] = uf
        else:
            log_fn(f"⚠️  Skipped (platform/level mismatch): {uf.name}")

    if not inputs:
        log_fn("❌  No valid files found. Check filenames and try again.")
        return {}, stats

    output_files = {}

    # ── Process each platform → date → level group ────────────────────────────
    for platform, dates in inputs.items():
        stats["platforms_processed"] += 1
        log_fn(f"⚙️  Processing: **{platform}**")

        for date, levels in dates.items():
            market_data = {}

            for level, uf in levels.items():
                try:
                    df = pd.read_excel(
                        io.BytesIO(uf.read()),
                        keep_default_na=False, na_values=[""]
                    )
                    uf.seek(0)  # Reset stream pointer for potential re-use
                    if df.empty or "Market" not in df.columns:
                        log_fn(f"   ↳ Skipped (empty or no Market col): {uf.name}")
                        continue
                    df = process_dataframe(df)
                    for market, group in df.groupby("Market"):
                        market_data.setdefault(market, {})[level] = group
                except Exception as e:
                    log_fn(f"   ↳ Error reading {uf.name}: {e}")

            # ── Write one output file per market ──────────────────────────────
            for market, levels_dict in market_data.items():
                output_filename = f"{platform}_{market}_{date}.xlsx"
                buf = io.BytesIO()
                with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                    for level in PLATFORM_CONFIG[platform]:
                        if level in levels_dict:
                            levels_dict[level].to_excel(
                                writer, sheet_name=level, index=False
                            )
                            stats["tabs_created"] += 1
                # Apply red highlighting (Free Text columns are skipped inside)
                raw = apply_formatting(buf.getvalue())
                output_files[output_filename] = raw
                stats["files_created"] += 1
                log_fn(f"   ✅  Created → {output_filename}")

    return output_files, stats


def build_zip(output_files: dict) -> bytes:
    """Bundles all output Excel files into one ZIP for a single download."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in output_files.items():
            zf.writestr(name, data)
    return buf.getvalue()


# =========================
# STREAMLIT UI
# =========================

def main():
    st.set_page_config(
        page_title="DG Taxonomy — Data Segregation",
        page_icon="🔷",
        layout="centered"
    )

    # ── Custom CSS ────────────────────────────────────────────────────────────
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap');

    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

    .hero {
        background: linear-gradient(135deg, #0F3460 0%, #1A1A2E 100%);
        border-radius: 14px;
        padding: 32px 36px 24px;
        margin-bottom: 28px;
    }
    .hero h1 { color: #E94560; font-size: 2.4rem; margin: 0; letter-spacing: 2px; }
    .hero p  { color: #A0A0B0; margin: 6px 0 0; font-size: 0.95rem; }

    .pill {
        display: inline-block;
        background: #0F3460;
        color: #A0C4FF;
        border-radius: 20px;
        padding: 3px 12px;
        font-size: 0.78rem;
        margin: 6px 4px 0 0;
    }

    .log-box {
        background: #0A0A1A;
        border-radius: 8px;
        padding: 14px 16px;
        font-family: 'Courier New', monospace;
        font-size: 0.82rem;
        color: #A8D8A8;
        max-height: 260px;
        overflow-y: auto;
        white-space: pre-wrap;
    }

    .stat-card {
        background: #16213E;
        border-left: 4px solid #E94560;
        border-radius: 8px;
        padding: 14px 20px;
        text-align: center;
    }
    .stat-num   { font-size: 2rem; font-weight: 700; color: #E94560; }
    .stat-label { font-size: 0.8rem; color: #A0A0B0; margin-top: 2px; }
    </style>
    """, unsafe_allow_html=True)

    # ── Hero header ───────────────────────────────────────────────────────────
    st.markdown("""
    <div class="hero">
        <h1>MAGNUM</h1>
        <p>Data Segregation Tool &nbsp;·&nbsp; v2.1</p>
        <div style="margin-top:12px">
            <span class="pill">AdServer</span>
            <span class="pill">Paid Social</span>
            <span class="pill">Programmatic</span>
            <span class="pill">Search</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Upload section ────────────────────────────────────────────────────────
    st.markdown("#### 📂 Upload Input Files")
    st.caption(
        "Upload one or more `.xlsx` files named in the format:  \n"
        "`Platform LEVEL - DATE.xlsx`  \n"
        "e.g. `Programmatic CREATIVE - Jan 2025.xlsx`"
    )

    uploaded_files = st.file_uploader(
        label="Drop files here or click to browse",
        type=["xlsx"],
        accept_multiple_files=True,
        label_visibility="collapsed"
    )

    if uploaded_files:
        st.success(f"{len(uploaded_files)} file(s) selected")
        with st.expander("View selected files"):
            for f in uploaded_files:
                st.text(f"  • {f.name}")

    st.divider()

    # ── Run button ────────────────────────────────────────────────────────────
    col_btn, col_spacer = st.columns([1, 3])
    run_clicked = col_btn.button(
        "▶  Run Segregation", type="primary", use_container_width=True
    )

    if run_clicked:
        if not uploaded_files:
            st.warning("Please upload at least one file before running.")
            st.stop()

        # Live activity log
        log_lines = []
        log_placeholder = st.empty()

        def log_fn(msg):
            log_lines.append(msg)
            log_placeholder.markdown(
                '<div class="log-box">' + "<br>".join(log_lines) + "</div>",
                unsafe_allow_html=True
            )

        with st.spinner("Processing files…"):
            output_files, stats = run_automation(uploaded_files, log_fn)

        st.divider()

        # ── Stats cards ───────────────────────────────────────────────────────
        if stats["files_created"] > 0:
            st.markdown("#### 📊 Summary")
            c1, c2, c3 = st.columns(3)
            for col, num, label in [
                (c1, stats["platforms_processed"], "Platforms"),
                (c2, stats["files_created"],       "Output Files"),
                (c3, stats["tabs_created"],        "Tabs Written"),
            ]:
                col.markdown(
                    f'<div class="stat-card">'
                    f'<div class="stat-num">{num}</div>'
                    f'<div class="stat-label">{label}</div>'
                    f'</div>',
                    unsafe_allow_html=True
                )

            st.divider()

            # ── Download section ──────────────────────────────────────────────
            st.markdown("#### 📥 Download Results")

            if len(output_files) == 1:
                # Single file — direct download button
                name, data = next(iter(output_files.items()))
                st.download_button(
                    label=f"⬇  Download  {name}",
                    data=data,
                    file_name=name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            else:
                # Multiple files — ZIP download + individual options
                zip_data = build_zip(output_files)
                st.download_button(
                    label=f"⬇  Download All as ZIP  ({len(output_files)} files)",
                    data=zip_data,
                    file_name="Magnum_Output.zip",
                    mime="application/zip",
                    use_container_width=True
                )
                with st.expander("Individual file downloads"):
                    for name, data in output_files.items():
                        st.download_button(
                            label=name,
                            data=data,
                            file_name=name,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=name
                        )
        else:
            st.error("No output files were generated. Review the activity log above.")

    # ── Footer ────────────────────────────────────────────────────────────────
    st.markdown("---")
    st.caption("© Data Governance  |  Internal use only")


if __name__ == "__main__":
    main()
