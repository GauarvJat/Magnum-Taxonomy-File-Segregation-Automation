"""
Magnum  —  Data Segregation Tool  (Streamlit / Web version)
────────────────────────────────────────────────────────────
Run locally : streamlit run magnum_segregation_STREAMLIT.py
Deploy      : Push repo to GitHub → connect to streamlit.app
              Set the main file to: magnum_segregation_STREAMLIT.py
IMPORTANT — requirements.txt must be in the GitHub repo ROOT:
    streamlit
    pandas
    openpyxl
"""
from __future__ import annotations
import io
import zipfile
from pathlib import PurePath

# ── Streamlit must be the very first import ───────────────────────────────────
import streamlit as st

# ── Page config must be the FIRST Streamlit call in the file ─────────────────
st.set_page_config(
    page_title = "Magnum — Data Segregation",
    page_icon  = "🔷",
    layout     = "centered",
)

# ── Safe imports — friendly error if packages are missing ────────────────────
_missing = []
try:
    import pandas as pd
except ImportError:
    _missing.append("pandas")
try:
    from openpyxl import load_workbook
    from openpyxl.styles import PatternFill
except ImportError:
    _missing.append("openpyxl")

if _missing:
    st.error(
        f"**Missing package(s): `{'`, `'.join(_missing)}`**\n\n"
        "Make sure your GitHub repo root contains a `requirements.txt` file with:\n\n"
        "```\nstreamlit\npandas\nopenpyxl\n```\n\n"
        "Streamlit Cloud will install them automatically on the next deploy."
    )
    st.stop()

# ═══════════════════════════════════════════════════════════════════════════════
#  CONFIGURATION
# ═══════════════════════════════════════════════════════════════════════════════
PLATFORM_CONFIG: dict[str, list[str]] = {
    "AdServer":     ["CAMPAIGN", "PLACEMENT", "CREATIVE"],
    "Paid Social":  ["CAMPAIGN", "PLACEMENT", "CREATIVE"],
    "Programmatic": ["CAMPAIGN", "PLACEMENT", "CREATIVE"],
    "Search":       ["CAMPAIGN", "PLACEMENT", "CREATIVE"],
}

# Columns dropped from every output file (case-insensitive)
COLS_TO_DROP: list[str] = ["clicks", "currency", "spend"]

# Known level tokens used when parsing filenames
KNOWN_LEVELS: list[str] = ["CAMPAIGN", "PLACEMENT", "PLACEMENTGROUP", "CREATIVE"]

# ═══════════════════════════════════════════════════════════════════════════════
#  CORE LOGIC
# ═══════════════════════════════════════════════════════════════════════════════
def parse_filename(filename: str) -> tuple[str | None, str | None, str | None]:
    """
    Parses filenames in the format:  Platform LEVEL - Date.xlsx
    e.g.  Programmatic CREATIVE - Jan 2025.xlsx
    Returns (platform_raw, level, date) or (None, None, None) on failure.
    """
    try:
        stem = PurePath(filename).stem.strip()
        name_part, date_part = stem.rsplit(" - ", 1)
        date = date_part.strip()
        for level in KNOWN_LEVELS:
            if name_part.upper().endswith(level):
                platform = name_part[: -len(level)].strip()
                return platform, level, date
    except Exception:
        pass
    return None, None, None


def get_nc_indices(df: "pd.DataFrame") -> tuple[int | None, int | None]:
    """Returns the first and last+1 column indices of NC-named columns."""
    nc_cols = [i for i, col in enumerate(df.columns) if "NC" in str(col)]
    return (nc_cols[0], nc_cols[-1] + 1) if nc_cols else (None, None)


def process_dataframe(df: "pd.DataFrame") -> "pd.DataFrame":
    """
    1. Fill empty cells with 'Is missing' across the NC column range,
       skipping any column whose name contains 'free text'.
    2. Rename a duplicate Market column to Market_MK.
    3. Drop all NC-named columns.
    4. Drop Clicks, Currency, Spend columns.
    """
    # Step 1 — fill missing, skip Free Text columns
    start, end_plus_one = get_nc_indices(df)
    if start is not None:
        for col_idx in range(start, end_plus_one + 1):
            col_name = df.columns[col_idx]
            if "free text" not in str(col_name).lower():
                df.iloc[:, col_idx] = df.iloc[:, col_idx].fillna("Is missing")

    # Step 2 — rename duplicate Market column
    market_cols = [c for c in df.columns if str(c).lower() == "market"]
    if len(market_cols) > 1:
        df.rename(columns={market_cols[1]: "Market_MK"}, inplace=True)

    # Step 3 — drop NC columns
    nc_cols = [c for c in df.columns if "NC" in str(c)]
    df.drop(columns=nc_cols, inplace=True, errors="ignore")

    # Step 4 — drop named columns
    lower_map = {str(c).lower(): c for c in df.columns}
    for target in COLS_TO_DROP:
        if target in lower_map:
            df.drop(columns=[lower_map[target]], inplace=True, errors="ignore")

    return df


def apply_formatting(file_bytes: bytes) -> bytes:
    """
    Red-highlights validation issues in the workbook:
      - CAMPAIGN / PLACEMENT sheets (and any non-CREATIVE sheet):
        every cell containing 'Is missing' / 'Invalid' / 'Invalid value',
        skipping any column whose header contains 'free text'.
      - CREATIVE sheet:
        ONLY the column named EXACTLY "Influencer" (case-insensitive,
        ignoring leading/trailing spaces) is checked and highlighted.
        Columns like "Influencer Name" or "Influencer Type" are left alone.
    Works entirely in memory — no files written to disk.
    """
    wb       = load_workbook(io.BytesIO(file_bytes))
    red_fill = PatternFill(start_color="FFFF0000", fill_type="solid")
    bad_vals = {"is missing", "invalid", "invalid value"}

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]

        # Build a column-index -> header-text map for row 1
        headers: dict[int, str] = {
            cell.column: (str(cell.value).strip() if cell.value else "")
            for cell in ws[1]
        }

        # Free Text columns are always skipped, on every sheet
        free_text_cols = {
            col for col, name in headers.items() if "free text" in name.lower()
        }

        is_creative = sheet_name.upper() == "CREATIVE"

        for row in ws.iter_rows(min_row=2):
            for cell in row:
                if cell.column in free_text_cols:
                    continue

                header = headers.get(cell.column, "")

                # CREATIVE Sheet Rule:
                # Highlight ONLY the column exactly named "Influencer"
                if is_creative and header.lower() != "influencer":
                    continue

                if cell.value and str(cell.value).strip().lower() in bad_vals:
                    cell.fill = red_fill

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def run_segregation(
    uploaded_files: list,
    progress_cb=None,
) -> tuple[dict[str, bytes], dict]:
    """
    Processes all uploaded files entirely in memory.
    Returns output_files dict and stats summary.
    """
    stats: dict = {
        "files_created":       0,
        "tabs_created":        0,
        "platforms_processed": 0,
        "skipped":             [],
    }
    inputs: dict    = {}
    platform_map    = {k.lower(): k for k in PLATFORM_CONFIG.keys()}

    # Phase 1 — group files
    for uf in uploaded_files:
        platform_raw, level, date = parse_filename(uf.name)
        if not platform_raw:
            stats["skipped"].append(f"{uf.name}  — filename not parseable")
            continue
        platform = platform_map.get(platform_raw.lower())
        if not platform:
            stats["skipped"].append(f"{uf.name}  — unknown platform: '{platform_raw}'")
            continue
        if level not in PLATFORM_CONFIG[platform]:
            stats["skipped"].append(f"{uf.name}  — '{level}' not valid for {platform}")
            continue
        inputs.setdefault(platform, {}).setdefault(date, {})[level] = uf

    output_files: dict[str, bytes] = {}
    total   = sum(len(d) for d in inputs.values())
    counter = 0

    # Phase 2 — process
    for platform, dates in inputs.items():
        stats["platforms_processed"] += 1
        for date, levels in dates.items():
            counter += 1
            if progress_cb:
                progress_cb(counter / max(total, 1))
            market_data: dict = {}
            for level, uf in levels.items():
                try:
                    uf.seek(0)
                    df = pd.read_excel(
                        io.BytesIO(uf.read()),
                        keep_default_na=False,
                        na_values=[""],
                    )
                    if df.empty or "Market" not in df.columns:
                        stats["skipped"].append(
                            f"{uf.name}  — empty or missing 'Market' column"
                        )
                        continue
                    df = process_dataframe(df)
                    for market, group in df.groupby("Market"):
                        market_data.setdefault(str(market), {})[level] = group
                except Exception as exc:
                    stats["skipped"].append(f"{uf.name}  — error: {exc}")

            for market, levels_dict in market_data.items():
                safe_market     = str(market).replace("/", "-").replace("\\", "-")
                output_filename = f"{platform}_{safe_market}_{date}.xlsx"
                buf = io.BytesIO()
                with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                    for level in PLATFORM_CONFIG[platform]:
                        if level in levels_dict:
                            levels_dict[level].to_excel(
                                writer, sheet_name=level, index=False
                            )
                            stats["tabs_created"] += 1
                output_files[output_filename] = apply_formatting(buf.getvalue())
                stats["files_created"] += 1

    return output_files, stats


def build_zip(output_files: dict[str, bytes]) -> bytes:
    """Packs all output files into a single downloadable ZIP."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for fname, data in output_files.items():
            zf.writestr(fname, data)
    return buf.getvalue()


# ═══════════════════════════════════════════════════════════════════════════════
#  CSS — full Streamlit chrome override
# ═══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<style>
/* ── Global background ── */
html, body, [data-testid="stAppViewContainer"] {
    background-color: #0E1628 !important;
}
[data-testid="stAppViewContainer"] > .main {
    background-color: #0E1628;
}
/* ── Hide Streamlit chrome ── */
[data-testid="stHeader"]  { background: transparent !important; }
[data-testid="stToolbar"] { display: none !important; }
footer                    { visibility: hidden !important; }
#MainMenu                 { display: none !important; }
/* ── Hero banner ── */
.mg-hero {
    background: linear-gradient(135deg, #111F3E 0%, #1A1A2E 100%);
    border-bottom: 4px solid #E94560;
    border-radius: 14px;
    padding: 34px 38px 28px;
    margin-bottom: 30px;
}
.mg-hero-name {
    font-size: 2.2rem; font-weight: 900;
    color: #E94560; letter-spacing: 2px; margin: 0 0 4px;
}
.mg-hero-sub  { font-size: .95rem; color: #7A90AA; margin: 0 0 18px; }
.mg-tags      { display: flex; flex-wrap: wrap; gap: 8px; }
.mg-tag {
    background: rgba(233,69,96,.12);
    border: 1px solid rgba(233,69,96,.30);
    color: #E94560; font-size: .78rem; font-weight: 700;
    padding: 3px 14px; border-radius: 20px; letter-spacing: .4px;
}
/* ── Section label ── */
.mg-label {
    font-size: .72rem; font-weight: 800;
    letter-spacing: 1.6px; text-transform: uppercase;
    color: #7A90AA; margin-bottom: 8px;
}
/* ── File uploader ── */
[data-testid="stFileUploader"] section {
    background: #111F3E !important;
    border: 1.5px dashed #1E3058 !important;
    border-radius: 10px !important;
}
[data-testid="stFileUploader"] section:hover {
    border-color: #E94560 !important;
}
[data-testid="stFileUploader"] label { color: #7A90AA !important; }
/* ── Run button ── */
.stButton > button {
    background: #E94560 !important;
    color: #fff !important;
    border: none !important;
    border-radius: 8px !important;
    font-weight: 800 !important;
    font-size: 1rem !important;
    padding: 12px 0 !important;
    width: 100% !important;
    letter-spacing: .4px;
    transition: background .2s;
}
.stButton > button:hover    { background: #C73350 !important; }
.stButton > button:disabled { opacity: .45 !important; }
/* ── Download button ── */
.stDownloadButton > button {
    background: #0066CC !important;
    color: #fff !important;
    border: none !important;
    border-radius: 8px !important;
    font-weight: 800 !important;
    font-size: .95rem !important;
    padding: 12px 0 !important;
    width: 100% !important;
    transition: background .2s;
}
.stDownloadButton > button:hover { background: #0052A3 !important; }
/* ── Metric boxes ── */
[data-testid="metric-container"] {
    background: #111F3E;
    border: 1px solid #1E3058;
    border-radius: 10px;
    padding: 14px !important;
}
[data-testid="metric-container"] label { color: #7A90AA !important; font-size: .8rem !important; }
[data-testid="stMetricValue"]          { color: #E94560 !important; font-weight: 800 !important; }
/* ── Alerts ── */
[data-testid="stAlert"] {
    background: #111F3E !important;
    border: 1px solid #1E3058 !important;
    border-radius: 8px !important;
    color: #7A90AA !important;
}
/* ── Progress bar ── */
[data-testid="stProgressBar"] > div > div { background: #E94560 !important; }
/* ── Expander ── */
[data-testid="stExpander"] {
    background: #111F3E !important;
    border: 1px solid #1E3058 !important;
    border-radius: 10px !important;
}
.streamlit-expanderHeader { color: #7A90AA !important; font-size: .9rem !important; }
/* ── Divider ── */
hr { border-color: #1E3058 !important; }
/* ── Text ── */
p, li, span { color: #7A90AA; }
strong      { color: #D0DCF0 !important; }
code {
    background: #07111F !important; color: #7DD4C0 !important;
    border-radius: 4px !important; font-size: .84rem !important;
    padding: 1px 6px !important;
}
/* ── Footer ── */
.mg-footer {
    text-align: center; color: #1E3058;
    font-size: .78rem; padding: 24px 0 8px;
    border-top: 1px solid #1E3058; margin-top: 40px;
}
</style>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
#  UI — HERO
# ═══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<div class="mg-hero">
    <div class="mg-hero-name">🔷 MAGNUM</div>
    <div class="mg-hero-sub">Data Segregation Tool &nbsp;·&nbsp; v2.1</div>
    <div class="mg-tags">
        <span class="mg-tag">● AdServer</span>
        <span class="mg-tag">● Paid Social</span>
        <span class="mg-tag">● Programmatic</span>
        <span class="mg-tag">● Search</span>
    </div>
</div>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
#  UI — FILE NAMING GUIDE
# ═══════════════════════════════════════════════════════════════════════════════
with st.expander("📋  File Naming Convention — click to expand"):
    st.markdown("""
Each uploaded file **must** follow this exact naming pattern:
```
Platform LEVEL - Date.xlsx
```
**Valid examples:**
```
AdServer CAMPAIGN - Jan 2025.xlsx
Paid Social PLACEMENT - Jan 2025.xlsx
Programmatic CREATIVE - Jan 2025.xlsx
Search CAMPAIGN - Jan 2025.xlsx
```
| Platform | Valid Levels |
|---|---|
| AdServer | CAMPAIGN · PLACEMENT · CREATIVE |
| Paid Social | CAMPAIGN · PLACEMENT · CREATIVE |
| Programmatic | CAMPAIGN · PLACEMENT · CREATIVE |
| Search | CAMPAIGN · PLACEMENT · CREATIVE |

**Output files are named:** `Platform_Market_Date.xlsx`
> e.g. `Programmatic_Germany_Jan 2025.xlsx`

**Notes:**
- The `Market` column must exist in every input file
- `Free Text` columns are excluded from **Is missing** flagging and red highlighting
- Columns **Clicks**, **Currency**, and **Spend** are removed from all outputs
- NC columns are removed after flagging
- On the **CREATIVE** tab, only the column named exactly **Influencer** is checked and highlighted
""")

# ═══════════════════════════════════════════════════════════════════════════════
#  UI — UPLOAD
# ═══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="mg-label">Upload Input Files</div>', unsafe_allow_html=True)
uploaded_files = st.file_uploader(
    label                 = "Drop your .xlsx files here or click to browse",
    type                  = ["xlsx"],
    accept_multiple_files = True,
    label_visibility      = "collapsed",
)
if uploaded_files:
    st.info(f"**{len(uploaded_files)}** file(s) selected and ready to process.")
st.markdown("<br>", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
#  UI — RUN BUTTON
# ═══════════════════════════════════════════════════════════════════════════════
run_clicked = st.button(
    "▶   Run Segregation",
    disabled = not bool(uploaded_files),
)

# ═══════════════════════════════════════════════════════════════════════════════
#  PROCESSING
# ═══════════════════════════════════════════════════════════════════════════════
if run_clicked and uploaded_files:
    progress_bar = st.progress(0, text="Initialising…")

    def update_progress(fraction: float):
        pct = min(int(fraction * 100), 99)
        progress_bar.progress(pct, text=f"Processing…  {pct}%")

    with st.spinner("Running segregation — please wait…"):
        try:
            output_files, stats = run_segregation(
                uploaded_files,
                progress_cb=update_progress,
            )
        except Exception as exc:
            st.error(f"Fatal error: {exc}")
            st.stop()

    progress_bar.progress(100, text="Done ✅")
    st.markdown("---")

    # ── Stats ─────────────────────────────────────────────────────────────────
    st.markdown("### ✅ Segregation Complete")
    c1, c2, c3 = st.columns(3)
    c1.metric("Platforms Processed",  stats["platforms_processed"])
    c2.metric("Output Files Created", stats["files_created"])
    c3.metric("Tabs Written",         stats["tabs_created"])

    # ── Skipped files ─────────────────────────────────────────────────────────
    if stats["skipped"]:
        with st.expander(
            f"⚠️  {len(stats['skipped'])} file(s) skipped — click to review",
            expanded=True,
        ):
            for msg in stats["skipped"]:
                st.markdown(f"- `{msg}`")

    # ── Download ──────────────────────────────────────────────────────────────
    if output_files:
        st.markdown("---")
        st.markdown('<div class="mg-label">Download Output</div>', unsafe_allow_html=True)
        if len(output_files) == 1:
            fname, data = next(iter(output_files.items()))
            st.download_button(
                label     = f"⬇️  Download  {fname}",
                data      = data,
                file_name = fname,
                mime      = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            zip_bytes = build_zip(output_files)
            st.download_button(
                label     = f"⬇️  Download All  ({len(output_files)} files)  as ZIP",
                data      = zip_bytes,
                file_name = "Magnum_Segregated_Output.zip",
                mime      = "application/zip",
            )
            with st.expander(f"📂  Files included in the ZIP  ({len(output_files)})"):
                for fname in sorted(output_files.keys()):
                    st.markdown(f"- `{fname}`")
    else:
        st.warning(
            "No output files were generated. "
            "Check that your filenames match the required format above."
        )

# ═══════════════════════════════════════════════════════════════════════════════
#  FOOTER
# ═══════════════════════════════════════════════════════════════════════════════
st.markdown(
    '<div class="mg-footer">© Magnum Analytics &nbsp;·&nbsp; Internal Use Only</div>',
    unsafe_allow_html=True,
)
