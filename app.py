import base64
import datetime
import tempfile
from pathlib import Path

import streamlit as st

_WIB = datetime.timezone(datetime.timedelta(hours=7))

def now_wib() -> datetime.datetime:
    return datetime.datetime.now(datetime.timezone.utc).astimezone(_WIB)


# =========================
# PAGE CONFIG
# =========================
st.set_page_config(
    page_title="HSD AGENT",
    page_icon="🧄",
    layout="wide",
    initial_sidebar_state="expanded",
)


# =========================
# SAFE IMPORTS
# =========================
try:
    import parser as hsd_parser
except Exception as e:
    hsd_parser = None
    PARSER_IMPORT_ERROR = e
else:
    PARSER_IMPORT_ERROR = None

try:
    import excel_writer as hsd_excel
except Exception as e:
    hsd_excel = None
    EXCEL_IMPORT_ERROR = e
else:
    EXCEL_IMPORT_ERROR = None


# =========================
# CONSTANTS
# =========================
ROLES = {
    "1234": "Gudang",
    "2345": "Konten",
    "3456": "Live",
    "0000": "Manager/BOD",
}

DIVISIONS = ["Gudang", "Konten", "Live"]

UPLOAD_GROUPS = [
    ("HSD Jakarta Pagi",   "HSD Jakarta",  "Pagi"),
    ("HSD Jakarta Siang",  "HSD Jakarta",  "Siang"),
    ("HSD Jakarta Sore",   "HSD Jakarta",  "Sore"),
    ("HSD Surabaya Pagi",  "HSD Surabaya", "Pagi"),
    ("HSD Surabaya Siang", "HSD Surabaya", "Siang"),
    ("HSD Surabaya Sore",  "HSD Surabaya", "Sore"),
    ("HSS Pagi",           "HSS",          "Pagi"),
    ("HSS Siang",          "HSS",          "Siang"),
    ("HSS Sore",           "HSS",          "Sore"),
]

PDF_SPLIT_THRESHOLD_MB = 10
PDF_TARGET_PAGES       = 300


# =========================
# LOGO HELPER
# =========================
def get_logo_b64() -> str | None:
    logo_path = Path("HSD-Logo1.png")
    if logo_path.exists():
        with open(str(logo_path), "rb") as f:
            return base64.b64encode(f.read()).decode()
    return None


# =========================
# CSS — MODERN REDESIGN
# =========================
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap');

    /* ── Base ─────────────────────────────────────────────── */
    html, body, [class*="css"] {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif !important;
    }

    .stApp {
        background: #F5F5F7;
        color: #09090B;
    }

    header[data-testid="stHeader"] {
        background: rgba(245, 245, 247, 0.8) !important;
        backdrop-filter: blur(16px);
        border-bottom: 1px solid rgba(0,0,0,0.07);
    }

    .block-container {
        padding-top: 2.2rem !important;
        padding-left: 2.4rem !important;
        padding-right: 2.4rem !important;
        max-width: 1440px !important;
    }

    /* ── Sidebar ──────────────────────────────────────────── */
    section[data-testid="stSidebar"] {
        display: block !important;
        visibility: visible !important;
        opacity: 1 !important;
        background: transparent !important;
        border-right: none !important;
        width: 260px !important;
        min-width: 260px !important;
        max-width: 260px !important;
        padding: 12px 10px !important;
    }

    section[data-testid="stSidebar"] > div {
        background: #0C0C0E !important;
        border-radius: 16px !important;
        padding: 24px 18px !important;
        margin: 8px 6px 8px 8px !important;
        box-shadow: 0 4px 24px rgba(0,0,0,0.18) !important;
        min-height: calc(100vh - 32px) !important;
        box-sizing: border-box !important;
    }

    section[data-testid="stSidebar"] * {
        color: #E4E4E7 !important;
    }

    section[data-testid="stSidebar"] hr {
        border-color: rgba(255,255,255,0.07) !important;
        margin: 14px 0 !important;
    }

    section[data-testid="stSidebar"] .stCaption,
    section[data-testid="stSidebar"] caption {
        color: #71717A !important;
        font-size: 12px !important;
    }

    /* Sidebar nav radio */
    section[data-testid="stSidebar"] [role="radiogroup"] label {
        background: rgba(255,255,255,0.03) !important;
        border: 1px solid rgba(255,255,255,0.07) !important;
        border-radius: 10px !important;
        padding: 10px 14px !important;
        margin-bottom: 6px !important;
        font-weight: 600 !important;
        font-size: 14px !important;
        transition: all 0.15s ease !important;
        cursor: pointer !important;
    }

    section[data-testid="stSidebar"] [role="radiogroup"] label:has(input:checked) {
        background: linear-gradient(135deg, #F97316 0%, #FBBF24 100%) !important;
        border-color: transparent !important;
        box-shadow: 0 4px 14px rgba(249, 115, 22, 0.35) !important;
        color: #FFFFFF !important;
    }

    section[data-testid="stSidebar"] [role="radiogroup"] label:has(input:checked) * {
        color: #FFFFFF !important;
    }

    section[data-testid="stSidebar"] [role="radiogroup"] label:hover:not(:has(input:checked)) {
        background: rgba(255,255,255,0.07) !important;
        border-color: rgba(255,255,255,0.12) !important;
    }

    /* ── SIDEBAR ARROW — FIX VISIBILITAS ─────────────────── */
    section[data-testid="stSidebar"] button[data-testid="stSidebarCollapseButton"],
    section[data-testid="stSidebar"] button[data-testid="stSidebarNavCollapseButton"],
    section[data-testid="stSidebar"] button[kind="header"] {
        background: rgba(249, 115, 22, 0.12) !important;
        border: 1px solid rgba(249, 115, 22, 0.25) !important;
        border-radius: 8px !important;
        padding: 6px !important;
    }
    section[data-testid="stSidebar"] button[data-testid="stSidebarCollapseButton"] svg,
    section[data-testid="stSidebar"] button[data-testid="stSidebarNavCollapseButton"] svg,
    section[data-testid="stSidebar"] button[kind="header"] svg {
        fill: #F97316 !important;
        stroke: #F97316 !important;
    }

    button[data-testid="stSidebarCollapseButton"],
    button[data-testid="stSidebarNavCollapseButton"],
    button[data-testid="collapsedControl"],
    button[data-testid="baseButton-headerNoPadding"] {
        background: #1C1C1E !important;
        border: 1px solid #3A3A3C !important;
        border-radius: 8px !important;
        padding: 6px !important;
        box-shadow: 0 2px 8px rgba(0,0,0,0.2) !important;
    }
    button[data-testid="stSidebarCollapseButton"] svg,
    button[data-testid="stSidebarNavCollapseButton"] svg,
    button[data-testid="collapsedControl"] svg,
    button[data-testid="baseButton-headerNoPadding"] svg {
        fill: #F97316 !important;
        stroke: #F97316 !important;
        color: #F97316 !important;
    }

    /* ── Cards ────────────────────────────────────────────── */
    .hsd-card {
        background: #FFFFFF;
        border: 1px solid #E4E4E7;
        border-radius: 16px;
        padding: 22px 24px;
        box-shadow: 0 1px 2px rgba(0,0,0,0.04), 0 4px 16px rgba(0,0,0,0.04);
        margin-bottom: 16px;
    }

    /* ── Typography ───────────────────────────────────────── */
    .hsd-title {
        font-size: 30px;
        font-weight: 800;
        color: #09090B;
        letter-spacing: -0.04em;
        line-height: 1.15;
        margin-bottom: 4px;
    }

    .hsd-subtitle {
        font-size: 14px;
        color: #71717A;
        font-weight: 500;
        margin-bottom: 0;
        line-height: 1.5;
    }

    /* ── Metric Cards ─────────────────────────────────────── */
    .metric-card {
        background: #FFFFFF;
        border: 1px solid #E4E4E7;
        border-radius: 14px;
        padding: 18px 20px;
        box-shadow: 0 1px 2px rgba(0,0,0,0.04);
    }

    .metric-label {
        font-size: 11px;
        font-weight: 600;
        color: #A1A1AA;
        text-transform: uppercase;
        letter-spacing: 0.06em;
        margin-bottom: 8px;
    }

    .metric-value {
        font-size: 22px;
        font-weight: 800;
        color: #09090B;
        letter-spacing: -0.025em;
    }

    /* ── Pills ────────────────────────────────────────────── */
    .hsd-pill {
        display: inline-flex;
        align-items: center;
        gap: 4px;
        padding: 3px 10px;
        border-radius: 6px;
        font-size: 12px;
        font-weight: 600;
        margin-right: 5px;
        margin-bottom: 4px;
        letter-spacing: 0.01em;
    }

    .pill-blue   { color: #1D4ED8; background: #EFF6FF; border: 1px solid #BFDBFE; }
    .pill-teal   { color: #0E7490; background: #ECFEFF; border: 1px solid #A5F3FC; }
    .pill-orange { color: #C2410C; background: #FFF7ED; border: 1px solid #FED7AA; }
    .pill-dark   { color: #F4F4F5; background: #18181B; }
    .pill-green  { color: #065F46; background: #ECFDF5; border: 1px solid #A7F3D0; }

    /* ── Buttons ──────────────────────────────────────────── */

    /* Secondary button (default / Reset) */
    div.stButton > button {
        background: #FFFFFF !important;
        color: #3F3F46 !important;
        border: 1.5px solid #D4D4D8 !important;
        border-radius: 10px !important;
        padding: 0.65rem 1rem !important;
        font-weight: 600 !important;
        font-size: 14px !important;
        box-shadow: 0 1px 3px rgba(0,0,0,0.06) !important;
        transition: all 0.15s ease !important;
        letter-spacing: -0.01em !important;
    }

    div.stButton > button * {
        color: #3F3F46 !important;
        font-weight: 600 !important;
    }

    div.stButton > button:hover {
        background: #FAFAFA !important;
        border-color: #A1A1AA !important;
        transform: translateY(-1px) !important;
        box-shadow: 0 3px 8px rgba(0,0,0,0.09) !important;
    }

    /* Primary button (Proses PDF) */
    button[kind="primary"] {
        background: linear-gradient(135deg, #F97316 0%, #FBBF24 100%) !important;
        color: #FFFFFF !important;
        border: 0 !important;
        border-radius: 12px !important;
        padding: 0.8rem 1.5rem !important;
        font-weight: 700 !important;
        font-size: 15px !important;
        box-shadow: 0 4px 16px rgba(249, 115, 22, 0.38) !important;
        transition: all 0.15s ease !important;
        letter-spacing: -0.01em !important;
    }

    button[kind="primary"] * {
        color: #FFFFFF !important;
        font-weight: 700 !important;
    }

    button[kind="primary"]:hover {
        background: linear-gradient(135deg, #EA580C 0%, #F59E0B 100%) !important;
        box-shadow: 0 6px 22px rgba(249, 115, 22, 0.48) !important;
        transform: translateY(-1px) !important;
    }

    button[kind="primary"]:disabled,
    button[kind="primary"][disabled] {
        background: #D4D4D8 !important;
        box-shadow: none !important;
        transform: none !important;
        cursor: not-allowed !important;
        opacity: 0.6 !important;
    }

    /* Download button */
    div.stDownloadButton > button,
    div[data-testid="stFormSubmitButton"] button {
        width: 100% !important;
        background: #09090B !important;
        color: #FFFFFF !important;
        border: 0 !important;
        border-radius: 12px !important;
        padding: 0.8rem 1rem !important;
        font-weight: 700 !important;
        font-size: 15px !important;
        box-shadow: 0 4px 14px rgba(9,9,11,0.22) !important;
        transition: all 0.15s ease !important;
    }

    div.stDownloadButton > button *,
    div[data-testid="stFormSubmitButton"] button * {
        color: #FFFFFF !important;
        font-weight: 700 !important;
    }

    div.stDownloadButton > button:hover {
        background: #18181B !important;
        box-shadow: 0 6px 20px rgba(9,9,11,0.3) !important;
        transform: translateY(-1px) !important;
    }

    /* ── File Uploader ────────────────────────────────────── */
    div[data-testid="stFileUploader"] {
        background: #FAFAFA;
        border: 1.5px dashed #D4D4D8;
        border-radius: 12px;
        padding: 8px;
        transition: border-color 0.15s ease;
    }

    div[data-testid="stFileUploader"]:hover {
        border-color: #F97316;
    }

    div[data-testid="stFileUploader"] section {
        background: transparent !important;
        border-radius: 10px !important;
    }

    div[data-testid="stFileUploader"] * {
        color: #3F3F46 !important;
        font-weight: 500 !important;
        opacity: 1 !important;
    }

    div[data-testid="stFileUploader"] button {
        background: #09090B !important;
        color: #FFFFFF !important;
        border-radius: 8px !important;
        font-weight: 600 !important;
        font-size: 13px !important;
        border: 0 !important;
    }

    div[data-testid="stFileUploader"] button * {
        color: #FFFFFF !important;
    }

    /* ── Upload Section Headers ───────────────────────────── */
    .upload-section-header {
        font-size: 13px;
        font-weight: 700;
        padding: 8px 14px;
        border-radius: 10px;
        margin-bottom: 10px;
        margin-top: 2px;
        letter-spacing: 0.01em;
    }

    .hdr-jkt {
        color: #1D4ED8;
        background: #EFF6FF;
        border: 1px solid #BFDBFE;
    }

    .hdr-sby {
        color: #0E7490;
        background: #ECFEFF;
        border: 1px solid #A5F3FC;
    }

    .hdr-hss {
        color: #B45309;
        background: #FFFBEB;
        border: 1px solid #FDE68A;
    }

    /* ── Progress bar ─────────────────────────────────────── */
    div[data-testid="stProgress"] > div {
        background: #E4E4E7 !important;
        border-radius: 99px !important;
    }

    div[data-testid="stProgress"] > div > div {
        background: linear-gradient(90deg, #F97316, #FBBF24) !important;
        border-radius: 99px !important;
        transition: width 0.3s ease !important;
    }

    /* ── Info/Status boxes ────────────────────────────────── */
    .success-box {
        background: #F0FDF4;
        border: 1px solid #BBF7D0;
        border-radius: 12px;
        padding: 14px 18px;
        font-weight: 600;
        font-size: 14px;
        color: #166534;
        margin-top: 12px;
    }

    .split-info-box {
        background: #FFF7ED;
        border: 1px solid #FED7AA;
        border-radius: 12px;
        padding: 12px 16px;
        font-size: 13px;
        font-weight: 500;
        color: #9A3412;
        margin-top: 8px;
        line-height: 1.6;
    }

    .small-muted {
        font-size: 13px;
        color: #71717A;
        font-weight: 500;
        line-height: 1.65;
    }

    /* ── Streamlit native alert overrides ─────────────────── */
    div[data-testid="stAlert"] {
        border-radius: 12px !important;
        font-weight: 500 !important;
        font-size: 14px !important;
    }

    /* ── Expander ─────────────────────────────────────────── */
    details summary {
        font-weight: 600 !important;
        font-size: 14px !important;
    }

    /* ── Status widget ────────────────────────────────────── */
    div[data-testid="stStatusWidget"] {
        border-radius: 12px !important;
        font-weight: 500 !important;
    }

    /* ── Sidebar Keluar button ────────────────────────────── */
    section[data-testid="stSidebar"] div.stButton > button {
        background: rgba(255,255,255,0.05) !important;
        color: #E4E4E7 !important;
        border: 1px solid rgba(255,255,255,0.10) !important;
        border-radius: 10px !important;
        font-weight: 600 !important;
    }

    section[data-testid="stSidebar"] div.stButton > button * {
        color: #E4E4E7 !important;
    }

    section[data-testid="stSidebar"] div.stButton > button:hover {
        background: rgba(249,115,22,0.15) !important;
        border-color: rgba(249,115,22,0.3) !important;
        color: #F97316 !important;
    }

    section[data-testid="stSidebar"] div.stButton > button:hover * {
        color: #F97316 !important;
    }

    /* ── Clock card ───────────────────────────────────────── */
    .clock-card {
        background: #FFFFFF;
        border: 1px solid #E4E4E7;
        border-radius: 14px;
        padding: 14px 18px;
        text-align: right;
        box-shadow: 0 1px 2px rgba(0,0,0,0.04);
    }

    .clock-time {
        font-size: 22px;
        font-weight: 800;
        color: #09090B;
        letter-spacing: -0.03em;
        line-height: 1.2;
    }

    .clock-label {
        font-size: 11px;
        font-weight: 600;
        color: #A1A1AA;
        text-transform: uppercase;
        letter-spacing: 0.06em;
    }

    /* ── Upload info banner ───────────────────────────────── */
    .upload-banner {
        background: #FFFFFF;
        border: 1px solid #E4E4E7;
        border-radius: 14px;
        padding: 16px 20px;
        margin-bottom: 20px;
    }

    .upload-banner-title {
        font-size: 16px;
        font-weight: 800;
        color: #09090B;
        letter-spacing: -0.02em;
        margin-bottom: 6px;
    }

    </style>
    """,
    unsafe_allow_html=True,
)


# =========================
# SESSION STATE
# =========================
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

if "role" not in st.session_state:
    st.session_state.role = None

if "selected_menu" not in st.session_state:
    st.session_state.selected_menu = "Gudang"

if "upload_reset_counter" not in st.session_state:
    st.session_state.upload_reset_counter = 0

if "last_excel_bytes" not in st.session_state:
    st.session_state.last_excel_bytes = None

if "last_excel_filename" not in st.session_state:
    st.session_state.last_excel_filename = None


# =========================
# AUTO SPLIT LARGE PDF
# =========================
def auto_split_large_pdfs(saved_files: list, status_writer=None) -> list:
    try:
        from pypdf import PdfReader, PdfWriter
    except ImportError:
        if status_writer:
            status_writer("⚠️ pypdf tidak tersedia, PDF besar diproses langsung.")
        return saved_files

    result = []
    temp_dir = Path(tempfile.mkdtemp(prefix="hsd_split_"))

    for item in saved_files:
        file_path = Path(item["path"])
        file_size_mb = file_path.stat().st_size / (1024 * 1024)

        if file_size_mb <= PDF_SPLIT_THRESHOLD_MB:
            result.append(item)
            continue

        if status_writer:
            status_writer(
                f"📦 PDF besar terdeteksi: **{item['filename']}** "
                f"({file_size_mb:.1f} MB) — sedang dioptimalkan..."
            )

        try:
            reader      = PdfReader(str(file_path))
            total_pages = len(reader.pages)
            pages_per_part = max(PDF_TARGET_PAGES, 1)
            total_parts    = (total_pages + pages_per_part - 1) // pages_per_part
            base_name      = file_path.stem

            for part_idx in range(total_parts):
                start_page = part_idx * pages_per_part
                end_page   = min(start_page + pages_per_part, total_pages)

                if status_writer:
                    status_writer(
                        f"✂️ Memotong **{item['filename']}** — "
                        f"Bagian {part_idx + 1}/{total_parts} "
                        f"(hal {start_page + 1}–{end_page} dari {total_pages})"
                    )

                writer = PdfWriter()
                for page_num in range(start_page, end_page):
                    writer.add_page(reader.pages[page_num])

                part_filename = f"{base_name}_part_{part_idx + 1}.pdf"
                part_path     = temp_dir / part_filename

                with open(str(part_path), "wb") as f:
                    writer.write(f)

                result.append({
                    "path":       str(part_path),
                    "filename":   part_filename,
                    "brand":      item["brand"],
                    "akun":       item["akun"],
                    "shift":      item["shift"],
                    "group":      item["group"],
                    "split_from": item["filename"],
                    "part":       f"{part_idx + 1}/{total_parts}",
                })

            if status_writer:
                status_writer(
                    f"✅ **{item['filename']}** berhasil dipotong jadi "
                    f"{total_parts} bagian ({total_pages} halaman total)"
                )

        except Exception as e:
            if status_writer:
                status_writer(f"⚠️ Split gagal untuk {item['filename']}: {e} — diproses langsung.")
            result.append(item)

    return result


# =========================
# HELPERS
# =========================
def save_uploaded_files(upload_map: dict) -> list:
    saved = []
    temp_dir = Path(tempfile.mkdtemp(prefix="hsd_agent_"))

    for group_name, payload in upload_map.items():
        brand = payload["brand"]
        shift = payload["shift"]
        files = payload["files"] or []

        for uploaded in files:
            safe_name = uploaded.name.replace("/", "_").replace("\\", "_")
            file_path = temp_dir / safe_name
            file_path.write_bytes(uploaded.getbuffer())

            saved.append({
                "path":     str(file_path),
                "filename": uploaded.name,
                "brand":    brand,
                "akun":     brand,
                "shift":    shift,
                "group":    group_name,
            })

    return saved


def call_parser(saved_files: list, progress_bar=None, status_text=None):
    if hsd_parser is None:
        raise RuntimeError(f"parser.py belum terbaca. Error: {PARSER_IMPORT_ERROR}")

    if not hasattr(hsd_parser, "process_pdf") or not callable(hsd_parser.process_pdf):
        raise RuntimeError("parser.py harus punya function process_pdf(pdf_path, progress_callback=None).")

    all_rows    = []
    all_errors  = []
    total_files = len(saved_files)

    for file_idx, item in enumerate(saved_files, 1):
        label = item["filename"]
        if item.get("split_from"):
            label = f"{item['split_from']} (bagian {item.get('part', file_idx)})"

        if status_text:
            status_text.info(f"📄 Membaca PDF {file_idx}/{total_files}: {label}")

        def page_progress(page_done, page_total, _idx=file_idx):
            if progress_bar and page_total:
                overall = ((_idx - 1) + (page_done / page_total)) / total_files
                progress_bar.progress(min(max(overall, 0), 1))

        result = hsd_parser.process_pdf(item["path"], progress_callback=page_progress)

        rows   = []
        errors = []

        if isinstance(result, tuple):
            rows   = result[0] if len(result) > 0 else []
            errors = result[1] if len(result) > 1 else []
        elif isinstance(result, list):
            rows = result
        elif isinstance(result, dict):
            rows = [result]

        for row in rows:
            if isinstance(row, dict):
                row.setdefault("brand",       item["brand"])
                row.setdefault("akun",        item["brand"])
                row.setdefault("shift",       item["shift"])
                row.setdefault("waktu",       item["shift"])
                row.setdefault("source_file", item["filename"])
                row.setdefault("group",       item["group"])
            all_rows.append(row)

        for err in errors:
            all_errors.append(f"{label} - {err}")

    if progress_bar:
        progress_bar.progress(1.0)

    return all_rows, all_errors


def call_excel_writer(parsed_data, saved_files: list) -> bytes:
    if hsd_excel is None:
        raise RuntimeError(f"excel_writer.py belum terbaca. Error: {EXCEL_IMPORT_ERROR}")

    output_path = Path(tempfile.mkdtemp(prefix="hsd_excel_")) / "rekap_resi_hsd.xlsx"

    candidate_names = [
        "write_excel", "write_excel_multi", "create_excel",
        "generate_excel", "build_excel", "make_excel", "export_excel",
    ]

    last_error = None

    for name in candidate_names:
        fn = getattr(hsd_excel, name, None)
        if not callable(fn):
            continue

        for args in [
            (parsed_data, str(output_path)),
            (parsed_data, saved_files, str(output_path)),
            (parsed_data,),
        ]:
            try:
                result = fn(*args)
                if isinstance(result, bytes):
                    return result
                if isinstance(result, (str, Path)) and Path(result).exists():
                    return Path(result).read_bytes()
                if output_path.exists():
                    return output_path.read_bytes()
            except TypeError as e:
                last_error = e
                continue

    if last_error:
        raise RuntimeError(f"Excel writer ditemukan, tapi argumennya tidak cocok: {last_error}")

    raise RuntimeError("Tidak menemukan function excel writer yang cocok.")


def render_metric(label: str, value: str):
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-label">{label}</div>
            <div class="metric-value">{value}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def do_logout():
    st.session_state.logged_in = False
    st.session_state.role = None
    st.rerun()


# =========================
# LOGIN PAGE
# =========================
def login_page():
    left, middle, right = st.columns([1, 1.1, 1])

    with middle:
        st.markdown("<br><br>", unsafe_allow_html=True)

        logo_b64 = get_logo_b64()
        if logo_b64:
            st.markdown(
                f"""
                <div class="hsd-card" style="text-align:center; padding:36px 28px;">
                    <img src="data:image/png;base64,{logo_b64}" width="88"
                         style="margin-bottom:16px; display:block; margin-left:auto; margin-right:auto; border-radius:18px;">
                    <div style="font-size:28px; font-weight:900; color:#09090B; letter-spacing:-0.04em;">HSD AGENT</div>
                    <div style="font-size:13px; color:#71717A; margin-top:4px; font-weight:500;">Operational System</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        else:
            st.markdown(
                """
                <div class="hsd-card" style="text-align:center; padding:36px 28px;">
                    <div style="font-size:48px; margin-bottom:12px;">🧄</div>
                    <div style="font-size:28px; font-weight:900; color:#09090B; letter-spacing:-0.04em;">HSD AGENT</div>
                    <div style="font-size:13px; color:#71717A; margin-top:4px; font-weight:500;">Operational System</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

        with st.form("login_form"):
            pin = st.text_input("Masukkan PIN", type="password", placeholder="PIN divisi kamu")
            login = st.form_submit_button("Masuk →")

        if login:
            if pin in ROLES:
                st.session_state.logged_in = True
                st.session_state.role = ROLES[pin]
                st.session_state.selected_menu = "Gudang" if ROLES[pin] == "Manager/BOD" else ROLES[pin]
                st.rerun()
            else:
                st.error("PIN salah. Coba lagi.")

        st.caption("Gudang: 1234 · Konten: 2345 · Live: 3456 · Manager/BOD: 0000")


# =========================
# SIDEBAR
# =========================
def render_sidebar():
    with st.sidebar:
        logo_b64 = get_logo_b64()
        if logo_b64:
            st.markdown(
                f"""
                <div style="text-align:center; margin-bottom:4px; padding-bottom:4px;">
                    <img src="data:image/png;base64,{logo_b64}" width="72"
                         style="display:block; margin:0 auto; border-radius:14px;">
                    <div style="font-size:15px; font-weight:800; color:#F4F4F5; margin-top:10px; letter-spacing:0.03em;">HSD AGENT</div>
                    <div style="font-size:11px; color:#71717A; font-weight:500; margin-top:2px;">Operational System</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        else:
            st.markdown(
                """
                <div style="text-align:center; margin-bottom:4px;">
                    <div style="font-size:36px;">🧄</div>
                    <div style="font-size:15px; font-weight:800; color:#F4F4F5; margin-top:8px;">HSD AGENT</div>
                    <div style="font-size:11px; color:#71717A; font-weight:500;">Operational System</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

        st.markdown("---")

        role = st.session_state.role or "-"
        st.markdown(
            f"""
            <div style="margin-bottom:4px;">
                <span style="font-size:11px; color:#71717A; font-weight:600; text-transform:uppercase; letter-spacing:0.05em;">Role aktif</span><br>
                <span style="font-size:14px; font-weight:700; color:#F4F4F5;">{role}</span>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.caption(now_wib().strftime("%A, %d %B %Y · %H:%M WIB"))

        st.markdown("---")

        if role == "Manager/BOD":
            available_menu = DIVISIONS
        else:
            available_menu = [role] if role in DIVISIONS else ["Gudang"]

        default_index = 0
        if st.session_state.selected_menu in available_menu:
            default_index = available_menu.index(st.session_state.selected_menu)

        st.markdown(
            "<div style='font-size:11px; font-weight:600; color:#71717A; text-transform:uppercase; letter-spacing:0.05em; margin-bottom:8px;'>Menu Divisi</div>",
            unsafe_allow_html=True,
        )

        selected = st.radio("", available_menu, index=default_index, label_visibility="collapsed")
        st.session_state.selected_menu = selected

        st.markdown("---")

        if st.button("Keluar"):
            do_logout()


# =========================
# PAGE COMPONENTS
# =========================
def page_header(title: str, subtitle: str):
    top_left, top_right = st.columns([3, 1])

    with top_left:
        st.markdown(f"<div class='hsd-title'>{title}</div>", unsafe_allow_html=True)
        st.markdown(f"<div class='hsd-subtitle'>{subtitle}</div>", unsafe_allow_html=True)

    with top_right:
        st.markdown(
            f"""
            <div class="clock-card">
                <div class="clock-label">Waktu sekarang</div>
                <div class="clock-time">{now_wib().strftime('%H:%M')}</div>
                <div class="clock-label">WIB</div>
            </div>
            """,
            unsafe_allow_html=True,
        )


# =========================
# GUDANG PAGE
# =========================
def gudang_page():
    page_header(
        "Gudang",
        "Upload PDF resi HSD Jakarta / HSD Surabaya / HSS, proses otomatis, lalu download Excel rekap.",
    )

    st.markdown("<br>", unsafe_allow_html=True)

    c1, c2, c3 = st.columns(3)
    with c1: render_metric("Limit Upload", "2 GB")
    with c2: render_metric("Format Output", "PDF → Excel")
    with c3: render_metric("Zona Waktu", "WIB")

    st.markdown("<br>", unsafe_allow_html=True)

    st.markdown(
        """
        <div class="upload-banner">
            <div style="display:flex; align-items:center; gap:8px; margin-bottom:10px; flex-wrap:wrap;">
                <span class="hsd-pill pill-blue">🏙️ HSD Jakarta</span>
                <span class="hsd-pill pill-teal">🌊 HSD Surabaya</span>
                <span class="hsd-pill pill-orange">⭐ HSS</span>
                <span class="hsd-pill pill-dark">Gudang</span>
            </div>
            <div class="upload-banner-title">Upload PDF Resi</div>
            <div class="small-muted">
                Upload sesuai brand, cabang, dan shift. Bisa upload lebih dari satu file di setiap bagian.<br>
                PDF besar (&gt;10 MB) dipotong otomatis — tidak perlu split manual.<br>
                <strong style="color:#09090B;">Order offline tidak masuk sistem</strong> — isi manual di Excel setelah download.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    upload_map = {}

    col_jkt, col_sby, col_hss = st.columns(3)

    with col_jkt:
        st.markdown("<div class='upload-section-header hdr-jkt'>🏙️ HSD Jakarta</div>", unsafe_allow_html=True)
        for label, brand, shift in UPLOAD_GROUPS[:3]:
            files = st.file_uploader(
                label, type=["pdf"], accept_multiple_files=True,
                key=f"upload_{brand}_{shift}_{st.session_state.upload_reset_counter}",
            )
            upload_map[label] = {"brand": brand, "shift": shift, "files": files}

    with col_sby:
        st.markdown("<div class='upload-section-header hdr-sby'>🌊 HSD Surabaya</div>", unsafe_allow_html=True)
        for label, brand, shift in UPLOAD_GROUPS[3:6]:
            files = st.file_uploader(
                label, type=["pdf"], accept_multiple_files=True,
                key=f"upload_{brand}_{shift}_{st.session_state.upload_reset_counter}",
            )
            upload_map[label] = {"brand": brand, "shift": shift, "files": files}

    with col_hss:
        st.markdown("<div class='upload-section-header hdr-hss'>⭐ HSS</div>", unsafe_allow_html=True)
        for label, brand, shift in UPLOAD_GROUPS[6:]:
            files = st.file_uploader(
                label, type=["pdf"], accept_multiple_files=True,
                key=f"upload_{brand}_{shift}_{st.session_state.upload_reset_counter}",
            )
            upload_map[label] = {"brand": brand, "shift": shift, "files": files}

    total_files = sum(len(p["files"] or []) for p in upload_map.values())
    large_files = []
    for payload in upload_map.values():
        for f in (payload["files"] or []):
            size_mb = len(f.getbuffer()) / (1024 * 1024)
            if size_mb > PDF_SPLIT_THRESHOLD_MB:
                large_files.append((f.name, size_mb))

    st.markdown("<br>", unsafe_allow_html=True)

    if total_files > 0:
        st.success(f"✅ {total_files} file PDF siap diproses.")
        if large_files:
            info_lines = "".join(
                f"<li style='margin:2px 0;'>{name} ({size:.1f} MB) — akan dipotong otomatis</li>"
                for name, size in large_files
            )
            st.markdown(
                f"""<div class="split-info-box">
                    ✂️ <strong>{len(large_files)} PDF besar terdeteksi</strong> — akan dioptimalkan sebelum diproses:
                    <ul style="margin:6px 0 0 16px; padding:0;">{info_lines}</ul>
                </div>""",
                unsafe_allow_html=True,
            )
    else:
        st.markdown(
            """
            <div style="background:#FAFAFA; border:1px solid #E4E4E7; border-radius:12px;
                        padding:14px 18px; font-size:14px; color:#71717A; font-weight:500;
                        display:flex; align-items:center; gap:8px;">
                📂 Belum ada PDF terupload — mulai upload di kolom di atas.
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown("<br>", unsafe_allow_html=True)

    reset_col, spacer_col, process_col = st.columns([1.2, 0.2, 2])

    with reset_col:
        reset_upload = st.button("🔄 Reset / Ganti File")

    with process_col:
        process = st.button(
            "🚀 Proses PDF Jadi Excel",
            disabled=(total_files == 0),
            help="Upload minimal 1 PDF dulu.",
            type="primary",
            use_container_width=True,
        )

    if reset_upload:
        st.session_state.upload_reset_counter += 1
        st.session_state.last_excel_bytes    = None
        st.session_state.last_excel_filename = None
        st.success("Upload dikosongkan. Silakan masukkan PDF baru.")
        st.rerun()

    if process:
        if total_files == 0:
            st.warning("Upload minimal 1 file PDF dulu.")
            return

        st.session_state.last_excel_bytes    = None
        st.session_state.last_excel_filename = None

        try:
            progress_bar = st.progress(0)
            status_text  = st.empty()

            with st.status("⏳ Memproses PDF menjadi Excel...", expanded=True) as status:

                st.write("📥 Menyimpan file sementara...")
                saved_files = save_uploaded_files(upload_map)
                st.write(f"📄 Total PDF terupload: {len(saved_files)}")

                has_large = any(
                    Path(item["path"]).stat().st_size / (1024 * 1024) > PDF_SPLIT_THRESHOLD_MB
                    for item in saved_files
                )

                if has_large:
                    st.write("✂️ Mengoptimalkan PDF besar...")
                    split_log = st.empty()

                    def split_status(msg):
                        split_log.info(msg)

                    saved_files = auto_split_large_pdfs(saved_files, status_writer=split_status)
                    split_log.success(
                        f"✅ Optimasi selesai. Total bagian PDF siap diproses: {len(saved_files)}"
                    )

                st.write(f"📄 Memulai proses parser ({len(saved_files)} PDF)...")
                parsed_data, parse_errors = call_parser(
                    saved_files,
                    progress_bar=progress_bar,
                    status_text=status_text,
                )

                parsed_count = len(parsed_data) if hasattr(parsed_data, "__len__") else "-"
                st.write(f"✅ Data terbaca: {parsed_count} baris")

                if parse_errors:
                    st.warning(f"Ada {len(parse_errors)} catatan parser.")
                    with st.expander("Lihat catatan parser"):
                        for err in parse_errors[:200]:
                            st.write(err)

                status_text.info("📊 Membuat file Excel...")
                st.write("📊 Membuat Excel rekap...")
                excel_bytes = call_excel_writer(parsed_data, saved_files)

                status.update(label="✅ Excel selesai dibuat", state="complete", expanded=False)

            filename = f"Rekap_Resi_HSD_{now_wib().strftime('%Y%m%d_%H%M')}_WIB.xlsx"
            st.session_state.last_excel_bytes    = excel_bytes
            st.session_state.last_excel_filename = filename

            st.markdown(
                """
                <div class="success-box">
                    ✅ Excel berhasil dibuat — klik tombol download di bawah.
                </div>
                """,
                unsafe_allow_html=True,
            )

        except Exception as e:
            st.error("Proses gagal.")
            st.exception(e)
            st.caption("Kalau error terjadi di parser/excel_writer, kirim isi error-nya.")

    if st.session_state.last_excel_bytes:
        st.markdown("<br>", unsafe_allow_html=True)
        st.download_button(
            label="⬇️ Download Excel Rekap",
            data=st.session_state.last_excel_bytes,
            file_name=st.session_state.last_excel_filename or "rekap_resi_hsd.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )


# =========================
# KONTEN PAGE
# =========================
def konten_page():
    page_header("Konten", "Area kerja divisi konten HSD.")
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(
        """
        <div class="hsd-card">
            <div style="font-size:18px; font-weight:800; color:#09090B; letter-spacing:-0.02em; margin-bottom:6px;">Coming Soon</div>
            <div class="small-muted">Kalender konten, ide script, approval, dan database asset.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# =========================
# LIVE PAGE
# =========================
def live_page():
    page_header("Live", "Area kerja divisi live HSD.")
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(
        """
        <div class="hsd-card">
            <div style="font-size:18px; font-weight:800; color:#09090B; letter-spacing:-0.02em; margin-bottom:6px;">Coming Soon</div>
            <div class="small-muted">Jadwal live, target GMV, host, produk, dan evaluasi performa.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# =========================
# MAIN ROUTER
# =========================
def main():
    if not st.session_state.logged_in:
        login_page()
        return

    render_sidebar()

    selected_menu = st.session_state.selected_menu

    if selected_menu == "Gudang":
        gudang_page()
    elif selected_menu == "Konten":
        konten_page()
    elif selected_menu == "Live":
        live_page()
    else:
        gudang_page()


if __name__ == "__main__":
    main()
