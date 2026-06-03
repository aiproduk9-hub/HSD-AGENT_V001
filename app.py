import streamlit as st
import datetime
import tempfile
from pathlib import Path
from zoneinfo import ZoneInfo

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
except Exception:
    hsd_parser = None

try:
    import excel_writer as hsd_excel
except Exception:
    hsd_excel = None


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
    ("HSD Pagi", "HSD", "Pagi"),
    ("HSD Siang", "HSD", "Siang"),
    ("HSD Sore", "HSD", "Sore"),
    ("HSS Pagi", "HSS", "Pagi"),
    ("HSS Siang", "HSS", "Siang"),
    ("HSS Sore", "HSS", "Sore"),
]


# =========================
# CSS
# =========================
st.markdown(
    """
    <style>
    .stApp {
        background: #F7F1E8;
        color: #111827;
    }

    header[data-testid="stHeader"] {
        background: rgba(247, 241, 232, 0.82) !important;
        backdrop-filter: blur(10px);
    }

    .block-container {
        padding-top: 2rem !important;
        padding-left: 2.4rem !important;
        padding-right: 2.4rem !important;
        max-width: 1400px !important;
    }

    /* SIDEBAR - JANGAN DIHAPUS */
    section[data-testid="stSidebar"] {
        display: block !important;
        visibility: visible !important;
        opacity: 1 !important;
        transform: translateX(0px) !important;
        background: #050505 !important;
        border-right: 1px solid #111827 !important;
        width: 290px !important;
        min-width: 290px !important;
        max-width: 290px !important;
    }

    section[data-testid="stSidebar"] > div {
        display: block !important;
        visibility: visible !important;
        opacity: 1 !important;
        background: #050505 !important;
        padding: 24px 18px !important;
    }

    section[data-testid="stSidebar"] * {
        color: #F9FAFB !important;
    }

    section[data-testid="stSidebar"] .stRadio label {
        color: #F9FAFB !important;
        font-weight: 600 !important;
    }

    section[data-testid="stSidebar"] [role="radiogroup"] label {
        background: #111827 !important;
        border: 1px solid #1F2937 !important;
        border-radius: 14px !important;
        padding: 10px 12px !important;
        margin-bottom: 8px !important;
    }

    section[data-testid="stSidebar"] [role="radiogroup"] label:hover {
        background: #1F2937 !important;
    }

    .hsd-card {
        background: #FFFFFF;
        border: 1px solid #E5E7EB;
        border-radius: 22px;
        padding: 22px;
        box-shadow: 0 8px 26px rgba(15, 23, 42, 0.07);
        margin-bottom: 18px;
    }

    .hsd-title {
        font-size: 34px;
        font-weight: 850;
        color: #111827;
        margin-bottom: 4px;
        letter-spacing: -0.03em;
    }

    .hsd-subtitle {
        font-size: 15px;
        color: #6B7280;
        margin-bottom: 20px;
    }

    .hsd-pill {
        display: inline-block;
        padding: 6px 12px;
        border-radius: 999px;
        font-size: 12px;
        font-weight: 800;
        margin-right: 6px;
    }

    .pill-blue {
        color: #1D4ED8;
        background: #DBEAFE;
    }

    .pill-orange {
        color: #C2410C;
        background: #FFEDD5;
    }

    .pill-dark {
        color: #F9FAFB;
        background: #111827;
    }

    .metric-card {
        background: #FFFFFF;
        border: 1px solid #E5E7EB;
        border-radius: 18px;
        padding: 18px;
        box-shadow: 0 8px 22px rgba(15, 23, 42, 0.05);
    }

    .metric-label {
        font-size: 13px;
        color: #6B7280;
        margin-bottom: 5px;
    }

    .metric-value {
        font-size: 26px;
        font-weight: 850;
        color: #111827;
    }

    div.stButton > button {
        width: 100%;
        background: #111827 !important;
        color: #FFFFFF !important;
        border: 0 !important;
        border-radius: 14px !important;
        padding: 0.75rem 1rem !important;
        font-weight: 800 !important;
    }

    div.stButton > button:hover {
        background: #0F172A !important;
        color: #FFFFFF !important;
        border: 0 !important;
    }

  div[data-testid="stFileUploader"] {
    background: #FFFFFF;
    border: 1px dashed #CBD5E1;
    border-radius: 18px;
    padding: 12px;
}

div[data-testid="stFileUploader"] section {
    background: #F8FAFC !important;
    border-radius: 14px !important;
}

/* FIX TEKS UPLOADER BIAR KELIHATAN */
div[data-testid="stFileUploader"] * {
    color: #111827 !important;
}

div[data-testid="stFileUploader"] label,
div[data-testid="stFileUploader"] small,
div[data-testid="stFileUploader"] p,
div[data-testid="stFileUploader"] span {
    color: #111827 !important;
    opacity: 1 !important;
}

div[data-testid="stFileUploader"] button {
    background: #111827 !important;
    color: #FFFFFF !important;
    border-radius: 12px !important;
    font-weight: 700 !important;
}

div[data-testid="stFileUploader"] button * {
    color: #FFFFFF !important;
}
    .small-muted {
        font-size: 13px;
        color: #6B7280;
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


# =========================
# HELPERS
# =========================
def now_wib() -> datetime.datetime:
    return datetime.datetime.now(ZoneInfo("Asia/Jakarta"))


def save_uploaded_files(upload_map: dict) -> list[dict]:
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

            saved.append(
                {
                    "path": str(file_path),
                    "filename": uploaded.name,
                    "brand": brand,
                    "shift": shift,
                    "group": group_name,
                }
            )

    return saved


def call_parser(saved_files: list[dict]):
    if hsd_parser is None:
        raise RuntimeError("File parser.py belum terbaca / error import. Pastikan parser.py ada di repo.")

    candidate_names = [
        "parse_uploaded_files",
        "parse_files",
        "parse_pdf_files",
        "parse_all_pdfs",
        "process_pdfs",
    ]

    for name in candidate_names:
        fn = getattr(hsd_parser, name, None)
        if callable(fn):
            try:
                return fn(saved_files)
            except TypeError:
                return fn([item["path"] for item in saved_files])

    single_candidates = ["parse_pdf", "extract_resi", "read_pdf"]

    for name in single_candidates:
        fn = getattr(hsd_parser, name, None)

        if callable(fn):
            rows = []

            for item in saved_files:
                result = fn(item["path"])

                if isinstance(result, list):
                    for row in result:
                        if isinstance(row, dict):
                            row.setdefault("brand", item["brand"])
                            row.setdefault("shift", item["shift"])
                            row.setdefault("source_file", item["filename"])
                        rows.append(row)

                elif isinstance(result, dict):
                    result.setdefault("brand", item["brand"])
                    result.setdefault("shift", item["shift"])
                    result.setdefault("source_file", item["filename"])
                    rows.append(result)

            return rows

    raise RuntimeError(
        "Tidak menemukan function parser yang cocok. Tambahkan salah satu: "
        "parse_uploaded_files(), parse_files(), parse_pdf_files(), parse_all_pdfs(), "
        "process_pdfs(), atau parse_pdf()."
    )


def call_excel_writer(parsed_data, saved_files: list[dict]) -> bytes:
    if hsd_excel is None:
        raise RuntimeError("File excel_writer.py belum terbaca / error import. Pastikan excel_writer.py ada di repo.")

    output_path = Path(tempfile.mkdtemp(prefix="hsd_excel_")) / "rekap_resi_hsd.xlsx"

    candidate_names = [
        "write_excel",
        "create_excel",
        "generate_excel",
        "build_excel",
        "make_excel",
        "export_excel",
    ]

    last_error = None

    for name in candidate_names:
        fn = getattr(hsd_excel, name, None)

        if callable(fn):
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
        raise RuntimeError(f"Excel writer ditemukan, tapi format argumennya tidak cocok: {last_error}")

    raise RuntimeError(
        "Tidak menemukan function excel writer yang cocok. Tambahkan salah satu: "
        "write_excel(), create_excel(), generate_excel(), build_excel(), make_excel(), atau export_excel()."
    )


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

        st.markdown(
            """
            <div class="hsd-card">
                <div style="font-size:42px; text-align:center;">🧄</div>
                <div style="font-size:32px; font-weight:900; text-align:center; color:#111827; letter-spacing:-0.04em;">HSD AGENT</div>
                <div style="font-size:14px; text-align:center; color:#6B7280; margin-bottom:18px;">Operational System</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        pin = st.text_input("Masukkan PIN", type="password", placeholder="PIN divisi")
        login = st.button("Masuk")

        if login:
            if pin in ROLES:
                st.session_state.logged_in = True
                st.session_state.role = ROLES[pin]

                if ROLES[pin] == "Manager/BOD":
                    st.session_state.selected_menu = "Gudang"
                else:
                    st.session_state.selected_menu = ROLES[pin]

                st.rerun()

            else:
                st.error("PIN salah. Coba lagi.")

        st.caption("PIN Gudang: 1234 | Konten: 2345 | Live: 3456 | Manager/BOD: 0000")


# =========================
# SIDEBAR
# =========================
def render_sidebar():
    with st.sidebar:
        st.markdown("# 🧄 HSD AGENT")
        st.caption("Operational System")
        st.markdown("---")

        role = st.session_state.role or "-"
        st.markdown(f"**Role aktif:** {role}")
        st.caption(now_wib().strftime("%A, %d %B %Y • %H:%M WIB"))
        st.markdown("---")

        if role == "Manager/BOD":
            available_menu = DIVISIONS
        else:
            available_menu = [role] if role in DIVISIONS else ["Gudang"]

        default_index = 0
        if st.session_state.selected_menu in available_menu:
            default_index = available_menu.index(st.session_state.selected_menu)

        selected = st.radio(
            "Menu Divisi",
            available_menu,
            index=default_index,
        )

        st.session_state.selected_menu = selected

        st.markdown("---")

        if st.button("Keluar"):
            do_logout()


# =========================
# PAGES
# =========================
def page_header(title: str, subtitle: str):
    top_left, top_right = st.columns([3, 1])

    with top_left:
        st.markdown(f"<div class='hsd-title'>{title}</div>", unsafe_allow_html=True)
        st.markdown(f"<div class='hsd-subtitle'>{subtitle}</div>", unsafe_allow_html=True)

    with top_right:
        st.markdown(
            f"""
            <div class="hsd-card" style="padding:16px; text-align:right;">
                <div class="small-muted">Waktu sekarang</div>
                <div style="font-weight:900; font-size:20px; color:#111827;">{now_wib().strftime('%H:%M')}</div>
                <div class="small-muted">WIB</div>
            </div>
            """,
            unsafe_allow_html=True,
        )


def gudang_page():
    page_header(
        "Gudang",
        "Upload PDF resi HSD/HSS, proses otomatis, lalu download Excel rekap.",
    )

    c1, c2, c3 = st.columns(3)

    with c1:
        render_metric("Limit Upload", "2GB")

    with c2:
        render_metric("Format", "PDF → Excel")

    with c3:
        render_metric("Zona Waktu", "WIB")

    st.markdown("<br>", unsafe_allow_html=True)

    st.markdown(
        """
        <div class="hsd-card">
            <span class="hsd-pill pill-blue">HSD</span>
            <span class="hsd-pill pill-orange">HSS</span>
            <span class="hsd-pill pill-dark">Gudang</span>
            <div style="font-size:20px; font-weight:850; color:#111827; margin-top:12px;">Upload PDF Resi</div>
            <div class="small-muted">Upload sesuai brand dan shift. Bisa upload lebih dari satu file di setiap bagian.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    upload_map = {}

    hsd_col, hss_col = st.columns(2)

    with hsd_col:
        st.markdown("### HSD")

        for label, brand, shift in UPLOAD_GROUPS[:3]:
            files = st.file_uploader(
                label,
                type=["pdf"],
                accept_multiple_files=True,
                key=f"upload_{brand}_{shift}",
            )

            upload_map[label] = {
                "brand": brand,
                "shift": shift,
                "files": files,
            }

    with hss_col:
        st.markdown("### HSS")

        for label, brand, shift in UPLOAD_GROUPS[3:]:
            files = st.file_uploader(
                label,
                type=["pdf"],
                accept_multiple_files=True,
                key=f"upload_{brand}_{shift}",
            )

            upload_map[label] = {
                "brand": brand,
                "shift": shift,
                "files": files,
            }

    total_files = sum(len(payload["files"] or []) for payload in upload_map.values())

    st.markdown("<br>", unsafe_allow_html=True)
    st.info(f"Total file terupload: {total_files} PDF")

    process = st.button("Proses PDF → Excel")

    if process:
        if total_files == 0:
            st.warning("Upload minimal 1 file PDF dulu.")
            return

        try:
            with st.status("Memproses PDF...", expanded=True) as status:
                st.write("Menyimpan file sementara...")
                saved_files = save_uploaded_files(upload_map)

                st.write("Membaca resi dari PDF...")
                parsed_data = call_parser(saved_files)

                parsed_count = len(parsed_data) if hasattr(parsed_data, "__len__") else "-"
                st.write(f"Data terbaca: {parsed_count} baris/resi")

                st.write("Membuat Excel...")
                excel_bytes = call_excel_writer(parsed_data, saved_files)

                status.update(label="Selesai", state="complete", expanded=False)

            filename = f"Rekap_Resi_HSD_{now_wib().strftime('%Y%m%d_%H%M')}_WIB.xlsx"

            st.success("Excel berhasil dibuat.")

            st.download_button(
                label="Download Excel Rekap",
                data=excel_bytes,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        except Exception as e:
            st.error("Proses gagal.")
            st.exception(e)
            st.caption(
                "Kalau error terjadi di bagian parser/excel_writer, kirim isi error-nya atau file parser.py dan excel_writer.py supaya bisa disesuaikan."
            )


def konten_page():
    page_header("Konten", "Area kerja divisi konten HSD.")

    st.markdown(
        """
        <div class="hsd-card">
            <div style="font-size:20px; font-weight:850; color:#111827;">Coming Soon</div>
            <div class="small-muted">Nanti bagian ini bisa diisi kalender konten, ide script, approval, dan database asset.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def live_page():
    page_header("Live", "Area kerja divisi live HSD.")

    st.markdown(
        """
        <div class="hsd-card">
            <div style="font-size:20px; font-weight:850; color:#111827;">Coming Soon</div>
            <div class="small-muted">Nanti bagian ini bisa diisi jadwal live, target GMV, host, produk, dan evaluasi performa.</div>
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
