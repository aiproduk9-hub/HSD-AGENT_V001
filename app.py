import streamlit as st
import tempfile
import os
import datetime
from io import BytesIO
from zoneinfo import ZoneInfo


st.set_page_config(
    page_title="HSD AGENT",
    page_icon="🧄",
    layout="wide",
    initial_sidebar_state="expanded"
)


# =========================
# STYLE
# =========================

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;600;700;800;900&display=swap');

html, body, [class*="css"] {
    font-family: 'Plus Jakarta Sans', sans-serif !important;
}

#MainMenu, footer, header, .stDeployButton {
    visibility: hidden;
    display: none;
}

.stApp {
    background: #F8F6F1;
}

.block-container {
    padding-top: 28px !important;
    padding-left: 36px !important;
    padding-right: 36px !important;
    max-width: 100% !important;
}

/* SIDEBAR ASLI STREAMLIT */
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
    font-family: 'Plus Jakarta Sans', sans-serif !important;
}

.sidebar-logo {
    width: 66px;
    height: 66px;
    background: #F5A623;
    border-radius: 18px;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 34px;
    margin-bottom: 16px;
}

.sidebar-title {
    color: #FFFFFF;
    font-size: 24px;
    font-weight: 900;
    line-height: 1.05;
    margin-bottom: 8px;
}

.sidebar-subtitle {
    color: #9CA3AF;
    font-size: 12px;
    font-weight: 700;
    margin-bottom: 14px;
}

.sidebar-role {
    display: inline-block;
    background: #F5A623;
    color: #FFFFFF;
    border-radius: 999px;
    padding: 6px 12px;
    font-size: 11px;
    font-weight: 900;
    margin-bottom: 26px;
}

.sidebar-menu-title {
    color: #6B7280;
    font-size: 10px;
    font-weight: 900;
    letter-spacing: 1px;
    margin: 18px 0 10px;
}

.sidebar-note {
    color: #6B7280;
    font-size: 11px;
    line-height: 1.6;
    margin-top: 24px;
}

/* BUTTON UMUM */
.stButton > button {
    border-radius: 12px !important;
    border: none !important;
    font-weight: 800 !important;
    transition: all .15s ease;
    font-family: 'Plus Jakarta Sans', sans-serif !important;
}

.stButton > button:hover {
    transform: translateY(-1px);
}

div[data-testid="stButton"] > button[kind="primary"] {
    background: #F5A623 !important;
    color: white !important;
    padding: 14px !important;
    font-size: 14px !important;
}

/* BUTTON DI SIDEBAR */
section[data-testid="stSidebar"] .stButton > button {
    background: #111827 !important;
    color: #E5E7EB !important;
    border: 1px solid #1F2937 !important;
    border-radius: 14px !important;
    padding: 13px 14px !important;
    text-align: left !important;
    justify-content: flex-start !important;
    margin-bottom: 6px !important;
}

section[data-testid="stSidebar"] .stButton > button:hover {
    background: #1F2937 !important;
    color: #FFFFFF !important;
    transform: none !important;
}

section[data-testid="stSidebar"] div[data-testid="stButton"] > button[kind="primary"] {
    background: #F5A623 !important;
    color: #FFFFFF !important;
    border: 1px solid #F5A623 !important;
}

section[data-testid="stSidebar"] .stButton > button:disabled {
    background: #0B0B0D !important;
    color: #4B5563 !important;
    border: 1px solid #111827 !important;
}

/* CARD */
.card {
    background: #FFFFFF;
    border: 1px solid #E5E7EB;
    border-radius: 18px;
    padding: 22px;
    margin-bottom: 18px;
    box-shadow: 0 8px 26px rgba(15, 23, 42, 0.04);
}

.top-card {
    background: #FFFFFF;
    border: 1px solid #E5E7EB;
    border-radius: 18px;
    padding: 18px 22px;
    margin-bottom: 18px;
    box-shadow: 0 8px 26px rgba(15, 23, 42, 0.04);
}

/* FILE UPLOADER */
[data-testid="stFileUploader"] {
    background: #111827 !important;
    border-radius: 16px !important;
    border: 1px solid #1F2937 !important;
    padding: 14px !important;
}

[data-testid="stFileUploader"] * {
    color: #E5E7EB !important;
}

[data-testid="stFileUploader"] button {
    background: #020617 !important;
    color: #FFFFFF !important;
    border: 1px solid #374151 !important;
    border-radius: 10px !important;
}

[data-testid="stFileUploaderDropzone"] {
    background: #111827 !important;
    border-radius: 14px !important;
}

.slot-hsd {
    background: #EFF6FF;
    border: 1.5px dashed #60A5FA;
    border-radius: 14px;
    padding: 12px;
    text-align: center;
    color: #2563EB;
    font-size: 12px;
    font-weight: 800;
    margin-top: 8px;
}

.slot-hss {
    background: #FFF7ED;
    border: 1.5px dashed #FB923C;
    border-radius: 14px;
    padding: 12px;
    text-align: center;
    color: #F97316;
    font-size: 12px;
    font-weight: 800;
    margin-top: 8px;
}

.slot-done {
    background: #F0FDF4;
    border: 1.5px solid #22C55E;
    border-radius: 14px;
    padding: 12px;
    text-align: center;
    color: #16A34A;
    font-size: 12px;
    font-weight: 900;
    margin-top: 8px;
}

.section-title {
    display: flex;
    align-items: center;
    gap: 8px;
    margin: 14px 0 12px;
}

.badge-hsd {
    background: #3B82F6;
    color: white;
    padding: 4px 10px;
    border-radius: 8px;
    font-size: 11px;
    font-weight: 900;
}

.badge-hss {
    background: #F97316;
    color: white;
    padding: 4px 10px;
    border-radius: 8px;
    font-size: 11px;
    font-weight: 900;
}

.metric-box {
    background: white;
    border: 1px solid #E5E7EB;
    border-radius: 16px;
    padding: 18px;
    text-align: center;
    box-shadow: 0 8px 26px rgba(15, 23, 42, 0.04);
}

.metric-val {
    font-size: 28px;
    font-weight: 900;
    color: #111827;
}

.metric-lbl {
    font-size: 11px;
    font-weight: 800;
    color: #9CA3AF;
    text-transform: uppercase;
}

.stProgress > div > div {
    background: #F5A623 !important;
}
</style>
""", unsafe_allow_html=True)


# =========================
# ROLE
# =========================

ROLES = {
    "gudang": {
        "pin": "1234",
        "role": "gudang",
        "label": "Divisi Gudang",
        "access": ["gudang"]
    },
    "konten": {
        "pin": "2345",
        "role": "konten",
        "label": "Divisi Konten",
        "access": ["konten"]
    },
    "live": {
        "pin": "3456",
        "role": "live",
        "label": "Divisi Live",
        "access": ["live"]
    },
    "manager": {
        "pin": "0000",
        "role": "admin",
        "label": "Manager / BOD",
        "access": ["gudang", "konten", "live"]
    },
}


def check_pin(pin):
    for item in ROLES.values():
        if item["pin"] == pin:
            return item["role"], item["label"], item["access"]
    return None, None, None


def can(role, page):
    if role == "admin":
        return True
    return role == page


for key, default in {
    "ok": False,
    "role": None,
    "label": None,
    "access": [],
    "page": "gudang"
}.items():
    if key not in st.session_state:
        st.session_state[key] = default


now = datetime.datetime.now(ZoneInfo("Asia/Jakarta"))


# =========================
# LOGIN
# =========================

if not st.session_state.ok:
    _, col, _ = st.columns([1, 1.1, 1])

    with col:
        st.markdown("<br><br>", unsafe_allow_html=True)

        st.markdown("""
        <div style="text-align:center;margin-bottom:24px;">
            <div style="width:82px;height:82px;background:#F5A623;border-radius:24px;
                margin:0 auto;font-size:44px;line-height:82px;">🧄</div>
            <h1 style="margin:18px 0 4px;color:#111827;font-weight:900;letter-spacing:-1px;">HSD AGENT</h1>
            <p style="color:#9CA3AF;font-size:13px;font-weight:600;">Sistem Manajemen Operasional</p>
        </div>
        """, unsafe_allow_html=True)

        with st.form("login"):
            pin = st.text_input(
                "PIN",
                "",
                type="password",
                placeholder="Masukkan PIN",
                label_visibility="collapsed"
            )

            login = st.form_submit_button(
                "🔐 Masuk",
                use_container_width=True,
                type="primary"
            )

        if login:
            role, label, access = check_pin(pin)

            if role:
                st.session_state.ok = True
                st.session_state.role = role
                st.session_state.label = label
                st.session_state.access = access
                st.session_state.page = access[0] if access else "gudang"
                st.rerun()
            else:
                st.error("PIN salah.")

        st.markdown("""
        <div class="card">
            <div style="font-size:11px;color:#9CA3AF;font-weight:900;margin-bottom:10px;">PIN DIVISI</div>
            <div style="display:flex;justify-content:space-between;border-bottom:1px solid #E5E7EB;padding:6px 0;">
                <span>Divisi Gudang</span><b>1234</b>
            </div>
            <div style="display:flex;justify-content:space-between;border-bottom:1px solid #E5E7EB;padding:6px 0;">
                <span>Divisi Konten</span><b>2345</b>
            </div>
            <div style="display:flex;justify-content:space-between;border-bottom:1px solid #E5E7EB;padding:6px 0;">
                <span>Divisi Live</span><b>3456</b>
            </div>
            <div style="display:flex;justify-content:space-between;padding:6px 0;">
                <span>Manager / BOD</span><b>0000</b>
            </div>
        </div>
        """, unsafe_allow_html=True)

    st.stop()


# =========================
# SIDEBAR
# =========================

with st.sidebar:
    st.markdown("""
    <div class="sidebar-logo">🧄</div>
    <div class="sidebar-title">HSD<br>AGENT</div>
    <div class="sidebar-subtitle">Sistem Operasional</div>
    """, unsafe_allow_html=True)

    st.markdown(
        f'<div class="sidebar-role">{st.session_state.label}</div>',
        unsafe_allow_html=True
    )

    st.markdown('<div class="sidebar-menu-title">MENU UTAMA</div>', unsafe_allow_html=True)

    if can(st.session_state.role, "gudang"):
        if st.button(
            "📦  Gudang",
            use_container_width=True,
            type="primary" if st.session_state.page == "gudang" else "secondary",
            key="menu_gudang"
        ):
            st.session_state.page = "gudang"
            st.rerun()
    else:
        st.button("🔒  Gudang", use_container_width=True, disabled=True, key="lock_gudang")

    if can(st.session_state.role, "konten"):
        if st.button(
            "🎬  Konten",
            use_container_width=True,
            type="primary" if st.session_state.page == "konten" else "secondary",
            key="menu_konten"
        ):
            st.session_state.page = "konten"
            st.rerun()
    else:
        st.button("🔒  Konten", use_container_width=True, disabled=True, key="lock_konten")

    if can(st.session_state.role, "live"):
        if st.button(
            "📡  Live",
            use_container_width=True,
            type="primary" if st.session_state.page == "live" else "secondary",
            key="menu_live"
        ):
            st.session_state.page = "live"
            st.rerun()
    else:
        st.button("🔒  Live", use_container_width=True, disabled=True, key="lock_live")

    st.markdown('<div class="sidebar-menu-title">AKUN</div>', unsafe_allow_html=True)

    if st.button("🚪  Keluar", use_container_width=True, key="logout_btn"):
        st.session_state.ok = False
        st.session_state.role = None
        st.session_state.label = None
        st.session_state.access = []
        st.session_state.page = "gudang"
        st.rerun()

    st.markdown("""
    <div class="sidebar-note">
        HSD Agent v2<br>
        Gudang · Konten · Live<br><br>
        Sidebar resmi aktif.
    </div>
    """, unsafe_allow_html=True)


# =========================
# TOP BAR
# =========================

st.markdown(f"""
<div class="top-card">
    <div style="display:flex;align-items:center;justify-content:space-between;">
        <div>
            <div style="font-size:22px;font-weight:900;color:#111827;">🧄 HSD AGENT</div>
            <div style="color:#9CA3AF;font-size:12px;font-weight:700;margin-top:4px;">
                {st.session_state.label}
            </div>
        </div>
        <div style="text-align:right;color:#6B7280;font-size:12px;font-weight:700;">
            {now.strftime('%A, %d %B %Y')}<br>
            {now.strftime('%H:%M:%S WIB')}
        </div>
    </div>
</div>
""", unsafe_allow_html=True)


page = st.session_state.page


# =========================
# PAGE GUDANG
# =========================

if page == "gudang":
    if not can(st.session_state.role, "gudang"):
        st.error("Akses ditolak.")
        st.stop()

    st.markdown("""
    <div class="card">
        <div style="font-size:18px;font-weight:900;color:#111827;">📦 Upload PDF Resi</div>
        <div style="font-size:13px;color:#6B7280;font-weight:600;margin-top:6px;">
            Mode PDF besar aktif. Cocok untuk ribuan resi dalam 1 PDF.
        </div>
    </div>
    """, unsafe_allow_html=True)

    uploads = {}

    st.markdown("""
    <div class="section-title">
        <span class="badge-hsd">HSD</span>
        <span style="color:#2563EB;font-size:18px;font-weight:900;">Botol Hitam</span>
    </div>
    """, unsafe_allow_html=True)

    c1, c2, c3 = st.columns(3)

    for col, waktu, key in [
        (c1, "PAGI", "hsd_p"),
        (c2, "SIANG", "hsd_s"),
        (c3, "SORE", "hsd_r"),
    ]:
        with col:
            st.markdown(
                f'<div style="font-size:11px;font-weight:900;color:#374151;margin-bottom:6px;">{waktu}</div>',
                unsafe_allow_html=True
            )

            f = st.file_uploader(
                f"HSD {waktu}",
                type="pdf",
                key=key,
                label_visibility="collapsed"
            )

            uploads[("HSD", waktu)] = f

            if f:
                st.markdown(
                    f'<div class="slot-done">✅ {f.name[:26]}</div>',
                    unsafe_allow_html=True
                )
            else:
                st.markdown(
                    '<div class="slot-hsd">📄 Pilih PDF</div>',
                    unsafe_allow_html=True
                )

    st.markdown("<br>", unsafe_allow_html=True)

    st.markdown("""
    <div class="section-title">
        <span class="badge-hss">HSS</span>
        <span style="color:#F97316;font-size:18px;font-weight:900;">Botol Merah</span>
    </div>
    """, unsafe_allow_html=True)

    c4, c5, c6 = st.columns(3)

    for col, waktu, key in [
        (c4, "PAGI", "hss_p"),
        (c5, "SIANG", "hss_s"),
        (c6, "SORE", "hss_r"),
    ]:
        with col:
            st.markdown(
                f'<div style="font-size:11px;font-weight:900;color:#374151;margin-bottom:6px;">{waktu}</div>',
                unsafe_allow_html=True
            )

            f = st.file_uploader(
                f"HSS {waktu}",
                type="pdf",
                key=key,
                label_visibility="collapsed"
            )

            uploads[("HSS", waktu)] = f

            if f:
                st.markdown(
                    f'<div class="slot-done">✅ {f.name[:26]}</div>',
                    unsafe_allow_html=True
                )
            else:
                st.markdown(
                    '<div class="slot-hss">📄 Pilih PDF</div>',
                    unsafe_allow_html=True
                )

    st.markdown("<br>", unsafe_allow_html=True)

    active = [
        (akun, waktu, file)
        for (akun, waktu), file in uploads.items()
        if file
    ]

    b1, b2 = st.columns([2, 3])

    with b1:
        proses = st.button(
            "⚡ PROSES PDF → EXCEL",
            type="primary",
            use_container_width=True,
            disabled=not active
        )

    with b2:
        if active:
            st.success(f"{len(active)} PDF siap diproses")
        else:
            st.info("Upload minimal 1 PDF")

    if proses and active:
        from parser import process_pdf
        from excel_writer import write_excel_multi

        all_rows = []
        progress = st.progress(0, "Memulai...")
        log_box = st.empty()
        logs = []

        for fi, (akun, waktu, file) in enumerate(active):
            file_size_mb = getattr(file, "size", 0) / (1024 * 1024)

            logs.append(f"📄 **{akun} {waktu}** — `{file.name}` ({file_size_mb:.1f} MB)")
            log_box.markdown("\n\n".join(logs))

            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                tmp.write(file.getbuffer())
                temp_pdf_path = tmp.name

            def prog(cur, tot, fi=fi, total_files=len(active), akun=akun, waktu=waktu):
                if tot and tot > 0:
                    percent = int((fi / total_files + cur / tot / total_files) * 100)
                else:
                    percent = int((fi / total_files) * 100)

                progress.progress(
                    min(percent, 100),
                    text=f"{akun} {waktu}: halaman {cur}/{tot}"
                )

            try:
                rows, _ = process_pdf(temp_pdf_path, progress_callback=prog)

                for r in rows:
                    r["akun"] = akun
                    r["waktu"] = waktu

                all_rows.extend(rows)

                logs.append(f"✅ {akun} {waktu}: **{len(rows)} baris**")
                log_box.markdown("\n\n".join(logs))

            except Exception as e:
                logs.append(f"❌ {akun} {waktu}: {e}")
                log_box.markdown("\n\n".join(logs))

            finally:
                if os.path.exists(temp_pdf_path):
                    os.unlink(temp_pdf_path)

        if all_rows:
            progress.progress(95, "Membuat Excel...")

            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                temp_xlsx_path = tmp.name

            write_excel_multi(all_rows, temp_xlsx_path)

            out = BytesIO()

            with open(temp_xlsx_path, "rb") as f:
                out.write(f.read())

            if os.path.exists(temp_xlsx_path):
                os.unlink(temp_xlsx_path)

            out.seek(0)
            progress.progress(100, "Selesai!")

            total_resi = len(set(r.get("no_resi") for r in all_rows))
            total_qty = sum(int(r.get("qty", 0) or 0) for r in all_rows)

            m1, m2, m3, m4 = st.columns(4)

            for col, val, label in [
                (m1, total_resi, "Total Resi"),
                (m2, total_qty, "Total Qty"),
                (m3, len(active), "PDF Diproses"),
                (m4, now.strftime("%H:%M"), "Waktu"),
            ]:
                with col:
                    st.markdown(f"""
                    <div class="metric-box">
                        <div class="metric-val">{val}</div>
                        <div class="metric-lbl">{label}</div>
                    </div>
                    """, unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)

            st.download_button(
                "📥 Download Excel Rekap",
                data=out,
                file_name=f"HSD_REKAP_{now.strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary"
            )

        else:
            st.error("Tidak ada data yang berhasil dibaca dari PDF.")


# =========================
# PAGE KONTEN
# =========================

elif page == "konten":
    st.markdown("""
    <div class="card">
        <div style="font-size:22px;font-weight:900;color:#111827;">🎬 Divisi Konten</div>
        <p style="color:#6B7280;font-size:14px;font-weight:600;">
            Fitur ini segera hadir. Nanti bisa dibuat untuk jadwal konten, ide konten, script, dan arsip video.
        </p>
    </div>
    """, unsafe_allow_html=True)


# =========================
# PAGE LIVE
# =========================

elif page == "live":
    st.markdown("""
    <div class="card">
        <div style="font-size:22px;font-weight:900;color:#111827;">📡 Divisi Live</div>
        <p style="color:#6B7280;font-size:14px;font-weight:600;">
            Fitur ini segera hadir. Nanti bisa dibuat untuk jadwal live, laporan live, host, dan performa penjualan.
        </p>
    </div>
    """, unsafe_allow_html=True)
