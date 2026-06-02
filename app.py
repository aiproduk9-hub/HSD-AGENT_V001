import streamlit as st
import tempfile, os, datetime, json
from io import BytesIO

st.set_page_config(
    page_title='HSD AGENT',
    page_icon='🧄',
    layout='wide',
    initial_sidebar_state='expanded'
)

# ── CSS ────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;600&display=swap');

html, body, [class*="css"] {
    font-family: 'Plus Jakarta Sans', sans-serif;
}

/* Hide default streamlit elements */
#MainMenu, footer, header { visibility: hidden; }
.stDeployButton { display: none; }

/* App background */
.stApp { background: #F8F6F1; }

/* Sidebar */
section[data-testid="stSidebar"] {
    background: #1C1C1E !important;
    border-right: none !important;
}
section[data-testid="stSidebar"] * { color: #E5E7EB !important; }
section[data-testid="stSidebar"] .stMarkdown p { color: #9CA3AF !important; font-size: 12px; }

/* Cards */
.hsd-card {
    background: white;
    border-radius: 12px;
    padding: 20px 24px;
    border: 1px solid #E5E7EB;
    margin-bottom: 12px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.06);
}
.hsd-card-title {
    font-size: 11px;
    font-weight: 700;
    color: #9CA3AF;
    text-transform: uppercase;
    letter-spacing: 0.08em;
    margin-bottom: 4px;
}

/* Slot boxes */
.slot-hsd {
    background: #EFF6FF;
    border: 1.5px dashed #93C5FD;
    border-radius: 10px;
    padding: 16px;
    text-align: center;
    color: #3B82F6;
    font-size: 13px;
    font-weight: 600;
    margin-bottom: 4px;
}
.slot-hss {
    background: #FFF7ED;
    border: 1.5px dashed #FDBA74;
    border-radius: 10px;
    padding: 16px;
    text-align: center;
    color: #F97316;
    font-size: 13px;
    font-weight: 600;
    margin-bottom: 4px;
}
.slot-done {
    background: #F0FDF4;
    border: 1.5px solid #86EFAC;
    border-radius: 10px;
    padding: 16px;
    text-align: center;
    color: #16A34A;
    font-size: 13px;
    font-weight: 600;
    margin-bottom: 4px;
}

/* Metric cards */
.metric-box {
    background: white;
    border-radius: 10px;
    padding: 16px 20px;
    border: 1px solid #E5E7EB;
    text-align: center;
}
.metric-val {
    font-size: 28px;
    font-weight: 800;
    color: #1C1C1E;
    line-height: 1;
}
.metric-lbl {
    font-size: 11px;
    color: #9CA3AF;
    font-weight: 600;
    margin-top: 4px;
    text-transform: uppercase;
    letter-spacing: 0.06em;
}

/* Buttons */
.stButton > button {
    font-family: 'Plus Jakarta Sans', sans-serif !important;
    font-weight: 700 !important;
    border-radius: 8px !important;
    border: none !important;
    transition: all 0.15s ease !important;
}
.stButton > button:hover { transform: translateY(-1px); box-shadow: 0 4px 12px rgba(0,0,0,0.15) !important; }

/* Primary button — amber */
div[data-testid="stButton"] > button[kind="primary"] {
    background: #F5A623 !important;
    color: white !important;
    padding: 14px 32px !important;
    font-size: 15px !important;
}

/* PIN input */
.pin-display {
    background: #FEF3DC;
    border-radius: 10px;
    padding: 16px;
    text-align: center;
    font-size: 28px;
    letter-spacing: 12px;
    color: #F5A623;
    font-weight: 800;
    margin: 12px 0;
    font-family: 'JetBrains Mono', monospace;
}

/* Badge */
.badge {
    display: inline-block;
    background: #F5A623;
    color: white;
    padding: 3px 10px;
    border-radius: 20px;
    font-size: 11px;
    font-weight: 700;
}
.badge-blue {
    background: #3B82F6;
}
.badge-orange {
    background: #F97316;
}
.badge-green {
    background: #22C55E;
}

/* File uploader */
[data-testid="stFileUploader"] {
    background: white !important;
    border-radius: 10px !important;
    border: 1.5px dashed #E5E7EB !important;
}

/* Progress */
.stProgress > div > div {
    background: #F5A623 !important;
    border-radius: 99px !important;
}

/* Divider */
hr { border-color: #2C2C2E !important; margin: 8px 0 !important; }

/* Table */
.stDataFrame { border-radius: 10px; overflow: hidden; }

/* Success/error messages */
.stSuccess { background: #F0FDF4 !important; border-color: #86EFAC !important; border-radius: 8px !important; }
.stError { background: #FEF2F2 !important; border-radius: 8px !important; }
.stWarning { background: #FFFBEB !important; border-radius: 8px !important; }
</style>
""", unsafe_allow_html=True)

# ── PIN System ─────────────────────────────────────────────────────────────────
DEFAULT_PINS = {
    'gudang':  {'pin':'1234','role':'gudang', 'label':'Divisi Gudang'},
    'konten':  {'pin':'2345','role':'konten', 'label':'Divisi Konten'},
    'live':    {'pin':'3456','role':'live',   'label':'Divisi Live'},
    'manager': {'pin':'0000','role':'admin',  'label':'Manager / BOD'},
}

def check_pin(entered):
    for k,v in DEFAULT_PINS.items():
        if v['pin']==entered: return v['role'],v['label']
    return None,None

# ── Session State ──────────────────────────────────────────────────────────────
if 'logged_in' not in st.session_state: st.session_state.logged_in=False
if 'role' not in st.session_state: st.session_state.role=None
if 'label' not in st.session_state: st.session_state.label=None
if 'pin_input' not in st.session_state: st.session_state.pin_input=''

# ── LOGIN PAGE ─────────────────────────────────────────────────────────────────
def show_login():
    col1,col2,col3=st.columns([1,1.2,1])
    with col2:
        st.markdown('<br><br>',unsafe_allow_html=True)
        # Logo placeholder
        st.markdown("""
        <div style="text-align:center;margin-bottom:24px;">
            <div style="width:72px;height:72px;background:#F5A623;border-radius:16px;
                        margin:0 auto;display:flex;align-items:center;justify-content:center;
                        font-size:36px;line-height:72px;text-align:center;">🧄</div>
            <h2 style="margin:12px 0 4px;color:#1C1C1E;font-weight:800;">HSD AGENT</h2>
            <p style="color:#9CA3AF;font-size:13px;margin:0;">Sistem Manajemen Operasional</p>
        </div>
        """,unsafe_allow_html=True)

        st.markdown('<div class="hsd-card">',unsafe_allow_html=True)
        st.markdown('<div class="hsd-card-title">Masukkan PIN Divisi</div>',unsafe_allow_html=True)

        pin=st.text_input('PIN','',max_chars=6,type='password',
                           placeholder='Ketik PIN lalu Enter',
                           label_visibility='collapsed')

        if pin:
            role,label=check_pin(pin)
            if role:
                st.session_state.logged_in=True
                st.session_state.role=role
                st.session_state.label=label
                st.rerun()
            elif len(pin)>=4:
                st.error('PIN salah. Coba lagi.')

        st.markdown('</div>',unsafe_allow_html=True)

        # PIN hint
        st.markdown("""
        <div style="background:#1C1C1E;border-radius:10px;padding:14px 16px;margin-top:12px;">
            <div style="font-size:10px;color:#6B7280;font-weight:700;
                        text-transform:uppercase;letter-spacing:0.08em;margin-bottom:8px;">PIN Divisi</div>
        """,unsafe_allow_html=True)
        for k,v in DEFAULT_PINS.items():
            st.markdown(f"""
            <div style="display:flex;justify-content:space-between;align-items:center;
                        padding:4px 0;border-bottom:1px solid #2C2C2E;">
                <span style="font-size:12px;color:#9CA3AF;">{v['label']}</span>
                <span style="background:#F5A623;color:white;padding:2px 8px;
                             border-radius:4px;font-size:11px;font-weight:700;
                             font-family:'JetBrains Mono',monospace;">{v['pin']}</span>
            </div>
            """,unsafe_allow_html=True)
        st.markdown('</div>',unsafe_allow_html=True)

# ── SIDEBAR ────────────────────────────────────────────────────────────────────
def show_sidebar():
    with st.sidebar:
        st.markdown(f"""
        <div style="text-align:center;padding:16px 0 8px;">
            <div style="font-size:32px;">🧄</div>
            <div style="font-size:16px;font-weight:800;color:white;margin:6px 0 2px;">HSD AGENT</div>
            <div style="background:#F5A623;color:white;padding:3px 12px;
                        border-radius:20px;font-size:11px;font-weight:700;
                        display:inline-block;">{st.session_state.label}</div>
        </div>
        """,unsafe_allow_html=True)
        st.divider()

        role=st.session_state.role
        pages=['📦  Gudang']
        if role in ('konten','admin'): pages.append('🎬  Konten')
        if role in ('live','admin'): pages.append('📡  Live')
        if role=='admin': pages.extend(['📦  Gudang','🎬  Konten','📡  Live'])
        pages=list(dict.fromkeys(pages))

        selected=st.radio('Menu',pages,label_visibility='collapsed')

        st.divider()
        st.markdown('<div style="font-size:10px;color:#6B7280;font-weight:700;text-transform:uppercase;letter-spacing:0.08em;margin-bottom:8px;">PIN DIVISI</div>',unsafe_allow_html=True)
        for k,v in DEFAULT_PINS.items():
            st.markdown(f"""
            <div style="display:flex;justify-content:space-between;align-items:center;
                        padding:3px 0;">
                <span style="font-size:11px;color:#6B7280;">{v['label']}</span>
                <span style="background:#F5A623;color:white;padding:1px 7px;
                             border-radius:4px;font-size:10px;font-weight:700;
                             font-family:monospace;">{v['pin']}</span>
            </div>
            """,unsafe_allow_html=True)

        st.divider()
        if st.button('🚪  Keluar',use_container_width=True):
            st.session_state.logged_in=False
            st.session_state.role=None
            st.session_state.label=None
            st.rerun()

        return selected

# ── GUDANG PAGE ────────────────────────────────────────────────────────────────
def show_gudang():
    st.markdown("""
    <h2 style="color:#1C1C1E;font-weight:800;margin-bottom:4px;">📦 Upload PDF Resi</h2>
    <p style="color:#9CA3AF;font-size:13px;margin-bottom:20px;">
        Upload PDF resi per akun dan waktu, lalu klik Proses.
    </p>
    """,unsafe_allow_html=True)

    # ── SLOTS HSD ────────────────────────────────────────────────────────────
    st.markdown("""
    <div style="display:flex;align-items:center;gap:8px;margin-bottom:8px;">
        <span style="background:#3B82F6;color:white;padding:3px 10px;
                     border-radius:6px;font-size:11px;font-weight:700;">HSD</span>
        <span style="color:#3B82F6;font-weight:700;font-size:13px;">Botol Hitam</span>
    </div>
    """,unsafe_allow_html=True)

    c1,c2,c3=st.columns(3)
    with c1:
        st.markdown('<div style="font-size:11px;font-weight:700;color:#6B7280;margin-bottom:4px;">PAGI</div>',unsafe_allow_html=True)
        hsd_pagi=st.file_uploader('HSD Pagi',type='pdf',key='hsd_pagi',label_visibility='collapsed')
        if hsd_pagi: st.markdown(f'<div class="slot-done">✅ {hsd_pagi.name[:22]}</div>',unsafe_allow_html=True)
        else: st.markdown('<div class="slot-hsd">📄 Pilih PDF</div>',unsafe_allow_html=True)
    with c2:
        st.markdown('<div style="font-size:11px;font-weight:700;color:#6B7280;margin-bottom:4px;">SIANG</div>',unsafe_allow_html=True)
        hsd_siang=st.file_uploader('HSD Siang',type='pdf',key='hsd_siang',label_visibility='collapsed')
        if hsd_siang: st.markdown(f'<div class="slot-done">✅ {hsd_siang.name[:22]}</div>',unsafe_allow_html=True)
        else: st.markdown('<div class="slot-hsd">📄 Pilih PDF</div>',unsafe_allow_html=True)
    with c3:
        st.markdown('<div style="font-size:11px;font-weight:700;color:#6B7280;margin-bottom:4px;">SORE</div>',unsafe_allow_html=True)
        hsd_sore=st.file_uploader('HSD Sore',type='pdf',key='hsd_sore',label_visibility='collapsed')
        if hsd_sore: st.markdown(f'<div class="slot-done">✅ {hsd_sore.name[:22]}</div>',unsafe_allow_html=True)
        else: st.markdown('<div class="slot-hsd">📄 Pilih PDF</div>',unsafe_allow_html=True)

    st.markdown('<br>',unsafe_allow_html=True)

    # ── SLOTS HSS ────────────────────────────────────────────────────────────
    st.markdown("""
    <div style="display:flex;align-items:center;gap:8px;margin-bottom:8px;">
        <span style="background:#F97316;color:white;padding:3px 10px;
                     border-radius:6px;font-size:11px;font-weight:700;">HSS</span>
        <span style="color:#F97316;font-weight:700;font-size:13px;">Botol Merah</span>
    </div>
    """,unsafe_allow_html=True)

    c4,c5,c6=st.columns(3)
    with c4:
        st.markdown('<div style="font-size:11px;font-weight:700;color:#6B7280;margin-bottom:4px;">PAGI</div>',unsafe_allow_html=True)
        hss_pagi=st.file_uploader('HSS Pagi',type='pdf',key='hss_pagi',label_visibility='collapsed')
        if hss_pagi: st.markdown(f'<div class="slot-done">✅ {hss_pagi.name[:22]}</div>',unsafe_allow_html=True)
        else: st.markdown('<div class="slot-hss">📄 Pilih PDF</div>',unsafe_allow_html=True)
    with c5:
        st.markdown('<div style="font-size:11px;font-weight:700;color:#6B7280;margin-bottom:4px;">SIANG</div>',unsafe_allow_html=True)
        hss_siang=st.file_uploader('HSS Siang',type='pdf',key='hss_siang',label_visibility='collapsed')
        if hss_siang: st.markdown(f'<div class="slot-done">✅ {hss_siang.name[:22]}</div>',unsafe_allow_html=True)
        else: st.markdown('<div class="slot-hss">📄 Pilih PDF</div>',unsafe_allow_html=True)
    with c6:
        st.markdown('<div style="font-size:11px;font-weight:700;color:#6B7280;margin-bottom:4px;">SORE</div>',unsafe_allow_html=True)
        hss_sore=st.file_uploader('HSS Sore',type='pdf',key='hss_sore',label_visibility='collapsed')
        if hss_sore: st.markdown(f'<div class="slot-done">✅ {hss_sore.name[:22]}</div>',unsafe_allow_html=True)
        else: st.markdown('<div class="slot-hss">📄 Pilih PDF</div>',unsafe_allow_html=True)

    st.markdown('<br>',unsafe_allow_html=True)

    # ── PROSES ────────────────────────────────────────────────────────────────
    slots=[
        ('HSD','PAGI',hsd_pagi),('HSD','SIANG',hsd_siang),('HSD','SORE',hsd_sore),
        ('HSS','PAGI',hss_pagi),('HSS','SIANG',hss_siang),('HSS','SORE',hss_sore),
    ]
    active=[(a,w,f) for a,w,f in slots if f is not None]

    col_btn,col_info=st.columns([2,3])
    with col_btn:
        proses=st.button('⚡   PROSES PDF → EXCEL',
                          type='primary',use_container_width=True,
                          disabled=len(active)==0)
    with col_info:
        if active:
            st.markdown(f'<div style="padding:14px 0;color:#22C55E;font-weight:600;font-size:13px;">✅ {len(active)} PDF siap diproses</div>',unsafe_allow_html=True)
        else:
            st.markdown('<div style="padding:14px 0;color:#9CA3AF;font-size:13px;">Upload minimal 1 PDF untuk memulai</div>',unsafe_allow_html=True)

    if proses and active:
        from parser import process_pdf
        from excel_writer import write_excel_multi

        all_rows=[]
        pb=st.progress(0,text='Memulai...')
        log_area=st.empty()
        logs=[]

        for fi,(akun,waktu,file_obj) in enumerate(active):
            logs.append(f'📄 Membaca **{akun} {waktu}** — `{file_obj.name}`')
            log_area.markdown('\n\n'.join(logs))

            # Simpan ke temp file
            with tempfile.NamedTemporaryFile(delete=False,suffix='.pdf') as tmp:
                tmp.write(file_obj.read()); tmp_path=tmp.name

            def prog(cur,tot,fi=fi,tf=len(active)):
                pct=int((fi/tf+cur/tot/tf)*100)
                pb.progress(pct,text=f'{akun} {waktu}: hal {cur}/{tot}')

            try:
                rows,errs=process_pdf(tmp_path,progress_callback=prog)
                for r in rows: r['akun']=akun; r['waktu']=waktu
                all_rows.extend(rows)
                logs.append(f'✅ {akun} {waktu}: **{len(rows)} baris** data')
                log_area.markdown('\n\n'.join(logs))
            except Exception as e:
                logs.append(f'❌ Error: {e}')
                log_area.markdown('\n\n'.join(logs))
            finally:
                os.unlink(tmp_path)

        if all_rows:
            pb.progress(95,text='Membuat Excel...')
            output=BytesIO()
            # Simpan ke temp file dulu
            with tempfile.NamedTemporaryFile(delete=False,suffix='.xlsx') as tmp:
                tmp_xlsx=tmp.name
            write_excel_multi(all_rows,tmp_xlsx)
            with open(tmp_xlsx,'rb') as f: output.write(f.read())
            os.unlink(tmp_xlsx)
            output.seek(0)
            pb.progress(100,text='Selesai!')

            # Metrics
            unique=len(set(r['no_resi'] for r in all_rows))
            total_qty=sum(r['qty'] for r in all_rows)
            st.markdown('<br>',unsafe_allow_html=True)
            m1,m2,m3,m4=st.columns(4)
            with m1:
                st.markdown(f'<div class="metric-box"><div class="metric-val">{unique}</div><div class="metric-lbl">Total Resi</div></div>',unsafe_allow_html=True)
            with m2:
                st.markdown(f'<div class="metric-box"><div class="metric-val">{total_qty}</div><div class="metric-lbl">Total Qty</div></div>',unsafe_allow_html=True)
            with m3:
                st.markdown(f'<div class="metric-box"><div class="metric-val">{len(active)}</div><div class="metric-lbl">PDF Diproses</div></div>',unsafe_allow_html=True)
            with m4:
                from collections import Counter
                produk=len(Counter(r['nama_produk'] for r in all_rows))
                st.markdown(f'<div class="metric-box"><div class="metric-val">{produk}</div><div class="metric-lbl">Jenis Produk</div></div>',unsafe_allow_html=True)

            st.markdown('<br>',unsafe_allow_html=True)
            dt=datetime.datetime.now().strftime('%Y%m%d_%H%M')
            st.download_button(
                label='📥   Download Excel Rekap',
                data=output,
                file_name=f'HSD_REKAP_{dt}.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                use_container_width=True,
                type='primary'
            )
        else:
            st.error('Tidak ada data yang berhasil diekstrak.')

# ── MAIN ───────────────────────────────────────────────────────────────────────
if not st.session_state.logged_in:
    show_login()
else:
    selected=show_sidebar()
    if 'Gudang' in selected:
        show_gudang()
    elif 'Konten' in selected:
        st.markdown('<h2 style="color:#1C1C1E;font-weight:800;">🎬 Divisi Konten</h2>',unsafe_allow_html=True)
        st.info('🚧 Fitur ini sedang dalam pengembangan.')
    elif 'Live' in selected:
        st.markdown('<h2 style="color:#1C1C1E;font-weight:800;">📡 Divisi Live</h2>',unsafe_allow_html=True)
        st.info('🚧 Fitur ini sedang dalam pengembangan.')
