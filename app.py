import streamlit as st
import tempfile, os, datetime
from io import BytesIO

st.set_page_config(page_title='HSD AGENT', page_icon='🧄', layout='wide')

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;600;700;800&display=swap');

html,[class*="css"]{font-family:'Plus Jakarta Sans',sans-serif!important;}
#MainMenu,footer,header,.stDeployButton{visibility:hidden;display:none;}
.stApp{background:#F8F6F1;}

.stButton>button{
    font-family:'Plus Jakarta Sans',sans-serif!important;
    font-weight:700!important;
    border-radius:8px!important;
    border:none!important;
    transition:all .15s;
}
.stButton>button:hover{transform:translateY(-1px);}
div[data-testid="stButton"]>button[kind="primary"]{
    background:#F5A623!important;
    color:white!important;
    font-size:15px!important;
    padding:14px!important;
}

.card{
    background:white;
    border-radius:12px;
    padding:20px;
    border:1px solid #E5E7EB;
    margin-bottom:12px;
}

.slot-hsd{
    background:#EFF6FF;
    border:1.5px dashed #93C5FD;
    border-radius:10px;
    padding:12px;
    text-align:center;
    color:#3B82F6;
    font-size:12px;
    font-weight:600;
}
.slot-hss{
    background:#FFF7ED;
    border:1.5px dashed #FDBA74;
    border-radius:10px;
    padding:12px;
    text-align:center;
    color:#F97316;
    font-size:12px;
    font-weight:600;
}
.slot-done{
    background:#F0FDF4;
    border:1.5px solid #86EFAC;
    border-radius:10px;
    padding:12px;
    text-align:center;
    color:#16A34A;
    font-size:12px;
    font-weight:600;
}

.metric-box{
    background:white;
    border-radius:10px;
    padding:16px;
    border:1px solid #E5E7EB;
    text-align:center;
}
.metric-val{
    font-size:26px;
    font-weight:800;
    color:#1C1C1E;
}
.metric-lbl{
    font-size:10px;
    color:#9CA3AF;
    font-weight:600;
    text-transform:uppercase;
}
[data-testid="stFileUploader"]{
    background:white!important;
    border-radius:8px!important;
}
.stProgress>div>div{
    background:#F5A623!important;
    border-radius:99px!important;
}

/* SIDEBAR HITAM */
section[data-testid="stSidebar"]{
    display:block!important;
    background:#050505!important;
    border-right:1px solid #111827!important;
}
section[data-testid="stSidebar"] > div{
    background:#050505!important;
    padding-top:18px!important;
}
section[data-testid="stSidebar"] *{
    font-family:'Plus Jakarta Sans',sans-serif!important;
}

.sidebar-card{
    background:#0B0B0D;
    border:1px solid #1F2937;
    border-radius:14px;
    padding:16px;
    margin-bottom:14px;
    color:white;
}
.sidebar-title{
    font-size:20px;
    font-weight:800;
    color:#FFFFFF;
    letter-spacing:.2px;
    margin-bottom:4px;
}
.sidebar-sub{
    font-size:11px;
    color:#9CA3AF;
    font-weight:600;
    margin-bottom:10px;
}
.sidebar-user{
    background:#F5A623;
    color:white;
    padding:5px 10px;
    border-radius:999px;
    font-size:11px;
    font-weight:800;
    display:inline-block;
}

section[data-testid="stSidebar"] .stButton>button{
    background:#111827!important;
    color:#E5E7EB!important;
    border:1px solid #1F2937!important;
    border-radius:10px!important;
    padding:12px 14px!important;
    text-align:left!important;
    justify-content:flex-start!important;
}
section[data-testid="stSidebar"] .stButton>button:hover{
    background:#1F2937!important;
    color:#FFFFFF!important;
    transform:none!important;
}
section[data-testid="stSidebar"] div[data-testid="stButton"]>button[kind="primary"]{
    background:#F5A623!important;
    color:#FFFFFF!important;
    border:1px solid #F5A623!important;
}
section[data-testid="stSidebar"] div[data-testid="stButton"]>button:disabled{
    background:#0B0B0D!important;
    color:#4B5563!important;
    border:1px solid #111827!important;
}
.sidebar-footer{
    position:fixed;
    bottom:18px;
    width:250px;
    color:#6B7280;
    font-size:10px;
    line-height:1.5;
}
</style>
""", unsafe_allow_html=True)


ROLES = {
    'gudang': {
        'pin': '1234',
        'role': 'gudang',
        'label': 'Divisi Gudang',
        'access': ['gudang']
    },
    'konten': {
        'pin': '2345',
        'role': 'konten',
        'label': 'Divisi Konten',
        'access': ['konten']
    },
    'live': {
        'pin': '3456',
        'role': 'live',
        'label': 'Divisi Live',
        'access': ['live']
    },
    'manager': {
        'pin': '0000',
        'role': 'admin',
        'label': 'Manager / BOD',
        'access': ['gudang', 'konten', 'live']
    },
}


def check_pin(p):
    for v in ROLES.values():
        if v['pin'] == p:
            return v['role'], v['label'], v['access']
    return None, None, None


def can(role, req):
    if role == 'admin':
        return True
    return req == role


for k, d in [
    ('ok', False),
    ('role', None),
    ('label', None),
    ('access', []),
    ('page', 'gudang')
]:
    if k not in st.session_state:
        st.session_state[k] = d


# LOGIN
if not st.session_state.ok:
    _, col, _ = st.columns([1, 1.2, 1])

    with col:
        st.markdown('<br><br>', unsafe_allow_html=True)
        st.markdown("""
        <div style="text-align:center;margin-bottom:20px;">
            <div style="width:80px;height:80px;background:#F5A623;border-radius:20px;
                margin:0 auto;font-size:44px;line-height:80px;">🧄</div>
            <h2 style="margin:12px 0 2px;color:#1C1C1E;font-weight:800;">HSD AGENT</h2>
            <p style="color:#9CA3AF;font-size:13px;">Sistem Manajemen Operasional</p>
        </div>
        """, unsafe_allow_html=True)

        with st.form('login'):
            pin = st.text_input(
                'PIN',
                '',
                type='password',
                placeholder='Masukkan PIN lalu Enter',
                label_visibility='collapsed'
            )
            ok = st.form_submit_button(
                '🔐  Masuk',
                use_container_width=True,
                type='primary'
            )

        if ok and pin:
            role, label, access = check_pin(pin)
            if role:
                st.session_state.ok = True
                st.session_state.role = role
                st.session_state.label = label
                st.session_state.access = access
                st.session_state.page = access[0] if access else 'gudang'
                st.rerun()
            else:
                st.error('PIN salah.')

        st.markdown(
            '<div style="background:#1C1C1E;border-radius:10px;padding:14px;margin-top:8px;">',
            unsafe_allow_html=True
        )
        st.markdown(
            '<div style="font-size:10px;color:#6B7280;font-weight:700;margin-bottom:8px;">PIN DIVISI</div>',
            unsafe_allow_html=True
        )

        for v in ROLES.values():
            st.markdown(
                f'''
                <div style="display:flex;justify-content:space-between;padding:3px 0;border-bottom:1px solid #2C2C2E;">
                    <span style="color:#9CA3AF;font-size:12px;">{v["label"]}</span>
                    <span style="background:#F5A623;color:white;padding:1px 8px;border-radius:4px;font-size:11px;font-weight:700;">{v["pin"]}</span>
                </div>
                ''',
                unsafe_allow_html=True
            )

        st.markdown('</div>', unsafe_allow_html=True)

    st.stop()


now = datetime.datetime.now()


# SIDEBAR
with st.sidebar:
    st.markdown(f"""
    <div class="sidebar-card">
        <div style="font-size:38px;line-height:1;margin-bottom:8px;">🧄</div>
        <div class="sidebar-title">HSD AGENT</div>
        <div class="sidebar-sub">Sistem Manajemen Operasional</div>
        <span class="sidebar-user">{st.session_state.label}</span>
    </div>
    """, unsafe_allow_html=True)

    for key, icon, label in [
        ('gudang', '📦', 'Gudang'),
        ('konten', '🎬', 'Konten'),
        ('live', '📡', 'Live')
    ]:
        has = can(st.session_state.role, key)
        is_active = st.session_state.page == key

        if has:
            if st.button(
                f'{icon}  {label}',
                use_container_width=True,
                type='primary' if is_active else 'secondary',
                key=f'side_{key}'
            ):
                st.session_state.page = key
                st.rerun()
        else:
            st.button(
                f'🔒  {label}',
                use_container_width=True,
                disabled=True,
                key=f'side_{key}'
            )

    st.markdown('<br>', unsafe_allow_html=True)

    if st.button('🚪  Keluar', use_container_width=True, key='side_logout'):
        for k in ['ok', 'role', 'label', 'access']:
            st.session_state[k] = False if k == 'ok' else (None if k != 'access' else [])
        st.session_state.page = 'gudang'
        st.rerun()

    st.markdown("""
    <div class="sidebar-footer">
        HSD AGENT<br>
        Sidebar Mode · Gudang / Konten / Live
    </div>
    """, unsafe_allow_html=True)


# TOP INFO
st.markdown(f"""
<div style="background:white;border-radius:12px;padding:12px 20px;
    border:1px solid #E5E7EB;margin-bottom:16px;">
    <span style="font-size:18px;font-weight:800;color:#1C1C1E;">🧄 HSD AGENT</span>
    <span style="background:#F5A623;color:white;padding:2px 10px;border-radius:20px;
        font-size:11px;font-weight:700;margin-left:10px;">{st.session_state.label}</span>
    <span style="color:#9CA3AF;font-size:12px;margin-left:12px;">
        {now.strftime('%A, %d %B %Y  ·  %H:%M WIB')}</span>
</div>
""", unsafe_allow_html=True)


page = st.session_state.page


# PAGE GUDANG
if page == 'gudang':
    if not can(st.session_state.role, 'gudang'):
        st.error('Akses ditolak.')
        st.stop()

    st.markdown(
        f'''
        <div class="card">
            <b style="font-size:16px;">📦 Upload PDF Resi</b>
            <span style="color:#9CA3AF;font-size:12px;margin-left:12px;">
                {now.strftime("%d/%m/%Y %H:%M WIB")}
            </span>
            <br>
            <span style="color:#6B7280;font-size:12px;">
                Mode PDF besar aktif. Cocok untuk ribuan resi dalam 1 PDF.
            </span>
        </div>
        ''',
        unsafe_allow_html=True
    )

    uploads = {}

    # HSD
    st.markdown(
        '''
        <div style="display:flex;align-items:center;gap:8px;margin-bottom:8px;">
            <span style="background:#3B82F6;color:white;padding:2px 10px;border-radius:6px;font-size:11px;font-weight:700;">HSD</span>
            <span style="color:#3B82F6;font-weight:700;">Botol Hitam</span>
        </div>
        ''',
        unsafe_allow_html=True
    )

    c1, c2, c3 = st.columns(3)

    for col, w, k in [
        (c1, 'PAGI', 'hsd_p'),
        (c2, 'SIANG', 'hsd_s'),
        (c3, 'SORE', 'hsd_r')
    ]:
        with col:
            st.markdown(
                f'<div style="font-size:11px;font-weight:700;color:#6B7280;margin-bottom:4px;">{w}</div>',
                unsafe_allow_html=True
            )
            f = st.file_uploader(
                f'HSD {w}',
                type='pdf',
                key=k,
                label_visibility='collapsed'
            )
            uploads[('HSD', w)] = f
            st.markdown(
                f'<div class="{"slot-done" if f else "slot-hsd"}">{"✅ " + f.name[:18] if f else "📄 Pilih PDF"}</div>',
                unsafe_allow_html=True
            )

    st.markdown('<br>', unsafe_allow_html=True)

    # HSS
    st.markdown(
        '''
        <div style="display:flex;align-items:center;gap:8px;margin-bottom:8px;">
            <span style="background:#F97316;color:white;padding:2px 10px;border-radius:6px;font-size:11px;font-weight:700;">HSS</span>
            <span style="color:#F97316;font-weight:700;">Botol Merah</span>
        </div>
        ''',
        unsafe_allow_html=True
    )

    c4, c5, c6 = st.columns(3)

    for col, w, k in [
        (c4, 'PAGI', 'hss_p'),
        (c5, 'SIANG', 'hss_s'),
        (c6, 'SORE', 'hss_r')
    ]:
        with col:
            st.markdown(
                f'<div style="font-size:11px;font-weight:700;color:#6B7280;margin-bottom:4px;">{w}</div>',
                unsafe_allow_html=True
            )
            f = st.file_uploader(
                f'HSS {w}',
                type='pdf',
                key=k,
                label_visibility='collapsed'
            )
            uploads[('HSS', w)] = f
            st.markdown(
                f'<div class="{"slot-done" if f else "slot-hss"}">{"✅ " + f.name[:18] if f else "📄 Pilih PDF"}</div>',
                unsafe_allow_html=True
            )

    st.markdown('<br>', unsafe_allow_html=True)

    active = [(a, w, f) for (a, w), f in uploads.items() if f]

    cb, ci = st.columns([2, 3])

    with cb:
        proses = st.button(
            '⚡  PROSES PDF → EXCEL',
            type='primary',
            use_container_width=True,
            disabled=not active
        )

    with ci:
        msg = f'✅ {len(active)} PDF siap' if active else 'Upload minimal 1 PDF'
        clr = '#22C55E' if active else '#9CA3AF'
        st.markdown(
            f'<div style="padding:14px 0;color:{clr};font-weight:600;font-size:13px;">{msg}</div>',
            unsafe_allow_html=True
        )

    if proses and active:
        from parser import process_pdf
        from excel_writer import write_excel_multi

        all_rows = []
        pb = st.progress(0, 'Memulai...')
        log = st.empty()
        logs = []

        for fi, (akun, waktu, fo) in enumerate(active):
            file_size_mb = getattr(fo, 'size', 0) / (1024 * 1024)

            logs.append(f'📄 **{akun} {waktu}** — `{fo.name}` ({file_size_mb:.1f} MB)')
            log.markdown('\n\n'.join(logs))

            with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp:
                # Lebih aman untuk PDF besar
                tmp.write(fo.getbuffer())
                tp = tmp.name

            def prog(cur, tot, fi=fi, tf=len(active), a=akun, w=waktu):
                if tot > 0:
                    percent = int((fi / tf + cur / tot / tf) * 100)
                else:
                    percent = int((fi / tf) * 100)

                pb.progress(
                    min(percent, 100),
                    text=f'{a} {w}: hal {cur}/{tot}'
                )

            try:
                rows, _ = process_pdf(tp, progress_callback=prog)

                for r in rows:
                    r['akun'] = akun
                    r['waktu'] = waktu

                all_rows.extend(rows)

                logs.append(f'✅ {akun} {waktu}: **{len(rows)} baris**')
                log.markdown('\n\n'.join(logs))

            except Exception as e:
                logs.append(f'❌ {akun} {waktu}: {e}')
                log.markdown('\n\n'.join(logs))

            finally:
                os.unlink(tp)

        if all_rows:
            pb.progress(95, 'Membuat Excel...')

            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
                txls = tmp.name

            write_excel_multi(all_rows, txls)

            out = BytesIO()
            with open(txls, 'rb') as f:
                out.write(f.read())

            os.unlink(txls)
            out.seek(0)

            pb.progress(100, '✅ Selesai!')

            total_resi = len(set(r['no_resi'] for r in all_rows))
            total_qty = sum(r['qty'] for r in all_rows)

            st.markdown('<br>', unsafe_allow_html=True)

            for col, val, lbl in zip(
                st.columns(4),
                [total_resi, total_qty, len(active), now.strftime('%H:%M')],
                ['Total Resi', 'Total Qty', 'PDF Diproses', 'Waktu Proses']
            ):
                with col:
                    st.markdown(
                        f'''
                        <div class="metric-box">
                            <div class="metric-val">{val}</div>
                            <div class="metric-lbl">{lbl}</div>
                        </div>
                        ''',
                        unsafe_allow_html=True
                    )

            st.markdown('<br>', unsafe_allow_html=True)

            st.download_button(
                '📥  Download Excel Rekap',
                data=out,
                file_name=f'HSD_REKAP_{now.strftime("%Y%m%d_%H%M")}.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                use_container_width=True,
                type='primary'
            )

        else:
            st.error('Tidak ada data yang berhasil dibaca dari PDF.')


# PAGE KONTEN
elif page == 'konten':
    if not can(st.session_state.role, 'konten'):
        st.error('Akses ditolak.')
        st.stop()

    st.markdown(
        '''
        <div class="card">
            <h3>🎬 Divisi Konten</h3>
            <p style="color:#9CA3AF;">Fitur ini segera hadir. 🚧</p>
        </div>
        ''',
        unsafe_allow_html=True
    )


# PAGE LIVE
elif page == 'live':
    if not can(st.session_state.role, 'live'):
        st.error('Akses ditolak.')
        st.stop()

    st.markdown(
        '''
        <div class="card">
            <h3>📡 Divisi Live</h3>
            <p style="color:#9CA3AF;">Fitur ini segera hadir. 🚧</p>
        </div>
        ''',
        unsafe_allow_html=True
    )
