import streamlit as st
import tempfile, os, datetime, json
from io import BytesIO

st.set_page_config(
    page_title='HSD AGENT',
    page_icon='🧄',
    layout='wide',
    initial_sidebar_state='expanded'
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=JetBrains+Mono:wght@500&display=swap');
html,[class*="css"]{font-family:'Plus Jakarta Sans',sans-serif;}
#MainMenu,footer,header{visibility:hidden;}
.stDeployButton{display:none;}
.stApp{background:#F8F6F1;}
section[data-testid="stSidebar"]{background:#1C1C1E!important;}
section[data-testid="stSidebar"] p,section[data-testid="stSidebar"] span,
section[data-testid="stSidebar"] label{color:#E5E7EB!important;}
.stButton>button{font-family:'Plus Jakarta Sans',sans-serif!important;font-weight:700!important;border-radius:8px!important;border:none!important;}
.stButton>button:hover{transform:translateY(-1px);box-shadow:0 4px 12px rgba(0,0,0,.15)!important;}
div[data-testid="stButton"]>button[kind="primary"]{background:#F5A623!important;color:white!important;padding:14px 32px!important;font-size:15px!important;}
.hsd-card{background:white;border-radius:12px;padding:20px 24px;border:1px solid #E5E7EB;margin-bottom:12px;box-shadow:0 1px 3px rgba(0,0,0,.06);}
.metric-box{background:white;border-radius:10px;padding:16px 20px;border:1px solid #E5E7EB;text-align:center;margin-bottom:8px;}
.metric-val{font-size:28px;font-weight:800;color:#1C1C1E;line-height:1;}
.metric-lbl{font-size:11px;color:#9CA3AF;font-weight:600;margin-top:4px;text-transform:uppercase;letter-spacing:.06em;}
.slot-hsd{background:#EFF6FF;border:1.5px dashed #93C5FD;border-radius:10px;padding:14px;text-align:center;color:#3B82F6;font-size:13px;font-weight:600;margin-bottom:4px;}
.slot-hss{background:#FFF7ED;border:1.5px dashed #FDBA74;border-radius:10px;padding:14px;text-align:center;color:#F97316;font-size:13px;font-weight:600;margin-bottom:4px;}
.slot-done{background:#F0FDF4;border:1.5px solid #86EFAC;border-radius:10px;padding:14px;text-align:center;color:#16A34A;font-size:13px;font-weight:600;margin-bottom:4px;}
.nav-item{padding:10px 16px;border-radius:8px;cursor:pointer;margin-bottom:4px;font-weight:600;font-size:14px;}
.nav-active{background:#F5A623;color:white;}
.nav-locked{color:#4B5563;opacity:.5;}
.nav-open{color:#E5E7EB;}
.topbar{background:white;border-radius:12px;padding:12px 20px;border:1px solid #E5E7EB;margin-bottom:16px;display:flex;justify-content:space-between;align-items:center;}
.stProgress>div>div{background:#F5A623!important;border-radius:99px!important;}
[data-testid="stFileUploader"]{background:white!important;border-radius:10px!important;}
.stSuccess{background:#F0FDF4!important;border-color:#86EFAC!important;border-radius:8px!important;}
.stError{background:#FEF2F2!important;border-radius:8px!important;}
</style>
""", unsafe_allow_html=True)

ROLES = {
    'gudang':  {'pin':'1234','role':'gudang', 'label':'Divisi Gudang',  'access':['gudang']},
    'konten':  {'pin':'2345','role':'konten', 'label':'Divisi Konten',  'access':['konten']},
    'live':    {'pin':'3456','role':'live',   'label':'Divisi Live',    'access':['live']},
    'manager': {'pin':'0000','role':'admin',  'label':'Manager / BOD',  'access':['gudang','konten','live','admin']},
}

MENUS = [
    {'key':'home',   'icon':'🏠', 'label':'Home',   'req':'all'},
    {'key':'gudang', 'icon':'📦', 'label':'Gudang', 'req':'gudang'},
    {'key':'konten', 'icon':'🎬', 'label':'Konten', 'req':'konten'},
    {'key':'live',   'icon':'📡', 'label':'Live',   'req':'live'},
]

def check_pin(pin):
    for k,v in ROLES.items():
        if v['pin']==pin: return v['role'],v['label'],v['access']
    return None,None,None

def can_access(role, req):
    if req=='all': return True
    if role=='admin': return True
    return req==role

# Session state
for k,d in [('logged_in',False),('role',None),('label',None),
             ('access',[]),('page','home')]:
    if k not in st.session_state: st.session_state[k]=d

# ── LOGIN ──────────────────────────────────────────────────────────────────────
def show_login():
    _,col,_=st.columns([1,1.1,1])
    with col:
        st.markdown('<br>',unsafe_allow_html=True)
        st.markdown("""
        <div style="text-align:center;margin-bottom:24px;">
            <div style="width:80px;height:80px;background:#F5A623;border-radius:20px;
                        margin:0 auto;font-size:44px;line-height:80px;text-align:center;">🧄</div>
            <h2 style="margin:12px 0 4px;color:#1C1C1E;font-weight:800;font-size:24px;">HSD AGENT</h2>
            <p style="color:#9CA3AF;font-size:13px;margin:0;">Sistem Manajemen Operasional</p>
        </div>
        """,unsafe_allow_html=True)

        with st.container():
            st.markdown('<div class="hsd-card">',unsafe_allow_html=True)
            st.markdown('**Masukkan PIN Divisi**')
            pin=st.text_input('','',max_chars=6,type='password',
                               placeholder='Ketik PIN lalu tekan Enter',
                               label_visibility='collapsed')
            if pin:
                role,label,access=check_pin(pin)
                if role:
                    st.session_state.logged_in=True
                    st.session_state.role=role
                    st.session_state.label=label
                    st.session_state.access=access
                    st.session_state.page='home'
                    st.rerun()
                elif len(pin)>=4:
                    st.error('PIN salah. Coba lagi.')
            st.markdown('</div>',unsafe_allow_html=True)

# ── SIDEBAR ────────────────────────────────────────────────────────────────────
def show_sidebar():
    with st.sidebar:
        # Brand
        st.markdown(f"""
        <div style="text-align:center;padding:16px 0 8px;">
            <div style="font-size:36px;">🧄</div>
            <div style="font-size:16px;font-weight:800;color:white;margin:6px 0 4px;">HSD AGENT</div>
            <div style="background:#F5A623;color:white;padding:3px 14px;border-radius:20px;
                        font-size:11px;font-weight:700;display:inline-block;">
                {st.session_state.label}</div>
        </div>
        """,unsafe_allow_html=True)

        # Jam realtime
        now=datetime.datetime.now()
        st.markdown(f"""
        <div style="text-align:center;background:#2C2C2E;border-radius:8px;
                    padding:8px;margin:8px 0 4px;">
            <div style="font-size:20px;font-weight:800;color:#F5A623;
                        font-family:'JetBrains Mono',monospace;">{now.strftime('%H:%M:%S')}</div>
            <div style="font-size:11px;color:#6B7280;">{now.strftime('%A, %d %B %Y')}</div>
        </div>
        """,unsafe_allow_html=True)

        st.divider()

        # Nav menu — semua menu tampil, dikunci kalau tidak punya akses
        st.markdown('<div style="font-size:10px;color:#6B7280;font-weight:700;letter-spacing:.08em;margin-bottom:6px;">MENU</div>',unsafe_allow_html=True)
        for m in MENUS:
            has_access=can_access(st.session_state.role, m['req'])
            is_active=st.session_state.page==m['key']

            if has_access:
                label_text=f"{m['icon']}  {m['label']}"
                if st.button(label_text, key=f"nav_{m['key']}",
                             use_container_width=True,
                             type='primary' if is_active else 'secondary'):
                    st.session_state.page=m['key']
                    st.rerun()
            else:
                st.markdown(f"""
                <div style="padding:10px 16px;border-radius:8px;margin-bottom:4px;
                            background:#2C2C2E;color:#4B5563;font-size:14px;
                            font-weight:600;cursor:not-allowed;border:1px solid #3C3C3E;">
                    🔒  {m['label']}
                    <span style="float:right;font-size:10px;color:#3C3C3E;">No Access</span>
                </div>
                """,unsafe_allow_html=True)

        st.divider()
        if st.button('🚪  Keluar',use_container_width=True):
            for k in ['logged_in','role','label','access','page']:
                st.session_state[k]=False if k=='logged_in' else (None if k!='page' else 'home') if k!='access' else []
            st.rerun()

# ── HOME PAGE ──────────────────────────────────────────────────────────────────
def show_home():
    now=datetime.datetime.now()
    st.markdown(f"""
    <div style="background:white;border-radius:12px;padding:20px 24px;
                border:1px solid #E5E7EB;margin-bottom:16px;
                display:flex;justify-content:space-between;align-items:center;">
        <div>
            <div style="font-size:22px;font-weight:800;color:#1C1C1E;">
                Selamat datang, {st.session_state.label}! 👋</div>
            <div style="color:#9CA3AF;font-size:13px;margin-top:2px;">
                {now.strftime('%A, %d %B %Y  ·  %H:%M WIB')}</div>
        </div>
        <div style="font-size:40px;">🧄</div>
    </div>
    """,unsafe_allow_html=True)

    st.markdown("""
    <div style="background:white;border-radius:12px;padding:20px 24px;border:1px solid #E5E7EB;margin-bottom:12px;">
        <div style="font-size:13px;font-weight:700;color:#9CA3AF;text-transform:uppercase;
                    letter-spacing:.08em;margin-bottom:12px;">FITUR TERSEDIA</div>
    """,unsafe_allow_html=True)

    for m in MENUS:
        if m['key']=='home': continue
        has=can_access(st.session_state.role,m['req'])
        status='✅ Bisa diakses' if has else '🔒 Perlu akses khusus'
        color='#16A34A' if has else '#9CA3AF'
        st.markdown(f"""
        <div style="display:flex;justify-content:space-between;align-items:center;
                    padding:10px 0;border-bottom:1px solid #F3F4F6;">
            <div style="font-size:14px;font-weight:600;color:#1C1C1E;">
                {m['icon']}  {m['label']}</div>
            <div style="font-size:12px;font-weight:600;color:{color};">{status}</div>
        </div>
        """,unsafe_allow_html=True)

    st.markdown('</div>',unsafe_allow_html=True)

    st.markdown("""
    <div style="background:#FEF3DC;border-radius:12px;padding:16px 20px;
                border:1px solid #FDE68A;">
        <div style="font-weight:700;color:#92400E;margin-bottom:4px;">📌 Panduan Penggunaan</div>
        <div style="color:#78350F;font-size:13px;line-height:1.6;">
        • <b>Gudang</b>: Upload PDF resi per akun (HSD/HSS) dan waktu (Pagi/Siang/Sore), klik Proses, download Excel<br>
        • Excel berisi: Stok Fisik, Semua Resi, Rekap Kurir, Rekap SKU, Packing List, Ringkasan Hari Ini, Ringkasan Order, Kode Pengambilan Gojek<br>
        • <b>Variasi ( - )</b> = data tidak tersedia di PDF resi kurir tersebut<br>
        • PDF besar (6000+ resi) membutuhkan waktu lebih lama, harap bersabar
        </div>
    </div>
    """,unsafe_allow_html=True)

# ── GUDANG PAGE ────────────────────────────────────────────────────────────────
def show_gudang():
    now=datetime.datetime.now()
    st.markdown(f"""
    <div style="background:white;border-radius:12px;padding:14px 20px;
                border:1px solid #E5E7EB;margin-bottom:16px;
                display:flex;justify-content:space-between;align-items:center;">
        <div style="font-size:18px;font-weight:800;color:#1C1C1E;">📦  Upload PDF Resi</div>
        <div style="color:#9CA3AF;font-size:13px;">{now.strftime('%d/%m/%Y  %H:%M:%S WIB')}</div>
    </div>
    """,unsafe_allow_html=True)

    # HSD slots
    st.markdown('<div style="display:flex;align-items:center;gap:8px;margin-bottom:8px;"><span style="background:#3B82F6;color:white;padding:3px 10px;border-radius:6px;font-size:11px;font-weight:700;">HSD</span><span style="color:#3B82F6;font-weight:700;font-size:13px;">Botol Hitam</span></div>',unsafe_allow_html=True)
    c1,c2,c3=st.columns(3)
    uploads={}
    for col,waktu,key in [(c1,'PAGI','hsd_pagi'),(c2,'SIANG','hsd_siang'),(c3,'SORE','hsd_sore')]:
        with col:
            st.markdown(f'<div style="font-size:11px;font-weight:700;color:#6B7280;margin-bottom:4px;">{waktu}</div>',unsafe_allow_html=True)
            f=st.file_uploader(f'HSD {waktu}',type='pdf',key=key,label_visibility='collapsed')
            uploads[('HSD',waktu)]=f
            if f: st.markdown(f'<div class="slot-done">✅ {f.name[:20]}</div>',unsafe_allow_html=True)
            else: st.markdown('<div class="slot-hsd">📄 Pilih PDF</div>',unsafe_allow_html=True)

    st.markdown('<br>',unsafe_allow_html=True)

    # HSS slots
    st.markdown('<div style="display:flex;align-items:center;gap:8px;margin-bottom:8px;"><span style="background:#F97316;color:white;padding:3px 10px;border-radius:6px;font-size:11px;font-weight:700;">HSS</span><span style="color:#F97316;font-weight:700;font-size:13px;">Botol Merah</span></div>',unsafe_allow_html=True)
    c4,c5,c6=st.columns(3)
    for col,waktu,key in [(c4,'PAGI','hss_pagi'),(c5,'SIANG','hss_siang'),(c6,'SORE','hss_sore')]:
        with col:
            st.markdown(f'<div style="font-size:11px;font-weight:700;color:#6B7280;margin-bottom:4px;">{waktu}</div>',unsafe_allow_html=True)
            f=st.file_uploader(f'HSS {waktu}',type='pdf',key=key,label_visibility='collapsed')
            uploads[('HSS',waktu)]=f
            if f: st.markdown(f'<div class="slot-done">✅ {f.name[:20]}</div>',unsafe_allow_html=True)
            else: st.markdown('<div class="slot-hss">📄 Pilih PDF</div>',unsafe_allow_html=True)

    st.markdown('<br>',unsafe_allow_html=True)
    active=[(a,w,f) for (a,w),f in uploads.items() if f]

    cb,ci=st.columns([2,3])
    with cb:
        proses=st.button('⚡   PROSES PDF → EXCEL',type='primary',
                          use_container_width=True,disabled=not active)
    with ci:
        if active:
            st.markdown(f'<div style="padding:14px 0;color:#22C55E;font-weight:600;font-size:13px;">✅ {len(active)} PDF siap diproses</div>',unsafe_allow_html=True)
        else:
            st.markdown('<div style="padding:14px 0;color:#9CA3AF;font-size:13px;">Upload minimal 1 PDF</div>',unsafe_allow_html=True)

    if proses and active:
        from parser import process_pdf
        from excel_writer import write_excel_multi

        all_rows=[]
        pb=st.progress(0,text='Memulai...')
        log=st.empty(); logs=[]

        for fi,(akun,waktu,fo) in enumerate(active):
            logs.append(f'📄 Membaca **{akun} {waktu}** — `{fo.name}`')
            log.markdown('\n\n'.join(logs))
            with tempfile.NamedTemporaryFile(delete=False,suffix='.pdf') as tmp:
                tmp.write(fo.read()); tp=tmp.name
            def prog(cur,tot,fi=fi,tf=len(active),a=akun,w=waktu):
                pct=int((fi/tf+cur/tot/tf)*100)
                pb.progress(pct,text=f'{a} {w}: hal {cur}/{tot}')
            try:
                rows,errs=process_pdf(tp,progress_callback=prog)
                for r in rows: r['akun']=akun; r['waktu']=waktu
                all_rows.extend(rows)
                logs.append(f'✅ {akun} {waktu}: **{len(rows)} baris**')
                log.markdown('\n\n'.join(logs))
            except Exception as e:
                logs.append(f'❌ Error {akun} {waktu}: {e}')
                log.markdown('\n\n'.join(logs))
            finally:
                os.unlink(tp)

        if all_rows:
            pb.progress(95,text='Membuat Excel...')
            with tempfile.NamedTemporaryFile(delete=False,suffix='.xlsx') as tmp:
                txls=tmp.name
            write_excel_multi(all_rows,txls)
            output=BytesIO()
            with open(txls,'rb') as f: output.write(f.read())
            os.unlink(txls); output.seek(0)
            pb.progress(100,text='✅ Selesai!')

            u=len(set(r['no_resi'] for r in all_rows))
            q=sum(r['qty'] for r in all_rows)
            st.markdown('<br>',unsafe_allow_html=True)
            m1,m2,m3,m4=st.columns(4)
            now2=datetime.datetime.now()
            for col,val,lbl in [(m1,u,'Total Resi'),(m2,q,'Total Qty'),
                                 (m3,len(active),'PDF Diproses'),
                                 (m4,now2.strftime('%H:%M'),'Waktu Proses')]:
                with col:
                    st.markdown(f'<div class="metric-box"><div class="metric-val">{val}</div><div class="metric-lbl">{lbl}</div></div>',unsafe_allow_html=True)

            st.markdown('<br>',unsafe_allow_html=True)
            dt=datetime.datetime.now().strftime('%Y%m%d_%H%M')
            st.download_button('📥   Download Excel Rekap',data=output,
                                file_name=f'HSD_REKAP_{dt}.xlsx',
                                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                                use_container_width=True,type='primary')

# ── MAIN ───────────────────────────────────────────────────────────────────────
if not st.session_state.logged_in:
    show_login()
else:
    show_sidebar()
    page=st.session_state.page
    if page=='home': show_home()
    elif page=='gudang':
        if can_access(st.session_state.role,'gudang'): show_gudang()
        else: st.error('Akses ditolak.')
    elif page=='konten':
        if can_access(st.session_state.role,'konten'):
            st.markdown('<h2>🎬 Divisi Konten</h2>',unsafe_allow_html=True)
            st.info('🚧 Segera hadir.')
        else: st.error('Akses ditolak.')
    elif page=='live':
        if can_access(st.session_state.role,'live'):
            st.markdown('<h2>📡 Divisi Live</h2>',unsafe_allow_html=True)
            st.info('🚧 Segera hadir.')
        else: st.error('Akses ditolak.')      
