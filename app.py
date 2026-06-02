import streamlit as st
import tempfile, os, datetime
from io import BytesIO

st.set_page_config(page_title='HSD AGENT',page_icon='🧄',layout='wide')

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;600;700;800&display=swap');
html,[class*="css"]{font-family:'Plus Jakarta Sans',sans-serif!important;}
#MainMenu,footer,header,.stDeployButton{visibility:hidden;display:none;}
.stApp{background:#F8F6F1;}
.stButton>button{font-family:'Plus Jakarta Sans',sans-serif!important;font-weight:700!important;border-radius:8px!important;border:none!important;transition:all .15s;}
.stButton>button:hover{transform:translateY(-1px);}
div[data-testid="stButton"]>button[kind="primary"]{background:#F5A623!important;color:white!important;font-size:15px!important;padding:14px!important;}
.card{background:white;border-radius:12px;padding:20px;border:1px solid #E5E7EB;margin-bottom:12px;}
.slot-hsd{background:#EFF6FF;border:1.5px dashed #93C5FD;border-radius:10px;padding:12px;text-align:center;color:#3B82F6;font-size:12px;font-weight:600;}
.slot-hss{background:#FFF7ED;border:1.5px dashed #FDBA74;border-radius:10px;padding:12px;text-align:center;color:#F97316;font-size:12px;font-weight:600;}
.slot-done{background:#F0FDF4;border:1.5px solid #86EFAC;border-radius:10px;padding:12px;text-align:center;color:#16A34A;font-size:12px;font-weight:600;}
.nav-btn{display:inline-block;padding:8px 18px;border-radius:8px;font-weight:700;font-size:13px;cursor:pointer;margin:2px;}
.metric-box{background:white;border-radius:10px;padding:16px;border:1px solid #E5E7EB;text-align:center;}
.metric-val{font-size:26px;font-weight:800;color:#1C1C1E;}
.metric-lbl{font-size:10px;color:#9CA3AF;font-weight:600;text-transform:uppercase;}
[data-testid="stFileUploader"]{background:white!important;border-radius:8px!important;}
.stProgress>div>div{background:#F5A623!important;border-radius:99px!important;}
section[data-testid="stSidebar"]{display:none!important;}
</style>
""",unsafe_allow_html=True)

ROLES={
    'gudang': {'pin':'1234','role':'gudang','label':'Divisi Gudang','access':['gudang']},
    'konten': {'pin':'2345','role':'konten','label':'Divisi Konten','access':['konten']},
    'live':   {'pin':'3456','role':'live',  'label':'Divisi Live',  'access':['live']},
    'manager':{'pin':'0000','role':'admin', 'label':'Manager / BOD','access':['gudang','konten','live']},
}
MENUS=[('home','🏠','Home'),('gudang','📦','Gudang'),
       ('konten','🎬','Konten'),('live','📡','Live')]

def check_pin(p):
    for v in ROLES.values():
        if v['pin']==p: return v['role'],v['label'],v['access']
    return None,None,None

def can(role,req):
    if req=='home': return True
    if role=='admin': return True
    return req==role

for k,d in [('ok',False),('role',None),('label',None),('access',[]),('page','home')]:
    if k not in st.session_state: st.session_state[k]=d

# ── LOGIN ──────────────────────────────────────────────────────────────────────
if not st.session_state.ok:
    _,col,_=st.columns([1,1.2,1])
    with col:
        st.markdown('<br><br>',unsafe_allow_html=True)
        st.markdown("""
        <div style="text-align:center;margin-bottom:20px;">
            <div style="width:80px;height:80px;background:#F5A623;border-radius:20px;
                margin:0 auto;font-size:44px;line-height:80px;">🧄</div>
            <h2 style="margin:12px 0 2px;color:#1C1C1E;font-weight:800;">HSD AGENT</h2>
            <p style="color:#9CA3AF;font-size:13px;">Sistem Manajemen Operasional</p>
        </div>""",unsafe_allow_html=True)
        with st.form('login'):
            pin=st.text_input('PIN','',type='password',placeholder='Masukkan PIN lalu Enter',label_visibility='collapsed')
            ok=st.form_submit_button('🔐  Masuk',use_container_width=True,type='primary')
        if ok and pin:
            role,label,access=check_pin(pin)
            if role:
                st.session_state.ok=True; st.session_state.role=role
                st.session_state.label=label; st.session_state.access=access
                st.rerun()
            else:
                st.error('PIN salah.')
        # PIN info
        st.markdown('<div style="background:#1C1C1E;border-radius:10px;padding:14px;margin-top:8px;">',unsafe_allow_html=True)
        st.markdown('<div style="font-size:10px;color:#6B7280;font-weight:700;margin-bottom:8px;">PIN DIVISI</div>',unsafe_allow_html=True)
        for v in ROLES.values():
            st.markdown(f'<div style="display:flex;justify-content:space-between;padding:3px 0;border-bottom:1px solid #2C2C2E;"><span style="color:#9CA3AF;font-size:12px;">{v["label"]}</span><span style="background:#F5A623;color:white;padding:1px 8px;border-radius:4px;font-size:11px;font-weight:700;">{v["pin"]}</span></div>',unsafe_allow_html=True)
        st.markdown('</div>',unsafe_allow_html=True)
    st.stop()

# ── TOPBAR ─────────────────────────────────────────────────────────────────────
now=datetime.datetime.now()
tcol1,tcol2=st.columns([3,1])
with tcol1:
    st.markdown(f"""
    <div style="background:white;border-radius:12px;padding:12px 20px;
        border:1px solid #E5E7EB;margin-bottom:12px;">
        <span style="font-size:18px;font-weight:800;color:#1C1C1E;">🧄 HSD AGENT</span>
        <span style="background:#F5A623;color:white;padding:2px 10px;border-radius:20px;
            font-size:11px;font-weight:700;margin-left:10px;">{st.session_state.label}</span>
        <span style="color:#9CA3AF;font-size:12px;margin-left:12px;">
            {now.strftime('%A, %d %B %Y  ·  %H:%M WIB')}</span>
    </div>""",unsafe_allow_html=True)
with tcol2:
    if st.button('🚪 Keluar',use_container_width=True):
        for k in ['ok','role','label','access']: st.session_state[k]=False if k=='ok' else (None if k!='access' else [])
        st.session_state.page='home'; st.rerun()

# ── NAVIGATION BAR ─────────────────────────────────────────────────────────────
nav_cols=st.columns(4)
for i,(key,icon,label) in enumerate(MENUS):
    with nav_cols[i]:
        has=can(st.session_state.role,key)
        is_active=st.session_state.page==key
        if has:
            if st.button(f'{icon}  {label}',use_container_width=True,
                        type='primary' if is_active else 'secondary',key=f'nav_{key}'):
                st.session_state.page=key; st.rerun()
        else:
            st.button(f'🔒  {label}',use_container_width=True,disabled=True,key=f'nav_{key}')

st.markdown('<hr style="margin:8px 0 16px;border-color:#E5E7EB;">',unsafe_allow_html=True)

# ── PAGES ──────────────────────────────────────────────────────────────────────
page=st.session_state.page

if page=='home':
    st.markdown(f"""
    <div class="card">
        <h3 style="margin:0 0 4px;color:#1C1C1E;">Selamat datang, {st.session_state.label}! 👋</h3>
        <p style="color:#9CA3AF;margin:0;font-size:13px;">{now.strftime('%A, %d %B %Y')}</p>
    </div>""",unsafe_allow_html=True)
    st.markdown("""
    <div class="card">
        <div style="font-size:11px;font-weight:700;color:#9CA3AF;text-transform:uppercase;margin-bottom:12px;">PANDUAN PENGGUNAAN</div>
        <div style="font-size:13px;color:#374151;line-height:1.8;">
        📦 <b>Gudang</b>: Upload PDF resi per akun (HSD/HSS) dan waktu (Pagi/Siang/Sore) → Klik Proses → Download Excel<br>
        📊 <b>Excel berisi</b>: Stok Fisik, Semua Resi, Rekap Kurir, Rekap SKU, Packing List, Ringkasan Hari Ini, Ringkasan Order, Kode Pengambilan Gojek<br>
        ℹ️ <b>Variasi ( - )</b> = data variasi tidak tercantum di PDF resi kurir tersebut<br>
        ⏳ <b>PDF besar</b> (6000+ resi) butuh waktu lebih lama, harap bersabar
        </div>
    </div>""",unsafe_allow_html=True)

elif page=='gudang':
    if not can(st.session_state.role,'gudang'): st.error('Akses ditolak.'); st.stop()

    st.markdown(f'<div class="card"><b style="font-size:16px;">📦 Upload PDF Resi</b><span style="color:#9CA3AF;font-size:12px;margin-left:12px;">{now.strftime("%d/%m/%Y %H:%M WIB")}</span></div>',unsafe_allow_html=True)

    uploads={}
    # HSD
    st.markdown('<div style="display:flex;align-items:center;gap:8px;margin-bottom:8px;"><span style="background:#3B82F6;color:white;padding:2px 10px;border-radius:6px;font-size:11px;font-weight:700;">HSD</span><span style="color:#3B82F6;font-weight:700;">Botol Hitam</span></div>',unsafe_allow_html=True)
    c1,c2,c3=st.columns(3)
    for col,w,k in [(c1,'PAGI','hsd_p'),(c2,'SIANG','hsd_s'),(c3,'SORE','hsd_r')]:
        with col:
            st.markdown(f'<div style="font-size:11px;font-weight:700;color:#6B7280;margin-bottom:4px;">{w}</div>',unsafe_allow_html=True)
            f=st.file_uploader(f'HSD {w}',type='pdf',key=k,label_visibility='collapsed')
            uploads[('HSD',w)]=f
            st.markdown(f'<div class="{"slot-done" if f else "slot-hsd"}">{"✅ "+f.name[:18] if f else "📄 Pilih PDF"}</div>',unsafe_allow_html=True)

    st.markdown('<br>',unsafe_allow_html=True)
    # HSS
    st.markdown('<div style="display:flex;align-items:center;gap:8px;margin-bottom:8px;"><span style="background:#F97316;color:white;padding:2px 10px;border-radius:6px;font-size:11px;font-weight:700;">HSS</span><span style="color:#F97316;font-weight:700;">Botol Merah</span></div>',unsafe_allow_html=True)
    c4,c5,c6=st.columns(3)
    for col,w,k in [(c4,'PAGI','hss_p'),(c5,'SIANG','hss_s'),(c6,'SORE','hss_r')]:
        with col:
            st.markdown(f'<div style="font-size:11px;font-weight:700;color:#6B7280;margin-bottom:4px;">{w}</div>',unsafe_allow_html=True)
            f=st.file_uploader(f'HSS {w}',type='pdf',key=k,label_visibility='collapsed')
            uploads[('HSS',w)]=f
            st.markdown(f'<div class="{"slot-done" if f else "slot-hss"}">{"✅ "+f.name[:18] if f else "📄 Pilih PDF"}</div>',unsafe_allow_html=True)

    st.markdown('<br>',unsafe_allow_html=True)
    active=[(a,w,f) for (a,w),f in uploads.items() if f]
    cb,ci=st.columns([2,3])
    with cb:
        proses=st.button('⚡  PROSES PDF → EXCEL',type='primary',use_container_width=True,disabled=not active)
    with ci:
        msg=f'✅ {len(active)} PDF siap' if active else 'Upload minimal 1 PDF'
        clr='#22C55E' if active else '#9CA3AF'
        st.markdown(f'<div style="padding:14px 0;color:{clr};font-weight:600;font-size:13px;">{msg}</div>',unsafe_allow_html=True)

    if proses and active:
        from parser import process_pdf
        from excel_writer import write_excel_multi
        all_rows=[]; pb=st.progress(0,'Memulai...'); log=st.empty(); logs=[]
        for fi,(akun,waktu,fo) in enumerate(active):
            logs.append(f'📄 **{akun} {waktu}** — `{fo.name}`')
            log.markdown('\n\n'.join(logs))
            with tempfile.NamedTemporaryFile(delete=False,suffix='.pdf') as tmp:
                tmp.write(fo.read()); tp=tmp.name
            def prog(cur,tot,fi=fi,tf=len(active),a=akun,w=waktu):
                pb.progress(int((fi/tf+cur/tot/tf)*100),text=f'{a} {w}: hal {cur}/{tot}')
            try:
                rows,_=process_pdf(tp,progress_callback=prog)
                for r in rows: r['akun']=akun; r['waktu']=waktu
                all_rows.extend(rows)
                logs.append(f'✅ {akun} {waktu}: **{len(rows)} baris**')
                log.markdown('\n\n'.join(logs))
            except Exception as e:
                logs.append(f'❌ {akun} {waktu}: {e}')
                log.markdown('\n\n'.join(logs))
            finally: os.unlink(tp)

        if all_rows:
            pb.progress(95,'Membuat Excel...')
            with tempfile.NamedTemporaryFile(delete=False,suffix='.xlsx') as tmp: txls=tmp.name
            write_excel_multi(all_rows,txls)
            out=BytesIO()
            with open(txls,'rb') as f: out.write(f.read())
            os.unlink(txls); out.seek(0)
            pb.progress(100,'✅ Selesai!')
            u=len(set(r['no_resi'] for r in all_rows)); q=sum(r['qty'] for r in all_rows)
            st.markdown('<br>',unsafe_allow_html=True)
            for col,val,lbl in zip(st.columns(4),[u,q,len(active),now.strftime('%H:%M')],
                                    ['Total Resi','Total Qty','PDF Diproses','Waktu Proses']):
                with col:
                    st.markdown(f'<div class="metric-box"><div class="metric-val">{val}</div><div class="metric-lbl">{lbl}</div></div>',unsafe_allow_html=True)
            st.markdown('<br>',unsafe_allow_html=True)
            st.download_button('📥  Download Excel Rekap',data=out,
                file_name=f'HSD_REKAP_{now.strftime("%Y%m%d_%H%M")}.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                use_container_width=True,type='primary')

elif page in ('konten','live'):
    label={'konten':'🎬 Divisi Konten','live':'📡 Divisi Live'}[page]
    if not can(st.session_state.role,page): st.error('Akses ditolak.'); st.stop()
    st.markdown(f'<div class="card"><h3>{label}</h3><p style="color:#9CA3AF;">Fitur ini segera hadir. 🚧</p></div>',unsafe_allow_html=True)
