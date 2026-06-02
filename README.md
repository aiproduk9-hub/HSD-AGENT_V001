# HSD AGENT — Web App

## Deploy ke Streamlit Cloud (Gratis)

### Langkah 1: Upload ke GitHub
1. Buat akun di [github.com](https://github.com) (gratis)
2. Buat repository baru, nama: `hsd-agent`
3. Upload semua file ini ke repository

### Langkah 2: Deploy ke Streamlit Cloud
1. Buka [share.streamlit.io](https://share.streamlit.io)
2. Login dengan GitHub
3. Klik **New app**
4. Pilih repository `hsd-agent`
5. Main file path: `app.py`
6. Klik **Deploy**

### Selesai!
Dalam 2-3 menit app online. Dapat link URL seperti:
`https://hsd-agent-xxxxx.streamlit.app`

Bisa dibuka dari HP, tablet, atau PC manapun.

---

## File yang Diperlukan
- `app.py` — UI web app
- `parser.py` — baca PDF resi
- `excel_writer.py` — buat Excel
- `requirements.txt` — library Python

## PIN Default
- Gudang: 1234
- Konten: 2345
- Live: 3456
- Manager/BOD: 0000
