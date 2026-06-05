import re
from collections import defaultdict
from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

FONT_NAME = "Calibri"

# =========================
# STYLE
# =========================
DARK      = PatternFill("solid", fgColor="111827")
HEADER    = PatternFill("solid", fgColor="1B1F2E")
WHITE     = PatternFill("solid", fgColor="FFFFFF")
GREEN     = PatternFill("solid", fgColor="E2F0D9")
GREEN_DK  = PatternFill("solid", fgColor="16A34A")
ORANGE    = PatternFill("solid", fgColor="FFE4CC")
ORANGE_DK = PatternFill("solid", fgColor="F97316")
BLUE      = PatternFill("solid", fgColor="DBEAFE")
BLUE_DK   = PatternFill("solid", fgColor="2563EB")
RED_LT    = PatternFill("solid", fgColor="FFEBEE")
RED_DK    = PatternFill("solid", fgColor="DC2626")
YELLOW    = PatternFill("solid", fgColor="FFF9C4")
GREY      = PatternFill("solid", fgColor="F3F4F6")
PURPLE    = PatternFill("solid", fgColor="EDE9FE")
PURPLE_DK = PatternFill("solid", fgColor="7C3AED")

CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT   = Alignment(horizontal="left",   vertical="center", wrap_text=True)

BORDER = Border(
    left=Side(style="thin", color="AAAAAA"),
    right=Side(style="thin", color="AAAAAA"),
    top=Side(style="thin", color="AAAAAA"),
    bottom=Side(style="thin", color="AAAAAA"),
)

def fnt(size=10, bold=True, color="111827"):
    return Font(name=FONT_NAME, size=size, bold=bold, color=color)

# =========================
# MASTER PAKET
# =========================
# Setiap paket → list komponen (nama_item, qty_per_paket)
# "tampilkan_nama_paket" = True → baris header paket ditampilkan sebelum komponen

PAKET_MAP = {
    # 3in1
    "BG 3in1 (100gr+220gr+500gr)": {
        "label": "PAKET 3in1 (100gr + 220gr + 500gr)",
        "komponen": [
            ("Black Garlic 100gr", 1),
            ("Black Garlic 220gr", 1),
            ("Black Garlic 500gr", 1),
        ]
    },
    # Hampers
    "Black Garlic 220gr x2 + Box Hampers Imlek": {
        "label": "PAKET HAMPERS IMLEK (220gr x2 + Box)",
        "komponen": [
            ("Black Garlic 220gr", 2),
            ("Box Hampers Imlek", 1),
        ]
    },
    "Black Garlic 220gr x2 + Box Hampers Lebaran": {
        "label": "PAKET HAMPERS LEBARAN (220gr x2 + Box)",
        "komponen": [
            ("Black Garlic 220gr", 2),
            ("Box Hampers Lebaran", 1),
        ]
    },
    "Black Garlic 220gr x2 + Box Hampers Natal": {
        "label": "PAKET HAMPERS NATAL (220gr x2 + Box)",
        "komponen": [
            ("Black Garlic 220gr", 2),
            ("Box Hampers Natal", 1),
        ]
    },
    "Black Garlic 220gr x2 + Box Hampers HSD": {
        "label": "PAKET HAMPERS HSD (220gr x2 + Box)",
        "komponen": [
            ("Black Garlic 220gr", 2),
            ("Box Hampers HSD", 1),
        ]
    },
    # Drink Mix
    "BG Drink Mix 7 Botol (4 Peach + 3 Original)": {
        "label": "PAKET DRINK MIX 7 BOTOL (4 Peach + 3 Original)",
        "komponen": [
            ("BG Drink Peach", 4),
            ("BG Drink Original", 3),
        ]
    },
}

# Urutan tampil di Sheet 1
URUTAN_ITEM = [
    "Black Garlic 100gr",
    "Black Garlic 220gr",
    "Black Garlic 500gr",
    "Black Garlic 84gr",
    "BG Drink Original",
    "BG Drink Peach",
    "Madu Bunga Kurma",
    "Madu Multi Floral",
    "Goodie Bag HSD",
    "Box Hampers Imlek",
    "Box Hampers HSD",
    "Box Hampers Natal",
    "Box Hampers Lebaran",
]

# =========================
# HELPERS
# =========================
def clean(v):
    if v is None: return ""
    v = re.sub(r"[\x00-\x1F]", "", str(v))
    return " ".join(v.split()).strip()

def to_int(v):
    try: return int(float(str(v).replace(",","").strip()))
    except: return 0

def safe(r, k, d=""):
    try:
        v = r.get(k, d)
        return v if v is not None else d
    except: return d

def get_akun(r):
    v = clean(safe(r,"akun")) or clean(safe(r,"brand"))
    v = v.upper()
    return v if v in ["HSD","HSS"] else (v or "-")

def get_shift(r):
    v = clean(safe(r,"shift")) or clean(safe(r,"waktu"))
    return v.title() if v else "-"

def set_view(ws):
    ws.sheet_view.showGridLines = False

def title_bar(ws, rng, text, fill=DARK, color="FFFFFF", size=13):
    ws.merge_cells(rng)
    c = ws[rng.split(":")[0]]
    c.value = text
    c.fill = fill
    c.font = fnt(size=size, bold=True, color=color)
    c.alignment = CENTER
    c.border = BORDER
    r0, c0_str = rng.split(":")[0], rng.split(":")[1]
    r_num = int("".join(filter(str.isdigit, r0)))
    for col_idx in range(ws[r0].column, ws[c0_str].column + 1):
        cell = ws.cell(row=r_num, column=col_idx)
        cell.fill = fill
        cell.border = BORDER

def hdr(ws, headers, row=1):
    """headers = list of (label, width)"""
    for col, (label, width) in enumerate(headers, 1):
        c = ws.cell(row=row, column=col, value=label)
        c.fill = HEADER
        c.font = fnt(size=10, bold=True, color="FFFFFF")
        c.alignment = CENTER
        c.border = BORDER
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.row_dimensions[row].height = 26
    ws.freeze_panes = f"A{row+1}"

def wrow(ws, row, vals, fill=WHITE, centers=None):
    centers = centers or []
    for col, val in enumerate(vals, 1):
        c = ws.cell(row=row, column=col, value=val)
        c.fill = fill
        c.font = fnt(size=10, bold=False, color="111827")
        c.alignment = CENTER if col in centers else LEFT
        c.border = BORDER

def total_row(ws, row, vals, fill=ORANGE_DK):
    for col, val in enumerate(vals, 1):
        c = ws.cell(row=row, column=col, value=val)
        c.fill = fill
        c.font = fnt(size=11, bold=True, color="FFFFFF")
        c.alignment = CENTER
        c.border = BORDER

# =========================
# EXPAND PRODUK → KOMPONEN
# =========================
def expand_to_komponen(nama_produk, qty_order):
    """
    Kembalikan list (nama_item, qty_total) untuk satu baris order.
    Untuk paket, breakdown ke komponen masing-masing.
    """
    qty_order = to_int(qty_order) or 1

    if nama_produk in PAKET_MAP:
        hasil = []
        for item, qty_per_paket in PAKET_MAP[nama_produk]["komponen"]:
            hasil.append((item, qty_per_paket * qty_order))
        return hasil, nama_produk  # kembalikan juga nama paket asli

    # Produk biasa yang punya suffix x2, x3 dst
    m = re.match(r"^(.+?)\s+x(\d+)$", nama_produk)
    if m:
        base = m.group(1).strip()
        mult = int(m.group(2))
        return [(base, mult * qty_order)], None

    return [(nama_produk, qty_order)], None

# =========================
# PREPARE DATA
# =========================
def prepare(rows):
    rows = rows or []

    # Struktur: {akun: {shift: {nama_item: total_qty}}}
    rekap = defaultdict(lambda: defaultdict(lambda: defaultdict(int)))

    # Untuk sheet 2: per akun+shift, kumpulkan kombinasi produk per resi
    # {akun: {shift: {combo_str: {"count": N, "resi": [...]}}}}
    variasi = defaultdict(lambda: defaultdict(lambda: defaultdict(lambda: {"count":0,"resi":[]})))

    # Track per resi agar tidak dobel
    resi_seen = set()
    resi_produk = defaultdict(list)  # resi → list nama_produk

    for r in rows:
        akun  = get_akun(r)
        shift = get_shift(r)
        resi  = clean(safe(r,"no_resi")) or f"no_resi_{id(r)}"
        nama  = clean(safe(r,"nama_produk"))
        qty   = to_int(safe(r,"qty",1)) or 1

        if not nama or nama == "(cek manual)":
            continue

        komponen, _ = expand_to_komponen(nama, qty)
        for item, total in komponen:
            rekap[akun][shift][item] += total

        # Kumpulkan produk per resi untuk sheet 2
        resi_produk[(akun, shift, resi)].append((nama, qty))

    # Bangun variasi per akun+shift
    for (akun, shift, resi), produk_list in resi_produk.items():
        # Buat string combo yang rapi
        parts = []
        for nama, qty in sorted(produk_list, key=lambda x: x[0]):
            parts.append(f"{nama} x{qty}" if qty > 1 else nama)
        combo = " + ".join(parts)
        variasi[akun][shift][combo]["count"] += 1
        variasi[akun][shift][combo]["resi"].append(resi)

    return rekap, variasi

# =========================
# SHEET 1: KEBUTUHAN BARANG
# =========================
def write_sheet1(wb, rekap, tgl_str):
    ws = wb.active
    ws.title = "KEBUTUHAN BARANG GUDANG"
    ws.sheet_tab_color = "F97316"
    set_view(ws)

    # Kumpulkan semua akun & shift
    akun_list  = sorted(rekap.keys())
    shift_list = ["Pagi","Siang","Sore"]

    # Header kolom: No | Nama Barang/Komponen | HSD Pagi | HSD Siang | HSD Sore | HSS Pagi | ... | TOTAL
    cols = [("No", 5), ("Nama Barang / Komponen", 38)]
    for akun in akun_list:
        for sh in shift_list:
            cols.append((f"{akun}\n{sh}", 10))
    cols.append(("TOTAL", 10))

    n_data_cols = len(cols) - 2  # jumlah kolom akun+shift

    title_bar(ws, f"A1:{get_column_letter(len(cols))}1",
              f"REKAP KEBUTUHAN BARANG GUDANG — {tgl_str}", DARK, size=13)
    title_bar(ws, f"A2:{get_column_letter(len(cols))}2",
              "Siapkan stok sesuai tabel di bawah sebelum tempel resi ke dus packaging",
              BLUE_DK, size=10)

    hdr(ws, cols, row=3)

    # Kumpulkan semua nama_item yang muncul
    all_items = set()
    for akun in rekap:
        for sh in rekap[akun]:
            all_items.update(rekap[akun][sh].keys())

    # Pisah: item biasa vs item yg merupakan komponen paket vs PERLU CEK
    item_biasa = [i for i in URUTAN_ITEM if i in all_items]
    item_lain  = sorted([i for i in all_items if i not in URUTAN_ITEM and "PERLU CEK" not in i])
    item_cek   = sorted([i for i in all_items if "PERLU CEK" in i])

    row = 4
    no  = 1

    def get_qty(akun, shift, item):
        return rekap.get(akun, {}).get(shift, {}).get(item, 0) or ""

    def write_item_row(item, fill=WHITE, bold=False, prefix=""):
        nonlocal row, no
        vals = [no, f"{prefix}{item}"]
        total = 0
        for akun in akun_list:
            for sh in shift_list:
                q = rekap.get(akun,{}).get(sh,{}).get(item, 0)
                vals.append(q or "")
                total += q
        vals.append(total or "")
        wrow(ws, row, vals,
             fill=fill,
             centers=list(range(3, len(vals)+1)))
        if bold:
            for col in range(1, len(vals)+1):
                ws.cell(row=row, column=col).font = fnt(size=10, bold=True, color="111827")
        row += 1
        no += 1

    def write_paket_block(nama_produk):
        """Tulis baris header paket + baris setiap komponen."""
        nonlocal row, no
        info = PAKET_MAP[nama_produk]
        label = info["label"]

        # Hitung qty paket per akun+shift
        # qty paket = rekap[akun][shift][nama_produk] jika ada
        # (parser sudah resolve nama_produk ke nama paket)
        paket_qty = {}
        for akun in akun_list:
            for sh in shift_list:
                q = rekap.get(akun,{}).get(sh,{}).get(nama_produk, 0)
                paket_qty[(akun,sh)] = q

        total_paket = sum(paket_qty.values())
        if total_paket == 0:
            return  # tidak ada order paket ini

        # Baris header paket (orange)
        vals_hdr = [no, f"▶ {label}"]
        for akun in akun_list:
            for sh in shift_list:
                q = paket_qty[(akun,sh)]
                vals_hdr.append(q or "")
        vals_hdr.append(total_paket or "")
        wrow(ws, row, vals_hdr, fill=ORANGE,
             centers=list(range(3, len(vals_hdr)+1)))
        for col in range(1, len(vals_hdr)+1):
            ws.cell(row=row, column=col).font = fnt(size=10, bold=True, color="C2410C")
        row += 1
        no += 1

        # Baris komponen
        for item, qty_per in info["komponen"]:
            vals_k = ["", f"   └ {item} (x{qty_per} per paket)"]
            total_k = 0
            for akun in akun_list:
                for sh in shift_list:
                    q = paket_qty[(akun,sh)] * qty_per
                    vals_k.append(q if q else "")
                    total_k += q
            vals_k.append(total_k if total_k else "")
            wrow(ws, row, vals_k, fill=YELLOW,
                 centers=list(range(3, len(vals_k)+1)))
            row += 1

    # ---- Tulis item biasa ----
    for item in item_biasa:
        write_item_row(item, fill=WHITE if no % 2 == 1 else GREY)

    # ---- Tulis paket yang ada ----
    paket_ada = [p for p in PAKET_MAP if p in all_items]
    if paket_ada:
        # Separator
        ws.merge_cells(f"A{row}:{get_column_letter(len(cols))}{row}")
        c = ws.cell(row=row, column=1, value="PAKET (nama paket + rincian komponen)")
        c.fill = PURPLE_DK
        c.font = fnt(size=10, bold=True, color="FFFFFF")
        c.alignment = CENTER
        c.border = BORDER
        for col in range(2, len(cols)+1):
            ws.cell(row=row, column=col).fill = PURPLE_DK
            ws.cell(row=row, column=col).border = BORDER
        row += 1
        for p in paket_ada:
            write_paket_block(p)

    # ---- Item lain yg tidak ada di daftar ----
    if item_lain:
        for item in item_lain:
            write_item_row(item, fill=BLUE)

    # ---- PERLU CEK ----
    if item_cek:
        ws.merge_cells(f"A{row}:{get_column_letter(len(cols))}{row}")
        c = ws.cell(row=row, column=1, value="⚠ BELUM TERIDENTIFIKASI — Cek manual sebelum ambil stok")
        c.fill = RED_DK
        c.font = fnt(size=10, bold=True, color="FFFFFF")
        c.alignment = CENTER
        c.border = BORDER
        for col in range(2, len(cols)+1):
            ws.cell(row=row, column=col).fill = RED_DK
            ws.cell(row=row, column=col).border = BORDER
        row += 1
        for item in item_cek:
            write_item_row(item, fill=RED_LT)

    # ---- TOTAL ROW ----
    total_vals = ["", "TOTAL SEMUA BARANG"]
    grand = 0
    for akun in akun_list:
        for sh in shift_list:
            subtotal = sum(rekap.get(akun,{}).get(sh,{}).get(i,0) for i in all_items)
            total_vals.append(subtotal or "")
            grand += subtotal
    total_vals.append(grand)
    total_row(ws, row, total_vals)

    ws.row_dimensions[1].height = 28
    ws.row_dimensions[2].height = 20

# =========================
# SHEET 2: VARIASI ORDER
# =========================
def write_sheet2(wb, variasi, tgl_str):
    ws = wb.create_sheet("VARIASI ORDER PER RESI")
    ws.sheet_tab_color = "2563EB"
    set_view(ws)

    title_bar(ws, "A1:F1",
              f"VARIASI ORDER PER RESI — {tgl_str}",
              BLUE_DK, size=13)
    title_bar(ws, "A2:F2",
              "Tabel ini menunjukkan pola pembelian: berapa resi beli 1 produk, berapa yang beli kombinasi",
              DARK, size=10)

    hdr(ws, [
        ("No",    5),
        ("Akun",  8),
        ("Shift", 10),
        ("Kombinasi Produk yang Dibeli (per 1 Resi)", 70),
        ("Jumlah Resi", 14),
        ("% dari Total", 14),
    ], row=3)

    row = 4
    no  = 1
    akun_list  = sorted(variasi.keys())
    shift_list = ["Pagi","Siang","Sore"]

    for akun in akun_list:
        for sh in shift_list:
            if sh not in variasi[akun]:
                continue

            data = variasi[akun][sh]
            if not data:
                continue

            total_resi = sum(v["count"] for v in data.values())

            # Separator akun+shift
            ws.merge_cells(f"A{row}:F{row}")
            c = ws.cell(row=row, column=1,
                        value=f"  {akun}  —  Shift {sh}  —  Total {total_resi} resi")
            fill_sep = GREEN_DK if akun == "HSD" else BLUE_DK
            c.fill = fill_sep
            c.font = fnt(size=11, bold=True, color="FFFFFF")
            c.alignment = LEFT
            c.border = BORDER
            for col in range(2, 7):
                ws.cell(row=row, column=col).fill = fill_sep
                ws.cell(row=row, column=col).border = BORDER
            row += 1

            # Sort: terbanyak dulu
            sorted_combos = sorted(data.items(), key=lambda x: -x[1]["count"])

            for combo, info in sorted_combos:
                count = info["count"]
                pct   = f"{100*count//total_resi}%" if total_resi else "-"
                fill  = GREEN if no % 2 == 1 else WHITE
                wrow(ws, row,
                     [no, akun, sh, combo, count, pct],
                     fill=fill, centers=[1,2,3,5,6])
                row += 1
                no  += 1

            # Subtotal shift
            total_row(ws, row,
                      ["", "", "", f"SUBTOTAL {akun} {sh}", total_resi, "100%"],
                      fill=ORANGE_DK)
            row += 1

    ws.column_dimensions["A"].width = 5
    ws.column_dimensions["B"].width = 8
    ws.column_dimensions["C"].width = 10
    ws.column_dimensions["D"].width = 70
    ws.column_dimensions["E"].width = 14
    ws.column_dimensions["F"].width = 14

# =========================
# MAIN
# =========================
def write_excel_multi(rows, output_path):
    wb = Workbook()
    rows = rows or []

    tgl_str = datetime.now().strftime("%d/%m/%Y %H:%M")
    rekap, variasi = prepare(rows)

    write_sheet1(wb, rekap, tgl_str)
    write_sheet2(wb, variasi, tgl_str)

    wb.save(output_path)
    return output_path

# Alias agar app.py tetap jalan
def write_excel(rows, output_path):       return write_excel_multi(rows, output_path)
def create_excel(rows, output_path):      return write_excel_multi(rows, output_path)
def generate_excel(rows, output_path):    return write_excel_multi(rows, output_path)
def build_excel(rows, output_path):       return write_excel_multi(rows, output_path)
def make_excel(rows, output_path):        return write_excel_multi(rows, output_path)
def export_excel(rows, output_path):      return write_excel_multi(rows, output_path)
