import re
from collections import defaultdict
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

FONT_NAME = "Calibri"

# =========================
# STYLE HELPERS
# =========================
def hdr_fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def make_font(bold=False, size=11, color="000000"):
    return Font(name=FONT_NAME, bold=bold, size=size, color=color)

def center():
    return Alignment(horizontal="center", vertical="center", wrap_text=True)

def left():
    return Alignment(horizontal="left", vertical="center", wrap_text=True)

def thin_border():
    s = Side(style="thin")
    return Border(left=s, right=s, top=s, bottom=s)

def style_cell(cell, bold=False, size=11, fg=None, font_color="000000", align="left"):
    cell.font = make_font(bold=bold, size=size, color=font_color)
    cell.alignment = center() if align == "center" else left()
    if fg:
        cell.fill = hdr_fill(fg)
    cell.border = thin_border()

# =========================
# MASTER PRODUK & PAKET
# =========================

URUTAN_PRODUK = [
    "Black Garlic 84gr",
    "Black Garlic 100gr",
    "Black Garlic 220gr",
    "Black Garlic 500gr",
    "BG Drink Original",
    "BG Drink Peach",
    "BG Madu Multi Floral",
    "BG Madu Kurma",
    "Box Hampers",
]

PAKET_KOMPONEN = {
    "Paket 3in1": [
        ("Black Garlic 100gr", 1),
        ("Black Garlic 220gr", 1),
        ("Black Garlic 500gr", 1),
    ],
    "Hampers Imlek": [
        ("Black Garlic 220gr", 2),
        ("Box Hampers", 1),
    ],
    "Hampers Lebaran": [
        ("Black Garlic 220gr", 2),
        ("Box Hampers", 1),
    ],
    "Hampers Natal": [
        ("Black Garlic 220gr", 2),
        ("Box Hampers", 1),
    ],
    "Hampers HSD": [
        ("Black Garlic 220gr", 2),
        ("Box Hampers", 1),
    ],
    "BG Drink Original x2": [("BG Drink Original", 2)],
    "BG Drink Original x4": [("BG Drink Original", 4)],
    "BG Drink Original 7 Botol": [("BG Drink Original", 7)],
    "BG Drink Original x7": [("BG Drink Original", 7)],
    "BG Drink Peach 7 Botol": [("BG Drink Peach", 7)],
    "BG Drink Peach x7": [("BG Drink Peach", 7)],
    "BG Drink Mix 7 Botol": [
        ("BG Drink Peach", 4),
        ("BG Drink Original", 3),
    ],
    "BG Drink Mix 7 Botol (4 Peach + 3 Original)": [
        ("BG Drink Peach", 4),
        ("BG Drink Original", 3),
    ],
    "Black Garlic 100gr x2": [("Black Garlic 100gr", 2)],
    "Black Garlic 100gr x3": [("Black Garlic 100gr", 3)],
    "Black Garlic 100gr x4": [("Black Garlic 100gr", 4)],
    "Black Garlic 100gr x5": [("Black Garlic 100gr", 5)],
    "Black Garlic 220gr x2": [("Black Garlic 220gr", 2)],
    "Black Garlic 220gr x3": [("Black Garlic 220gr", 3)],
    "Black Garlic 220gr x4": [("Black Garlic 220gr", 4)],
    "Black Garlic 220gr x5": [("Black Garlic 220gr", 5)],
    "Black Garlic 500gr x2": [("Black Garlic 500gr", 2)],
    "Black Garlic 500gr x3": [("Black Garlic 500gr", 3)],
    "Black Garlic 500gr x4": [("Black Garlic 500gr", 4)],
    "Black Garlic 500gr x5": [("Black Garlic 500gr", 5)],
    "BG 3in1 (100gr+220gr+500gr)": [
        ("Black Garlic 100gr", 1),
        ("Black Garlic 220gr", 1),
        ("Black Garlic 500gr", 1),
    ],
    "Black Garlic 220gr x2 + Box Hampers Imlek":   [("Black Garlic 220gr", 2), ("Box Hampers", 1)],
    "Black Garlic 220gr x2 + Box Hampers Lebaran": [("Black Garlic 220gr", 2), ("Box Hampers", 1)],
    "Black Garlic 220gr x2 + Box Hampers Natal":   [("Black Garlic 220gr", 2), ("Box Hampers", 1)],
    "Black Garlic 220gr x2 + Box Hampers HSD":     [("Black Garlic 220gr", 2), ("Box Hampers", 1)],
}

PRODUK_TUNGGAL = set(URUTAN_PRODUK)


def expand_to_komponen(nama_produk, qty):
    if not nama_produk:
        return []
    match_x = re.search(r'\s*[xX](\d+)\s*$', nama_produk)
    if match_x:
        multiplier = int(match_x.group(1))
        base_name = nama_produk[:match_x.start()].strip()
        if base_name in PAKET_KOMPONEN:
            return [(k, q * multiplier * qty) for k, q in PAKET_KOMPONEN[base_name]]
        else:
            return [(base_name, multiplier * qty)]
    if nama_produk in PAKET_KOMPONEN:
        return [(k, q * qty) for k, q in PAKET_KOMPONEN[nama_produk]]
    return [(nama_produk, qty)]


# =========================
# PREPARE DATA
# =========================

def prepare(orders):
    rekap       = defaultdict(lambda: defaultdict(lambda: defaultdict(int)))
    rekap_paket = defaultdict(lambda: defaultdict(lambda: defaultdict(int)))
    orders_bersih = []
    orders_cek    = []

    for order in orders:
        akun        = order.get('akun', 'HSD')
        shift       = order.get('shift', 'Pagi')
        nama_produk = order.get('nama_produk', '')
        qty         = int(order.get('qty', 1) or 1)
        perlu_cek   = order.get('perlu_cek', False)

        if perlu_cek or not nama_produk or nama_produk.strip() == '(cek manual)':
            orders_cek.append(order)
            continue

        orders_bersih.append(order)

        if nama_produk in PAKET_KOMPONEN:
            rekap_paket[akun][shift][nama_produk] += qty
        match_x = re.search(r'\s*[xX](\d+)\s*$', nama_produk)
        if match_x:
            base = nama_produk[:match_x.start()].strip()
            if base in PAKET_KOMPONEN:
                rekap_paket[akun][shift][nama_produk] += qty

        for nama_komponen, qty_komponen in expand_to_komponen(nama_produk, qty):
            if nama_komponen:
                rekap[akun][shift][nama_komponen] += qty_komponen

    return rekap, rekap_paket, orders_bersih, orders_cek


# =========================
# SHEET 1 — KEBUTUHAN BARANG GUDANG
# =========================

def write_sheet1(ws, rekap, rekap_paket):
    ws.title = "Kebutuhan Barang Gudang"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions['A'].width = 35
    ws.column_dimensions['B'].width = 15
    ws.column_dimensions['C'].width = 15
    ws.column_dimensions['D'].width = 20

    row = 1

    # Header utama
    ws.merge_cells(f"A{row}:D{row}")
    c = ws.cell(row=row, column=1, value="KEBUTUHAN BARANG GUDANG")
    style_cell(c, bold=True, size=14, fg="1F3864", font_color="FFFFFF", align="center")
    row += 1

    ws.merge_cells(f"A{row}:D{row}")
    c = ws.cell(row=row, column=1,
                value=f"Generated: {datetime.now().strftime('%d %B %Y %H:%M WIB')}")
    style_cell(c, size=10, fg="D9E1F2", align="center")
    row += 1

    # ── KETERANGAN OFFLINE ──────────────────────────────────────────────
    ws.merge_cells(f"A{row}:D{row}")
    c = ws.cell(row=row, column=1,
                value="⚠️  DATA DI BAWAH HANYA DARI PDF MARKETPLACE")
    style_cell(c, bold=True, size=10, fg="FFC000", font_color="000000", align="center")
    row += 1

    ws.merge_cells(f"A{row}:D{row}")
    c = ws.cell(row=row, column=1,
                value="Order OFFLINE tidak masuk sistem — harap tambahkan / isi manual ke rekap gudang")
    style_cell(c, size=10, fg="FFF2CC", font_color="7F6000", align="center")
    row += 2
    # ────────────────────────────────────────────────────────────────────

    akun_list = sorted(rekap.keys())
    if not akun_list:
        ws.cell(row=row, column=1, value="Tidak ada data order valid.")
        return

    shift_list = ["Pagi", "Siang", "Sore"]

    for akun in akun_list:
        ws.merge_cells(f"A{row}:D{row}")
        c = ws.cell(row=row, column=1, value=f"AKUN: {akun}")
        style_cell(c, bold=True, size=12, fg="2E75B6", font_color="FFFFFF", align="center")
        row += 1

        for shift in shift_list:
            if shift not in rekap[akun]:
                continue

            ws.merge_cells(f"A{row}:D{row}")
            c = ws.cell(row=row, column=1, value=f"  Shift {shift}")
            style_cell(c, bold=True, size=11, fg="BDD7EE", align="left")
            row += 1

            headers = ["Nama Produk", "Qty Siapkan", "Satuan", "Keterangan"]
            for col_idx, h in enumerate(headers, 1):
                c = ws.cell(row=row, column=col_idx, value=h)
                style_cell(c, bold=True, size=10, fg="9DC3E6", align="center")
            row += 1

            data_shift  = rekap[akun][shift]
            paket_shift = rekap_paket[akun].get(shift, {})

            # Paket dulu
            for nama_paket, qty_paket in paket_shift.items():
                c_nama = ws.cell(row=row, column=1, value=f"📦 {nama_paket}")
                c_qty  = ws.cell(row=row, column=2, value=qty_paket)
                c_sat  = ws.cell(row=row, column=3, value="paket")
                c_ket  = ws.cell(row=row, column=4, value="")
                style_cell(c_nama, bold=True, size=10, fg="FFF2CC")
                style_cell(c_qty,  bold=True, size=10, fg="FFF2CC", align="center")
                style_cell(c_sat,  size=10,   fg="FFF2CC", align="center")
                style_cell(c_ket,  size=10,   fg="FFF2CC")
                row += 1

                match_x  = re.search(r'\s*[xX](\d+)\s*$', nama_paket)
                base_name = nama_paket[:match_x.start()].strip() if match_x else nama_paket
                multiplier = int(match_x.group(1)) if match_x else 1

                if base_name in PAKET_KOMPONEN:
                    for komponen, qty_per_unit in PAKET_KOMPONEN[base_name]:
                        qty_total_komponen = qty_per_unit * multiplier * qty_paket
                        c_k = ws.cell(row=row, column=1, value=f"     → {komponen}")
                        c_q = ws.cell(row=row, column=2, value=qty_total_komponen)
                        c_s = ws.cell(row=row, column=3, value="botol/pcs")
                        c_e = ws.cell(row=row, column=4, value=f"isi {qty_per_unit * multiplier} per paket")
                        style_cell(c_k, size=10, fg="FFFACD")
                        style_cell(c_q, size=10, fg="FFFACD", align="center")
                        style_cell(c_s, size=10, fg="FFFACD", align="center")
                        style_cell(c_e, size=10, fg="FFFACD")
                        row += 1

            # Produk tunggal
            for nama_produk in URUTAN_PRODUK:
                if nama_produk not in data_shift:
                    continue
                qty_total = data_shift[nama_produk]
                c_nama = ws.cell(row=row, column=1, value=nama_produk)
                c_qty  = ws.cell(row=row, column=2, value=qty_total)
                c_sat  = ws.cell(row=row, column=3, value="botol/pcs")
                c_ket  = ws.cell(row=row, column=4, value="")

                if "100gr" in nama_produk:   bg = "E2EFDA"
                elif "220gr" in nama_produk: bg = "DDEBF7"
                elif "500gr" in nama_produk: bg = "FCE4D6"
                elif "Drink" in nama_produk: bg = "EAD1DC"
                elif "Madu"  in nama_produk: bg = "D9D2E9"
                else:                        bg = "F2F2F2"

                style_cell(c_nama, size=10,   fg=bg)
                style_cell(c_qty,  bold=True, size=11, fg=bg, align="center")
                style_cell(c_sat,  size=10,   fg=bg, align="center")
                style_cell(c_ket,  size=10,   fg=bg)
                row += 1

            # Produk fallback (tidak di URUTAN_PRODUK)
            for nama_produk, qty_total in data_shift.items():
                if nama_produk in URUTAN_PRODUK:
                    continue
                c_nama = ws.cell(row=row, column=1, value=nama_produk)
                c_qty  = ws.cell(row=row, column=2, value=qty_total)
                c_sat  = ws.cell(row=row, column=3, value="botol/pcs")
                c_ket  = ws.cell(row=row, column=4, value="⚠️ cek mapping")
                style_cell(c_nama, size=10, fg="FFE699")
                style_cell(c_qty,  bold=True, size=10, fg="FFE699", align="center")
                style_cell(c_sat,  size=10, fg="FFE699", align="center")
                style_cell(c_ket,  size=10, fg="FFE699")
                row += 1

            # Total shift
            total_shift = sum(data_shift.values())
            c_t  = ws.cell(row=row, column=1, value=f"  TOTAL SHIFT {shift.upper()}")
            c_tq = ws.cell(row=row, column=2, value=total_shift)
            style_cell(c_t,  bold=True, size=10, fg="70AD47", font_color="FFFFFF")
            style_cell(c_tq, bold=True, size=11, fg="70AD47", font_color="FFFFFF", align="center")
            for col in [3, 4]:
                style_cell(ws.cell(row=row, column=col), fg="70AD47")
            row += 2

        # Total keseluruhan per akun
        total_akun = defaultdict(int)
        for shift in shift_list:
            for nama, qty in rekap[akun].get(shift, {}).items():
                total_akun[nama] += qty

        ws.merge_cells(f"A{row}:D{row}")
        c_ta = ws.cell(row=row, column=1, value=f"TOTAL KESELURUHAN — {akun}")
        style_cell(c_ta, bold=True, size=11, fg="1F3864", font_color="FFFFFF", align="center")
        row += 1

        for nama_produk in URUTAN_PRODUK:
            if nama_produk not in total_akun:
                continue
            c_n = ws.cell(row=row, column=1, value=nama_produk)
            c_q = ws.cell(row=row, column=2, value=total_akun[nama_produk])
            style_cell(c_n, bold=True, size=10, fg="D6DCE4")
            style_cell(c_q, bold=True, size=11, fg="D6DCE4", align="center")
            for col in [3, 4]:
                style_cell(ws.cell(row=row, column=col), fg="D6DCE4")
            row += 1

        row += 2

    # ── KETERANGAN OFFLINE di bagian bawah (pengingat kedua) ────────────
    ws.merge_cells(f"A{row}:D{row}")
    c = ws.cell(row=row, column=1, value="📋  CATATAN PENTING")
    style_cell(c, bold=True, size=11, fg="FFC000", font_color="000000", align="center")
    row += 1

    catatan = [
        "• Data di atas HANYA dari PDF marketplace (Shopee, TikTok, Lazada, dll).",
        "• Order OFFLINE (dari Mbak Fitri via WA) TIDAK masuk di sini.",
        "• Tambahkan jumlah offline secara MANUAL ke rekap gudang setelah melihat file ini.",
    ]
    for baris in catatan:
        ws.merge_cells(f"A{row}:D{row}")
        c = ws.cell(row=row, column=1, value=baris)
        style_cell(c, size=10, fg="FFF2CC", font_color="7F6000", align="left")
        row += 1
    # ────────────────────────────────────────────────────────────────────


# =========================
# SHEET 2 — VARIASI ORDER PER RESI
# =========================

def write_sheet2(ws, orders_bersih):
    ws.title = "Variasi Order per Resi"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions['A'].width = 40
    ws.column_dimensions['B'].width = 15
    ws.column_dimensions['C'].width = 20
    ws.column_dimensions['D'].width = 25

    row = 1

    ws.merge_cells(f"A{row}:D{row}")
    c = ws.cell(row=row, column=1, value="VARIASI ORDER PER RESI")
    style_cell(c, bold=True, size=13, fg="1F3864", font_color="FFFFFF", align="center")
    row += 1

    ws.merge_cells(f"A{row}:D{row}")
    c = ws.cell(row=row, column=1,
                value=f"Generated: {datetime.now().strftime('%d %B %Y %H:%M WIB')}")
    style_cell(c, size=10, fg="D9E1F2", align="center")
    row += 2

    variasi    = defaultdict(list)
    resi_produk = defaultdict(list)

    for order in orders_bersih:
        resi = order.get('resi', '') or order.get('no_resi', '')
        nama = order.get('nama_produk', '')
        qty  = int(order.get('qty', 1) or 1)
        if resi and nama:
            resi_produk[resi].append((nama, qty))

    for resi, produk_list in resi_produk.items():
        key = tuple(sorted([(n, q) for n, q in produk_list]))
        variasi[key].append(resi)

    variasi_sorted = sorted(variasi.items(), key=lambda x: -len(x[1]))

    headers = ["Kombinasi Produk yang Dibeli", "Jumlah Resi", "% dari Total", "Contoh Resi"]
    for col_idx, h in enumerate(headers, 1):
        c = ws.cell(row=row, column=col_idx, value=h)
        style_cell(c, bold=True, size=10, fg="2E75B6", font_color="FFFFFF", align="center")
    row += 1

    total_resi = len(resi_produk)
    colors = ["DDEEFF", "EEF5FF"]

    for i, (key, resi_list) in enumerate(variasi_sorted):
        bg = colors[i % 2]
        if len(key) == 1:
            nama, qty = key[0]
            kombinasi = f"{nama} × {qty}"
        else:
            kombinasi = " | ".join(f"{nama} × {qty}" for nama, qty in key)

        jumlah = len(resi_list)
        pct    = f"{jumlah/total_resi*100:.1f}%" if total_resi > 0 else "0%"
        contoh = resi_list[0] if resi_list else ""

        c1 = ws.cell(row=row, column=1, value=kombinasi)
        c2 = ws.cell(row=row, column=2, value=jumlah)
        c3 = ws.cell(row=row, column=3, value=pct)
        c4 = ws.cell(row=row, column=4, value=contoh)
        style_cell(c1, size=10, fg=bg)
        style_cell(c2, bold=True, size=10, fg=bg, align="center")
        style_cell(c3, size=10, fg=bg, align="center")
        style_cell(c4, size=10, fg=bg)
        row += 1

    row += 1
    c_t = ws.cell(row=row, column=1, value="TOTAL RESI")
    c_v = ws.cell(row=row, column=2, value=total_resi)
    style_cell(c_t, bold=True, size=10, fg="1F3864", font_color="FFFFFF")
    style_cell(c_v, bold=True, size=11, fg="1F3864", font_color="FFFFFF", align="center")
    for col in [3, 4]:
        style_cell(ws.cell(row=row, column=col), fg="1F3864", font_color="FFFFFF")


# =========================
# SHEET PERLU TINDAKAN
# =========================

def write_sheet_perlu_cek(ws, orders_cek):
    ws.title = "Perlu Tindakan"
    ws.sheet_view.showGridLines = False

    cols   = ["Resi", "Platform", "SKU Raw", "Nama Produk", "Qty", "Alasan", "Akun", "Shift"]
    widths = [25, 15, 35, 30, 8, 30, 10, 10]

    for i, (col, w) in enumerate(zip(cols, widths), 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    row = 1
    ws.merge_cells(f"A{row}:H{row}")
    c = ws.cell(row=row, column=1, value="ORDER PERLU TINDAKAN MANUAL")
    style_cell(c, bold=True, size=12, fg="C00000", font_color="FFFFFF", align="center")
    row += 1

    ws.merge_cells(f"A{row}:H{row}")
    c = ws.cell(row=row, column=1,
                value=f"Total: {len(orders_cek)} order | {datetime.now().strftime('%d %B %Y %H:%M WIB')}")
    style_cell(c, size=10, fg="FCE4D6", align="center")
    row += 1

    for col_idx, h in enumerate(cols, 1):
        c = ws.cell(row=row, column=col_idx, value=h)
        style_cell(c, bold=True, size=10, fg="C00000", font_color="FFFFFF", align="center")
    row += 1

    for i, order in enumerate(orders_cek):
        bg = "FFF2CC" if i % 2 == 0 else "FFFFFF"
        values = [
            order.get('resi', '') or order.get('no_resi', ''),
            order.get('platform', ''),
            order.get('sku_raw', '') or order.get('sku', ''),
            order.get('nama_produk', ''),
            order.get('qty', ''),
            order.get('alasan_cek', 'SKU tidak dikenali'),
            order.get('akun', ''),
            order.get('shift', ''),
        ]
        for col_idx, val in enumerate(values, 1):
            c = ws.cell(row=row, column=col_idx, value=val)
            style_cell(c, size=10, fg=bg)
        row += 1


# =========================
# MAIN WRITE FUNCTION
# =========================

def write_excel(orders, output_path):
    wb = Workbook()
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]

    rekap, rekap_paket, orders_bersih, orders_cek = prepare(orders)

    ws1 = wb.create_sheet("Kebutuhan Barang Gudang")
    write_sheet1(ws1, rekap, rekap_paket)

    ws2 = wb.create_sheet("Variasi Order per Resi")
    write_sheet2(ws2, orders_bersih)

    if orders_cek:
        ws3 = wb.create_sheet("Perlu Tindakan")
        write_sheet_perlu_cek(ws3, orders_cek)

    wb.save(output_path)
    return {
        "total_order":  len(orders),
        "order_valid":  len(orders_bersih),
        "order_cek":    len(orders_cek),
        "output_path":  output_path,
    }


# =========================
# TEST / DEBUG
# =========================

if __name__ == "__main__":
    test_orders = [
        {"akun": "HSD", "shift": "Pagi", "resi": "TK001", "platform": "TikTok",
         "nama_produk": "Black Garlic 100gr", "qty": 1, "perlu_cek": False},
        {"akun": "HSD", "shift": "Pagi", "resi": "TK002", "platform": "TikTok",
         "nama_produk": "Black Garlic 220gr", "qty": 2, "perlu_cek": False},
        {"akun": "HSD", "shift": "Pagi", "resi": "TK003", "platform": "Shopee",
         "nama_produk": "Black Garlic 500gr x2", "qty": 1, "perlu_cek": False},
        {"akun": "HSD", "shift": "Pagi", "resi": "TK004", "platform": "TikTok",
         "nama_produk": "Paket 3in1", "qty": 2, "perlu_cek": False},
        {"akun": "HSD", "shift": "Pagi", "resi": "TK005", "platform": "Shopee",
         "nama_produk": "Hampers Lebaran", "qty": 1, "perlu_cek": False},
        {"akun": "HSD", "shift": "Siang", "resi": "SP001", "platform": "Shopee",
         "nama_produk": "Black Garlic 100gr", "qty": 3, "perlu_cek": False},
        {"akun": "HSD", "shift": "Siang", "resi": "SP002", "platform": "Shopee",
         "nama_produk": "BG Drink Original x2", "qty": 1, "perlu_cek": False},
        {"akun": "HSS", "shift": "Pagi", "resi": "LZ001", "platform": "Lazada",
         "nama_produk": "Black Garlic 220gr", "qty": 1, "perlu_cek": False},
        {"akun": "HSS", "shift": "Pagi", "resi": "LZ002", "platform": "Lazada",
         "nama_produk": "BG Drink Mix 7 Botol", "qty": 1, "perlu_cek": False},
        {"akun": "HSD", "shift": "Pagi", "resi": "XX001", "platform": "Unknown",
         "nama_produk": "(cek manual)", "qty": 1, "perlu_cek": True,
         "sku_raw": "BG-UNKNOWN-SKU", "alasan_cek": "SKU tidak dikenali"},
    ]

    out = "/tmp/test_excel_writer.xlsx"
    result = write_excel(test_orders, out)
    print(f"✅ Excel berhasil: {out}")
    print(f"   Total: {result['total_order']} | Valid: {result['order_valid']} | Cek: {result['order_cek']}")
