from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from collections import defaultdict
from datetime import datetime


# =========================
# BASIC STYLE
# =========================

HDR_FILL = PatternFill("solid", fgColor="1B1F2E")
HDR_FONT = Font(bold=True, color="FFFFFF", size=10)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)

BORDER = Border(
    left=Side(style="thin", color="CCCCCC"),
    right=Side(style="thin", color="CCCCCC"),
    top=Side(style="thin", color="CCCCCC"),
    bottom=Side(style="thin", color="CCCCCC")
)

AMBER = PatternFill("solid", fgColor="F5A623")
GREEN = PatternFill("solid", fgColor="EAF3DE")
BLUE = PatternFill("solid", fgColor="D6EAFF")
ORANGE = PatternFill("solid", fgColor="FFE4CC")
WHITE = PatternFill("solid", fgColor="FFFFFF")
GREY = PatternFill("solid", fgColor="F5F5F5")
YELLOW = PatternFill("solid", fgColor="FFF9C4")
RED_LIGHT = PatternFill("solid", fgColor="FFEBEE")


def safe(r, key, default=""):
    try:
        value = r.get(key, default)
        if value is None:
            return default
        return value
    except Exception:
        return default


def clean(v):
    if v is None:
        return ""
    return str(v).rstrip(",").strip()


def to_int(v):
    try:
        return int(v)
    except Exception:
        return 0


def set_header(ws, headers):
    for col, item in enumerate(headers, 1):
        title, width = item

        cell = ws.cell(row=1, column=col, value=title)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.alignment = CENTER
        cell.border = BORDER

        ws.column_dimensions[get_column_letter(col)].width = width

    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 28


def write_cells(ws, row, values, fill=WHITE):
    for col, value in enumerate(values, 1):
        cell = ws.cell(row=row, column=col, value=value)
        cell.fill = fill
        cell.border = BORDER
        cell.alignment = CENTER if col in [1, 2, 3, 7, 8] else LEFT
        cell.font = Font(size=10)


def add_total_row(ws, row, values):
    for col, value in enumerate(values, 1):
        cell = ws.cell(row=row, column=col, value=value)
        cell.fill = AMBER
        cell.font = Font(bold=True, size=11)
        cell.alignment = CENTER
        cell.border = BORDER


# =========================
# MAIN WRITER
# =========================

def write_excel_multi(rows, output_path):
    wb = Workbook()

    # Safety: pastikan rows list
    if rows is None:
        rows = []

    # Sort
    rows_sorted = sorted(
        rows,
        key=lambda r: (
            str(safe(r, "akun", "")),
            str(safe(r, "waktu", "")),
            str(safe(r, "courier", "")),
            str(safe(r, "no_resi", ""))
        )
    )

    # =========================
    # SHEET 1: STOK FISIK HARI INI
    # =========================

    ws0 = wb.active
    ws0.title = "STOK FISIK HARI INI"
    ws0.sheet_tab_color = "F5A623"

    ws0.merge_cells("A1:D1")
    title = ws0.cell(
        row=1,
        column=1,
        value=f"STOK FISIK YANG HARUS DISIAPKAN - {datetime.now().strftime('%d/%m/%Y')}"
    )
    title.fill = HDR_FILL
    title.font = Font(bold=True, color="FFFFFF", size=14)
    title.alignment = CENTER
    title.border = BORDER

    ws0.column_dimensions["A"].width = 40
    ws0.column_dimensions["B"].width = 18
    ws0.column_dimensions["C"].width = 18
    ws0.column_dimensions["D"].width = 28

    stok_produk = defaultdict(int)
    resi_map = defaultdict(list)

    for r in rows_sorted:
        nama = clean(safe(r, "nama_produk")) or clean(safe(r, "sku")) or "(tanpa nama)"
        qty = to_int(safe(r, "qty", 0))

        stok_produk[nama] += qty
        resi_map[clean(safe(r, "no_resi"))].append(r)

    total_produk = sum(stok_produk.values())
    total_resi = len(resi_map)
    kardus_campur = sum(1 for items in resi_map.values() if len(items) > 1)
    kardus_tunggal = total_resi - kardus_campur

    metrics = [
        ("TOTAL PRODUK", total_produk),
        ("TOTAL RESI", total_resi),
        ("KARDUS CAMPUR", kardus_campur),
        ("KARDUS TUNGGAL", kardus_tunggal),
    ]

    for col, item in enumerate(metrics, 1):
        label, value = item

        c1 = ws0.cell(row=3, column=col, value=label)
        c1.fill = PatternFill("solid", fgColor="F0F3FA")
        c1.font = Font(bold=True, size=10)
        c1.alignment = CENTER
        c1.border = BORDER

        c2 = ws0.cell(row=4, column=col, value=value)
        c2.fill = WHITE
        c2.font = Font(bold=True, size=22)
        c2.alignment = CENTER
        c2.border = BORDER

    ws0.row_dimensions[4].height = 40

    headers = ["Nama Produk", "Qty Disiapkan", "% Total", "Keterangan"]
    for col, h in enumerate(headers, 1):
        cell = ws0.cell(row=6, column=col, value=h)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.alignment = CENTER
        cell.border = BORDER

    row = 7
    for idx, item in enumerate(sorted(stok_produk.items(), key=lambda x: -x[1])):
        nama, qty = item
        pct = f"{qty / total_produk * 100:.1f}%" if total_produk else "0%"
        fill = GREEN if idx % 2 == 0 else WHITE

        values = [nama, qty, pct, "Ambil dari gudang"]
        for col, value in enumerate(values, 1):
            cell = ws0.cell(row=row, column=col, value=value)
            cell.fill = fill
            cell.border = BORDER
            cell.alignment = LEFT if col in [1, 4] else CENTER
            cell.font = Font(bold=True if col == 2 else False, size=11)

        row += 1

    add_total_row(ws0, row, ["TOTAL", total_produk, "100%", ""])
    ws0.freeze_panes = "A7"

    # =========================
    # SHEET 2: SEMUA RESI
    # =========================

    ws1 = wb.create_sheet("Semua Resi")

    set_header(ws1, [
        ("No", 6),
        ("Akun", 10),
        ("Waktu", 12),
        ("Kurir", 16),
        ("Layanan", 12),
        ("No. Resi", 24),
        ("Nama Produk", 38),
        ("Seller SKU", 26),
        ("Variasi", 18),
        ("Qty", 8),
        ("Nama Penerima", 24),
        ("Alamat", 46),
        ("Platform", 16),
    ])

    prev_resi = ""
    flip = False

    for idx, r in enumerate(rows_sorted, 1):
        no_resi = clean(safe(r, "no_resi"))

        if no_resi != prev_resi:
            flip = not flip
            prev_resi = no_resi

        fill = BLUE if flip else WHITE

        qty = to_int(safe(r, "qty", 0))

        values = [
            idx,
            safe(r, "akun", ""),
            safe(r, "waktu", ""),
            safe(r, "courier", ""),
            safe(r, "layanan", ""),
            no_resi,
            clean(safe(r, "nama_produk")),
            clean(safe(r, "sku")),
            clean(safe(r, "variasi")) or "-",
            qty,
            clean(safe(r, "nama_penerima")),
            clean(safe(r, "alamat")),
            clean(safe(r, "platform")),
        ]

        for col, value in enumerate(values, 1):
            cell = ws1.cell(row=idx + 1, column=col, value=value)
            cell.fill = fill
            cell.border = BORDER
            cell.font = Font(size=9)
            cell.alignment = CENTER if col in [1, 2, 3, 5, 9, 10] else LEFT

        if qty > 1:
            ws1.cell(row=idx + 1, column=10).font = Font(bold=True, color="CC0000", size=11)

    ws1.auto_filter.ref = f"A1:M1"

    # =========================
    # SHEET 3: REKAP KURIR
    # =========================

    ws2 = wb.create_sheet("Rekap Kurir")

    set_header(ws2, [
        ("Kurir", 20),
        ("Total Resi", 14),
        ("Total Qty", 14),
        ("Akun", 14),
        ("Waktu", 18),
    ])

    by_courier = defaultdict(lambda: {
        "resi": set(),
        "qty": 0,
        "akun": set(),
        "waktu": set(),
    })

    for r in rows_sorted:
        kurir = clean(safe(r, "courier")) or "Unknown"

        by_courier[kurir]["resi"].add(clean(safe(r, "no_resi")))
        by_courier[kurir]["qty"] += to_int(safe(r, "qty", 0))
        by_courier[kurir]["akun"].add(clean(safe(r, "akun")))
        by_courier[kurir]["waktu"].add(clean(safe(r, "waktu")))

    row = 2
    for kurir, data in sorted(by_courier.items()):
        fill = BLUE if row % 2 == 0 else WHITE

        values = [
            kurir,
            len(data["resi"]),
            data["qty"],
            ", ".join(sorted(data["akun"])),
            ", ".join(sorted(data["waktu"])),
        ]

        for col, value in enumerate(values, 1):
            cell = ws2.cell(row=row, column=col, value=value)
            cell.fill = fill
            cell.border = BORDER
            cell.alignment = CENTER if col in [2, 3] else LEFT
            cell.font = Font(bold=True if col in [2, 3] else False, size=10)

        row += 1

    add_total_row(ws2, row, [
        "TOTAL",
        len(set(clean(safe(r, "no_resi")) for r in rows_sorted)),
        sum(to_int(safe(r, "qty", 0)) for r in rows_sorted),
        "",
        "",
    ])

    # =========================
    # SHEET 4: REKAP SKU
    # =========================

    ws3 = wb.create_sheet("Rekap SKU")

    set_header(ws3, [
        ("Seller SKU", 30),
        ("Nama Produk", 40),
        ("Variasi", 18),
        ("Total Resi", 14),
        ("Total Qty", 14),
    ])

    by_sku = defaultdict(lambda: {
        "nama": "",
        "variasi": "",
        "resi": set(),
        "qty": 0,
    })

    for r in rows_sorted:
        sku = clean(safe(r, "sku")) or "(tanpa SKU)"

        if not by_sku[sku]["nama"]:
            by_sku[sku]["nama"] = clean(safe(r, "nama_produk"))

        if not by_sku[sku]["variasi"]:
            by_sku[sku]["variasi"] = clean(safe(r, "variasi"))

        by_sku[sku]["resi"].add(clean(safe(r, "no_resi")))
        by_sku[sku]["qty"] += to_int(safe(r, "qty", 0))

    row = 2
    for sku, data in sorted(by_sku.items(), key=lambda x: -x[1]["qty"]):
        fill = GREEN if row % 2 == 0 else WHITE

        values = [
            sku,
            data["nama"],
            data["variasi"] or "-",
            len(data["resi"]),
            data["qty"],
        ]

        for col, value in enumerate(values, 1):
            cell = ws3.cell(row=row, column=col, value=value)
            cell.fill = fill
            cell.border = BORDER
            cell.alignment = CENTER if col in [3, 4, 5] else LEFT
            cell.font = Font(bold=True if col == 5 else False, size=10)

        row += 1

    add_total_row(ws3, row, [
        "TOTAL",
        "",
        "",
        len(set(clean(safe(r, "no_resi")) for r in rows_sorted)),
        sum(to_int(safe(r, "qty", 0)) for r in rows_sorted),
    ])

    # =========================
    # SHEET 5: PACKING LIST GUDANG
    # =========================

    ws4 = wb.create_sheet("Packing List Gudang")

    set_header(ws4, [
        ("No Urut", 8),
        ("No. Resi", 24),
        ("Kurir", 16),
        ("Nama Produk", 40),
        ("Seller SKU", 26),
        ("Variasi", 18),
        ("Qty", 8),
        ("Cek", 10),
    ])

    row = 2
    no_urut = 1

    for no_resi, items in resi_map.items():
        is_multi = len(items) > 1
        fill = YELLOW if is_multi else (GREY if row % 2 == 0 else WHITE)

        for r in items:
            qty = to_int(safe(r, "qty", 0))

            values = [
                no_urut,
                clean(safe(r, "no_resi")),
                clean(safe(r, "courier")),
                clean(safe(r, "nama_produk")),
                clean(safe(r, "sku")),
                clean(safe(r, "variasi")) or "-",
                qty,
                "",
            ]

            for col, value in enumerate(values, 1):
                cell = ws4.cell(row=row, column=col, value=value)
                cell.fill = fill
                cell.border = BORDER
                cell.alignment = CENTER if col in [1, 6, 7, 8] else LEFT
                cell.font = Font(bold=True if col == 7 and qty > 1 else False, size=10)

                if col == 7 and qty > 1:
                    cell.font = Font(bold=True, color="CC0000", size=11)

            row += 1
            no_urut += 1

    # =========================
    # SHEET 6: RINGKASAN ORDER
    # =========================

    ws5 = wb.create_sheet("Ringkasan Order")

    set_header(ws5, [
        ("Akun", 10),
        ("Waktu", 12),
        ("Kombinasi Produk", 50),
        ("Jumlah Resi", 14),
        ("Total Qty", 14),
        ("Contoh Resi", 24),
    ])

    combo_map = defaultdict(lambda: {
        "resi": [],
        "qty": 0,
        "akun": "",
        "waktu": "",
    })

    for no_resi, items in resi_map.items():
        nama_qty = defaultdict(int)

        for r in items:
            nama = clean(safe(r, "nama_produk")) or clean(safe(r, "sku")) or "(tanpa nama)"
            nama_qty[nama] += to_int(safe(r, "qty", 0))

        parts = []
        for nama in sorted(nama_qty.keys()):
            parts.append(f"{nama} x{nama_qty[nama]}")

        combo = " | ".join(parts)

        akun = clean(safe(items[0], "akun"))
        waktu = clean(safe(items[0], "waktu"))
        key = (akun, waktu, combo)

        combo_map[key]["resi"].append(no_resi)
        combo_map[key]["qty"] += sum(to_int(safe(r, "qty", 0)) for r in items)
        combo_map[key]["akun"] = akun
        combo_map[key]["waktu"] = waktu

    row = 2
    for key, data in sorted(combo_map.items(), key=lambda x: -len(x[1]["resi"])):
        akun, waktu, combo = key
        is_multi = "|" in combo
        fill = RED_LIGHT if is_multi else GREEN

        values = [
            akun,
            waktu,
            combo,
            len(data["resi"]),
            data["qty"],
            data["resi"][0] if data["resi"] else "",
        ]

        for col, value in enumerate(values, 1):
            cell = ws5.cell(row=row, column=col, value=value)
            cell.fill = fill
            cell.border = BORDER
            cell.alignment = LEFT if col in [3, 6] else CENTER
            cell.font = Font(bold=True if col in [4, 5] else False, size=10)

        row += 1

    # =========================
    # SHEET 7: KODE PENGAMBILAN GOJEK
    # =========================

    ws6 = wb.create_sheet("Kode Pengambilan Gojek")

    set_header(ws6, [
        ("No. Resi", 24),
        ("Kurir", 16),
        ("Layanan", 14),
        ("Kode Pengambilan", 24),
        ("Produk & Qty", 40),
    ])

    row = 2
    found = False

    for no_resi, items in resi_map.items():
        kode = clean(safe(items[0], "kode_pengambilan"))

        if not kode:
            continue

        found = True
        produk = " | ".join(
            f"{clean(safe(r, 'sku'))} x{to_int(safe(r, 'qty', 0))}"
            for r in items
        )

        values = [
            no_resi,
            clean(safe(items[0], "courier")),
            clean(safe(items[0], "layanan")),
            kode,
            produk,
        ]

        for col, value in enumerate(values, 1):
            cell = ws6.cell(row=row, column=col, value=value)
            cell.fill = GREEN
            cell.border = BORDER
            cell.alignment = LEFT if col in [1, 4, 5] else CENTER
            cell.font = Font(bold=True if col == 4 else False, size=10)

        row += 1

    if not found:
        ws6.merge_cells("A2:E2")
        cell = ws6.cell(
            row=2,
            column=1,
            value="Tidak ada kode pengambilan di PDF ini."
        )
        cell.alignment = LEFT
        cell.font = Font(italic=True, size=10)

    # =========================
    # SAVE
    # =========================

    wb.save(output_path)
    return output_path


# Biar kalau app lama masih manggil write_excel(), tetap aman
def write_excel(rows, output_path):
    return write_excel_multi(rows, output_path)
