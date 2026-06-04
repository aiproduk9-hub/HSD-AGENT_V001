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
DARK = PatternFill("solid", fgColor="111827")
HEADER = PatternFill("solid", fgColor="1B1F2E")
WHITE = PatternFill("solid", fgColor="FFFFFF")
CREAM = PatternFill("solid", fgColor="FFF7E8")
GREY = PatternFill("solid", fgColor="F3F4F6")
BLUE = PatternFill("solid", fgColor="D6EAFF")
BLUE_DARK = PatternFill("solid", fgColor="2563EB")
ORANGE = PatternFill("solid", fgColor="FFE4CC")
ORANGE_DARK = PatternFill("solid", fgColor="F97316")
GREEN = PatternFill("solid", fgColor="EAF3DE")
GREEN_DARK = PatternFill("solid", fgColor="16A34A")
YELLOW = PatternFill("solid", fgColor="FFF9C4")
RED_LIGHT = PatternFill("solid", fgColor="FFEBEE")
RED_DARK = PatternFill("solid", fgColor="DC2626")
PURPLE = PatternFill("solid", fgColor="EDE9FE")
BOS_YELLOW = PatternFill("solid", fgColor="FFF200")
BOS_GREY = PatternFill("solid", fgColor="D9D9D9")
BOS_GREEN = PatternFill("solid", fgColor="E2F0D9")
BOS_ORANGE = PatternFill("solid", fgColor="F4B183")
BOS_BLUE = PatternFill("solid", fgColor="DDEBF7")

CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
RIGHT = Alignment(horizontal="right", vertical="center", wrap_text=True)

BORDER = Border(
    left=Side(style="thin", color="000000"),
    right=Side(style="thin", color="000000"),
    top=Side(style="thin", color="000000"),
    bottom=Side(style="thin", color="000000"),
)


def font(size=10, bold=True, color="111827"):
    return Font(name=FONT_NAME, size=size, bold=bold, color=color)


FONT_HEADER = font(size=10, bold=True, color="FFFFFF")
FONT_NORMAL = font(size=10, bold=True, color="111827")
FONT_SMALL = font(size=9, bold=True, color="111827")
FONT_TITLE = font(size=16, bold=True, color="FFFFFF")
FONT_BIG = font(size=22, bold=True, color="111827")


# =========================
# BASIC HELPER
# =========================
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
    v = str(v)
    v = re.sub(r"[\x00-\x08\x0B-\x0C\x0E-\x1F]", "", v)
    v = v.replace("\uFFFE", " ")
    v = " ".join(v.split())
    return v.strip().rstrip(",")


def to_int(v):
    try:
        return int(float(str(v).replace(",", "").strip()))
    except Exception:
        return 0


def normalize_token(v):
    v = clean(v).upper()
    v = v.replace("_", "-")
    v = re.sub(r"\s+", "-", v)
    v = re.sub(r"-+", "-", v)
    v = v.strip("-")
    return v


def set_sheet_view(ws):
    ws.sheet_view.showGridLines = False


def set_col_widths(ws, widths):
    for col, width in widths.items():
        ws.column_dimensions[col].width = width


def set_header(ws, headers, row=1):
    for col, item in enumerate(headers, 1):
        title, width = item
        cell = ws.cell(row=row, column=col, value=title)
        cell.fill = HEADER
        cell.font = FONT_HEADER
        cell.alignment = CENTER
        cell.border = BORDER
        ws.column_dimensions[get_column_letter(col)].width = width

    ws.row_dimensions[row].height = 28
    ws.freeze_panes = f"A{row + 1}"
    ws.auto_filter.ref = f"A{row}:{get_column_letter(len(headers))}{row}"


def write_row(ws, row, values, fill=WHITE, center_cols=None):
    center_cols = center_cols or []

    for col, value in enumerate(values, 1):
        cell = ws.cell(row=row, column=col, value=value)
        cell.fill = fill
        cell.border = BORDER
        cell.font = FONT_NORMAL
        cell.alignment = CENTER if col in center_cols else LEFT


def title_bar(ws, cell_range, title, fill=DARK, color="FFFFFF"):
    ws.merge_cells(cell_range)
    cell = ws[cell_range.split(":")[0]]
    cell.value = title
    cell.fill = fill
    cell.font = font(size=16, bold=True, color=color)
    cell.alignment = CENTER
    cell.border = BORDER

    min_cell, max_cell = cell_range.split(":")
    start_col = ws[min_cell].column
    end_col = ws[max_cell].column
    row = ws[min_cell].row

    for col in range(start_col, end_col + 1):
        ws.cell(row=row, column=col).fill = fill
        ws.cell(row=row, column=col).border = BORDER


def add_total_row(ws, row, values):
    for col, value in enumerate(values, 1):
        cell = ws.cell(row=row, column=col, value=value)
        cell.fill = ORANGE_DARK
        cell.font = font(size=11, bold=True, color="FFFFFF")
        cell.alignment = CENTER
        cell.border = BORDER


def box_metric(ws, start_row, start_col, label, value, fill, value_suffix=""):
    ws.merge_cells(start_row=start_row, start_column=start_col, end_row=start_row, end_column=start_col + 1)
    ws.merge_cells(start_row=start_row + 1, start_column=start_col, end_row=start_row + 1, end_column=start_col + 1)

    c1 = ws.cell(row=start_row, column=start_col, value=label)
    c1.fill = fill
    c1.font = font(size=10, bold=True, color="FFFFFF")
    c1.alignment = CENTER
    c1.border = BORDER

    c2 = ws.cell(row=start_row + 1, column=start_col, value=f"{value} {value_suffix}".strip())
    c2.fill = WHITE
    c2.font = font(size=20, bold=True, color="111827")
    c2.alignment = CENTER
    c2.border = BORDER

    for r in [start_row, start_row + 1]:
        for c in [start_col, start_col + 1]:
            ws.cell(row=r, column=c).border = BORDER
            ws.cell(row=r, column=c).alignment = CENTER
            if r == start_row:
                ws.cell(row=r, column=c).fill = fill
            else:
                ws.cell(row=r, column=c).fill = WHITE


# =========================
# FIELD HELPER
# =========================
def get_akun(r):
    akun = clean(safe(r, "akun")) or clean(safe(r, "brand"))
    akun = akun.upper()
    if akun in ["HSD", "HSS"]:
        return akun
    return akun or "-"


def get_waktu(r):
    waktu = clean(safe(r, "waktu")) or clean(safe(r, "shift"))
    return waktu.title() if waktu else "-"


def get_platform(r):
    platform = clean(safe(r, "platform"))
    p = platform.upper()

    if "SHOPEE" in p:
        return "Shopee"
    if "TIKTOK" in p:
        return "TikTok"
    if "TOKOPEDIA" in p:
        return "TikTok"
    if "LAZADA" in p:
        return "Lazada"
    if "BLIBLI" in p:
        return "Blibli"

    return platform or "-"


def get_pembayaran(r):
    raw = " ".join([
        clean(safe(r, "pembayaran")),
        clean(safe(r, "payment")),
        clean(safe(r, "cod")),
        clean(safe(r, "layanan")),
        clean(safe(r, "courier")),
        clean(safe(r, "catatan")),
    ]).upper()

    if "NON-COD" in raw or "NON COD" in raw:
        return "Non-COD"
    if "COD" in raw:
        return "COD"
    if "CASHLESS" in raw:
        return "Non-COD"

    return "Belum Terbaca"


def get_kode_pengambilan(r):
    return clean(safe(r, "kode_pengambilan")) or clean(safe(r, "kode"))


def platform_akun_label(akun, platform):
    akun = clean(akun).upper()
    platform = clean(platform)
    p = platform.upper()

    if "TIKTOK" in p or "TOKOPEDIA" in p:
        if akun and akun != "-":
            return f"TIKTOK {akun}"
        return "TIKTOK"

    if "SHOPEE" in p:
        return "SHOPEE"

    if "LAZADA" in p:
        return "LAZADA"

    if "BLIBLI" in p:
        return "BLIBLI"

    if akun and akun != "-":
        return f"{platform} {akun}".strip()

    return platform or "-"


# =========================
# MASTER SKU HSD
# =========================
def comp(item, qty, ringkasan=True, packing=True):
    return {
        "item": item,
        "qty": int(qty),
        "ringkasan": bool(ringkasan),
        "packing": bool(packing),
    }


def build_master_sku():
    m = {}

    def add(keys, components):
        for k in keys:
            m[normalize_token(k)] = components

    add(["BGH-MULTI-FLORAL-1-BOTOL"], [comp("Madu Multi Floral", 1)])
    add(["BGH-BUNGA-KURMA-1-BOTOL"], [comp("Madu Bunga Kurma", 1)])

    add([
        "BG-DRINK-ORI-1-BOTOL-PROMO",
        "BG-DRINK-PROMO-1-BTL-ORI",
        "BG-PROMO-ORI-1-BOTOL",
        "BG-Drink-Original-FS",
        "BG-Drink-Original",
        "BG-DRINK-ORIGINAL",
        "BLACKGARLIC-DRINK-ORIGINAL",
        "BLACKGARLIC-DRINK-ORIGINAL-FS",
    ], [comp("BG Drink Original", 1)])

    add([
        "BG-DRINK-ORI-2-BOTOL-PROMO",
        "BG-DRINK-PROMO-2-BTL-ORI",
    ], [comp("BG Drink Original", 2)])

    add(["BG-DRINK-PROMO-4-BTL-ORI"], [comp("BG Drink Original", 4)])

    add([
        "BG-Drink-Original-7-Botol",
        "BG-DRINK-ORIGINAL-7-BOTOL",
    ], [comp("BG Drink Original", 7)])

    add([
        "BG-Drink-Peach",
        "BG-Drink-Peach-FS",
        "BG-DRINK-PEACH",
        "BLACKGARLIC-DRINK-PEACH",
        "BLACKGARLIC-DRINK-PEACH-FS",
    ], [comp("BG Drink Peach", 1)])

    add([
        "BG-Drink-Peach-7-Botol",
        "BG-DRINK-PEACH-7-BOTOL",
    ], [comp("BG Drink Peach", 7)])

    add([
        "BG-Drink-Mix-7-Botol",
        "BG-DRINK-MIX-7-BOTOL",
    ], [
        comp("BG Drink Peach", 4),
        comp("BG Drink Original", 3),
    ])

    add(["BOX-HAMPERS-IMLEK"], [comp("Box Hampers Imlek", 1, ringkasan=False)])
    add(["BOX-HAMPERS-HSD"], [comp("Box Hampers HSD", 1, ringkasan=False)])
    add(["BOX-HAMPERS-NATAL"], [comp("Box Hampers Natal", 1, ringkasan=False)])
    add(["BOX-HAMPERS-LEBARAN"], [comp("Box Hampers Lebaran", 1, ringkasan=False)])

    add(["BG-220GR-2-BOTOL-HAMPERS-IMLEK"], [
        comp("Black Garlic 220gr", 2),
        comp("Box Hampers Imlek", 1, ringkasan=False),
    ])

    add(["BG-220GR-2-BOTOL-HAMPERS-LEBARAN"], [
        comp("Black Garlic 220gr", 2),
        comp("Box Hampers Lebaran", 1, ringkasan=False),
    ])

    add(["BG-220GR-2-BOTOL-HAMPERS-NATAL"], [
        comp("Black Garlic 220gr", 2),
        comp("Box Hampers Natal", 1, ringkasan=False),
    ])

    add(["BG-220GR-2-BOTOL-HAMPERS-HSD"], [
        comp("Black Garlic 220gr", 2),
        comp("Box Hampers HSD", 1, ringkasan=False),
    ])

    add([
        "BG-3IN1",
        "Black Garlic Tunggal 3 in 1 + Goodie Bag HSD",
    ], [
        comp("Black Garlic 100gr", 1),
        comp("Black Garlic 220gr", 1),
        comp("Black Garlic 500gr", 1),
        comp("Goodie Bag HSD", 1, ringkasan=False),
    ])

    add([
        "BLACKGARLIC-HONAN-HSD-84G-3PCS",
        "BG-84GR-CLOVER",
        "BG-84GR-1-CLOVER",
    ], [comp("Black Garlic 84gr", 1)])

    add([
        "BG-100GR-1-BOTOL",
        "BG-100GR-1-BOTOL-3CM",
        "BG-100GR-1-BOTOL-BG-SKU",
        "BG-100GR-1-BOTOL-BH",
        "BG-100GR-1 BOTOL-BH-SKU",
        "BG-100GR-SKU",
        "BG-100GR-3CM-SKU",
        "BG-100GR-1 BOTOL-BG-SKU",
        "BG-100GR-1- BOTOL-BG-SKU",
        "BG-100GR-1- BOTOL-BH",
    ], [comp("Black Garlic 100gr", 1)])

    add(["BG-100GR-2-BOTOL", "BG-100GR-2-BOTOL-3CM"], [comp("Black Garlic 100gr", 2)])
    add(["BG-100GR-3-BOTOL", "BG-100GR-3-BOTOL-3CM"], [comp("Black Garlic 100gr", 3)])
    add(["BG-100GR-4-BOTOL", "BG-100GR-4-BOTOL-3CM"], [comp("Black Garlic 100gr", 4)])
    add(["BG-100GR-5-BOTOL", "BG-100GR-5-BOTOL-3CM"], [comp("Black Garlic 100gr", 5)])

    add([
        "BG-220GR-1-BOTOL",
        "BG-220GR-1-BOTOL-BG-SKU",
        "BG-220GR-1-BOTOL-BH",
        "BG-220GR-1 BOTOL-BH-SKU",
        "BG-220GR-SKU",
        "BG-220GR-1-BOTOL-V2",
        "BG-220GR-1- BOTOL-BG-SKU",
        "BG-220GR-1- BOTOL-BH",
        "BG-220GR-1 BOTOL-BG-SKU",
        "BG-220GR-SK U",
    ], [comp("Black Garlic 220gr", 1)])

    add(["BG-220GR-2-BOTOL"], [comp("Black Garlic 220gr", 2)])
    add(["BG-220GR-3-BOTOL"], [comp("Black Garlic 220gr", 3)])
    add(["BG-220GR-4-BOTOL"], [comp("Black Garlic 220gr", 4)])
    add(["BG-220GR-5-BOTOL"], [comp("Black Garlic 220gr", 5)])

    add([
        "BG-500GR-1-BOTOL",
        "BG-500GR-1-BOTOL-BH",
        "BG-500GR-1-BOTOL-BG-SKU",
        "BG-500GR-1 BOTOL-BH-SKU",
        "BG-500GR-SKU",
        "BG-500GR-1- BOTOL",
        "BG-500GR-1 BOTOL",
    ], [comp("Black Garlic 500gr", 1)])

    add(["BG-500GR-2-BOTOL"], [comp("Black Garlic 500gr", 2)])
    add(["BG-500GR-3-BOTOL"], [comp("Black Garlic 500gr", 3)])
    add(["BG-500GR-4-BOTOL"], [comp("Black Garlic 500gr", 4)])
    add(["BG-500GR-5-BOTOL"], [comp("Black Garlic 500gr", 5)])

    return m


MASTER_SKU = build_master_sku()


# =========================
# EXPAND PRODUCT
# =========================
def multiply_components(components, multiplier):
    multiplier = max(0, to_int(multiplier))
    result = []

    for c in components:
        result.append({
            "item": c["item"],
            "qty": c["qty"] * multiplier,
            "ringkasan": c.get("ringkasan", True),
            "packing": c.get("packing", True),
        })

    return result


def expand_product(row):
    sku_raw = clean(safe(row, "sku"))
    nama_raw = clean(safe(row, "nama_produk"))
    nama_asli = clean(safe(row, "nama_produk_asli"))
    variasi = clean(safe(row, "variasi"))
    qty = to_int(safe(row, "qty", 1)) or 1

    candidates = [
        sku_raw,
        nama_raw,
        nama_asli,
        variasi,
        f"{nama_raw} {sku_raw}",
        f"{nama_asli} {sku_raw}",
    ]

    normalized_candidates = [normalize_token(x) for x in candidates if clean(x)]

    for token in normalized_candidates:
        if token in MASTER_SKU:
            return multiply_components(MASTER_SKU[token], qty), "OK", ""

    full_text = normalize_token(" ".join(candidates))
    full_readable = " ".join(candidates).upper()

    if "3IN1" in full_text or "3-IN-1" in full_text or "3 IN 1" in full_readable:
        components = [
            comp("Black Garlic 100gr", 1),
            comp("Black Garlic 220gr", 1),
            comp("Black Garlic 500gr", 1),
            comp("Goodie Bag HSD", 1, ringkasan=False),
        ]
        return multiply_components(components, qty), "OK", ""

    if "DRINK" in full_text and "MIX" in full_text and "7-BOTOL" in full_text:
        components = [comp("BG Drink Peach", 4), comp("BG Drink Original", 3)]
        return multiply_components(components, qty), "OK", ""

    if "DRINK" in full_text and "ORIGINAL" in full_text and "7-BOTOL" in full_text:
        return multiply_components([comp("BG Drink Original", 7)], qty), "OK", ""

    if "DRINK" in full_text and "PEACH" in full_text and "7-BOTOL" in full_text:
        return multiply_components([comp("BG Drink Peach", 7)], qty), "OK", ""

    if "DRINK" in full_text and "PEACH" in full_text:
        return multiply_components([comp("BG Drink Peach", 1)], qty), "OK", ""

    if "DRINK" in full_text and ("ORIGINAL" in full_text or "-ORI" in full_text or "ORI-" in full_text):
        return multiply_components([comp("BG Drink Original", 1)], qty), "OK", ""

    if "MULTI" in full_text and ("FLORAL" in full_text or "FLORA" in full_text):
        return multiply_components([comp("Madu Multi Floral", 1)], qty), "OK", ""

    if "KURMA" in full_text:
        return multiply_components([comp("Madu Bunga Kurma", 1)], qty), "OK", ""

    if "84GR" in full_text or "84G" in full_text:
        return multiply_components([comp("Black Garlic 84gr", 1)], qty), "OK", ""

    m = re.search(r"BG-(100|220|500)GR-(\d+)-BOTOL", full_text)
    if m:
        size = m.group(1)
        jumlah = int(m.group(2))
        return multiply_components([comp(f"Black Garlic {size}gr", jumlah)], qty), "OK", ""

    if "100GR" in full_text or "100-GR" in full_text or "100 GR" in full_readable:
        return multiply_components([comp("Black Garlic 100gr", 1)], qty), "OK", ""

    if "220GR" in full_text or "220-GR" in full_text or "220 GR" in full_readable:
        return multiply_components([comp("Black Garlic 220gr", 1)], qty), "OK", ""

    if "500GR" in full_text or "500-GR" in full_text or "500 GR" in full_readable:
        return multiply_components([comp("Black Garlic 500gr", 1)], qty), "OK", ""

    unknown_name = nama_raw or nama_asli or sku_raw or "(produk belum terbaca)"

    return [
        {
            "item": f"PERLU CEK: {unknown_name}",
            "qty": qty,
            "ringkasan": False,
            "packing": True,
        }
    ], "PERLU TINDAKAN", "SKU/Nama produk belum ada di Master SKU HSD"


def combine_components(components, packing_only=False, ringkasan_only=False):
    total = defaultdict(int)

    for c in components:
        if packing_only and not c.get("packing", True):
            continue
        if ringkasan_only and not c.get("ringkasan", True):
            continue
        total[c["item"]] += to_int(c["qty"])

    parts = []
    for item in sorted(total.keys()):
        qty = total[item]
        if qty:
            parts.append(f"{item} x{qty}")

    return " + ".join(parts)


# =========================
# PREPARE DATA
# =========================
def prepare_data(rows):
    rows = rows or []

    expanded_rows = []
    perlu_tindakan = []

    for idx, r in enumerate(rows, 1):
        components, status, reason = expand_product(r)

        base = dict(r)
        base["_row_no"] = idx
        base["_akun"] = get_akun(r)
        base["_waktu"] = get_waktu(r)
        base["_platform"] = get_platform(r)
        base["_pembayaran"] = get_pembayaran(r)
        base["_components"] = components
        base["_status"] = status
        base["_reason"] = reason

        expanded_rows.append(base)

        if status != "OK":
            perlu_tindakan.append(base)

    grouped = defaultdict(list)

    for r in expanded_rows:
        no_resi = clean(safe(r, "no_resi")) or f"TANPA_RESI_{r['_row_no']}"
        grouped[no_resi].append(r)

    grouped_orders = []

    for no_resi, items in grouped.items():
        first = items[0]
        components = []

        for r in items:
            components.extend(r["_components"])

        order = {
            "no_resi": no_resi,
            "akun": first["_akun"],
            "waktu": first["_waktu"],
            "platform": first["_platform"],
            "label": platform_akun_label(first["_akun"], first["_platform"]),
            "courier": clean(safe(first, "courier")),
            "layanan": clean(safe(first, "layanan")),
            "pembayaran": first["_pembayaran"],
            "kode_pengambilan": clean(safe(first, "kode_pengambilan")),
            "no_pesanan": clean(safe(first, "no_pesanan")),
            "nama_penerima": clean(safe(first, "nama_penerima")),
            "alamat": clean(safe(first, "alamat")),
            "source_file": clean(safe(first, "source_file")),
            "items": items,
            "components": components,
            "produk_packing": combine_components(components, packing_only=True),
            "produk_ringkasan": combine_components(components, ringkasan_only=True),
            "total_item": sum(to_int(c["qty"]) for c in components if c.get("packing", True)),
            "status": "PERLU TINDAKAN" if any(i["_status"] != "OK" for i in items) else "OK",
        }

        grouped_orders.append(order)

    grouped_orders = sorted(
        grouped_orders,
        key=lambda x: (
            x["akun"],
            x["waktu"],
            x["platform"],
            x["courier"],
            x["no_resi"],
        )
    )

    rekap_barang = defaultdict(int)
    rekap_ringkasan = defaultdict(int)

    for order in grouped_orders:
        for c in order["components"]:
            if c.get("packing", True):
                rekap_barang[c["item"]] += to_int(c["qty"])
            if c.get("ringkasan", True):
                rekap_ringkasan[c["item"]] += to_int(c["qty"])

    combo_map = defaultdict(lambda: {
        "resi": [],
        "qty": 0,
        "akun": set(),
        "platform": set(),
        "pembayaran": set(),
        "kurir": set(),
    })

    for order in grouped_orders:
        combo = order["produk_packing"] or "(produk belum terbaca)"
        combo_map[combo]["resi"].append(order["no_resi"])
        combo_map[combo]["qty"] += order["total_item"]
        combo_map[combo]["akun"].add(order["akun"])
        combo_map[combo]["platform"].add(order["platform"])
        combo_map[combo]["pembayaran"].add(order["pembayaran"])
        combo_map[combo]["kurir"].add(order["courier"])

    return {
        "expanded_rows": expanded_rows,
        "grouped_orders": grouped_orders,
        "rekap_barang": rekap_barang,
        "rekap_ringkasan": rekap_ringkasan,
        "perlu_tindakan": perlu_tindakan,
        "combo_map": combo_map,
    }


# =========================
# SHEET 1
# =========================
def write_ringkasan_operasional(wb, data):
    ws = wb.active
    ws.title = "RINGKASAN OPERASIONAL"
    ws.sheet_tab_color = "F97316"
    set_sheet_view(ws)

    orders = data["grouped_orders"]
    rekap = data["rekap_ringkasan"]
    perlu = data["perlu_tindakan"]

    total_resi = len(orders)

    bg100 = rekap.get("Black Garlic 100gr", 0)
    bg220 = rekap.get("Black Garlic 220gr", 0)
    bg500 = rekap.get("Black Garlic 500gr", 0)
    bg84 = rekap.get("Black Garlic 84gr", 0)
    total_bg = bg100 + bg220 + bg500 + bg84

    drink_ori = rekap.get("BG Drink Original", 0)
    drink_peach = rekap.get("BG Drink Peach", 0)
    total_drink = drink_ori + drink_peach

    madu_kurma = rekap.get("Madu Bunga Kurma", 0)
    madu_multi = rekap.get("Madu Multi Floral", 0)
    total_madu = madu_kurma + madu_multi

    total_produk = total_bg + total_drink + total_madu

    set_col_widths(ws, {
        "A": 22,
        "B": 14,
        "C": 14,
        "D": 14,
        "E": 14,
        "F": 14,
        "G": 14,
        "H": 14,
        "I": 14,
    })

    title_bar(ws, "A1:I1", datetime.now().strftime("%-m/%-d/%Y") if hasattr(datetime.now(), "strftime") else datetime.now().strftime("%m/%d/%Y"), BOS_YELLOW, color="000000")

    box_metric(ws, 3, 1, "TOTAL RESI", total_resi, BOS_GREY, "Pcs")
    box_metric(ws, 3, 3, "TOTAL BG", total_bg, BOS_GREEN, "Pcs")
    box_metric(ws, 3, 5, "BG 100", bg100, BOS_GREEN, "Pcs")
    box_metric(ws, 3, 7, "BG 220", bg220, BOS_GREEN, "Pcs")

    box_metric(ws, 6, 1, "BG 500", bg500, BOS_GREEN, "Pcs")
    box_metric(ws, 6, 3, "BG 84", bg84, BOS_GREEN, "Pcs")
    box_metric(ws, 6, 5, "TOTAL DRINK", total_drink, BOS_ORANGE, "Pcs")
    box_metric(ws, 6, 7, "BG DRINK ORIGINAL", drink_ori, BOS_ORANGE, "Pcs")

    box_metric(ws, 9, 1, "BG DRINK PEACH", drink_peach, BOS_ORANGE, "Pcs")
    box_metric(ws, 9, 3, "TOTAL MADU", total_madu, BOS_BLUE, "Pcs")
    box_metric(ws, 9, 5, "Madu Bunga Kurma", madu_kurma, BOS_BLUE, "Pcs")
    box_metric(ws, 9, 7, "Madu Multi Floral", madu_multi, BOS_BLUE, "Pcs")

    if perlu:
        title_bar(ws, "A12:I12", f"Status: PERLU REVIEW ADMIN - {len(perlu)} data perlu tindakan", ORANGE_DARK)
    else:
        title_bar(ws, "A12:I12", "Status: SIAP DIGUNAKAN", GREEN_DARK)

    by_label = defaultdict(lambda: {
        "resi": set(),
        "qty": 0,
        "bg100": 0,
        "bg220": 0,
        "bg500": 0,
        "bg84": 0,
        "ori": 0,
        "peach": 0,
        "kurma": 0,
        "multi": 0,
    })

    for order in orders:
        label = order["label"]
        by_label[label]["resi"].add(order["no_resi"])

        for c in order["components"]:
            if not c.get("ringkasan", True):
                continue

            item = c["item"]
            qty = to_int(c["qty"])
            by_label[label]["qty"] += qty

            if item == "Black Garlic 100gr":
                by_label[label]["bg100"] += qty
            elif item == "Black Garlic 220gr":
                by_label[label]["bg220"] += qty
            elif item == "Black Garlic 500gr":
                by_label[label]["bg500"] += qty
            elif item == "Black Garlic 84gr":
                by_label[label]["bg84"] += qty
            elif item == "BG Drink Original":
                by_label[label]["ori"] += qty
            elif item == "BG Drink Peach":
                by_label[label]["peach"] += qty
            elif item == "Madu Bunga Kurma":
                by_label[label]["kurma"] += qty
            elif item == "Madu Multi Floral":
                by_label[label]["multi"] += qty

    start = 14
    headers = ["BARANG / PLATFORM", "RESI", "TOTAL", "BG100", "BG220", "BG500", "BG84", "ORI", "PEACH"]

    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=start, column=col, value=h)
        cell.fill = BOS_YELLOW
        cell.font = font(size=10, bold=True, color="000000")
        cell.alignment = CENTER
        cell.border = BORDER

    row = start + 1
    preferred = ["TIKTOK HSS", "SHOPEE", "TIKTOK HSD", "LAZADA", "BLIBLI"]
    labels = []

    for p in preferred:
        if p in by_label:
            labels.append(p)

    for k in sorted(by_label.keys()):
        if k not in labels:
            labels.append(k)

    for idx, label in enumerate(labels):
        d = by_label[label]
        if "HSS" in label:
            fill = GREEN
        elif "SHOPEE" in label:
            fill = ORANGE
        elif "HSD" in label:
            fill = BOS_BLUE
        elif "LAZADA" in label:
            fill = BLUE
        else:
            fill = WHITE

        values = [
            label,
            len(d["resi"]),
            d["qty"],
            d["bg100"],
            d["bg220"],
            d["bg500"],
            d["bg84"],
            d["ori"],
            d["peach"],
        ]

        write_row(ws, row, values, fill=fill, center_cols=[2, 3, 4, 5, 6, 7, 8, 9])
        row += 1

    ws.freeze_panes = "A14"


# =========================
# SHEET 2
# =========================
def write_rekap_kebutuhan_barang(wb, data):
    ws = wb.create_sheet("REKAP KEBUTUHAN BARANG")
    ws.sheet_tab_color = "2563EB"
    set_sheet_view(ws)

    title_bar(ws, "A1:E1", "REKAP KEBUTUHAN BARANG - UNTUK AMBIL STOK GUDANG", BLUE_DARK)

    perlu = len(data["perlu_tindakan"])
    status = "SIAP DIGUNAKAN" if perlu == 0 else f"PERLU REVIEW ADMIN: {perlu} DATA"
    fill = GREEN_DARK if perlu == 0 else ORANGE_DARK
    title_bar(ws, "A2:E2", f"Status Rekap: {status}", fill)

    set_header(ws, [
        ("No", 6),
        ("Nama Barang", 36),
        ("Total Qty", 14),
        ("Kategori", 24),
        ("Keterangan", 36),
    ], row=4)

    main_order = [
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

    rekap = data["rekap_barang"]
    names = [n for n in main_order if rekap.get(n, 0) > 0]

    for n in sorted(rekap.keys()):
        if n not in names and rekap.get(n, 0) > 0:
            names.append(n)

    row = 5

    for idx, name in enumerate(names, 1):
        qty = rekap.get(name, 0)

        if "Black Garlic" in name:
            kategori = "Black Garlic"
        elif "Drink" in name:
            kategori = "Drink"
        elif "Madu" in name:
            kategori = "Madu"
        elif "Box" in name or "Goodie" in name:
            kategori = "Perlengkapan Packing"
        elif "PERLU CEK" in name:
            kategori = "Perlu Review"
        else:
            kategori = "Lainnya"

        note = "Ambil dari stok gudang"
        if "PERLU CEK" in name:
            note = "Jangan dijadikan final sebelum dicek"

        fill_row = RED_LIGHT if "PERLU CEK" in name else (GREEN if idx % 2 == 0 else WHITE)

        values = [idx, name, qty, kategori, note]
        write_row(ws, row, values, fill=fill_row, center_cols=[1, 3])
        row += 1

    add_total_row(ws, row, ["", "TOTAL", sum(rekap.values()), "", ""])


# =========================
# SHEET 3
# =========================
def write_panduan_siap_packing(wb, data):
    ws = wb.create_sheet("PANDUAN SIAP PACKING")
    ws.sheet_tab_color = "16A34A"
    set_sheet_view(ws)

    title_bar(ws, "A1:G1", "PANDUAN SIAP PACKING - KELOMPOK PAKET", GREEN_DARK)

    set_header(ws, [
        ("No", 6),
        ("Paket yang Harus Disiapkan", 70),
        ("Jumlah Resi/Paket", 16),
        ("Total Barang", 14),
        ("Akun", 18),
        ("Platform", 22),
        ("Catatan", 35),
    ], row=3)

    combo_map = data["combo_map"]

    sorted_combos = sorted(
        combo_map.items(),
        key=lambda x: (-len(x[1]["resi"]), x[0])
    )

    row = 4

    for idx, (combo, d) in enumerate(sorted_combos, 1):
        is_review = "PERLU CEK" in combo
        fill = RED_LIGHT if is_review else (GREEN if idx % 2 == 0 else WHITE)

        values = [
            idx,
            combo,
            len(d["resi"]),
            d["qty"],
            ", ".join(sorted([x for x in d["akun"] if x])),
            ", ".join(sorted([x for x in d["platform"] if x])),
            "Perlu dicek" if is_review else "Siapkan sesuai jumlah paket",
        ]

        write_row(ws, row, values, fill=fill, center_cols=[1, 3, 4])
        row += 1


# =========================
# SHEET 4
# =========================
def write_detail_packing_resi(wb, data):
    ws = wb.create_sheet("DETAIL PACKING RESI")
    ws.sheet_tab_color = "84CC16"
    set_sheet_view(ws)

    set_header(ws, [
        ("No", 6),
        ("Akun", 10),
        ("Waktu", 12),
        ("Platform", 16),
        ("No. Resi", 24),
        ("Produk yang Harus Disiapkan", 70),
        ("Total Item", 12),
        ("Kurir", 16),
        ("Layanan", 14),
        ("Pembayaran", 16),
        ("Status", 18),
        ("Cek", 10),
    ])

    row = 2

    for idx, order in enumerate(data["grouped_orders"], 1):
        fill = RED_LIGHT if order["status"] != "OK" else (GREEN if idx % 2 == 0 else WHITE)

        values = [
            idx,
            order["akun"],
            order["waktu"],
            order["platform"],
            order["no_resi"],
            order["produk_packing"],
            order["total_item"],
            order["courier"],
            order["layanan"],
            order["pembayaran"],
            "Perlu Review" if order["status"] != "OK" else "Siap",
            "",
        ]

        write_row(ws, row, values, fill=fill, center_cols=[1, 2, 3, 7, 9, 10, 11, 12])
        row += 1


# =========================
# SHEET 5
# =========================
def write_data_order_customer(wb, data):
    ws = wb.create_sheet("DATA ORDER CUSTOMER")
    ws.sheet_tab_color = "F97316"
    set_sheet_view(ws)

    title_bar(ws, "A1:M1", "DATA ORDER CUSTOMER", ORANGE_DARK)

    set_header(ws, [
        ("No", 6),
        ("Akun", 10),
        ("Waktu", 12),
        ("Platform", 16),
        ("No. Resi", 24),
        ("No. Pesanan", 24),
        ("Nama Pembeli", 26),
        ("Alamat", 60),
        ("Produk Dibeli", 70),
        ("Total Item", 12),
        ("Pembayaran", 16),
        ("Kurir", 16),
        ("Layanan", 14),
    ], row=3)

    row = 4

    for idx, order in enumerate(data["grouped_orders"], 1):
        fill = BLUE if idx % 2 == 0 else WHITE

        values = [
            idx,
            order["akun"],
            order["waktu"],
            order["platform"],
            order["no_resi"],
            order["no_pesanan"],
            order["nama_penerima"],
            order["alamat"],
            order["produk_packing"],
            order["total_item"],
            order["pembayaran"],
            order["courier"],
            order["layanan"],
        ]

        write_row(ws, row, values, fill=fill, center_cols=[1, 2, 3, 10, 11, 13])
        row += 1


# =========================
# SHEET 6
# =========================
def write_order_instant(wb, data):
    ws = wb.create_sheet("ORDER INSTANT")
    ws.sheet_tab_color = "F5A623"
    set_sheet_view(ws)

    set_header(ws, [
        ("No", 6),
        ("Akun", 10),
        ("Waktu", 12),
        ("Platform", 16),
        ("No. Resi", 24),
        ("Kode Pengambilan", 22),
        ("Nama Pembeli", 26),
        ("Produk", 70),
        ("Kurir", 16),
        ("Layanan", 16),
        ("Catatan", 35),
    ])

    row = 2
    no = 1
    found = False

    for order in data["grouped_orders"]:
        text = " ".join([
            order["courier"],
            order["layanan"],
            order["kode_pengambilan"],
            order["no_resi"],
        ]).upper()

        is_instant = (
            "INSTANT" in text
            or "GOSEND" in text
            or "GOJEK" in text
            or "GRAB" in text
            or bool(order["kode_pengambilan"])
        )

        if not is_instant:
            continue

        found = True

        values = [
            no,
            order["akun"],
            order["waktu"],
            order["platform"],
            order["no_resi"],
            order["kode_pengambilan"],
            order["nama_penerima"],
            order["produk_packing"],
            order["courier"],
            order["layanan"],
            "Butuh perhatian admin" if order["kode_pengambilan"] else "Instant tanpa kode terbaca",
        ]

        write_row(ws, row, values, fill=YELLOW, center_cols=[1, 2, 3, 6, 10])
        row += 1
        no += 1

    if not found:
        ws.merge_cells("A2:K2")
        cell = ws["A2"]
        cell.value = "Tidak ada order instant / kode pengambilan yang terbaca di batch ini."
        cell.fill = WHITE
        cell.font = font(size=10, bold=True, color="111827")
        cell.alignment = LEFT
        cell.border = BORDER


# =========================
# SHEET 7
# =========================
def write_kontrol_validasi_pdf(wb, data):
    ws = wb.create_sheet("KONTROL VALIDASI PDF")
    ws.sheet_tab_color = "6B7280"
    set_sheet_view(ws)

    title_bar(ws, "A1:G1", "KONTROL VALIDASI PDF", DARK)

    orders = data["grouped_orders"]
    expanded = data["expanded_rows"]
    perlu = data["perlu_tindakan"]

    total_resi = len(orders)
    total_baris_produk = len(expanded)
    total_perlu = len(perlu)

    status = "SIAP DIGUNAKAN" if total_perlu == 0 else "PERLU REVIEW ADMIN"
    fill = GREEN_DARK if total_perlu == 0 else ORANGE_DARK

    title_bar(ws, "A2:G2", f"Status Validasi: {status}", fill)

    set_header(ws, [
        ("Metrik", 35),
        ("Jumlah", 16),
        ("Keterangan", 70),
    ], row=4)

    rows = [
        ("Total Resi Terbaca", total_resi, "Jumlah resi yang berhasil masuk ke Excel"),
        ("Total Baris Produk Terbaca", total_baris_produk, "Jumlah baris produk dari parser"),
        ("Data Perlu Tindakan", total_perlu, "Produk/SKU yang belum dikenali Master SKU"),
        ("Catatan", "", "Kalau angka manual berbeda jauh, kemungkinan parser belum membaca semua halaman/resi dari PDF."),
    ]

    row = 5
    for idx, values in enumerate(rows):
        fill_row = RED_LIGHT if values[0] == "Data Perlu Tindakan" and total_perlu > 0 else (GREY if idx % 2 == 0 else WHITE)
        write_row(ws, row, values, fill=fill_row, center_cols=[2])
        row += 1

    row += 2
    title_bar(ws, f"A{row}:G{row}", "RINGKASAN PER FILE / SUMBER PDF", DARK)
    row += 1

    headers2 = [
        ("Source File", 40),
        ("Total Resi", 14),
        ("Total Baris Produk", 18),
        ("Perlu Tindakan", 16),
        ("Catatan", 55),
    ]

    for col, (h, w) in enumerate(headers2, 1):
        cell = ws.cell(row=row, column=col, value=h)
        cell.fill = HEADER
        cell.font = FONT_HEADER
        cell.alignment = CENTER
        cell.border = BORDER
        ws.column_dimensions[get_column_letter(col)].width = w

    by_file = defaultdict(lambda: {"resi": set(), "rows": 0, "perlu": 0})

    for r in expanded:
        file_name = clean(safe(r, "source_file")) or "(tanpa source file)"
        by_file[file_name]["rows"] += 1
        by_file[file_name]["resi"].add(clean(safe(r, "no_resi")))
        if r["_status"] != "OK":
            by_file[file_name]["perlu"] += 1

    row += 1

    for idx, file_name in enumerate(sorted(by_file.keys()), 1):
        d = by_file[file_name]
        fill_row = RED_LIGHT if d["perlu"] > 0 else (GREEN if idx % 2 == 0 else WHITE)

        values = [
            file_name,
            len(d["resi"]),
            d["rows"],
            d["perlu"],
            "Perlu dicek" if d["perlu"] > 0 else "OK",
        ]

        write_row(ws, row, values, fill=fill_row, center_cols=[2, 3, 4])
        row += 1


# =========================
# SHEET 8
# =========================
def write_perlu_tindakan(wb, data):
    ws = wb.create_sheet("PERLU TINDAKAN")
    ws.sheet_tab_color = "DC2626"
    set_sheet_view(ws)

    set_header(ws, [
        ("No", 6),
        ("Akun", 10),
        ("Waktu", 12),
        ("Platform", 16),
        ("No. Resi", 24),
        ("Nama Produk", 45),
        ("SKU", 32),
        ("Qty", 10),
        ("Masalah", 42),
        ("Saran Tindakan", 50),
        ("Source File", 32),
    ])

    perlu = data["perlu_tindakan"]

    if not perlu:
        ws.merge_cells("A2:K2")
        cell = ws["A2"]
        cell.value = "Tidak ada data yang perlu tindakan. Semua SKU/produk berhasil diterjemahkan oleh Master SKU HSD."
        cell.fill = GREEN
        cell.font = font(size=11, bold=True, color="111827")
        cell.alignment = LEFT
        cell.border = BORDER
        return

    row = 2

    for idx, r in enumerate(perlu, 1):
        values = [
            idx,
            r["_akun"],
            r["_waktu"],
            r["_platform"],
            clean(safe(r, "no_resi")),
            clean(safe(r, "nama_produk")) or clean(safe(r, "nama_produk_asli")),
            clean(safe(r, "sku")),
            to_int(safe(r, "qty", 0)),
            r["_reason"],
            "Tambahkan SKU/nama produk ini ke Master SKU HSD",
            clean(safe(r, "source_file")),
        ]

        write_row(ws, row, values, fill=RED_LIGHT, center_cols=[1, 2, 3, 8])
        row += 1


# =========================
# SHEET 9
# =========================
def write_master_sku(wb):
    ws = wb.create_sheet("MASTER SKU HSD")
    ws.sheet_tab_color = "7C3AED"
    set_sheet_view(ws)

    title_bar(ws, "A1:F1", "MASTER SKU HSD - KAMUS PRODUK SISTEM", fill=PURPLE, color="111827")

    set_header(ws, [
        ("No", 6),
        ("SKU / Nama yang Dikenali", 42),
        ("Isi Barang", 55),
        ("Masuk Ringkasan", 18),
        ("Masuk Packing", 16),
        ("Catatan", 45),
    ], row=3)

    row = 4

    for idx, token in enumerate(sorted(MASTER_SKU.keys()), 1):
        components = MASTER_SKU[token]
        isi = " + ".join([f"{c['item']} x{c['qty']}" for c in components])
        masuk_ringkasan = ", ".join(sorted(set(["Ya" if c.get("ringkasan", True) else "Tidak" for c in components])))
        masuk_packing = ", ".join(sorted(set(["Ya" if c.get("packing", True) else "Tidak" for c in components])))

        values = [
            idx,
            token,
            isi,
            masuk_ringkasan,
            masuk_packing,
            "Master SKU internal HSD",
        ]

        fill = PURPLE if idx % 2 == 0 else WHITE
        write_row(ws, row, values, fill=fill, center_cols=[1, 4, 5])
        row += 1


# =========================
# MAIN WRITER
# =========================
def write_excel_multi(rows, output_path):
    wb = Workbook()

    if rows is None:
        rows = []

    data = prepare_data(rows)

    write_ringkasan_operasional(wb, data)
    write_rekap_kebutuhan_barang(wb, data)
    write_panduan_siap_packing(wb, data)
    write_detail_packing_resi(wb, data)
    write_data_order_customer(wb, data)
    write_order_instant(wb, data)
    write_kontrol_validasi_pdf(wb, data)
    write_perlu_tindakan(wb, data)
    write_master_sku(wb)

    wb.save(output_path)
    return output_path


# =========================
# COMPATIBILITY UNTUK APP.PY
# =========================
def write_excel(rows, output_path):
    return write_excel_multi(rows, output_path)


def create_excel(rows, output_path):
    return write_excel_multi(rows, output_path)


def generate_excel(rows, output_path):
    return write_excel_multi(rows, output_path)


def build_excel(rows, output_path):
    return write_excel_multi(rows, output_path)


def make_excel(rows, output_path):
    return write_excel_multi(rows, output_path)


def export_excel(rows, output_path):
    return write_excel_multi(rows, output_path)
