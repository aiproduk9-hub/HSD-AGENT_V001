import re
from collections import defaultdict
from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.exceptions import IllegalCharacterError


# =========================================================
# STYLE DASAR
# =========================================================

HDR_FILL = PatternFill("solid", fgColor="1B1F2E")
HDR_FONT = Font(bold=True, color="FFFFFF", size=10)

WHITE = PatternFill("solid", fgColor="FFFFFF")
CREAM = PatternFill("solid", fgColor="FFF7E8")
GREY = PatternFill("solid", fgColor="F3F4F6")
DARK = PatternFill("solid", fgColor="111827")

BLUE = PatternFill("solid", fgColor="D6EAFF")
BLUE_DARK = PatternFill("solid", fgColor="2563EB")

ORANGE = PatternFill("solid", fgColor="FFE4CC")
ORANGE_DARK = PatternFill("solid", fgColor="F97316")

GREEN = PatternFill("solid", fgColor="EAF3DE")
GREEN_DARK = PatternFill("solid", fgColor="16A34A")

YELLOW = PatternFill("solid", fgColor="FFF9C4")
RED_LIGHT = PatternFill("solid", fgColor="FFEBEE")
PURPLE = PatternFill("solid", fgColor="EDE9FE")

CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
RIGHT = Alignment(horizontal="right", vertical="center", wrap_text=True)

BORDER = Border(
    left=Side(style="thin", color="D1D5DB"),
    right=Side(style="thin", color="D1D5DB"),
    top=Side(style="thin", color="D1D5DB"),
    bottom=Side(style="thin", color="D1D5DB"),
)


# =========================================================
# HELPER DASAR
# =========================================================

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
    return " ".join(v.replace("\uFFFE", " ").split()).strip().rstrip(",")


def to_int(v):
    try:
        return int(float(str(v).replace(",", "").strip()))
    except Exception:
        return 0


def normalize_token(v):
    """
    Biar SKU yang kepisah spasi / beda penulisan tetap kebaca.
    Contoh:
    BG-220GR-1- BOTOL-BG-SKU -> BG-220GR-1-BOTOL-BG-SKU
    BG-100GR-1 BOTOL-BH-SKU -> BG-100GR-1-BOTOL-BH-SKU
    """
    v = clean(v).upper()
    v = v.replace("_", "-")
    v = re.sub(r"\s+", "-", v)
    v = re.sub(r"-+", "-", v)
    v = v.strip("-")
    return v


def set_header(ws, headers, row=1):
    for col, item in enumerate(headers, 1):
        title, width = item

        cell = ws.cell(row=row, column=col, value=title)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.alignment = CENTER
        cell.border = BORDER

        ws.column_dimensions[get_column_letter(col)].width = width

    ws.freeze_panes = f"A{row + 1}"
    ws.row_dimensions[row].height = 28


def write_row(ws, row, values, fill=WHITE, bold_cols=None, center_cols=None):
    bold_cols = bold_cols or []
    center_cols = center_cols or []

    for col, value in enumerate(values, 1):
        cell = ws.cell(row=row, column=col, value=value)
        cell.fill = fill
        cell.border = BORDER
        cell.alignment = CENTER if col in center_cols else LEFT
        cell.font = Font(size=10, bold=col in bold_cols)


def add_total_row(ws, row, values):
    for col, value in enumerate(values, 1):
        cell = ws.cell(row=row, column=col, value=value)
        cell.fill = ORANGE_DARK
        cell.font = Font(bold=True, color="FFFFFF", size=11)
        cell.alignment = CENTER
        cell.border = BORDER


def get_akun(r):
    akun = clean(safe(r, "akun")) or clean(safe(r, "brand"))
    return akun.upper() if akun else "-"


def get_waktu(r):
    waktu = clean(safe(r, "waktu")) or clean(safe(r, "shift"))
    return waktu.title() if waktu else "-"


def get_platform(r):
    platform = clean(safe(r, "platform"))
    if not platform:
        return "-"
    p = platform.upper()
    if "TIKTOK" in p:
        return "TikTok"
    if "SHOPEE" in p:
        return "Shopee"
    if "LAZADA" in p:
        return "Lazada"
    if "BLIBLI" in p:
        return "Blibli"
    if "TOKOPEDIA" in p:
        return "Tokopedia/TikTok"
    return platform


def get_pembayaran(r):
    raw = " ".join(
        [
            clean(safe(r, "pembayaran")),
            clean(safe(r, "payment")),
            clean(safe(r, "cod")),
            clean(safe(r, "layanan")),
            clean(safe(r, "courier")),
            clean(safe(r, "catatan")),
        ]
    ).upper()

    if "COD" in raw:
        return "COD"

    if "NON-COD" in raw or "NON COD" in raw:
        return "Non-COD"

    # Kalau parser belum kasih data pembayaran, jangan ditebak terlalu agresif.
    return "Belum Terbaca"


def get_kode_pengambilan(r):
    return clean(safe(r, "kode_pengambilan")) or clean(safe(r, "kode"))


def account_platform_label(akun, platform):
    akun = clean(akun).upper()
    platform = clean(platform)

    p = platform.upper()

    if "TIKTOK" in p:
        return f"TIKTOK {akun}" if akun and akun != "-" else "TIKTOK"
    if "TOKOPEDIA" in p:
        return f"TIKTOK {akun}" if akun and akun != "-" else "TOKOPEDIA/TIKTOK"
    if "SHOPEE" in p:
        return "SHOPEE"
    if "LAZADA" in p:
        return "LAZADA"
    if "BLIBLI" in p:
        return "BLIBLI"

    if akun and akun != "-":
        return f"{platform} {akun}".strip()

    return platform or "-"


# =========================================================
# MASTER SKU HSD
# =========================================================

def comp(item, qty, ringkasan=True, packing=True):
    return {
        "item": item,
        "qty": int(qty),
        "ringkasan": bool(ringkasan),
        "packing": bool(packing),
    }


def exact_master_components():
    """
    MASTER SKU resmi tahap pertama.
    Sumber: info internal HSD dari user/Mbak Fitri.
    """
    m = {}

    def add(keys, components):
        for k in keys:
            m[normalize_token(k)] = components

    # -------------------------
    # MADU
    # -------------------------
    add(["BGH-MULTI-FLORAL-1-BOTOL"], [
        comp("Madu Multi Floral", 1),
    ])

    add(["BGH-BUNGA-KURMA-1-BOTOL"], [
        comp("Madu Bunga Kurma", 1),
    ])

    # -------------------------
    # DRINK ORIGINAL 1 BOTOL
    # -------------------------
    add([
        "BG-DRINK-ORI-1-BOTOL-PROMO",
        "BG-DRINK-PROMO-1-BTL-ORI",
        "BG-Drink-Original-FS",
        "BG-Drink-Original",
        "BG-DRINK-ORIGINAL",
        "BLACKGARLIC-DRINK-ORIGINAL",
        "BLACKGARLIC-DRINK-ORIGINAL-FS",
    ], [
        comp("BG Drink Original", 1),
    ])

    # DRINK ORIGINAL 2 BOTOL
    add([
        "BG-DRINK-ORI-2-BOTOL-PROMO",
        "BG-DRINK-PROMO-2-BTL-ORI",
    ], [
        comp("BG Drink Original", 2),
    ])

    # DRINK ORIGINAL 4 BOTOL
    add([
        "BG-DRINK-PROMO-4-BTL-ORI",
    ], [
        comp("BG Drink Original", 4),
    ])

    # DRINK ORIGINAL 7 BOTOL
    add([
        "BG-Drink-Original-7-Botol",
        "BG-DRINK-ORIGINAL-7-BOTOL",
    ], [
        comp("BG Drink Original", 7),
    ])

    # -------------------------
    # DRINK PEACH
    # -------------------------
    add([
        "BG-Drink-Peach",
        "BG-Drink-Peach-FS",
        "BG-DRINK-PEACH",
        "BLACKGARLIC-DRINK-PEACH",
        "BLACKGARLIC-DRINK-PEACH-FS",
    ], [
        comp("BG Drink Peach", 1),
    ])

    add([
        "BG-Drink-Peach-7-Botol",
        "BG-DRINK-PEACH-7-BOTOL",
    ], [
        comp("BG Drink Peach", 7),
    ])

    # -------------------------
    # DRINK MIX 7 BOTOL
    # -------------------------
    add([
        "BG-Drink-Mix-7-Botol",
        "BG-DRINK-MIX-7-BOTOL",
    ], [
        comp("BG Drink Peach", 4),
        comp("BG Drink Original", 3),
    ])

    # -------------------------
    # BOX HAMPERS
    # -------------------------
    add(["BOX-HAMPERS-IMLEK"], [
        comp("Box Hampers Imlek", 1, ringkasan=False),
    ])

    add(["BOX-HAMPERS-HSD"], [
        comp("Box Hampers HSD", 1, ringkasan=False),
    ])

    add(["BOX-HAMPERS-NATAL"], [
        comp("Box Hampers Natal", 1, ringkasan=False),
    ])

    add(["BOX-HAMPERS-LEBARAN"], [
        comp("Box Hampers Lebaran", 1, ringkasan=False),
    ])

    # -------------------------
    # HAMPERS 220GR 2 BOTOL + BOX
    # -------------------------
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

    # -------------------------
    # 3 IN 1 + GOODIE BAG
    # -------------------------
    add([
        "BG-3IN1",
        "Black Garlic Tunggal 3 in 1 + Goodie Bag HSD",
    ], [
        comp("Black Garlic 100gr", 1),
        comp("Black Garlic 220gr", 1),
        comp("Black Garlic 500gr", 1),
        comp("Goodie Bag HSD", 1, ringkasan=False),
    ])

    # -------------------------
    # BG 84GR
    # Catatan: sesuai info internal, BLACKGARLIC-HONAN-HSD-84G-3PCS dihitung 1 botol 84gr.
    # -------------------------
    add([
        "BLACKGARLIC-HONAN-HSD-84G-3PCS",
        "BG-84GR-CLOVER",
        "BG-84GR-1-CLOVER",
    ], [
        comp("Black Garlic 84gr", 1),
    ])

    # -------------------------
    # BG 100GR SATUAN DAN PAKET 1-5 BOTOL
    # -------------------------
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
    ], [
        comp("Black Garlic 100gr", 1),
    ])

    add([
        "BG-100GR-2-BOTOL",
        "BG-100GR-2-BOTOL-3CM",
    ], [
        comp("Black Garlic 100gr", 2),
    ])

    add([
        "BG-100GR-3-BOTOL",
        "BG-100GR-3-BOTOL-3CM",
    ], [
        comp("Black Garlic 100gr", 3),
    ])

    add([
        "BG-100GR-4-BOTOL",
        "BG-100GR-4-BOTOL-3CM",
    ], [
        comp("Black Garlic 100gr", 4),
    ])

    add([
        "BG-100GR-5-BOTOL",
        "BG-100GR-5-BOTOL-3CM",
    ], [
        comp("Black Garlic 100gr", 5),
    ])

    # -------------------------
    # BG 220GR SATUAN DAN PAKET 1-5 BOTOL
    # -------------------------
    add([
        "BG-220GR-1-BOTOL",
        "BG-220GR-1-BOTOL-BG-SKU",
        "BG-220GR-1-BOTOL-BH",
        "BG-220GR-1 BOTOL-BH-SKU",
        "BG-220GR-SKU",
        "BG-220GR-SKU",
        "BG-220GR-1-BOTOL-V2",
        "BG-220GR-1- BOTOL-BG-SKU",
        "BG-220GR-1- BOTOL-BH",
        "BG-220GR-1 BOTOL-BG-SKU",
        "BG-220GR-SK U",
    ], [
        comp("Black Garlic 220gr", 1),
    ])

    add(["BG-220GR-2-BOTOL"], [
        comp("Black Garlic 220gr", 2),
    ])

    add(["BG-220GR-3-BOTOL"], [
        comp("Black Garlic 220gr", 3),
    ])

    add(["BG-220GR-4-BOTOL"], [
        comp("Black Garlic 220gr", 4),
    ])

    add(["BG-220GR-5-BOTOL"], [
        comp("Black Garlic 220gr", 5),
    ])

    # -------------------------
    # BG 500GR SATUAN DAN PAKET 1-5 BOTOL
    # -------------------------
    add([
        "BG-500GR-1-BOTOL",
        "BG-500GR-1-BOTOL-BH",
        "BG-500GR-1-BOTOL-BG-SKU",
        "BG-500GR-1 BOTOL-BH-SKU",
        "BG-500GR-SKU",
        "BG-500GR-1- BOTOL",
        "BG-500GR-1 BOTOL",
    ], [
        comp("Black Garlic 500gr", 1),
    ])

    add(["BG-500GR-2-BOTOL"], [
        comp("Black Garlic 500gr", 2),
    ])

    add(["BG-500GR-3-BOTOL"], [
        comp("Black Garlic 500gr", 3),
    ])

    add(["BG-500GR-4-BOTOL"], [
        comp("Black Garlic 500gr", 4),
    ])

    add(["BG-500GR-5-BOTOL"], [
        comp("Black Garlic 500gr", 5),
    ])

    return m


MASTER_SKU = exact_master_components()


def multiply_components(components, multiplier):
    result = []
    multiplier = max(0, to_int(multiplier))

    for c in components:
        result.append({
            "item": c["item"],
            "qty": c["qty"] * multiplier,
            "ringkasan": c.get("ringkasan", True),
            "packing": c.get("packing", True),
        })

    return result


def expand_product(row):
    """
    Terjemahkan 1 baris produk dari parser menjadi barang nyata untuk gudang.
    Return:
    - components: daftar barang real
    - status: OK / PERLU TINDAKAN
    - reason: alasan kalau perlu tindakan
    """
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

    # 1) Cek exact master SKU
    for token in normalized_candidates:
        if token in MASTER_SKU:
            return multiply_components(MASTER_SKU[token], qty), "OK", ""

    # 2) Cek nama paket 3 in 1
    full_text = normalize_token(" ".join(candidates))
    full_readable = " ".join(candidates).upper()

    if "3-IN-1" in full_text or "3IN1" in full_text or ("3 IN 1" in full_readable and "BLACK" in full_readable):
        components = [
            comp("Black Garlic 100gr", 1),
            comp("Black Garlic 220gr", 1),
            comp("Black Garlic 500gr", 1),
            comp("Goodie Bag HSD", 1, ringkasan=False),
        ]
        return multiply_components(components, qty), "OK", ""

    # 3) Cek drink mix
    if "DRINK" in full_text and "MIX" in full_text and "7-BOTOL" in full_text:
        components = [
            comp("BG Drink Peach", 4),
            comp("BG Drink Original", 3),
        ]
        return multiply_components(components, qty), "OK", ""

    # 4) Cek drink original/peach 7 botol
    if "DRINK" in full_text and "ORIGINAL" in full_text and "7-BOTOL" in full_text:
        return multiply_components([comp("BG Drink Original", 7)], qty), "OK", ""

    if "DRINK" in full_text and "PEACH" in full_text and "7-BOTOL" in full_text:
        return multiply_components([comp("BG Drink Peach", 7)], qty), "OK", ""

    # 5) Cek drink satuan berdasarkan kata
    if "DRINK" in full_text and "PEACH" in full_text:
        return multiply_components([comp("BG Drink Peach", 1)], qty), "OK", ""

    if "DRINK" in full_text and ("ORIGINAL" in full_text or "-ORI" in full_text or "ORI-" in full_text):
        return multiply_components([comp("BG Drink Original", 1)], qty), "OK", ""

    # 6) Cek madu
    if "MULTI" in full_text and ("FLORAL" in full_text or "FLORA" in full_text):
        return multiply_components([comp("Madu Multi Floral", 1)], qty), "OK", ""

    if "KURMA" in full_text:
        return multiply_components([comp("Madu Bunga Kurma", 1)], qty), "OK", ""

    # 7) Cek 84gr
    if "84GR" in full_text or "84G" in full_text:
        return multiply_components([comp("Black Garlic 84gr", 1)], qty), "OK", ""

    # 8) Cek pola BG ukuran 100/220/500 dan jumlah botol
    # Contoh: BG-220GR-3-BOTOL
    m = re.search(r"BG-(100|220|500)GR-(\d+)-BOTOL", full_text)

    if m:
        size = m.group(1)
        jumlah_botol = int(m.group(2))
        return multiply_components([comp(f"Black Garlic {size}gr", jumlah_botol)], qty), "OK", ""

    # 9) Cek pola ukuran dari SKU/nama produk
    if "100GR" in full_text or "100-GR" in full_text or "100 GR" in full_readable:
        return multiply_components([comp("Black Garlic 100gr", 1)], qty), "OK", ""

    if "220GR" in full_text or "220-GR" in full_text or "220 GR" in full_readable:
        return multiply_components([comp("Black Garlic 220gr", 1)], qty), "OK", ""

    if "500GR" in full_text or "500-GR" in full_text or "500 GR" in full_readable:
        return multiply_components([comp("Black Garlic 500gr", 1)], qty), "OK", ""

    # 10) Kalau tidak dikenal, jangan dibuang.
    unknown_name = nama_raw or nama_asli or sku_raw or "(produk belum terbaca)"
    return [
        {
            "item": f"PERLU CEK: {unknown_name}",
            "qty": qty,
            "ringkasan": False,
            "packing": True,
        }
    ], "PERLU TINDAKAN", "SKU/Nama produk belum ada di Master SKU"


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


def prepare_data(rows):
    """
    Buat data turunan:
    - expanded_rows
    - grouped_orders
    - rekap_barang
    - perlu_tindakan
    """
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

    # Group by resi
    grouped = defaultdict(list)

    for r in expanded_rows:
        no_resi = clean(safe(r, "no_resi")) or f"TANPA_RESI_{r['_row_no']}"
        grouped[no_resi].append(r)

    grouped_orders = []

    for no_resi, items in grouped.items():
        all_components = []

        for r in items:
            all_components.extend(r["_components"])

        first = items[0]

        grouped_orders.append({
            "no_resi": no_resi,
            "akun": first["_akun"],
            "waktu": first["_waktu"],
            "platform": first["_platform"],
            "label": account_platform_label(first["_akun"], first["_platform"]),
            "courier": clean(safe(first, "courier")),
            "layanan": clean(safe(first, "layanan")),
            "pembayaran": first["_pembayaran"],
            "kode_pengambilan": get_kode_pengambilan(first),
            "no_pesanan": clean(safe(first, "no_pesanan")),
            "nama_penerima": clean(safe(first, "nama_penerima")),
            "alamat": clean(safe(first, "alamat")),
            "source_file": clean(safe(first, "source_file")),
            "items": items,
            "components": all_components,
            "produk_packing": combine_components(all_components, packing_only=True),
            "produk_ringkasan": combine_components(all_components, ringkasan_only=True),
            "total_item": sum(to_int(c["qty"]) for c in all_components if c.get("packing", True)),
            "status": "PERLU TINDAKAN" if any(i["_status"] != "OK" for i in items) else "OK",
        })

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

    # Rekap barang
    rekap_barang = defaultdict(int)
    rekap_ringkasan = defaultdict(int)

    for order in grouped_orders:
        for c in order["components"]:
            if c.get("packing", True):
                rekap_barang[c["item"]] += to_int(c["qty"])
            if c.get("ringkasan", True):
                rekap_ringkasan[c["item"]] += to_int(c["qty"])

    return {
        "expanded_rows": expanded_rows,
        "grouped_orders": grouped_orders,
        "rekap_barang": rekap_barang,
        "rekap_ringkasan": rekap_ringkasan,
        "perlu_tindakan": perlu_tindakan,
    }


# =========================================================
# SHEET 1: RINGKASAN OPERASIONAL
# =========================================================

def write_ringkasan_operasional(wb, data):
    ws = wb.active
    ws.title = "RINGKASAN OPERASIONAL"
    ws.sheet_tab_color = "F97316"

    orders = data["grouped_orders"]
    rekap = data["rekap_ringkasan"]
    perlu_tindakan = data["perlu_tindakan"]

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

    # Width
    widths = {
        "A": 18,
        "B": 16,
        "C": 16,
        "D": 16,
        "E": 16,
        "F": 16,
        "G": 16,
        "H": 16,
        "I": 16,
    }

    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    # Title
    ws.merge_cells("A1:I1")
    c = ws["A1"]
    c.value = f"RINGKASAN OPERASIONAL HSD - {datetime.now().strftime('%d/%m/%Y')}"
    c.fill = DARK
    c.font = Font(bold=True, color="FFFFFF", size=16)
    c.alignment = CENTER
    c.border = BORDER
    ws.row_dimensions[1].height = 32

    status_text = "SIAP DIGUNAKAN" if len(perlu_tindakan) == 0 else "PERLU REVIEW ADMIN"
    status_fill = GREEN_DARK if len(perlu_tindakan) == 0 else ORANGE_DARK

    ws.merge_cells("A2:I2")
    c = ws["A2"]
    c.value = f"Status Proses: {status_text}"
    c.fill = status_fill
    c.font = Font(bold=True, color="FFFFFF", size=11)
    c.alignment = CENTER
    c.border = BORDER

    # Big metrics
    metrics = [
        ("TOTAL RESI", total_resi, BLUE_DARK),
        ("TOTAL PRODUK", total_produk, ORANGE_DARK),
        ("TOTAL BG", total_bg, GREEN_DARK),
        ("TOTAL DRINK", total_drink, BLUE_DARK),
        ("TOTAL MADU", total_madu, ORANGE_DARK),
    ]

    start_col = 1
    for idx, (label, value, fill) in enumerate(metrics):
        col = start_col + idx * 2
        if col + 1 > 9:
            break

        ws.merge_cells(start_row=4, start_column=col, end_row=4, end_column=col + 1)
        ws.merge_cells(start_row=5, start_column=col, end_row=5, end_column=col + 1)

        c1 = ws.cell(row=4, column=col, value=label)
        c1.fill = fill
        c1.font = Font(bold=True, color="FFFFFF", size=11)
        c1.alignment = CENTER
        c1.border = BORDER

        c2 = ws.cell(row=5, column=col, value=value)
        c2.fill = WHITE
        c2.font = Font(bold=True, color="111827", size=24)
        c2.alignment = CENTER
        c2.border = BORDER

        # Border for merged partner
        ws.cell(row=4, column=col + 1).fill = fill
        ws.cell(row=4, column=col + 1).border = BORDER
        ws.cell(row=5, column=col + 1).fill = WHITE
        ws.cell(row=5, column=col + 1).border = BORDER

    ws.row_dimensions[5].height = 42

    # Detail product boxes
    detail_rows = [
        ("BLACK GARLIC", [
            ("BG 100gr", bg100),
            ("BG 220gr", bg220),
            ("BG 500gr", bg500),
            ("BG 84gr", bg84),
        ], GREEN),
        ("DRINK", [
            ("Original", drink_ori),
            ("Peach", drink_peach),
        ], BLUE),
        ("MADU", [
            ("Bunga Kurma", madu_kurma),
            ("Multi Floral", madu_multi),
        ], ORANGE),
    ]

    row = 8
    for title, items, fill in detail_rows:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=9)
        c = ws.cell(row=row, column=1, value=title)
        c.fill = DARK
        c.font = Font(bold=True, color="FFFFFF", size=12)
        c.alignment = CENTER
        c.border = BORDER

        row += 1
        col = 1

        for label, value in items:
            ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col + 1)
            ws.merge_cells(start_row=row + 1, start_column=col, end_row=row + 1, end_column=col + 1)

            c1 = ws.cell(row=row, column=col, value=label)
            c1.fill = fill
            c1.font = Font(bold=True, color="111827", size=10)
            c1.alignment = CENTER
            c1.border = BORDER

            c2 = ws.cell(row=row + 1, column=col, value=value)
            c2.fill = WHITE
            c2.font = Font(bold=True, size=20)
            c2.alignment = CENTER
            c2.border = BORDER

            ws.cell(row=row, column=col + 1).fill = fill
            ws.cell(row=row, column=col + 1).border = BORDER
            ws.cell(row=row + 1, column=col + 1).fill = WHITE
            ws.cell(row=row + 1, column=col + 1).border = BORDER

            col += 2

        row += 3

    # Ringkasan by platform/akun
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
            item = c["item"]
            qty = to_int(c["qty"])

            if not c.get("ringkasan", True):
                continue

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

    row += 1
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=9)
    c = ws.cell(row=row, column=1, value="RINGKASAN PER PLATFORM / AKUN")
    c.fill = DARK
    c.font = Font(bold=True, color="FFFFFF", size=12)
    c.alignment = CENTER
    c.border = BORDER

    row += 1
    headers = [
        ("Platform/Akun", 22),
        ("Resi", 10),
        ("Total Qty", 12),
        ("BG 100", 10),
        ("BG 220", 10),
        ("BG 500", 10),
        ("BG 84", 10),
        ("Drink Ori", 12),
        ("Drink Peach", 12),
    ]

    for col, (h, _) in enumerate(headers, 1):
        cell = ws.cell(row=row, column=col, value=h)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.alignment = CENTER
        cell.border = BORDER

    row += 1

    for idx, label in enumerate(sorted(by_label.keys())):
        d = by_label[label]
        fill = BLUE if idx % 2 == 0 else WHITE

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

        for col, value in enumerate(values, 1):
            cell = ws.cell(row=row, column=col, value=value)
            cell.fill = fill
            cell.border = BORDER
            cell.alignment = LEFT if col == 1 else CENTER
            cell.font = Font(bold=True if col in [2, 3] else False, size=10)

        row += 1

    ws.freeze_panes = "A4"


# =========================================================
# SHEET 2: DAFTAR SIAP PACKING
# =========================================================

def write_daftar_siap_packing(wb, data):
    ws = wb.create_sheet("DAFTAR SIAP PACKING")
    ws.sheet_tab_color = "16A34A"

    set_header(ws, [
        ("No", 6),
        ("Akun", 10),
        ("Waktu", 12),
        ("Platform", 16),
        ("No. Resi", 24),
        ("Produk yang Harus Disiapkan", 65),
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

        write_row(
            ws,
            row,
            values,
            fill=fill,
            bold_cols=[6, 7, 11],
            center_cols=[1, 2, 3, 7, 9, 10, 11, 12],
        )

        row += 1


# =========================================================
# SHEET 3: REKAP KEBUTUHAN BARANG
# =========================================================

def write_rekap_kebutuhan_barang(wb, data):
    ws = wb.create_sheet("REKAP KEBUTUHAN BARANG")
    ws.sheet_tab_color = "2563EB"

    set_header(ws, [
        ("No", 6),
        ("Nama Barang", 36),
        ("Total Qty", 14),
        ("Kategori", 22),
        ("Keterangan", 32),
    ])

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

    row = 2

    for idx, name in enumerate(names, 1):
        qty = rekap.get(name, 0)

        if "Black Garlic" in name:
            kategori = "Black Garlic"
        elif "Drink" in name:
            kategori = "Drink"
        elif "Madu" in name:
            kategori = "Madu"
        elif "Box" in name:
            kategori = "Perlengkapan Packing"
        elif "Goodie" in name:
            kategori = "Perlengkapan Packing"
        else:
            kategori = "Lainnya"

        fill = GREEN if idx % 2 == 0 else WHITE

        values = [
            idx,
            name,
            qty,
            kategori,
            "Ambil dari stok gudang",
        ]

        write_row(ws, row, values, fill=fill, bold_cols=[2, 3], center_cols=[1, 3])

        row += 1

    total_qty = sum(rekap.values())
    add_total_row(ws, row, ["", "TOTAL", total_qty, "", ""])


# =========================================================
# SHEET 4: DATA ORDER CUSTOMER
# =========================================================

def write_data_order_customer(wb, data):
    ws = wb.create_sheet("DATA ORDER CUSTOMER")
    ws.sheet_tab_color = "F97316"

    set_header(ws, [
        ("No", 6),
        ("Akun", 10),
        ("Waktu", 12),
        ("Platform", 16),
        ("No. Resi", 24),
        ("No. Pesanan", 22),
        ("Nama Pembeli", 24),
        ("Alamat", 55),
        ("Produk Dibeli", 65),
        ("Total Item", 12),
        ("Pembayaran", 16),
        ("Kurir", 16),
        ("Layanan", 14),
    ])

    row = 2

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

        write_row(
            ws,
            row,
            values,
            fill=fill,
            bold_cols=[9, 10],
            center_cols=[1, 2, 3, 10, 11, 13],
        )

        row += 1


# =========================================================
# SHEET 5: ORDER INSTANT
# =========================================================

def write_order_instant(wb, data):
    ws = wb.create_sheet("ORDER INSTANT")
    ws.sheet_tab_color = "F5A623"

    set_header(ws, [
        ("No", 6),
        ("Akun", 10),
        ("Waktu", 12),
        ("Platform", 16),
        ("No. Resi", 24),
        ("Kode Pengambilan", 22),
        ("Nama Pembeli", 24),
        ("Produk", 60),
        ("Kurir", 16),
        ("Layanan", 16),
        ("Catatan", 28),
    ])

    row = 2
    found = False

    for idx, order in enumerate(data["grouped_orders"], 1):
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
            row - 1,
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

        write_row(
            ws,
            row,
            values,
            fill=YELLOW,
            bold_cols=[6, 8],
            center_cols=[1, 2, 3, 6, 10],
        )

        row += 1

    if not found:
        ws.merge_cells("A2:K2")
        cell = ws["A2"]
        cell.value = "Tidak ada order instant / kode pengambilan yang terbaca di batch ini."
        cell.alignment = LEFT
        cell.font = Font(italic=True, size=10)
        cell.fill = WHITE


# =========================================================
# SHEET 6: AUDIT DATA
# =========================================================

def write_audit_data(wb, data):
    ws = wb.create_sheet("AUDIT DATA")
    ws.sheet_tab_color = "6B7280"

    set_header(ws, [
        ("No", 6),
        ("Akun", 10),
        ("Waktu", 12),
        ("Platform", 16),
        ("Kurir", 16),
        ("Layanan", 14),
        ("No. Resi", 24),
        ("No. Pesanan", 22),
        ("Nama Produk Parser", 38),
        ("Nama Produk Asli", 38),
        ("SKU Parser", 30),
        ("Variasi", 20),
        ("Qty Parser", 12),
        ("Hasil Terjemahan Master SKU", 65),
        ("Nama Pembeli", 24),
        ("Alamat", 55),
        ("Source File", 30),
        ("Status", 18),
        ("Catatan", 35),
    ])

    row = 2

    for idx, r in enumerate(data["expanded_rows"], 1):
        fill = RED_LIGHT if r["_status"] != "OK" else (GREY if idx % 2 == 0 else WHITE)

        translated = combine_components(r["_components"], packing_only=True)

        values = [
            idx,
            r["_akun"],
            r["_waktu"],
            r["_platform"],
            clean(safe(r, "courier")),
            clean(safe(r, "layanan")),
            clean(safe(r, "no_resi")),
            clean(safe(r, "no_pesanan")),
            clean(safe(r, "nama_produk")),
            clean(safe(r, "nama_produk_asli")),
            clean(safe(r, "sku")),
            clean(safe(r, "variasi")),
            to_int(safe(r, "qty", 0)),
            translated,
            clean(safe(r, "nama_penerima")),
            clean(safe(r, "alamat")),
            clean(safe(r, "source_file")),
            r["_status"],
            r["_reason"],
        ]

        write_row(
            ws,
            row,
            values,
            fill=fill,
            bold_cols=[14, 18],
            center_cols=[1, 2, 3, 6, 13, 18],
        )

        row += 1


# =========================================================
# SHEET 7: PERLU TINDAKAN
# =========================================================

def write_perlu_tindakan(wb, data):
    ws = wb.create_sheet("PERLU TINDAKAN")
    ws.sheet_tab_color = "DC2626"

    set_header(ws, [
        ("No", 6),
        ("Akun", 10),
        ("Waktu", 12),
        ("Platform", 16),
        ("No. Resi", 24),
        ("Nama Produk", 42),
        ("SKU", 32),
        ("Qty", 10),
        ("Masalah", 38),
        ("Saran Tindakan", 45),
        ("Source File", 30),
    ])

    perlu = data["perlu_tindakan"]

    if not perlu:
        ws.merge_cells("A2:K2")
        cell = ws["A2"]
        cell.value = "Tidak ada data yang perlu tindakan. Semua SKU/produk berhasil diterjemahkan oleh Master SKU."
        cell.fill = GREEN
        cell.font = Font(bold=True, size=11)
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

        write_row(
            ws,
            row,
            values,
            fill=RED_LIGHT,
            bold_cols=[6, 7, 9],
            center_cols=[1, 2, 3, 8],
        )

        row += 1


# =========================================================
# MAIN WRITER
# =========================================================

def write_excel_multi(rows, output_path):
    wb = Workbook()

    if rows is None:
        rows = []

    data = prepare_data(rows)

    write_ringkasan_operasional(wb, data)
    write_daftar_siap_packing(wb, data)
    write_rekap_kebutuhan_barang(wb, data)
    write_data_order_customer(wb, data)
    write_order_instant(wb, data)
    write_audit_data(wb, data)
    write_perlu_tindakan(wb, data)

    wb.save(output_path)
    return output_path


# =========================================================
# COMPATIBILITY UNTUK APP.PY
# =========================================================

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
