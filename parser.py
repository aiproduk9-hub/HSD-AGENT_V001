import re
import pdfplumber


# =========================
# MASTER SKU LOOKUP
# =========================

def _build_sku_lookup():
    m = {}

    def add(keys, nama):
        for k in keys:
            m[_norm_key(k)] = nama

    add([
        "BG-100GR-1-BOTOL", "BG-100GR-1-BOTOL-3CM", "BG-100GR-1-BOTOL-BG-SKU",
        "BG-100GR-1-BOTOL-BG-SKU1", "BG-100GR-1-BOTOL-BH", "BG-100GR-1-BOTOL-BH-SKU",
        "BG-100GR-1 BOTOL-BH-SKU", "BG-100GR-1 BOTOL-BG-SKU", "BG-100GR-1- BOTOL-BG-SKU",
        "BG-100GR-1- BOTOL-BH", "BG-100GR-SKU", "BG-100GR-3CM-SKU", "BG-100GR-1BOTOL",
        "BG-100GR-1-BOTOL-BG-SKU-1", "BG-100GR-1-BOTOL-BH-SKU-1", "BLACKGARLIC-100GR-1-BOTOL",
    ], "Black Garlic 100gr")

    add(["BG-100GR-2-BOTOL", "BG-100GR-2-BOTOL-3CM", "BG-100GR-1-BOTOL-BG-SKU2", "BG-100GR-1-BOTOL-BH-SKU2"], "Black Garlic 100gr x2")
    add(["BG-100GR-3-BOTOL", "BG-100GR-3-BOTOL-3CM", "BG-100GR-1-BOTOL-BG-SKU3", "BG-100GR-1-BOTOL-BH-SKU3"], "Black Garlic 100gr x3")
    add(["BG-100GR-4-BOTOL", "BG-100GR-1-BOTOL-BG-SKU4", "BG-100GR-1-BOTOL-BH-SKU4"], "Black Garlic 100gr x4")
    add(["BG-100GR-5-BOTOL", "BG-100GR-1-BOTOL-BG-SKU5", "BG-100GR-1-BOTOL-BH-SKU5"], "Black Garlic 100gr x5")

    add([
        "BG-220GR-1-BOTOL", "BG-220GR-1-BOTOL-BG-SKU", "BG-220GR-1-BOTOL-BG-SKU1",
        "BG-220GR-1-BOTOL-BH", "BG-220GR-1-BOTOL-BH-SKU", "BG-220GR-1 BOTOL-BH-SKU",
        "BG-220GR-1 BOTOL-BG-SKU", "BG-220GR-1- BOTOL-BG-SKU", "BG-220GR-1- BOTOL-BH",
        "BG-220GR-SKU", "BG-220GR-SK U", "BG-220GR-1-BOTOL-V2", "BG-220GR-1BOTOL",
        "BLACKGARLIC-220GR-1-BOTOL",
    ], "Black Garlic 220gr")

    add(["BG-220GR-2-BOTOL", "BG-220GR-1-BOTOL-BG-SKU2", "BG-220GR-1-BOTOL-BH-SKU2"], "Black Garlic 220gr x2")
    add(["BG-220GR-3-BOTOL", "BG-220GR-1-BOTOL-BG-SKU3"], "Black Garlic 220gr x3")
    add(["BG-220GR-4-BOTOL", "BG-220GR-1-BOTOL-BG-SKU4"], "Black Garlic 220gr x4")
    add(["BG-220GR-5-BOTOL", "BG-220GR-1-BOTOL-BG-SKU5"], "Black Garlic 220gr x5")

    add([
        "BG-500GR-1-BOTOL", "BG-500GR-1-BOTOL-BH", "BG-500GR-1-BOTOL-BG-SKU",
        "BG-500GR-1-BOTOL-BG-SKU1", "BG-500GR-1-BOTOL-BH-SKU", "BG-500GR-1 BOTOL-BH-SKU",
        "BG-500GR-SKU", "BG-500GR-1- BOTOL", "BG-500GR-1 BOTOL", "BG-500GR-1BOTOL",
        "BG-500GR1-BOTOL", "BLACKGARLIC-500GR-1-BOTOL",
    ], "Black Garlic 500gr")

    add(["BG-500GR-2-BOTOL", "BG-500GR-1-BOTOL-BG-SKU2"], "Black Garlic 500gr x2")
    add(["BG-500GR-3-BOTOL", "BG-500GR-1-BOTOL-BG-SKU3"], "Black Garlic 500gr x3")
    add(["BG-500GR-4-BOTOL", "BG-500GR-1-BOTOL-BG-SKU4"], "Black Garlic 500gr x4")
    add(["BG-500GR-5-BOTOL", "BG-500GR-1-BOTOL-BG-SKU5"], "Black Garlic 500gr x5")

    add([
        "BLACKGARLIC-HONAN-HSD-84G-3PCS", "BG-84GR-CLOVER",
        "BG-84GR-1-CLOVER", "BG-84GR-1-BOTOL",
    ], "Black Garlic 84gr")

    add(["BG-3IN1", "BG-3-IN-1"], "Paket 3in1")
    add(["BG-220GR-2-BOTOL-HAMPERS-IMLEK"], "Hampers Imlek")
    add(["BG-220GR-2-BOTOL-HAMPERS-LEBARAN"], "Hampers Lebaran")
    add(["BG-220GR-2-BOTOL-HAMPERS-NATAL"], "Hampers Natal")
    add(["BG-220GR-2-BOTOL-HAMPERS-HSD"], "Hampers HSD")
    add(["BOX-HAMPERS-IMLEK"], "Box Hampers")
    add(["BOX-HAMPERS-HSD"], "Box Hampers")
    add(["BOX-HAMPERS-NATAL"], "Box Hampers")
    add(["BOX-HAMPERS-LEBARAN"], "Box Hampers")

    add([
        "BG-DRINK-ORI-1-BOTOL-PROMO", "BG-DRINK-PROMO-1-BTL-ORI", "BG-PROMO-ORI-1-BOTOL",
        "BG-DRINK-ORIGINAL", "BG-DRINK-ORI", "BG-DRINK-PROMO", "BG-DRINK-PROMO-ORI",
        "BG-Drink-Original", "BG-Drink-Original-FS", "BLACKGARLIC-DRINK-ORIGINAL",
        "BLACKGARLIC-DRINK-ORIGINAL-FS", "BG-DRINK-ORIGINAL-1-BOTOL",
    ], "BG Drink Original")

    add(["BG-DRINK-ORI-2-BOTOL-PROMO", "BG-DRINK-PROMO-2-BTL-ORI", "BG-DRINK-ORIGINAL-2-BOTOL"], "BG Drink Original x2")
    add(["BG-DRINK-PROMO-4-BTL-ORI", "BG-DRINK-ORIGINAL-4-BOTOL"], "BG Drink Original x4")
    add(["BG-DRINK-ORIGINAL-7-BOTOL", "BG-Drink-Original-7-Botol"], "BG Drink Original 7 Botol")

    add([
        "BG-DRINK-PEACH", "BG-Drink-Peach", "BG-Drink-Peach-FS",
        "BLACKGARLIC-DRINK-PEACH", "BLACKGARLIC-DRINK-PEACH-FS", "BG-DRINK-PEACH-1-BOTOL",
    ], "BG Drink Peach")

    add(["BG-DRINK-PEACH-7-BOTOL", "BG-Drink-Peach-7-Botol"], "BG Drink Peach 7 Botol")
    add(["BG-DRINK-MIX-7-BOTOL", "BG-Drink-Mix-7-Botol"], "BG Drink Mix 7 Botol")

    add([
        "BGH-MULTI-FLORAL-1-BOTOL", "BGH-MULTI-FLORAL", "MADU-MULTI-FLORAL",
        "BGH-MULTI", "BGH-MULTI-1-BOTOL",
        "BGH-MULTI-FL", "BGH-MULTI-FL-1-BOTOL",
        "BGH-MULTI-FLORA", "BGH-MULTI-FLORA-1-BOTOL",
    ], "BG Madu Multi Floral")

    add([
        "BGH-BUNGA-KURMA-1-BOTOL", "BGH-BUNGA-KURMA", "MADU-BUNGA-KURMA",
        "BGH-BUNGA", "BGH-BUNGA-1-BOTOL",
        "BGH-BNG-KURMA", "BGH-BNG-KURMA-1-BOTOL",
        "BGH-KURMA", "BGH-KURMA-1-BOTOL",
    ], "BG Madu Kurma")

    add(["BLACKGARLIC-LANANG-HSD-100G", "BLACKGARLIC-LANANG-BAWANG-HSD-100G", "BLACKGARLIC-HSD-100G"], "Black Garlic 100gr")

    return m


def _norm_key(s):
    if not s:
        return ""
    s = str(s).upper().strip()
    s = s.replace("\uFFFE", "")
    s = re.sub(r"[\x00-\x1F]", "", s)
    s = s.replace("_", "-")
    s = s.replace("\u2013", "-").replace("\u2014", "-")
    s = re.sub(r"\s+", "-", s)
    s = re.sub(r"-+", "-", s)
    s = s.strip("-")
    return s


SKU_LOOKUP = _build_sku_lookup()


def lookup_sku(raw_sku, fallback_nama=""):
    candidates = [raw_sku, fallback_nama, f"{raw_sku} {fallback_nama}"]
    for c in candidates:
        if not c:
            continue
        key = _norm_key(c)
        if key in SKU_LOOKUP:
            return SKU_LOOKUP[key], None

    raw_upper = _norm_key(raw_sku)
    m = re.search(r"(BG-(?:100|220|500)GR-1-BOTOL(?:-BG|-BH)?-?SKU?)(\d+)$", raw_upper)
    if m:
        base = m.group(1).rstrip("-")
        qty_override = int(m.group(2))
        if base in SKU_LOOKUP:
            return SKU_LOOKUP[base], qty_override

    for suffix in ["-BG-SKU1", "-BH-SKU1", "-BG-SKU", "-BH-SKU", "-BG-SKU2", "-BH-SKU2", "-BG", "-BH", "-V2", "-3CM"]:
        if raw_upper.endswith(suffix):
            trimmed = raw_upper[: -len(suffix)]
            if trimmed in SKU_LOOKUP:
                return SKU_LOOKUP[trimmed], None

    return None, None


# =========================
# TEXT UTILITIES
# =========================

def clean(s):
    if s is None:
        return ""
    s = str(s).replace("\uFFFE", " ")
    s = re.sub(r"[\x00-\x08\x0B-\x0C\x0E-\x1F]", " ", s)
    return " ".join(s.split()).strip()


def normalize_sku(s):
    if not s:
        return ""
    s = clean(s).upper()
    s = s.replace("_", "-")
    s = s.replace("\u2013", "-").replace("\u2014", "-")
    s = re.sub(r"\s+", "-", s)
    s = re.sub(r"-+", "-", s)
    s = s.strip("-")
    return s


def to_int(v, default=0):
    try:
        return int(float(str(v).replace(",", "").strip()))
    except Exception:
        return default


def _extract_text_smart(page):
    """
    Ekstrak teks dari halaman PDF.
    Jika halaman terdeteksi rusak karena barcode SPX ECO, coba crop
    area bersih (hilangkan bagian barcode di atas).
    """
    text = page.extract_text() or ""

    # Hitung karakter sampah dari barcode
    garbage = (
        text.count("\uFFFE")
        + text.count("\x00")
        + text.count("\uFFFD")
        + len(re.findall(r"[\x80-\x9F]", text))
    )

    # Halaman bersih: langsung kembalikan
    if garbage <= 15:
        return text

    # Halaman rusak: coba crop dari atas ke bawah bertahap
    try:
        w = float(page.width)
        h = float(page.height)
        for top_frac in [0.18, 0.28, 0.38, 0.48]:
            try:
                cropped_text = page.crop((0, h * top_frac, w, h)).extract_text() or ""
                ct_garbage = (
                    cropped_text.count("\uFFFE")
                    + cropped_text.count("\x00")
                    + cropped_text.count("\uFFFD")
                )
                if len(cropped_text.strip()) > 60 and ct_garbage < garbage:
                    # Pertahankan 300 karakter pertama asli (biasanya ada nomor resi di sana)
                    header = text[:300]
                    return header + "\n" + cropped_text
            except Exception:
                continue
    except Exception:
        pass

    return text


# =========================
# MERGE BROKEN LINES (Perbaikan SKU terpotong antar kolom/baris)
# =========================

def merge_broken_lines(text):
    if not text:
        return ""
    text = text.replace("\uFFFE", " ")

    # ============================================================
    # FIX CROSS-COLUMN SHOPEE 11003 / SPXID (format baru)
    # ============================================================

    # POLA 1: BGH- BUNGA KURMA QTY \n ... BUNGA- \n KURMA-N- \n BOTOL
    text = re.sub(
        r"BGH-\s+BUNGA\s+KURMA\s+(\d+)\n[^\n]*BUNGA-\n[^\n]*KURMA-\d+-\n[^\n]*BOTOL",
        lambda m: f"BGH-BUNGA-KURMA-1-BOTOL __CCQTY{m.group(1)}__",
        text, flags=re.I
    )
    # POLA 2: BGH- QTY \n ... BUNGA- \n KURMA-N- \n BOTOL
    text = re.sub(
        r"BGH-\s+(\d+)\n[^\n]*BUNGA-\n[^\n]*KURMA-\d+-\n[^\n]*BOTOL",
        lambda m: f"BGH-BUNGA-KURMA-1-BOTOL __CCQTY{m.group(1)}__",
        text, flags=re.I
    )
    # POLA 3: BGH- MULTI FLORAL QTY \n ... FLORAL- \n N-BOTOL
    text = re.sub(
        r"BGH-\s+MULTI\s+FLORAL\s+(\d+)\n[^\n]*FLORAL-\n[^\n]*\d+-BOTOL",
        lambda m: f"BGH-MULTI-FLORAL-1-BOTOL __CCQTY{m.group(1)}__",
        text, flags=re.I
    )
    # POLA 3b: BGH- MULTI FL QTY \n ... FL- \n N-BOTOL (versi singkatan)
    text = re.sub(
        r"BGH-\s+MULTI\s+FL\s+(\d+)\n[^\n]*FL-\n[^\n]*\d+-BOTOL",
        lambda m: f"BGH-MULTI-FL-1-BOTOL __CCQTY{m.group(1)}__",
        text, flags=re.I
    )
    # POLA 4: BG- 500GR QTY \n ... 500GR- \n SKU
    text = re.sub(
        r"BG-\s+(100|220|500)\s*GR\s+(\d+)\s*\n[^\n]*(100|220|500)GR-\s*\n\s*SKU",
        lambda m: f"BG-{m.group(1)}GR-SKU __CCQTY{m.group(2)}__",
        text, flags=re.I
    )
    # POLA 5: BG- 500GR QTY \n ... 500GR-N- \n ... BOTOL
    text = re.sub(
        r"BG-\s+(100|220|500)\s*GR\s+(\d+)\s*\n[^\n]*(100|220|500)GR-(\d+)-\s*\n[^\n]*BOTOL",
        lambda m: f"BG-{m.group(1)}GR-{m.group(4)}-BOTOL __CCQTY{m.group(2)}__",
        text, flags=re.I
    )
    # POLA 6: BG-100GR- QTY \n ... SKU (SPXID, SKU prefix sudah sebagian)
    text = re.sub(
        r"(BG-(?:100|220|500)GR)-\s+(?:100|220|500)\s*GR[^\n]*?(\d+)\s*\n[^\n]*\bSKU\b",
        lambda m: f"{m.group(1)}-SKU __CCQTY{m.group(2)}__",
        text, flags=re.I
    )
    # POLA 7: BG- 100GR 3CM QTY \n ... 100GR-N- \n ... BOTOL- (new packaging)
    text = re.sub(
        r"BG-\s+(100|220|500)GR\s+\w+\s+(\d+)\n[^\n]*(100|220|500)GR-(\d+)-\n[^\n]*BOTOL-?",
        lambda m: f"BG-{m.group(1)}GR-{m.group(4)}-BOTOL __CCQTY{m.group(2)}__",
        text, flags=re.I
    )
    # POLA 8: BG- 100 GR 3 CM QTY \n ... 100GR- \n N Hari SKU
    text = re.sub(
        r"BG-\s+(100|220|500)\s*GR[^\n]*?(\d+)\n[^\n]*(100|220|500)GR-\n[^\n]*SKU",
        lambda m: f"BG-{m.group(1)}GR-SKU __CCQTY{m.group(2)}__",
        text, flags=re.I
    )
    # POLA 9: BG- 100 GR 3 cm QTY \n ... 100GR-N- \n BOTOL (multi botol)
    text = re.sub(
        r"BG-\s+(100|220|500)\s*GR[^\n]*?(\d+)\n[^\n]*(100|220|500)GR-(\d+)-\n[^\n]*BOTOL",
        lambda m: f"BG-{m.group(1)}GR-{m.group(4)}-BOTOL __CCQTY{m.group(2)}__",
        text, flags=re.I
    )
    # POLA 10 (HSS): BG-SIZEGR- QTY \n ... N-BOTOL (ada kolom Lokasi di tengah)
    text = re.sub(
        r"(BG-(?:100|220|500)GR)-\s+(\d+)\s*\n[^\n]*(\d+)-BOTOL",
        lambda m: f"{m.group(1)}-{m.group(3)}-BOTOL __CCQTY{m.group(2)}__",
        text, flags=re.I
    )

    # ============================================================
    # FIX STANDAR — SKU terpotong antar baris biasa
    # ============================================================
    fixes = [
        (r"BG-(100|220|500)GR-\s*\n\s*(\d+)-BOTOL", r"BG-\1GR-\2-BOTOL"),
        (r"BG-(100|220|500)GR-1-\s*\n\s*BOTOL", r"BG-\1GR-1-BOTOL"),
        (r"BG-(100|220|500)GR-1\s*\n\s*BOTOL", r"BG-\1GR-1-BOTOL"),
        (r"BG-(100|220|500)GR-SK\s*\n\s*U\b", r"BG-\1GR-SKU"),
        (r"BG-(100|220|500)GR-\s*\n\s*SKU", r"BG-\1GR-SKU"),
        (r"BG-(100|220|500)GR-1-BOTOL-BG-\s*\n\s*SKU", r"BG-\1GR-1-BOTOL-BG-SKU"),
        (r"BG-(100|220|500)GR-1-BOTOL-BH-\s*\n\s*SKU", r"BG-\1GR-1-BOTOL-BH-SKU"),
        (r"BG-(100|220|500)GR-1-BOTOL-\s*\n\s*(BG|BH)-SKU", r"BG-\1GR-1-BOTOL-\2-SKU"),
        (r"BG-(100|220|500)GR-\s*\n\s*1-BOTOL-(BG|BH)-SKU", r"BG-\1GR-1-BOTOL-\2-SKU"),
        (r"BG-DRINK-\s+ORIGINAL", r"BG-DRINK-ORIGINAL"),
        (r"BG-DRINK-\s+PEACH", r"BG-DRINK-PEACH"),
        (r"ORI-1-\s*BOTOL", r"ORI-1-BOTOL"),
        (r"BOTOL-\s*\n\s*PROMO", r"BOTOL-PROMO"),
        (r"ORIGINAL\s*\n\s*ORI-1-\s*BOTOL", r"ORIGINAL ORI-1-BOTOL"),
        (r"BG-DRINK-PROMO-\s*\n\s*(\d+)-BTL-ORI", r"BG-DRINK-PROMO-\1-BTL-ORI"),
        (r"BG-PROMO-ORI-1-\s*\n\s*BOTOL", r"BG-PROMO-ORI-1-BOTOL"),
        (r"BG-DRINK-ORI-1-\s*\n\s*BOTOL-PROMO", r"BG-DRINK-ORI-1-BOTOL-PROMO"),
        (r"BG-(100|220|500)GR\s+-\s*\n\s*(\d+)-BOTOL", r"BG-\1GR-\2-BOTOL"),
        (r"(BG-(?:100|220|500)GR-1)-\s*\n\s*BOTOL", r"\1-BOTOL"),
        (r"(BG-(?:100|220|500)GR-1-B)\s*\n\s*OTOL", r"\1OTOL"),
        (r"(BLACKGARLIC-LANANG(?:-BAWANG)?)-\s*\n\s*(HSD-100G)", r"\1-\2"),
        (r"(BG-)\s*\n\s*((?:100|220|500)GR-\S+)", r"\1\2"),
        (r"(BG-(?:100|220|500)GR-)\s*\n\s*(SKU\b)", r"\1\2"),
    ]
    for pattern, repl in fixes:
        text = re.sub(pattern, repl, text, flags=re.I)
    return text


# =========================
# DETEKSI PLATFORM & KURIR
# =========================

def detect(text):
    t = text.upper()
    if "BLIBLI" in t or "BLIBLI.COM" in t:
        kurir = "J&T Express" if ("JNT" in t or "J&T" in t) else "Unknown"
        return "Blibli", kurir, "CASHLESS"
    if "LAZADA" in t or "LXAD" in t or "LEX" in t:
        return "Lazada", "LEX", "STANDARD"
    if "GOSEND" in t or "GO-SEND" in t:
        return "Shopee", "Gosend", "Instant"
    if "SHOPEE" in t:
        if "SICEPAT" in t:
            return "Shopee", "SiCepat", "REG"
        if "ANTERAJA" in t or re.search(r"\bAAJ\w+", text):
            return "Shopee", "Anteraja", "REG"
        if "SPX" in t or "SHOPEE EXPRESS" in t:
            return "Shopee", "Shopee Express", "ECO"
        if "WAHANA" in t or re.search(r"\bBDO\d+", text):
            return "Shopee", "Wahana", "Reguler"
        return "Shopee", "Unknown", "REG"
    if re.search(r"\bSPXID\d+", text):
        return "Shopee", "Shopee Express", "ECO"
    if re.search(r"\bJX\d{10}\b", text) or "JET.CO.ID" in t or "J&T" in t or "JNT" in t:
        layanan = "EZ"
        for s in ["NDD", "ECO", "EZ", "REG"]:
            if re.search(r"\b" + s + r"\b", t[:600]):
                layanan = s
                break
        return "TikTok", "J&T Express", layanan
    if re.search(r"\bGTL\d{8,12}\b", text):
        layanan = "NDD" if "NDD" in t else "REG"
        return "TikTok", "GTL", layanan
    if re.search(r"\b(0046\d{8,12}|00296\d{7,12})\b", text) or "SICEPAT" in t:
        return "TikTok", "SiCepat", "REG"
    if re.search(r"\bCM\d{10,15}\b", text):
        return "TikTok", "Wahana", "REG"
    return "Unknown", "Unknown", "-"


def get_resi(text):
    patterns = [
        r"No\.\s*Resi\s*[::]?\s*(\d{10,15})",
        r"Resi\s*[::]?\s*(SPXID\d{8,18})",
        r"Air\s*waybill\s*[::]?\s*([A-Z]{2}\d{8,15})",
        r"\b(JX\d{10})\b",
        r"\b(GTL\d{8,12})\b",
        r"\b(SPXID\d{10,18})\b",
        r"\b(LXAD-\d+)\b",
        r"\b(AAJ\w{8,})\b",
        r"\b(BDO\d{6,})\b",
        r"\b(CM\d{10,15})\b",
        r"\b(0046\d{8,12})\b",
        r"\b(00296\d{7,12})\b",
        r"\b(JJ\d{8,15})\b",
        # Shopee Express format 11003xxxxxx — lebih spesifik daripada 11\d{12,14}
        r"\b(110\d{12,13})\b",
        # Fallback Shopee numerik panjang
        r"\b(11\d{12,13})\b",
    ]
    for p in patterns:
        m = re.search(p, text, re.I)
        if m:
            return clean(m.group(1))
    # Fallback ke nomor pesanan hanya jika tidak ada pilihan lain
    m = re.search(r"(?:No\.?\s*Pesanan|Nomor\s*Order)\s*[::]?\s*([A-Z0-9]{8,})", text, re.I)
    if m:
        return clean(m.group(1))
    return ""


def get_pesanan(text):
    patterns = [
        r"No\.?\s*Pesanan\s*[::]?\s*([A-Z0-9\-]+)",
        r"Order\s*ID\s*[::]?\s*([0-9]+)",
        r"TT\s*Order\s*ID\s*[::]?\s*([0-9]+)",
        r"Nomor\s*Order\s*[::]?\s*([0-9]+)",
        r"Nomor\s*Pesanan\s*[::]?\s*([0-9]+)",
    ]
    for p in patterns:
        m = re.search(p, text, re.I)
        if m:
            return clean(m.group(1))
    return ""


def get_kode_pengambilan(text):
    m = re.search(r"Kode\s*Pengambilan\s*[::]?\s*([A-Z0-9]+)", text, re.I)
    if m:
        return clean(m.group(1))
    return ""


def get_pembayaran(text):
    t = text.upper()
    if "NON-COD" in t or "NON COD" in t:
        return "Non-COD"
    if re.search(r"\bCOD\b", t):
        return "COD"
    if "CASHLESS" in t:
        return "Non-COD"
    return "Belum Terbaca"


def get_penerima(text):
    patterns = [
        r"Penerima\s*[:]\s*([^\n(]+)",
        r"Receiver\s+([^\n(]+)",
        r"Ke\(penerima\)\s*([^\n(]+)",
    ]
    for p in patterns:
        m = re.search(p, text, re.I)
        if m:
            name = clean(m.group(1))
            name = re.sub(r"^\+?62[0-9\*\s]+", "", name).strip()
            if 1 < len(name) < 80:
                return name
    return ""


def get_alamat(text):
    lines = [clean(x) for x in text.split("\n") if clean(x)]
    for i, line in enumerate(lines):
        if line.lower().startswith("penerima:"):
            chunk = []
            for j in range(i + 1, min(i + 6, len(lines))):
                lj = lines[j]
                if re.search(r"Nama Produk|Pengirim|Diserahkan|Diantar|STANDARD|Total Qty", lj, re.I):
                    break
                if len(lj) > 5:
                    chunk.append(lj)
            if chunk:
                return " ".join(chunk)
    for i, line in enumerate(lines):
        if re.search(r"Penerima\s*:", line, re.I):
            chunk = []
            for j in range(i + 1, min(i + 8, len(lines))):
                lj = lines[j]
                if re.search(r"Berat|COD|Batas Kirim|No\. Pesanan|# Nama Produk|Pengirim", lj, re.I):
                    break
                if len(lj) > 5 and not re.match(r"^\+?62", lj):
                    chunk.append(lj)
            if chunk:
                return " ".join(chunk)
    capture = False
    chunk = []
    for line in lines:
        if re.search(r"Receiver|Pene\s*rima|Penerima", line, re.I):
            capture = True
            continue
        if capture:
            if re.search(r"Sender|Pengirim|Product Name|TT Order|Order Id|Weight|Estimated|In transit", line, re.I):
                break
            if len(line) > 8 and not re.match(r"^\+?62", line):
                chunk.append(line)
            if len(chunk) >= 5:
                break
    if chunk:
        return " ".join(chunk)
    return ""


# =========================
# PRODUK — LOOKUP & RESOLVE
# =========================

# Mapping multi-botol ke jumlah botol satuan
_MULTI_BOTOL = {
    "Black Garlic 100gr x2": ("Black Garlic 100gr", 2),
    "Black Garlic 100gr x3": ("Black Garlic 100gr", 3),
    "Black Garlic 100gr x4": ("Black Garlic 100gr", 4),
    "Black Garlic 100gr x5": ("Black Garlic 100gr", 5),
    "Black Garlic 220gr x2": ("Black Garlic 220gr", 2),
    "Black Garlic 220gr x3": ("Black Garlic 220gr", 3),
    "Black Garlic 220gr x4": ("Black Garlic 220gr", 4),
    "Black Garlic 220gr x5": ("Black Garlic 220gr", 5),
    "Black Garlic 500gr x2": ("Black Garlic 500gr", 2),
    "Black Garlic 500gr x3": ("Black Garlic 500gr", 3),
    "Black Garlic 500gr x4": ("Black Garlic 500gr", 4),
    "Black Garlic 500gr x5": ("Black Garlic 500gr", 5),
    "BG Drink Original x2": ("BG Drink Original", 2),
    "BG Drink Original x4": ("BG Drink Original", 4),
}


def resolve_nama_produk(sku_raw, nama_produk_raw, qty_raw):
    qty_raw = to_int(qty_raw, 1) or 1
    nama, qty_override = lookup_sku(sku_raw, nama_produk_raw)
    if nama:
        qty_final = qty_override if qty_override else qty_raw
        if nama in _MULTI_BOTOL:
            nama_base, multiplier = _MULTI_BOTOL[nama]
            return nama_base, qty_final * multiplier
        return nama, qty_final

    if nama_produk_raw:
        nama2, qty_override2 = lookup_sku(nama_produk_raw, "")
        if nama2:
            qty_final = qty_override2 if qty_override2 else qty_raw
            if nama2 in _MULTI_BOTOL:
                nama_base, multiplier = _MULTI_BOTOL[nama2]
                return nama_base, qty_final * multiplier
            return nama2, qty_final

    text = normalize_sku(f"{sku_raw} {nama_produk_raw}").upper()

    if "DRINK" in text and "MIX" in text and "7" in text:
        return "BG Drink Mix 7 Botol", qty_raw
    if "DRINK" in text and "PEACH" in text and "7" in text:
        return "BG Drink Peach 7 Botol", qty_raw
    if "DRINK" in text and "ORIGINAL" in text and "7" in text:
        return "BG Drink Original 7 Botol", qty_raw
    if "DRINK" in text and "PEACH" in text:
        return "BG Drink Peach", qty_raw
    if "DRINK" in text and ("ORIGINAL" in text or "-ORI" in text or "ORI-" in text or "PROMO" in text):
        return "BG Drink Original", qty_raw
    if "DRINK" in text:
        return "BG Drink Original", qty_raw
    if "PROMOSI" in text:
        return "BG Drink Original", qty_raw
    if "3IN1" in text or "3-IN-1" in text:
        return "Paket 3in1", qty_raw
    if "HAMPERS" in text and "LEBARAN" in text:
        return "Hampers Lebaran", qty_raw
    if "HAMPERS" in text and "IMLEK" in text:
        return "Hampers Imlek", qty_raw
    if "HAMPERS" in text and "NATAL" in text:
        return "Hampers Natal", qty_raw
    if "HAMPERS" in text:
        return "Hampers HSD", qty_raw
    if "MULTI" in text and ("FLORAL" in text or "FLORA" in text or
                             text.endswith("-FL") or "-FL-" in text or
                             text.endswith("FL")):
        return "BG Madu Multi Floral", qty_raw
    if ("KURMA" in text or "BUNGA" in text) and ("BGH" in text or "MADU" in text or "HONEY" in text):
        return "BG Madu Kurma", qty_raw
    if "KURMA" in text:
        return "BG Madu Kurma", qty_raw
    if "HONEY" in text or "MADU" in text:
        if "MULTI" in text or "FLORAL" in text:
            return "BG Madu Multi Floral", qty_raw
        return "BG Madu Kurma", qty_raw

    m = re.search(r"BG-(100|220|500)GR-(\d+)-BOTOL", text)
    if m:
        size = m.group(1)
        jumlah = int(m.group(2))
        return f"Black Garlic {size}gr", qty_raw * jumlah

    if "500GR" in text or "500-GR" in text:
        return "Black Garlic 500gr", qty_raw
    if "220GR" in text or "220-GR" in text:
        return "Black Garlic 220gr", qty_raw
    if "100GR" in text or "100-GR" in text:
        return "Black Garlic 100gr", qty_raw

    return None, qty_raw


def _extract_qty_from_line(line):
    """
    Ambil angka qty dari baris tabel.
    Strategi: cari angka yang muncul SETELAH token SKU terakhir di baris,
    yang tidak diikuti huruf (bukan satuan seperti 100GR, 220ML).
    Tidak ada batas maksimum — mendukung pesanan grosir berapapun.
    """
    # Cari posisi akhir SKU terakhir di baris
    sku_end = 0
    for m in re.finditer(
        r"(?:BGH|BG|BLACKGARLIC|MADU|BOX)-[A-Z0-9\-]+", line, re.I
    ):
        sku_end = max(sku_end, m.end())

    if sku_end > 0:
        after_sku = line[sku_end:]
        # Cari angka standalone setelah SKU (bukan diikuti huruf = bukan satuan)
        for m in re.finditer(r"\b(\d{1,4})\b", after_sku):
            num = int(m.group(1))
            if num < 1:
                continue
            end_pos = m.end()
            # Pastikan tidak diikuti huruf (misal "100GR" → bukan qty)
            if end_pos >= len(after_sku) or not after_sku[end_pos].isalpha():
                return num

    # Fallback: ambil angka terakhir di baris yang tidak diikuti huruf
    for m in reversed(list(re.finditer(r"\b(\d{1,4})\b", line))):
        num = int(m.group(1))
        end_pos = m.end()
        if end_pos >= len(line) or not line[end_pos].isalpha():
            # Abaikan angka ukuran produk yang berdiri sendiri
            if num not in (84, 100, 220, 500, 1000):
                return num

    return 1


def extract_sku_candidates(text):
    """
    Temukan semua kandidat SKU dalam teks.
    Mencakup semua prefix yang dikenal: BGH-, MADU-, BG-, BLACKGARLIC-, BOX-.
    """
    text = merge_broken_lines(text)
    patterns = [
        r"BGH-[A-Z0-9\-]+",
        r"MADU-[A-Z0-9\-]+",              # Prefix Madu langsung
        r"BLACKGARLIC-[A-Z0-9\-]+",
        r"BG-DRINK-[A-Z0-9\-]+",
        r"BG-PROMO-[A-Z0-9\-]+",
        r"BG-(?:100|220|500)GR-[A-Z0-9\-]+",
        r"BG-(?:100|220|500)GR-SKU",
        r"BG-(?:100|220|500)GR-\d+-BOTOL",
        r"BG-84GR-[A-Z0-9\-]+",
        r"BOX-HAMPERS-[A-Z0-9\-]+",
        r"BG-3IN1",
        r"BG-3-IN-1",                      # Variasi ejaan
    ]
    found = []
    for p in patterns:
        for m in re.finditer(p, text, re.I):
            sku = normalize_sku(m.group(0))
            if sku and sku not in found:
                found.append(sku)
    return found


def build_product_row(platform, courier, layanan, resi, pesanan, penerima, alamat, pembayaran, kode, nama, sku, variasi, qty):
    nama_resolved, qty_final = resolve_nama_produk(sku, nama, qty)
    if nama_resolved is None:
        nama_final = clean(nama) or clean(sku) or "(cek manual)"
        sku_final = normalize_sku(sku)
    else:
        nama_final = nama_resolved
        sku_final = normalize_sku(sku)

    return {
        "platform": platform,
        "courier": courier,
        "layanan": layanan,
        "resi": resi,
        "no_resi": resi,
        "sku_raw": normalize_sku(sku),
        "perlu_cek": nama_final == "(cek manual)",
        "alasan_cek": "SKU tidak dikenali" if nama_final == "(cek manual)" else "",
        "no_pesanan": pesanan,
        "nama_produk": nama_final,
        "nama_produk_asli": clean(nama),
        "sku": sku_final,
        "variasi": clean(variasi),
        "qty": to_int(qty_final, 1) or 1,
        "nama_penerima": penerima,
        "alamat": alamat,
        "pembayaran": pembayaran,
        "kode_pengambilan": kode,
    }


# =========================
# PARSER TABEL PER PLATFORM
# =========================

def parse_tiktok_table(text):
    products = []
    t = merge_broken_lines(text)

    # Header check lebih fleksibel: cukup salah satu header yang ada
    has_tt_header = (
        ("Product Name" in t or "Nama Produk" in t)
        and ("Seller SKU" in t or "SKU" in t)
    )
    if not has_tt_header:
        return products

    m = re.search(
        r"(?:Product Name|Nama Produk)[\s\S]*?(?:Seller\s+)?SKU\s+Qty\s*\n([\s\S]+?)(?:Qty Total:|Order ID:|$)",
        t, re.I,
    )
    if not m:
        # Fallback: cari blok setelah header
        m = re.search(
            r"Product Name\s+SKU\s+Seller SKU\s+Qty\s*\n([\s\S]+?)(?:Qty Total:|Order ID:|$)",
            t, re.I,
        )
    if not m:
        return products

    block = m.group(1)
    lines = [clean(x) for x in block.split("\n") if clean(x)]
    current_name = []

    for line in lines:
        # Pola lengkap: nama variasi SKU qty
        mm = re.search(
            r"^(.*?)\s*(Default|ORI[\s\-]PROMO[\s\-]\d|ORIGINAL|PEACH|100[\s\-]GR|220[\s\-]GR|500[\s\-]GR)?\s+"
            r"((?:BGH|BG|BLACKGARLIC|BOX|MADU)-[A-Z0-9\-]+)\s+(\d{1,4})$",
            line, re.I,
        )
        if mm:
            prefix  = clean(mm.group(1))
            variasi = clean(mm.group(2))
            sku     = normalize_sku(mm.group(3))
            qty     = to_int(mm.group(4), 1)
            name    = " ".join(current_name).strip()
            if prefix and len(prefix) > 3 and prefix.upper() not in ["DEFAULT"]:
                name = (name + " " + prefix).strip()
            products.append({"nama_produk": name or prefix or sku, "sku": sku, "variasi": variasi, "qty": qty})
            current_name = []
            continue

        # Pola ringkas: SKU qty
        mm2 = re.search(
            r"^((?:BGH|BG|BLACKGARLIC|BOX|MADU)-[A-Z0-9\-]+)\s+(\d{1,4})$",
            line, re.I,
        )
        if mm2:
            sku  = normalize_sku(mm2.group(1))
            qty  = to_int(mm2.group(2), 1)
            name = " ".join(current_name).strip()
            products.append({"nama_produk": name or sku, "sku": sku, "variasi": "", "qty": qty})
            current_name = []
            continue

        if not re.match(r"^\d+$", line) and len(line) > 2:
            current_name.append(line)

    if products:
        return products

    # Fallback: ambil semua SKU dari teks
    skus = extract_sku_candidates(t)
    for sku in skus:
        products.append({"nama_produk": sku, "sku": sku, "variasi": "", "qty": 1})
    return products


def parse_shopee_table(text):
    """
    Parser tabel Shopee.
    Perbaikan utama:
    - Proses baris per baris (bukan blok penuh), sehingga setiap SKU
      mendapat qty-nya sendiri (fix multi-SKU bug).
    - Gunakan _extract_qty_from_line yang tidak ada batas cap 20.
    - CCQTY dicari per baris, bukan seluruh blok.
    """
    products = []
    t = merge_broken_lines(text)

    has_header = (
        "# Nama Produk" in t
        or "Nama Produk SKU" in t
        or re.search(r"#\s*\n\s*Nama Produk", t) is not None
        or re.search(r"Nama Produk\s*\n\s*SKU", t) is not None
    )
    if not has_header:
        return products

    # Ekstrak blok tabel
    m = re.search(
        r"(?:#\s*)?Nama Produk\s+SKU[\s\S]*?(?:Variasi|Lokasi)?\s*Qty\s*\n([\s\S]+?)"
        r"(?:Pesan:|Pengirim:|CASHLESS|$)",
        t, re.I,
    )
    block = m.group(1) if m else t
    lines = [clean(x) for x in block.split("\n") if clean(x)]

    name_buffer = []

    for line in lines:
        # Hentikan di bagian footer resi
        if re.search(r"(?:Pengirim|CASHLESS|Pesan:|Total Qty|Batas Kirim)", line, re.I):
            break

        # --- Cek CCQTY per baris (format Shopee cross-column 11003/SPXID) ---
        ccqty_m = re.search(r"__CCQTY(\d+)__", line)
        qty_cc = to_int(ccqty_m.group(1), 1) if ccqty_m else None
        line_clean = re.sub(r"__CCQTY\d+__", "", line).strip()

        skus_in_line = extract_sku_candidates(line_clean)

        if skus_in_line:
            if qty_cc is not None:
                qty = qty_cc
            else:
                qty = _extract_qty_from_line(line_clean)

            name = " ".join(name_buffer).strip()
            for sku in skus_in_line:
                products.append({
                    "nama_produk": name or sku,
                    "sku": sku,
                    "variasi": "",
                    "qty": qty,
                })
            name_buffer = []
        else:
            # Akumulasi sebagai nama produk
            if line_clean and not re.match(r"^\d+$", line_clean):
                if not re.search(
                    r"(?:Pengirim|CASHLESS|Pesan:|Total|Batas|Penerima|Variasi|Lokasi)",
                    line_clean, re.I
                ):
                    name_buffer.append(line_clean)

    return products


def parse_lazada_table(text):
    products = []
    t = merge_broken_lines(text)
    if "LAZADA" not in t.upper() and "LXAD" not in t.upper():
        return products
    m = re.search(
        r"Nama Produk\s+Qty\s+SKU\s+Item Variant\s*\n([\s\S]+?)(?:Pengirim:|Penerima:|$)",
        t, re.I,
    )
    if m:
        block = clean(m.group(1))
        skus = extract_sku_candidates(block)
        for sku in skus:
            qty = 1
            mq = re.search(r"\s(\d{1,3})\s+" + re.escape(sku), normalize_sku(block), re.I)
            if mq:
                qty = to_int(mq.group(1), 1)
            products.append({"nama_produk": block, "sku": sku, "variasi": "", "qty": qty})
    return products


def parse_blibli_table(text):
    products = []
    t = merge_broken_lines(text)
    if "BLIBLI" not in t.upper():
        return products
    skus = extract_sku_candidates(t)
    for sku in skus:
        qty = _extract_qty_from_line(
            t[max(0, t.upper().find(sku.upper()) - 5): t.upper().find(sku.upper()) + len(sku) + 60]
        )
        products.append({"nama_produk": sku, "sku": sku, "variasi": "", "qty": qty})
    return products


def parse_generic_products(text):
    """Fallback universal: ambil semua SKU dari teks apapun."""
    products = []
    t = merge_broken_lines(text)
    skus = extract_sku_candidates(t)
    for sku in skus:
        # Cari qty di sekitar SKU (100 karakter setelahnya)
        token_text = normalize_sku(t)
        pos = token_text.find(sku)
        qty = 1
        if pos >= 0:
            region = t[pos: pos + len(sku) + 80]
            qty = _extract_qty_from_line(region)
        products.append({"nama_produk": sku, "sku": sku, "variasi": "", "qty": qty})
    return products


# =========================
# PARSER HALAMAN
# =========================

def parse_page(text):
    text = merge_broken_lines(text or "")
    platform, courier, layanan = detect(text)
    resi = get_resi(text)
    if not resi:
        return []
    pesanan = get_pesanan(text)
    penerima = get_penerima(text)
    alamat = get_alamat(text)
    pembayaran = get_pembayaran(text)
    kode = get_kode_pengambilan(text)

    product_candidates = parse_shopee_table(text)
    if not product_candidates:
        product_candidates = parse_lazada_table(text)
    if not product_candidates:
        product_candidates = parse_blibli_table(text)
    if not product_candidates:
        product_candidates = parse_tiktok_table(text)
    if not product_candidates:
        product_candidates = parse_generic_products(text)
    if not product_candidates:
        product_candidates = [{"nama_produk": "(cek manual)", "sku": "(cek manual)", "variasi": "", "qty": 1}]

    rows = []
    for p in product_candidates:
        rows.append(build_product_row(
            platform=platform, courier=courier, layanan=layanan,
            resi=resi, pesanan=pesanan, penerima=penerima, alamat=alamat,
            pembayaran=pembayaran, kode=kode,
            nama=p.get("nama_produk", ""), sku=p.get("sku", ""),
            variasi=p.get("variasi", ""), qty=p.get("qty", 1),
        ))
    return rows


def _has_sku_in_text(text):
    """Deteksi apakah halaman mengandung SKU produk (untuk page merge logic)."""
    if re.search(
        r"BG-(?:100|220|500)GR|BGH-|BG-DRINK|BG-3IN1|BOX-HAMPERS|BLACKGARLIC-|BG-84GR|MADU-",
        text, re.I
    ):
        return True
    # Shopee cross-column: BG- dengan spasi sebelum ukuran
    if re.search(r"BG-\s+(?:100|220|500)\s*GR", text, re.I):
        return True
    return False


# =========================
# PROSES PDF UTAMA
# =========================

def process_pdf(pdf_path, progress_callback=None):
    """
    Baca seluruh PDF dan kembalikan daftar baris produk.

    Perbaikan:
    - Gunakan _extract_text_smart untuk halaman barcode-rusak (SPX ECO).
    - Page merge hingga 5 halaman ke depan (resi pecah banyak halaman).
    - Dedup aman untuk baris dengan resi kosong (pakai index halaman).
    """
    all_rows = []
    errors = []
    seen = set()

    with pdfplumber.open(pdf_path) as pdf:
        total = len(pdf.pages)
        i = 0
        while i < total:
            try:
                text = _extract_text_smart(pdf.pages[i])
                resi = get_resi(text)

                # Page merge: kalau halaman ini punya resi tapi belum ada SKU,
                # baca halaman berikutnya (sampai 5 halaman ke depan)
                if resi and not _has_sku_in_text(text):
                    pages_merged = 0
                    for lookahead in range(1, 6):
                        if i + lookahead >= total:
                            break
                        next_text = _extract_text_smart(pdf.pages[i + lookahead])
                        text = text + "\n" + next_text
                        pages_merged = lookahead
                        if progress_callback:
                            progress_callback(i + lookahead, total)
                        # Stop merge kalau halaman berikutnya sudah punya resi baru
                        # (artinya ini order baru, bukan lanjutan)
                        next_resi = get_resi(next_text)
                        if _has_sku_in_text(next_text):
                            i += lookahead
                            break
                        if next_resi and next_resi != resi:
                            # Halaman berikutnya adalah order baru — jangan merge
                            i += lookahead - 1
                            break
                    _ = pages_merged  # suppress unused warning

                page_rows = parse_page(text)
                for r in page_rows:
                    r_resi = r.get("no_resi", "")
                    r_sku  = r.get("sku", "")
                    r_nama = r.get("nama_produk", "")
                    r_qty  = str(r.get("qty", ""))

                    if r_resi:
                        # Resi ada: dedup normal
                        key = (r_resi, r_sku, r_nama, r_qty)
                    else:
                        # Resi kosong: pakai index halaman agar tidak over-dedup
                        key = (f"__page_{i}__", r_sku, r_nama, r_qty)

                    if key not in seen:
                        seen.add(key)
                        all_rows.append(r)

            except Exception as e:
                errors.append(f"Halaman {i + 1}: {e}")

            if progress_callback:
                progress_callback(i + 1, total)
            i += 1

    return all_rows, errors
