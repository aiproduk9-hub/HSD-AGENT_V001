import re
import pdfplumber


# =========================
# BASIC HELPER
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
    s = s.replace("\u2013", "-")
    s = s.replace("\u2014", "-")
    s = re.sub(r"\s+", "-", s)
    s = re.sub(r"-+", "-", s)
    s = s.strip("-")

    # Fix typo OCR / pecahan umum
    s = s.replace("SK-U", "SKU")
    s = s.replace("S-KU", "SKU")
    s = s.replace("BOTOL-BGSKU", "BOTOL-BG-SKU")
    s = s.replace("BOTOL-BHSKU", "BOTOL-BH-SKU")
    s = s.replace("BOTOL-BG SKU", "BOTOL-BG-SKU")
    s = s.replace("BOTOL-BH SKU", "BOTOL-BH-SKU")

    return s


def merge_broken_lines(text):
    """
    Gabung pecahan SKU yang terpotong antar baris di PDF.
    Contoh:
      BG-500GR-     →  BG-500GR-1-BOTOL
      1-BOTOL
    """
    if not text:
        return ""

    text = text.replace("\uFFFE", " ")

    fixes = [
        # BG-xxxGR- lanjut angka-BOTOL
        (r"BG-(100|220|500)GR-\s*\n\s*(\d+)-BOTOL", r"BG-\1GR-\2-BOTOL"),
        # BG-xxxGR-1- lanjut BOTOL
        (r"BG-(100|220|500)GR-(\d+)-\s*\n\s*BOTOL", r"BG-\1GR-\2-BOTOL"),
        # BG-xxxGR-1 lanjut BOTOL (tanpa strip)
        (r"BG-(100|220|500)GR-(\d+)\s*\n\s*BOTOL", r"BG-\1GR-\2-BOTOL"),
        # BG-xxxGR- lanjut SKU
        (r"BG-(100|220|500)GR-\s*\n\s*SKU", r"BG-\1GR-SKU"),
        # BG-xxxGR-1-BOTOL- lanjut BG-SKU atau BH-SKU
        (r"(BG-(?:100|220|500)GR-\d+-BOTOL)-\s*\n\s*(BG-SKU|BH-SKU)", r"\1-\2"),
        # BG-xxxGR-1-BOTOL- lanjut BG atau BH (tanpa SKU)
        (r"(BG-(?:100|220|500)GR-\d+-BOTOL)-\s*\n\s*(BG|BH)\b", r"\1-\2"),
        # DRINK promo pecah
        (r"BG-DRINK-PROMO-\s*\n\s*(\d+)-BTL-ORI", r"BG-DRINK-PROMO-\1-BTL-ORI"),
        (r"BG-DRINK-ORI-(\d+)-\s*\n\s*BOTOL-PROMO", r"BG-DRINK-ORI-\1-BOTOL-PROMO"),
        (r"BG-PROMO-ORI-(\d+)-\s*\n\s*BOTOL", r"BG-PROMO-ORI-\1-BOTOL"),
        # Drink 7 botol pecah
        (r"BG-DRINK-(PEACH|ORIGINAL|MIX)-\s*\n\s*7-BOTOL", r"BG-DRINK-\1-7-BOTOL"),
        # BGH pecah
        (r"BGH-(MULTI|BUNGA)-\s*\n\s*(FLORAL|KURMA)-(\d+)-BOTOL", r"BGH-\1-\2-\3-BOTOL"),
    ]

    for pattern, repl in fixes:
        text = re.sub(pattern, repl, text, flags=re.I)

    return text


def to_int(v, default=0):
    try:
        return int(float(str(v).replace(",", "").strip()))
    except Exception:
        return default


# =========================
# SKU → QTY BUNDLE RESOLVER
# Ini jantung perbaikan: SKU sendiri yang tentukan qty barang
# =========================

# Lookup dictionary lengkap: SKU (normalized) → (nama_produk, qty_per_unit)
# qty_per_unit = jumlah unit fisik barang yang dikirim per 1 SKU ini
SKU_LOOKUP = {
    # ── DRINK ORIGINAL (1 botol) ─────────────────────────────────────
    "BG-DRINK-ORI-1-BOTOL-PROMO":          ("BG Drink Original", 1),
    "BG-DRINK-PROMO-1-BTL-ORI":            ("BG Drink Original", 1),
    "BG-DRINK-PROMO-2-BTL-ORI":            ("BG Drink Original", 2),
    "BG-DRINK-PROMO-4-BTL-ORI":            ("BG Drink Original", 4),
    "BG-DRINK-ORI-2-BOTOL-PROMO":          ("BG Drink Original", 2),
    "BG-DRINK-PEACH":                      ("BG Drink Peach", 1),
    "BG-DRINK-ORIGINAL":                   ("BG Drink Original", 1),
    "BG-DRINK-ORIGINAL-FS":                ("BG Drink Original", 1),
    "BG-DRINK-PEACH-FS":                   ("BG Drink Peach", 1),
    "BLACKGARLIC-DRINK-ORIGINAL":          ("BG Drink Original", 1),
    "BLACKGARLIC-DRINK-PEACH":             ("BG Drink Peach", 1),
    "BLACKGARLIC-DRINK-ORIGINAL-FS":       ("BG Drink Original", 1),
    "BLACKGARLIC-DRINK-PEACH-FS":          ("BG Drink Peach", 1),
    # alias tanpa FS
    "BG-DRINK-ORI":                        ("BG Drink Original", 1),
    "BG-DRINK-PEACH-1":                    ("BG Drink Peach", 1),

    # ── DRINK 7 BOTOL ────────────────────────────────────────────────
    "BG-DRINK-PEACH-7-BOTOL":              ("BG Drink Peach", 7),
    "BG-DRINK-ORIGINAL-7-BOTOL":           ("BG Drink Original", 7),
    "BG-DRINK-MIX-7-BOTOL":               ("BG Drink Mix (4 Peach + 3 Original)", 7),

    # ── MADU ─────────────────────────────────────────────────────────
    "BGH-MULTI-FLORAL-1-BOTOL":            ("Madu Multi Floral", 1),
    "BGH-BUNGA-KURMA-1-BOTOL":             ("Madu Bunga Kurma", 1),

    # ── BLACK GARLIC 84GR ────────────────────────────────────────────
    "BLACKGARLIC-HONAN-HSD-84G-3PCS":      ("Black Garlic 84gr", 1),
    "BG-84GR-CLOVER":                      ("Black Garlic 84gr", 1),
    "BG-84GR-1-CLOVER":                    ("Black Garlic 84gr", 1),

    # ── BLACK GARLIC 100GR ───────────────────────────────────────────
    "BG-100GR-1-BOTOL":                    ("Black Garlic 100gr", 1),
    "BG-100GR-1-BOTOL-3CM":               ("Black Garlic 100gr", 1),
    "BG-100GR-1-BOTOL-BG-SKU":            ("Black Garlic 100gr", 1),
    "BG-100GR-1-BOTOL-BH":               ("Black Garlic 100gr", 1),
    "BG-100GR-1-BOTOL-BH-SKU":           ("Black Garlic 100gr", 1),
    "BG-100GR-1-BOTOL-BG":              ("Black Garlic 100gr", 1),
    "BG-100GR-SKU":                       ("Black Garlic 100gr", 1),
    "BG-100GR-2-BOTOL":                   ("Black Garlic 100gr", 2),
    "BG-100GR-2-BOTOL-BG-SKU":           ("Black Garlic 100gr", 2),
    "BG-100GR-2-BOTOL-BH-SKU":           ("Black Garlic 100gr", 2),
    "BG-100GR-3-BOTOL":                   ("Black Garlic 100gr", 3),
    "BG-100GR-3-BOTOL-BG-SKU":           ("Black Garlic 100gr", 3),

    # ── BLACK GARLIC 220GR ───────────────────────────────────────────
    "BG-220GR-1-BOTOL":                   ("Black Garlic 220gr", 1),
    "BG-220GR-1-BOTOL-BG-SKU":           ("Black Garlic 220gr", 1),
    "BG-220GR-1-BOTOL-BH":              ("Black Garlic 220gr", 1),
    "BG-220GR-1-BOTOL-BH-SKU":          ("Black Garlic 220gr", 1),
    "BG-220GR-1-BOTOL-BG":             ("Black Garlic 220gr", 1),
    "BG-220GR-SKU":                      ("Black Garlic 220gr", 1),
    "BG-220GR-2-BOTOL":                  ("Black Garlic 220gr", 2),
    "BG-220GR-2-BOTOL-BG-SKU":          ("Black Garlic 220gr", 2),
    "BG-220GR-2-BOTOL-BH-SKU":          ("Black Garlic 220gr", 2),
    "BG-220GR-3-BOTOL":                  ("Black Garlic 220gr", 3),
    "BG-220GR-3-BOTOL-BG-SKU":          ("Black Garlic 220gr", 3),

    # ── BLACK GARLIC 500GR ───────────────────────────────────────────
    "BG-500GR-1-BOTOL":                  ("Black Garlic 500gr", 1),
    "BG-500GR-1-BOTOL-BH":             ("Black Garlic 500gr", 1),
    "BG-500GR-1-BOTOL-BG-SKU":         ("Black Garlic 500gr", 1),
    "BG-500GR-1-BOTOL-BH-SKU":         ("Black Garlic 500gr", 1),
    "BG-500GR-1-BOTOL-BG":            ("Black Garlic 500gr", 1),
    "BG-500GR-SKU":                     ("Black Garlic 500gr", 1),
    "BG-500GR-2-BOTOL":                 ("Black Garlic 500gr", 2),
    "BG-500GR-2-BOTOL-BG-SKU":         ("Black Garlic 500gr", 2),
    "BG-500GR-2-BOTOL-BH-SKU":         ("Black Garlic 500gr", 2),

    # ── HAMPERS BOX ONLY ─────────────────────────────────────────────
    "BOX-HAMPERS-IMLEK":               ("Box Hampers Imlek", 1),
    "BOX-HAMPERS-HSD":                 ("Box Hampers HSD", 1),
    "BOX-HAMPERS-NATAL":               ("Box Hampers Natal", 1),
    "BOX-HAMPERS-LEBARAN":             ("Box Hampers Lebaran", 1),

    # ── HAMPERS ISI ──────────────────────────────────────────────────
    "BG-220GR-2-BOTOL-HAMPERS-IMLEK":  ("Hampers Imlek (2 botol 220gr + box)", 2),
    "BG-220GR-2-BOTOL-HAMPERS-LEBARAN":("Hampers Lebaran (2 botol 220gr + box)", 2),
    "BG-220GR-2-BOTOL-HAMPERS-NATAL":  ("Hampers Natal (2 botol 220gr + box)", 2),
    "BG-220GR-2-BOTOL-HAMPERS-HSD":    ("Hampers HSD (2 botol 220gr + box)", 2),

    # ── 3IN1 ─────────────────────────────────────────────────────────
    "BG-3IN1": ("Black Garlic 3in1 (100gr + 220gr + 500gr + Goodie Bag)", 3),
}


def lookup_sku(sku_raw):
    """
    Cari SKU di dictionary. Return (nama_produk, qty_bundle) atau None.
    Prioritas: exact match → prefix match dari panjang ke pendek.
    """
    s = normalize_sku(sku_raw)
    if not s:
        return None

    # Exact match
    if s in SKU_LOOKUP:
        return SKU_LOOKUP[s]

    # Prefix match: coba hilangkan suffix satu per satu dari kanan
    # Berguna untuk SKU seperti BG-100GR-1-BOTOL-BG-SKU1 (ada angka tambahan di ujung)
    parts = s.split("-")
    for i in range(len(parts), 1, -1):
        candidate = "-".join(parts[:i])
        if candidate in SKU_LOOKUP:
            return SKU_LOOKUP[candidate]

    return None


def get_qty_from_sku(sku_raw):
    """
    Ekstrak qty bundle dari nama SKU itu sendiri (sebelum pakai qty dari PDF).
    BG-220GR-2-BOTOL → 2
    BG-DRINK-PROMO-4-BTL-ORI → 4
    BG-DRINK-MIX-7-BOTOL → 7
    Return None jika tidak ada pola.
    """
    s = normalize_sku(sku_raw)
    if not s:
        return None

    # Pola: angka sebelum -BOTOL
    m = re.search(r"-(\d+)-BOTOL", s)
    if m:
        val = int(m.group(1))
        if val < 50:  # sanity check, bukan ukuran gram
            return val

    # Pola: angka sebelum -BTL
    m = re.search(r"-(\d+)-BTL", s)
    if m:
        val = int(m.group(1))
        if val < 50:
            return val

    return None


def normalize_product(sku, nama_produk):
    """
    Resolusi nama produk: lookup dict → qty bundle → keyword fallback.
    Return (nama_produk_bersih, alasan_cek)
    """
    # 1. Coba lookup dictionary
    result = lookup_sku(sku)
    if result:
        return result[0], ""

    # 2. Fallback keyword dari SKU dan nama produk
    s = normalize_sku(sku)
    n = clean(nama_produk).upper()
    text = f"{s} {n}"

    if "DRINK" in text and "PEACH" in text:
        return "BG Drink Peach", ""
    if "DRINK" in text and ("ORIGINAL" in text or "ORI" in text):
        return "BG Drink Original", ""
    if "MULTI" in text and ("FLORAL" in text or "FLORA" in text):
        return "Madu Multi Floral", ""
    if "KURMA" in text:
        return "Madu Bunga Kurma", ""
    if "84GR" in text or "84G" in text:
        return "Black Garlic 84gr", ""
    if "100GR" in text or "100 GR" in text:
        return "Black Garlic 100gr", ""
    if "220GR" in text or "220 GR" in text:
        return "Black Garlic 220gr", ""
    if "500GR" in text or "500 GR" in text:
        return "Black Garlic 500gr", ""
    if "3IN1" in text:
        return "Black Garlic 3in1", ""
    if "HAMPERS" in text:
        return "Hampers HSD", ""

    # 3. Tidak dikenali
    alasan = "SKU tidak dikenali"
    if not s or s == "(CEK-MANUAL)":
        alasan = "SKU kosong/tidak terbaca"
    return clean(nama_produk) or clean(sku) or "(cek manual)", alasan


# =========================
# DETECT PLATFORM / KURIR
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
        return "TikTok", "GTL", "NDD" if "NDD" in t else "REG"

    if re.search(r"\b(0046\d{8,12}|00296\d{7,12})\b", text) or "SICEPAT" in t:
        return "TikTok", "SiCepat", "REG"

    if re.search(r"\bCM\d{6,}\b", text):
        return "Shopee", "Gosend", "Instant"

    return "Unknown", "Unknown", "-"


# =========================
# BASIC FIELD EXTRACTOR
# =========================

def get_resi(text):
    patterns = [
        r"No\.\s*Resi\s*[:：]?\s*(\d{10,15})",
        r"Resi\s*[:：]?\s*(SPXID\d{8,18})",
        r"Air\s*waybill\s*[:：]?\s*([A-Z]{2}\d{8,15})",
        r"\b(JX\d{10})\b",
        r"\b(GTL\d{8,12})\b",
        r"\b(SPXID\d{10,18})\b",
        r"\b(LXAD-\d+)\b",
        r"\b(AAJ\w{8,})\b",
        r"\b(BDO\d{6,})\b",
        r"\b(CM\d{6,})\b",
        r"\b(0046\d{8,12})\b",
        r"\b(00296\d{7,12})\b",
        r"\b(JJ\d{8,15})\b",
    ]

    for p in patterns:
        m = re.search(p, text, re.I)
        if m:
            return clean(m.group(1))

    # Fallback: pakai No. Pesanan agar order tidak hilang sama sekali
    fallback_patterns = [
        r"No\.?\s*Pesanan\s*[:：]?\s*([A-Z0-9\-]{6,})",
        r"Order\s*ID\s*[:：]?\s*([0-9]{6,})",
        r"TT\s*Order\s*ID\s*[:：]?\s*([0-9]{6,})",
    ]
    for p in fallback_patterns:
        m = re.search(p, text, re.I)
        if m:
            return "PESANAN-" + clean(m.group(1))

    return ""


def get_pesanan(text):
    patterns = [
        r"No\.?\s*Pesanan\s*[:：]?\s*([A-Z0-9\-]+)",
        r"Order\s*ID\s*[:：]?\s*([0-9]+)",
        r"TT\s*Order\s*ID\s*[:：]?\s*([0-9]+)",
        r"Nomor\s*Order\s*[:：]?\s*([0-9]+)",
        r"Pesan\s*[:：]?\s*\(?([A-Z0-9\-]+)\)?",
    ]
    for p in patterns:
        m = re.search(p, text, re.I)
        if m:
            return clean(m.group(1))
    return ""


def get_kode_pengambilan(text):
    m = re.search(r"Kode\s*Pengambilan\s*[:：]?\s*([A-Z0-9]+)", text, re.I)
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
        r"Penerima\s*[:：]\s*([^\n(]+)",
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

    m = re.search(r"Penerima\s*:\s*([^\n]+)", text, re.I)
    if m:
        return clean(m.group(1))

    return ""


def get_alamat(text):
    lines = [clean(x) for x in text.split("\n") if clean(x)]

    # Lazada
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

    # Shopee
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

    # TikTok / J&T / GTL
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
# PRODUCT HELPERS
# =========================

def extract_sku_candidates(text):
    """
    Ambil semua kandidat SKU dari teks, termasuk setelah merge_broken_lines.
    """
    text = merge_broken_lines(text)

    patterns = [
        r"BGH-[A-Z0-9\-]+",
        r"BLACKGARLIC-[A-Z0-9\-]+",
        r"BG-DRINK-[A-Z0-9\-]+",
        r"BG-PROMO-[A-Z0-9\-]+",
        r"BG-(?:100|220|500)GR-[A-Z0-9\-]+",
        r"BG-(?:100|220|500)GR-SKU",
        r"BG-(?:100|220|500)GR-\d+-BOTOL(?:-[A-Z0-9\-]+)?",
        r"BG-84GR-[A-Z0-9\-]+",
        r"BOX-HAMPERS-[A-Z0-9\-]+",
        r"BG-3IN1",
    ]

    found = []
    for p in patterns:
        for m in re.finditer(p, text, re.I):
            sku = normalize_sku(m.group(0))
            if sku and sku not in found:
                found.append(sku)

    return found


def resolve_qty(sku, qty_from_pdf):
    """
    Tentukan qty final dengan prioritas:
    1. Lookup dictionary (paling akurat)
    2. Ekstrak dari nama SKU
    3. Qty dari PDF
    4. Default 1
    """
    # 1. Dari lookup dict
    result = lookup_sku(sku)
    if result:
        qty_bundle = result[1]
        # qty_from_pdf adalah berapa kali SKU ini dipesan
        # qty final = qty_bundle (isi per SKU) × qty_from_pdf (berapa SKU dipesan)
        pdf_qty = to_int(qty_from_pdf, 1) or 1
        return qty_bundle * pdf_qty

    # 2. Dari nama SKU
    bundle = get_qty_from_sku(sku)
    if bundle:
        pdf_qty = to_int(qty_from_pdf, 1) or 1
        return bundle * pdf_qty

    # 3. Dari PDF
    q = to_int(qty_from_pdf, 1)
    return q if q > 0 else 1


def build_product_row(platform, courier, layanan, resi, pesanan, penerima,
                      alamat, pembayaran, kode, nama, sku, variasi, qty):
    nama_bersih, alasan_cek = normalize_product(sku, nama)
    final_qty = resolve_qty(sku, qty)

    return {
        "platform": platform,
        "courier": courier,
        "layanan": layanan,
        "no_resi": resi,
        "no_pesanan": pesanan,
        "nama_produk": nama_bersih,
        "nama_produk_asli": clean(nama),
        "sku": normalize_sku(sku),
        "variasi": clean(variasi),
        "qty": final_qty,
        "nama_penerima": penerima,
        "alamat": alamat,
        "pembayaran": pembayaran,
        "kode_pengambilan": kode,
        "alasan_cek": alasan_cek,
    }


# =========================
# PARSER PRODUCT FORMAT
# =========================

def parse_tiktok_table(text):
    """
    TikTok/J&T/GTL format:
    Product Name  SKU  Seller SKU  Qty
    """
    products = []
    t = merge_broken_lines(text)

    if "Product Name" not in t or "Seller SKU" not in t:
        return products

    m = re.search(
        r"Product Name\s+SKU\s+Seller SKU\s+Qty\s*\n([\s\S]+?)(?:Qty Total:|Order ID:|$)",
        t,
        re.I,
    )

    if not m:
        return products

    block = m.group(1)
    lines = [clean(x) for x in block.split("\n") if clean(x)]
    current_name = []

    for line in lines:
        mm = re.search(
            r"^(.*?)\s*(Default|ORI\s*PROMO\s*1|ORIGINAL|PEACH|100\s*GR|220\s*GR|500\s*GR)?\s+((?:BGH|BG|BLACKGARLIC|BOX)-[A-Z0-9\-\s]+)\s+(\d{1,3})$",
            line,
            re.I,
        )

        if mm:
            prefix = clean(mm.group(1))
            variasi = clean(mm.group(2))
            sku = normalize_sku(mm.group(3))
            qty = to_int(mm.group(4), 1)

            name = " ".join(current_name).strip()
            if prefix and len(prefix) > 4 and prefix.upper() not in ["DEFAULT"]:
                name = (name + " " + prefix).strip()

            products.append({
                "nama_produk": name or prefix or sku,
                "sku": sku,
                "variasi": variasi,
                "qty": qty,
            })
            current_name = []
        else:
            current_name.append(line)

    if products:
        return products

    # Fallback SKU kandidat
    skus = extract_sku_candidates(t)
    for sku in skus:
        products.append({"nama_produk": sku, "sku": sku, "variasi": "", "qty": 1})

    return products


def parse_shopee_table(text):
    """
    Shopee standard table: # Nama Produk  SKU  Lokasi  Variasi  Qty
    """
    products = []
    t = merge_broken_lines(text)

    if "# Nama Produk" not in t and "Nama Produk SKU" not in t:
        return products

    m = re.search(
        r"(?:#\s*)?Nama Produk\s+SKU[\s\S]*?(?:Variasi\s+)?Qty\s*\n([\s\S]+?)(?:Pesan:|Pengirim:|CASHLESS|$)",
        t,
        re.I,
    )
    block = m.group(1) if m else t
    lines = [clean(x) for x in block.split("\n") if clean(x)]
    block_text = "\n".join(lines)

    skus = extract_sku_candidates(block_text)

    for sku in skus:
        qty = 1
        token_text = normalize_sku(block_text)
        pos = token_text.find(sku)

        if pos >= 0:
            tail = token_text[pos + len(sku):pos + len(sku) + 60]
            nums = re.findall(r"\b(\d{1,3})\b", tail)
            # Ambil angka pertama yang bukan bagian SKU (bukan 100/220/500/84)
            for n in nums:
                if int(n) not in [100, 220, 500, 84, 7]:
                    qty = to_int(n, 1)
                    break

        name = ""
        for line in lines:
            if "BG-" in normalize_sku(line) or "BGH-" in normalize_sku(line):
                break
            if not re.match(r"^\d+$", line):
                name += " " + line

        products.append({
            "nama_produk": clean(name) or sku,
            "sku": sku,
            "variasi": "",
            "qty": qty,
        })

    return products


def parse_gosend_table(text):
    """
    Shopee Gosend / Instant — tidak punya header tabel standar.
    Langsung cari SKU dari teks bebas.
    """
    products = []
    t = merge_broken_lines(text)

    if "GOSEND" not in t.upper() and "GO-SEND" not in t.upper() and "CM" not in t.upper():
        return products

    skus = extract_sku_candidates(t)

    for sku in skus:
        qty = 1
        token_text = normalize_sku(t)
        pos = token_text.find(sku)

        if pos >= 0:
            tail = token_text[pos + len(sku):pos + len(sku) + 60]
            nums = re.findall(r"\b(\d{1,3})\b", tail)
            for n in nums:
                if int(n) not in [100, 220, 500, 84, 7]:
                    qty = to_int(n, 1)
                    break

        products.append({
            "nama_produk": sku,
            "sku": sku,
            "variasi": "",
            "qty": qty,
        })

    return products


def parse_lazada_table(text):
    products = []
    t = merge_broken_lines(text)

    if "LAZADA" not in t.upper() and "LXAD" not in t.upper():
        return products

    m = re.search(
        r"Nama Produk\s+Qty\s+SKU\s+Item Variant\s*\n([\s\S]+?)(?:Pengirim:|Penerima:|$)",
        t,
        re.I,
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
        qty = 1
        token_text = normalize_sku(t)
        pos = token_text.find(sku)

        if pos >= 0:
            tail = token_text[pos + len(sku):pos + len(sku) + 100]
            nums = re.findall(r"\b(\d{1,3})\b", tail)
            for n in nums:
                if int(n) not in [100, 220, 500, 84]:
                    qty = to_int(n, 1)
                    break

        products.append({"nama_produk": sku, "sku": sku, "variasi": "", "qty": qty})

    return products


def parse_generic_products(text):
    products = []
    t = merge_broken_lines(text)
    skus = extract_sku_candidates(t)

    for sku in skus:
        qty = 1
        token_text = normalize_sku(t)
        pos = token_text.find(sku)

        if pos >= 0:
            tail = token_text[pos + len(sku):pos + len(sku) + 70]
            nums = re.findall(r"\b(\d{1,3})\b", tail)
            for n in nums:
                if int(n) not in [100, 220, 500, 84, 7]:
                    qty = to_int(n, 1)
                    break

        products.append({"nama_produk": sku, "sku": sku, "variasi": "", "qty": qty})

    return products


# =========================
# PAGE PARSER
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

    product_candidates = []

    # Gosend/Instant duluan karena tidak ada header tabel standar
    if "GOSEND" in text.upper() or "GO-SEND" in text.upper():
        product_candidates = parse_gosend_table(text)

    if not product_candidates:
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
        product_candidates = [{
            "nama_produk": "(cek manual)",
            "sku": "(cek manual)",
            "variasi": "",
            "qty": 1,
        }]

    rows = []
    for p in product_candidates:
        rows.append(
            build_product_row(
                platform=platform,
                courier=courier,
                layanan=layanan,
                resi=resi,
                pesanan=pesanan,
                penerima=penerima,
                alamat=alamat,
                pembayaran=pembayaran,
                kode=kode,
                nama=p.get("nama_produk", ""),
                sku=p.get("sku", ""),
                variasi=p.get("variasi", ""),
                qty=p.get("qty", 1),
            )
        )

    return rows


# =========================
# MAIN PROCESS PDF
# =========================

def process_pdf(pdf_path, progress_callback=None):
    all_rows = []
    errors = []
    seen = set()

    with pdfplumber.open(pdf_path) as pdf:
        total = len(pdf.pages)

        for i, page in enumerate(pdf.pages):
            try:
                text = page.extract_text() or ""
                page_rows = parse_page(text)

                for r in page_rows:
                    key = (
                        r.get("no_resi", ""),
                        r.get("sku", ""),
                        r.get("nama_produk", ""),
                        str(r.get("qty", "")),
                    )
                    if key not in seen:
                        seen.add(key)
                        all_rows.append(r)

            except Exception as e:
                errors.append(f"Halaman {i + 1}: {e}")

            if progress_callback:
                progress_callback(i + 1, total)

    return all_rows, errors
