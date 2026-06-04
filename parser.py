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
    s = s.replace("–", "-")
    s = s.replace("—", "-")
    s = re.sub(r"\s+", "-", s)
    s = re.sub(r"-+", "-", s)
    s = s.strip("-")

    # Fix typo OCR / pecahan umum
    s = s.replace("SK-U", "SKU")
    s = s.replace("S-KU", "SKU")
    s = s.replace("BOTOL-BGSKU", "BOTOL-BG-SKU")
    s = s.replace("BOTOL-BHSKU", "BOTOL-BH-SKU")

    return s


def merge_broken_lines(text):
    """
    Gabung pecahan SKU dari PDF.
    Contoh:
    BG-500GR-
    1-BOTOL
    jadi BG-500GR-1-BOTOL
    """
    if not text:
        return ""

    text = text.replace("\uFFFE", " ")

    fixes = [
        (r"BG-(100|220|500)GR-\s*\n\s*(\d+)-BOTOL", r"BG-\1GR-\2-BOTOL"),
        (r"BG-(100|220|500)GR-1-\s*\n\s*BOTOL", r"BG-\1GR-1-BOTOL"),
        (r"BG-(100|220|500)GR-1\s*\n\s*BOTOL", r"BG-\1GR-1-BOTOL"),
        (r"BG-(100|220|500)GR-\s*\n\s*SKU", r"BG-\1GR-SKU"),
        (r"BG-DRINK-PROMO-\s*\n\s*(\d+)-BTL-ORI", r"BG-DRINK-PROMO-\1-BTL-ORI"),
        (r"BG-PROMO-ORI-1-\s*\n\s*BOTOL", r"BG-PROMO-ORI-1-BOTOL"),
        (r"BG-DRINK-ORI-1-\s*\n\s*BOTOL-PROMO", r"BG-DRINK-ORI-1-BOTOL-PROMO"),
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
# DETECT PLATFORM / KURIR
# =========================

def detect(text):
    t = text.upper()

    if "BLIBLI" in t or "BLIBLI.COM" in t:
        return "Blibli", "J&T Express" if "JNT" in t or "J&T" in t else "Unknown", "CASHLESS"

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

    # Blibli
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

def normalize_product(sku, nama_produk):
    s = normalize_sku(sku)
    n = clean(nama_produk).upper()

    text = f"{s} {n}"

    if "DRINK" in text and "PEACH" in text:
        return "BG Drink Peach"

    if "DRINK" in text and ("ORIGINAL" in text or "ORI" in text):
        return "BG Drink Original"

    if "MULTI" in text and ("FLORAL" in text or "FLORA" in text):
        return "Madu Multi Floral"

    if "KURMA" in text:
        return "Madu Bunga Kurma"

    if "84GR" in text or "84G" in text:
        return "Black Garlic 84gr"

    if "100GR" in text or "100 GR" in text:
        return "Black Garlic 100gr"

    if "220GR" in text or "220 GR" in text:
        return "Black Garlic 220gr"

    if "500GR" in text or "500 GR" in text:
        return "Black Garlic 500gr"

    return clean(nama_produk) or clean(sku) or "(cek manual)"


def extract_sku_candidates(text):
    """
    Ambil SKU dari teks walaupun pecah.
    """
    text = merge_broken_lines(text)

    patterns = [
        r"BGH-[A-Z0-9\-]+",
        r"BLACKGARLIC-[A-Z0-9\-]+",
        r"BG-DRINK-[A-Z0-9\-]+",
        r"BG-PROMO-[A-Z0-9\-]+",
        r"BG-(?:100|220|500)GR-[A-Z0-9\-]+",
        r"BG-(?:100|220|500)GR-SKU",
        r"BG-(?:100|220|500)GR-\d+-BOTOL",
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


def build_product_row(platform, courier, layanan, resi, pesanan, penerima, alamat, pembayaran, kode, nama, sku, variasi, qty):
    return {
        "platform": platform,
        "courier": courier,
        "layanan": layanan,
        "no_resi": resi,
        "no_pesanan": pesanan,
        "nama_produk": normalize_product(sku, nama),
        "nama_produk_asli": clean(nama),
        "sku": normalize_sku(sku),
        "variasi": clean(variasi),
        "qty": to_int(qty, 1) or 1,
        "nama_penerima": penerima,
        "alamat": alamat,
        "pembayaran": pembayaran,
        "kode_pengambilan": kode,
    }


# =========================
# PARSER PRODUCT FORMAT
# =========================

def parse_tiktok_table(text):
    """
    TikTok/Tokopedia format:
    Product Name SKU Seller SKU Qty
    nama produk
    Default BG-220GR-1-BOTOL 1
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
        # Pola paling umum: Default BG-xxx 1 atau 220 GR BG-xxx 1
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
            if prefix and len(prefix) > 4 and not prefix.upper() in ["DEFAULT"]:
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

    # Fallback: ambil SKU kandidat dan Qty Total
    skus = extract_sku_candidates(t)
    for sku in skus:
        products.append({
            "nama_produk": sku,
            "sku": sku,
            "variasi": "",
            "qty": 1,
        })

    return products


def parse_shopee_table(text):
    """
    Shopee table:
    # Nama Produk SKU Lokasi Variasi Qty
    1 Bawang Hitam ...
    BG-500GR-
    1-BOTOL
    2
    """
    products = []
    t = merge_broken_lines(text)

    if "# Nama Produk" not in t and "Nama Produk SKU" not in t:
        return products

    # Ambil block mulai header produk sampai Pesan/Pengirim/akhir
    m = re.search(
        r"(?:#\s*)?Nama Produk\s+SKU[\s\S]*?(?:Variasi\s+)?Qty\s*\n([\s\S]+?)(?:Pesan:|Pengirim:|CASHLESS|$)",
        t,
        re.I,
    )

    block = m.group(1) if m else t
    lines = [clean(x) for x in block.split("\n") if clean(x)]

    # Gabung semua untuk cari SKU + qty dekat setelahnya
    block_text = "\n".join(lines)
    skus = extract_sku_candidates(block_text)

    for sku in skus:
        sku_pattern = re.escape(sku).replace("\\-", r"[-\s]*")
        qty = 1

        # cari qty setelah SKU
        mqty = re.search(sku_pattern + r"[\s\S]{0,80}?\b(\d{1,3})\b", normalize_sku(block_text), re.I)
        if mqty:
            qty = to_int(mqty.group(1), 1)

        # cara khusus: setelah sku di teks original ada angka sendiri
        pos = normalize_sku(block_text).find(sku)
        if pos >= 0:
            tail = normalize_sku(block_text)[pos + len(sku):pos + len(sku) + 60]
            nums = re.findall(r"\b\d{1,3}\b", tail)
            if nums:
                qty = to_int(nums[-1], qty)

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

    if products:
        return products

    return []


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

            products.append({
                "nama_produk": block,
                "sku": sku,
                "variasi": "",
                "qty": qty,
            })

    return products


def parse_blibli_table(text):
    products = []
    t = merge_broken_lines(text)

    if "BLIBLI" not in t.upper():
        return products

    skus = extract_sku_candidates(t)

    for sku in skus:
        qty = 1

        # Blibli biasanya ada kolom Jml setelah SKU item
        # Ambil angka terdekat setelah SKU
        token_text = normalize_sku(t)
        pos = token_text.find(sku)

        if pos >= 0:
            tail = token_text[pos + len(sku):pos + len(sku) + 100]
            nums = re.findall(r"\b\d{1,3}\b", tail)
            if nums:
                qty = to_int(nums[0], 1)

        products.append({
            "nama_produk": sku,
            "sku": sku,
            "variasi": "",
            "qty": qty,
        })

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
            nums = re.findall(r"\b\d{1,3}\b", tail)
            if nums:
                qty = to_int(nums[0], 1)

        products.append({
            "nama_produk": sku,
            "sku": sku,
            "variasi": "",
            "qty": qty,
        })

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

    # Urutan penting
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
                    # Dedup aman: jangan hilangkan item beda qty/sku di resi sama
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
