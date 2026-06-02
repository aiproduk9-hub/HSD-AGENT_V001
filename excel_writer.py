  from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from collections import defaultdict
from datetime import datetime


# =========================
# STYLE DASAR
# =========================

C_COLORS = {
    'J&T Express': 'D6EAFF',
    'GTL': 'D6F5E3',
    'SiCepat': 'FFF0CC',
    'Shopee Express': 'FCE4D6',
    'Wahana': 'E8D5F5',
    'Anteraja': 'D5ECF5',
    'LEX': 'FFD6D6',
    'Unknown': 'F2F2F2',
}

HDR = PatternFill('solid', fgColor='1B1F2E')
HFNT = Font(bold=True, color='FFFFFF', size=10, name='Calibri')
C = Alignment(horizontal='center', vertical='center', wrap_text=True)
L = Alignment(horizontal='left', vertical='center', wrap_text=True)

BD = Border(
    left=Side(style='thin', color='CCCCCC'),
    right=Side(style='thin', color='CCCCCC'),
    top=Side(style='thin', color='CCCCCC'),
    bottom=Side(style='thin', color='CCCCCC')
)

AMBER = PatternFill('solid', fgColor='F5A623')


def safe(r, key, default=''):
    return r.get(key, default) if isinstance(r, dict) else default


def clean_text(v):
    if v is None:
        return ''
    return str(v).rstrip(',').strip()


def header(ws, cols):
    for i, (lbl, w) in enumerate(cols, 1):
        c = ws.cell(row=1, column=i, value=lbl)
        c.fill = HDR
        c.font = HFNT
        c.alignment = C
        c.border = BD
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.row_dimensions[1].height = 30
    ws.freeze_panes = 'A2'


def write_row(ws, row, vals, fgcolor, aligns):
    f = PatternFill('solid', fgColor=fgcolor)

    for i, (v, al) in enumerate(zip(vals, aligns), 1):
        c = ws.cell(row=row, column=i, value=v)
        c.fill = f
        c.alignment = al
        c.border = BD
        c.font = Font(size=9, name='Calibri')


def total_row(ws, row, vals, color='F5A623'):
    fill = PatternFill('solid', fgColor=color)

    for col, val in enumerate(vals, 1):
        c = ws.cell(row=row, column=col, value=val)
        c.fill = fill
        c.font = Font(bold=True, size=11, name='Calibri', color='FFFFFF' if color == '1B1F2E' else '000000')
        c.alignment = C
        c.border = BD


# =========================
# EXCEL UTAMA
# =========================

def write_excel(rows, output_path):
    wb = Workbook()

    # =========================
    # SHEET 1: SEMUA RESI
    # =========================

    ws1 = wb.active
    ws1.title = "Semua Resi"

    header(ws1, [
        ('No', 5),
        ('Kurir', 15),
        ('Layanan', 9),
        ('No. Resi', 22),
        ('Nama Produk', 34),
        ('Seller SKU', 24),
        ('Variasi', 16),
        ('Qty', 6),
        ('Nama Penerima', 22),
        ('Alamat', 42),
        ('Platform', 14),
    ])

    srows = sorted(rows, key=lambda r: (safe(r, 'courier'), safe(r, 'no_resi')))

    prev_resi = None
    flip = True

    for i, r in enumerate(srows, 1):
        no_resi = safe(r, 'no_resi')

        if no_resi != prev_resi:
            flip = not flip
            prev_resi = no_resi

        courier = safe(r, 'courier', 'Unknown')
        base = C_COLORS.get(courier, 'F2F2F2')
        clr = base if flip else 'FFFFFF'

        qty = safe(r, 'qty', 0)
        try:
            qty = int(qty)
        except Exception:
            qty = 0

        vals = [
            i,
            courier,
            safe(r, 'layanan'),
            no_resi,
            clean_text(safe(r, 'nama_produk')),
            safe(r, 'sku'),
            safe(r, 'variasi') or '-',
            qty,
            safe(r, 'nama_penerima'),
            safe(r, 'alamat'),
            safe(r, 'platform'),
        ]

        write_row(ws1, i + 1, vals, clr, [C, L, C, L, L, L, C, C, L, L, L])

        if qty > 1:
            ws1.cell(row=i + 1, column=8).font = Font(
                bold=True,
                color='CC0000',
                size=11,
                name='Calibri'
            )

    ws1.auto_filter.ref = f"A1:{get_column_letter(11)}1"

    # =========================
    # SHEET 2: REKAP KURIR
    # =========================

    ws2 = wb.create_sheet("Rekap Kurir")

    header(ws2, [
        ('Kurir', 18),
        ('Total Resi', 14),
        ('Total Qty', 12),
        ('Platform', 18),
        ('Layanan', 12),
    ])

    by_courier = defaultdict(lambda: {
        'resi': set(),
        'qty': 0,
        'platform': set(),
        'layanan': set(),
    })

    for r in rows:
        courier = safe(r, 'courier', 'Unknown')
        by_courier[courier]['resi'].add(safe(r, 'no_resi'))
        by_courier[courier]['qty'] += int(safe(r, 'qty', 0) or 0)

        if safe(r, 'platform'):
            by_courier[courier]['platform'].add(safe(r, 'platform'))

        if safe(r, 'layanan'):
            by_courier[courier]['layanan'].add(safe(r, 'layanan'))

    row_num = 2

    for cur, v in sorted(by_courier.items()):
        clr = C_COLORS.get(cur, 'F2F2F2')

        vals = [
            cur,
            len(v['resi']),
            v['qty'],
            ', '.join(sorted(v['platform'])),
            ', '.join(sorted(v['layanan'])),
        ]

        write_row(ws2, row_num, vals, clr, [L, C, C, L, L])

        ws2.cell(row=row_num, column=2).font = Font(bold=True, size=11, name='Calibri')
        ws2.cell(row=row_num, column=3).font = Font(bold=True, size=11, name='Calibri')

        row_num += 1

    total_row(ws2, row_num, [
        'TOTAL',
        len(set(safe(r, 'no_resi') for r in rows)),
        sum(int(safe(r, 'qty', 0) or 0) for r in rows),
        '',
        '',
    ])

    # =========================
    # SHEET 3: REKAP SKU
    # =========================

    ws3 = wb.create_sheet("Rekap SKU")

    header(ws3, [
        ('Seller SKU', 26),
        ('Variasi / Ukuran', 18),
        ('Nama Produk', 32),
        ('Total Resi', 12),
        ('Total Qty', 12),
    ])

    sku_s = defaultdict(lambda: {
        'resi': set(),
        'qty': 0,
        'nama': '',
        'variasi': '',
    })

    for r in rows:
        sku = safe(r, 'sku') or '(tanpa SKU)'
        sku_s[sku]['resi'].add(safe(r, 'no_resi'))
        sku_s[sku]['qty'] += int(safe(r, 'qty', 0) or 0)

        if not sku_s[sku]['nama'] and safe(r, 'nama_produk'):
            sku_s[sku]['nama'] = clean_text(safe(r, 'nama_produk'))

        if not sku_s[sku]['variasi'] and safe(r, 'variasi'):
            sku_s[sku]['variasi'] = safe(r, 'variasi')

    row_num = 2

    for sku, v in sorted(sku_s.items(), key=lambda x: -x[1]['qty']):
        vals = [
            sku,
            v['variasi'] or '-',
            v['nama'],
            len(v['resi']),
            v['qty'],
        ]

        write_row(ws3, row_num, vals, 'EAF3DE', [L, C, L, C, C])
        ws3.cell(row=row_num, column=5).font = Font(bold=True, size=11, name='Calibri')

        row_num += 1

    total_row(ws3, row_num, [
        'TOTAL',
        '',
        '',
        len(set(safe(r, 'no_resi') for r in rows)),
        sum(int(safe(r, 'qty', 0) or 0) for r in rows),
    ])

    # =========================
    # SHEET 4: PACKING LIST GUDANG
    # =========================

    ws4 = wb.create_sheet("Packing List Gudang")

    header(ws4, [
        ('No Urut', 7),
        ('No. Resi', 22),
        ('Kurir', 13),
        ('Nama Produk', 34),
        ('Seller SKU', 24),
        ('Variasi', 14),
        ('Qty', 6),
        ('✓ Ambil', 8),
    ])

    resi_groups = defaultdict(list)

    for r in srows:
        resi_groups[safe(r, 'no_resi')].append(r)

    row_num = 2
    no_urut = 1

    multi_colors = [
        'FFF9C4',
        'E8F5E9',
        'E3F2FD',
        'FCE4D6',
        'F3E5F5',
        'E0F7FA',
    ]

    single_colors = [
        'F5F5F5',
        'FFFFFF',
    ]

    mc_idx = 0
    single_idx = 0

    for resi, items in resi_groups.items():
        if len(items) > 1:
            grp_color = multi_colors[mc_idx % len(multi_colors)]
            mc_idx += 1

            for r in items:
                qty = int(safe(r, 'qty', 0) or 0)

                vals = [
                    no_urut,
                    safe(r, 'no_resi'),
                    safe(r, 'courier'),
                    clean_text(safe(r, 'nama_produk')),
                    safe(r, 'sku'),
                    safe(r, 'variasi') or '-',
                    qty,
                    '',
                ]

                write_row(ws4, row_num, vals, grp_color, [C, L, L, L, L, C, C, C])

                if qty > 1:
                    ws4.cell(row=row_num, column=7).font = Font(
                        bold=True,
                        color='CC0000',
                        size=11,
                        name='Calibri'
                    )

                row_num += 1
                no_urut += 1

            for col in range(1, 9):
                ws4.cell(row=row_num - 1, column=col).border = Border(
                    left=Side(style='thin', color='CCCCCC'),
                    right=Side(style='thin', color='CCCCCC'),
                    top=Side(style='thin', color='CCCCCC'),
                    bottom=Side(style='medium', color='888888')
                )

        else:
            r = items[0]
            qty = int(safe(r, 'qty', 0) or 0)

            clr = single_colors[single_idx % 2]
            single_idx += 1

            vals = [
                no_urut,
                safe(r, 'no_resi'),
                safe(r, 'courier'),
                clean_text(safe(r, 'nama_produk')),
                safe(r, 'sku'),
                safe(r, 'variasi') or '-',
                qty,
                '',
            ]

            write_row(ws4, row_num, vals, clr, [C, L, L, L, L, C, C, C])

            if qty > 1:
                ws4.cell(row=row_num, column=7).font = Font(
                    bold=True,
                    color='CC0000',
                    size=11,
                    name='Calibri'
                )

            row_num += 1
            no_urut += 1

    # =========================
    # SHEET 5: RINGKASAN HARI INI
    # =========================

    ws5 = wb.create_sheet("Ringkasan Hari Ini")

    ws5.column_dimensions['A'].width = 30
    ws5.column_dimensions['B'].width = 45
    ws5.column_dimensions['C'].width = 14

    ws5.merge_cells('A1:C1')

    jdl = ws5.cell(row=1, column=1, value='📦  RINGKASAN PRODUK HARI INI')
    jdl.font = Font(bold=True, size=14, name='Calibri', color='1B1F2E')
    jdl.alignment = C
    jdl.fill = AMBER

    ws5.merge_cells('A2:C2')

    h2 = ws5.cell(row=2, column=1, value='STOK YANG HARUS DISIAPKAN')
    h2.font = Font(bold=True, size=10, name='Calibri', color='FFFFFF')
    h2.fill = HDR
    h2.alignment = C

    for col, lbl in enumerate(['Produk / Seller SKU', 'Variasi', 'Qty Disiapkan'], 1):
        c = ws5.cell(row=3, column=col, value=lbl)
        c.fill = HDR
        c.font = HFNT
        c.alignment = C
        c.border = BD

    row5 = 4

    for sku, v in sorted(sku_s.items(), key=lambda x: -x[1]['qty']):
        vals = [
            sku,
            v['variasi'] or '-',
            v['qty'],
        ]

        for col, val in enumerate(vals, 1):
            c = ws5.cell(row=row5, column=col, value=val)
            c.fill = PatternFill('solid', fgColor='EAF3DE')
            c.alignment = L if col < 3 else C
            c.border = BD
            c.font = Font(size=10, name='Calibri', bold=(col == 3))

        row5 += 1

    for col, val in enumerate(['TOTAL PRODUK', '', sum(int(safe(r, 'qty', 0) or 0) for r in rows)], 1):
        c = ws5.cell(row=row5, column=col, value=val)
        c.fill = AMBER
        c.font = Font(bold=True, size=11, name='Calibri')
        c.alignment = C
        c.border = BD

    row5 += 2

    ws5.merge_cells(f'A{row5}:C{row5}')

    h3 = ws5.cell(row=row5, column=1, value='RESI DENGAN LEBIH DARI 1 PRODUK / VARIASI')
    h3.font = Font(bold=True, size=10, name='Calibri', color='FFFFFF')
    h3.fill = PatternFill('solid', fgColor='CC0000')
    h3.alignment = C

    row5 += 1

    for col, lbl in enumerate(['No. Resi', 'Produk & Qty', 'Kurir'], 1):
        c = ws5.cell(row=row5, column=col, value=lbl)
        c.fill = HDR
        c.font = HFNT
        c.alignment = C
        c.border = BD

    row5 += 1

    multi_resi = {
        resi: items
        for resi, items in resi_groups.items()
        if len(items) > 1
    }

    for resi, items in sorted(multi_resi.items()):
        detail = ' | '.join(
            f"{(safe(r, 'sku') or safe(r, 'nama_produk'))[:15]} x{safe(r, 'qty')}"
            for r in items
        )

        vals = [
            resi,
            detail,
            safe(items[0], 'courier'),
        ]

        for col, val in enumerate(vals, 1):
            c = ws5.cell(row=row5, column=col, value=val)
            c.fill = PatternFill('solid', fgColor='FFF9C4')
            c.alignment = L
            c.border = BD
            c.font = Font(size=9, name='Calibri')

        row5 += 1

    wb.save(output_path)
    return output_path


# =========================
# EXCEL MULTI AKUN
# =========================

def write_excel_multi(rows, output_path):
    """
    rows wajib sudah punya:
    akun: HSD / HSS
    waktu: PAGI / SIANG / SORE
    """

    write_excel(rows, output_path)

    from openpyxl import load_workbook

    wb = load_workbook(output_path)

    # =========================
    # SHEET PALING DEPAN: STOK FISIK HARI INI
    # =========================

    ws0 = wb.create_sheet("STOK FISIK HARI INI", 0)
    ws0.sheet_tab_color = "F5A623"

    for col, width in zip('ABCD', [36, 18, 18, 22]):
        ws0.column_dimensions[col].width = width

    title_fill = PatternFill('solid', fgColor='1B1F2E')
    title_font = Font(bold=True, color='FFFFFF', size=14, name='Calibri')

    sec_fill = PatternFill('solid', fgColor='F5A623')
    sec_font = Font(bold=True, color='FFFFFF', size=11, name='Calibri')

    row_fill1 = PatternFill('solid', fgColor='FFFFFF')
    row_fill2 = PatternFill('solid', fgColor='F8F6F1')

    tgl = datetime.now().strftime('%d %B %Y')

    ws0.merge_cells('A1:D1')
    t = ws0.cell(row=1, column=1, value=f'STOK FISIK YANG HARUS DISIAPKAN — {tgl}')
    t.fill = title_fill
    t.font = title_font
    t.alignment = C
    t.border = BD
    ws0.row_dimensions[1].height = 36

    stok = defaultdict(int)

    for r in rows:
        nama = clean_text(safe(r, 'nama_produk')) or safe(r, 'sku') or '(tanpa nama produk)'
        stok[nama] += int(safe(r, 'qty', 0) or 0)

    resi_map = defaultdict(list)

    for r in rows:
        resi_map[safe(r, 'no_resi')].append(r)

    kardus_single = sum(
        1 for items in resi_map.values()
        if len(items) == 1 and sum(int(safe(i, 'qty', 0) or 0) for i in items) == 1
    )

    kardus_multi_qty = sum(
        1 for items in resi_map.values()
        if len(items) == 1 and sum(int(safe(i, 'qty', 0) or 0) for i in items) > 1
    )

    kardus_campur = sum(
        1 for items in resi_map.values()
        if len(items) > 1
    )

    total_kardus = len(resi_map)
    total_produk = sum(stok.values())

    ws0.merge_cells('A2:D2')
    s1 = ws0.cell(row=2, column=1, value='RINGKASAN CEPAT')
    s1.fill = sec_fill
    s1.font = sec_font
    s1.alignment = C
    s1.border = BD

    metrics = [
        ('TOTAL BOTOL/PRODUK\nHarus Disiapkan', total_produk),
        ('TOTAL KARDUS/RESI\nHarus Dikirim', total_kardus),
        ('KARDUS ISI CAMPUR\n(>1 jenis produk)', kardus_campur),
        ('KARDUS ISI TUNGGAL\n(1 jenis produk)', kardus_single + kardus_multi_qty),
    ]

    for col, (label, val) in enumerate(metrics, 1):
        lc = ws0.cell(row=3, column=col, value=label)
        lc.fill = PatternFill('solid', fgColor='F0F3FA')
        lc.font = Font(size=9, name='Calibri', color='6B7280', bold=True)
        lc.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        lc.border = BD

        vc = ws0.cell(row=4, column=col, value=val)
        vc.fill = PatternFill('solid', fgColor='FFFFFF')
        vc.font = Font(bold=True, size=22, name='Calibri', color='1B1F2E')
        vc.alignment = C
        vc.border = BD

    ws0.row_dimensions[3].height = 34
    ws0.row_dimensions[4].height = 42

    ws0.merge_cells('A5:D5')
    s2 = ws0.cell(row=5, column=1, value='DETAIL STOK — AMBIL DARI GUDANG')
    s2.fill = sec_fill
    s2.font = sec_font
    s2.alignment = C
    s2.border = BD

    for col, lbl in enumerate(['Nama Produk', 'Jumlah Botol Disiapkan', '% dari Total', 'Keterangan'], 1):
        c = ws0.cell(row=6, column=col, value=lbl)
        c.fill = HDR
        c.font = HFNT
        c.alignment = C
        c.border = BD

    row_n = 7

    for i, (prod, qty) in enumerate(sorted(stok.items(), key=lambda x: -x[1])):
        fill = row_fill1 if i % 2 == 0 else row_fill2
        pct = f'{qty / total_produk * 100:.1f}%' if total_produk > 0 else '0%'
        bar_len = int(qty / total_produk * 20) if total_produk > 0 else 0
        bar = '█' * bar_len + '░' * (20 - bar_len)

        vals = [prod, qty, pct, bar]

        for col, val in enumerate(vals, 1):
            c = ws0.cell(row=row_n, column=col, value=val)
            c.fill = fill
            c.alignment = L if col in [1, 4] else C
            c.border = BD
            c.font = Font(size=11, name='Calibri')

        ws0.cell(row=row_n, column=2).font = Font(
            bold=True,
            size=14,
            name='Calibri',
            color='1B1F2E'
        )

        row_n += 1

    for col, val in enumerate(['TOTAL', total_produk, '100%', ''], 1):
        c = ws0.cell(row=row_n, column=col, value=val)
        c.fill = AMBER
        c.font = Font(bold=True, size=12, name='Calibri', color='FFFFFF')
        c.alignment = C
        c.border = BD

    row_n += 2

    ws0.merge_cells(f'A{row_n}:D{row_n}')
    s3 = ws0.cell(row=row_n, column=1, value='RINCIAN KARDUS / PACKAGING')
    s3.fill = sec_fill
    s3.font = sec_font
    s3.alignment = C
    s3.border = BD

    row_n += 1

    for col, lbl in enumerate(['Jenis Kardus', 'Jumlah', 'Keterangan', ''], 1):
        c = ws0.cell(row=row_n, column=col, value=lbl)
        c.fill = HDR
        c.font = HFNT
        c.alignment = C
        c.border = BD

    row_n += 1

    kardus_data = [
        ('Kardus isi 1 produk, qty 1', kardus_single, 'Packing standar, 1 produk langsung'),
        ('Kardus isi 1 produk, qty >1', kardus_multi_qty, 'Packing >1 botol produk sama'),
        ('Kardus isi CAMPUR (>1 jenis)', kardus_campur, '⚠️ Cek kombinasi di sheet Ringkasan Order'),
        ('TOTAL KARDUS', total_kardus, 'Total resi yang harus dikirim hari ini'),
    ]

    for i, (label, qty, ket) in enumerate(kardus_data):
        is_total = 'TOTAL' in label
        fill = AMBER if is_total else (row_fill1 if i % 2 == 0 else row_fill2)
        font = Font(
            bold=True if is_total else False,
            size=12,
            name='Calibri',
            color='FFFFFF' if is_total else '1B1F2E'
        )

        vals = [label, qty, ket, '']

        for col, val in enumerate(vals, 1):
            c = ws0.cell(row=row_n, column=col, value=val)
            c.fill = fill
            c.font = font
            c.alignment = C if col == 2 else L
            c.border = BD

        row_n += 1

    row_n += 1

    ws0.merge_cells(f'A{row_n}:D{row_n}')
    note = ws0.cell(
        row=row_n,
        column=1,
        value='📌 Catatan: Variasi (-) artinya data variasi tidak tersedia di PDF resi. Cek sheet Semua Resi untuk detail SKU.'
    )
    note.font = Font(italic=True, size=9, name='Calibri', color='9CA3AF')
    note.alignment = L

    ws0.freeze_panes = 'A7'

    # =========================
    # SHEET: REKAP GUDANG
    # =========================

    ws = wb.create_sheet("Rekap Gudang")

    for col, width in zip('ABCDE', [10, 12, 32, 18, 10]):
        ws.column_dimensions[col].width = width

    ws.merge_cells('A1:E1')

    jdl = ws.cell(row=1, column=1, value='📦  REKAP GUDANG')
    jdl.font = Font(bold=True, size=13, name='Calibri', color='FFFFFF')
    jdl.fill = PatternFill('solid', fgColor='1B1F2E')
    jdl.alignment = C

    for col, lbl in enumerate(['Akun', 'Waktu', 'Nama Produk / SKU', 'Variasi', 'Total Qty'], 1):
        c = ws.cell(row=2, column=col, value=lbl)
        c.fill = HDR
        c.font = HFNT
        c.alignment = C
        c.border = BD

    ws.freeze_panes = 'A3'

    data = defaultdict(lambda: defaultdict(lambda: defaultdict(int)))
    nama_map = {}
    variasi_map = {}

    for r in rows:
        akun = safe(r, 'akun', 'HSD')
        waktu = safe(r, 'waktu', 'PAGI')
        sku = safe(r, 'sku') or '(cek manual)'

        data[akun][waktu][sku] += int(safe(r, 'qty', 0) or 0)

        if sku not in nama_map and safe(r, 'nama_produk'):
            nama_map[sku] = clean_text(safe(r, 'nama_produk'))

        if sku not in variasi_map and safe(r, 'variasi'):
            variasi_map[sku] = safe(r, 'variasi')

    hsd_fill = PatternFill('solid', fgColor='1A3A5C')
    hss_fill = PatternFill('solid', fgColor='5C2A00')

    hsd_rows = {
        'PAGI': 'D6E4FF',
        'SIANG': 'C5D8F5',
        'SORE': 'B8CAE8',
    }

    hss_rows = {
        'PAGI': 'FFE4CC',
        'SIANG': 'F5D4B8',
        'SORE': 'E8C4A8',
    }

    waktu_order = ['PAGI', 'SIANG', 'SORE']

    row_n = 3

    for akun in ['HSD', 'HSS']:
        if akun not in data:
            continue

        akun_color = hsd_fill if akun == 'HSD' else hss_fill
        row_colors = hsd_rows if akun == 'HSD' else hss_rows

        for waktu in waktu_order:
            if waktu not in data[akun]:
                continue

            sku_qty = sorted(data[akun][waktu].items(), key=lambda x: -x[1])

            for sku, qty in sku_qty:
                nama = nama_map.get(sku, '')
                variasi = variasi_map.get(sku, '') or '-'
                clr = row_colors.get(waktu, 'FFFFFF')
                fill = PatternFill('solid', fgColor=clr)

                vals = [akun, waktu, nama or sku, variasi, qty]

                for col, val in enumerate(vals, 1):
                    c = ws.cell(row=row_n, column=col, value=val)
                    c.fill = fill
                    c.alignment = L if col == 3 else C
                    c.border = BD
                    c.font = Font(size=10, name='Calibri')

                c_akun = ws.cell(row=row_n, column=1)
                c_akun.fill = akun_color
                c_akun.font = Font(bold=True, color='FFFFFF', size=10, name='Calibri')

                if not nama:
                    c_nama = ws.cell(row=row_n, column=3)
                    c_nama.value = f'{sku} ← nama tidak terbaca di PDF'
                    c_nama.font = Font(color='CC0000', size=9, italic=True, name='Calibri')

                row_n += 1

            subtotal = sum(q for _, q in sku_qty)

            ws.merge_cells(f'A{row_n}:D{row_n}')
            c = ws.cell(row=row_n, column=1, value=f'  SUBTOTAL {akun} {waktu}')
            c.fill = AMBER
            c.font = Font(bold=True, size=10, name='Calibri')
            c.alignment = L
            c.border = BD

            ct = ws.cell(row=row_n, column=5, value=subtotal)
            ct.fill = AMBER
            ct.font = Font(bold=True, size=12, name='Calibri')
            ct.alignment = C
            ct.border = BD

            row_n += 1

        row_n += 1

    ws.merge_cells(f'A{row_n}:D{row_n}')
    c = ws.cell(row=row_n, column=1, value='  GRAND TOTAL SEMUA')
    c.fill = PatternFill('solid', fgColor='1B1F2E')
    c.font = Font(bold=True, color='FFFFFF', size=11, name='Calibri')
    c.alignment = L
    c.border = BD

    ct = ws.cell(row=row_n, column=5, value=sum(int(safe(r, 'qty', 0) or 0) for r in rows))
    ct.fill = PatternFill('solid', fgColor='1B1F2E')
    ct.font = Font(bold=True, color='FFFFFF', size=13, name='Calibri')
    ct.alignment = C
    ct.border = BD

    # =========================
    # SHEET: RINGKASAN ORDER
    # =========================

    ws6 = wb.create_sheet("Ringkasan Order")
    ws6.freeze_panes = 'A2'

    for col, width in zip('ABCDEF', [10, 10, 42, 14, 14, 22]):
        ws6.column_dimensions[col].width = width

    for col, lbl in enumerate([
        'Akun',
        'Waktu',
        'Kombinasi Produk yang Dibeli',
        'Jumlah Resi',
        'Total Qty',
        'Contoh No. Resi'
    ], 1):
        c = ws6.cell(row=1, column=col, value=lbl)
        c.fill = HDR
        c.font = HFNT
        c.alignment = C
        c.border = BD

    resi_map2 = defaultdict(list)

    for r in rows:
        resi_map2[safe(r, 'no_resi')].append(r)

    f_single = PatternFill('solid', fgColor='E8F5E9')
    f_multi = PatternFill('solid', fgColor='FFEBEE')
    f_qty = PatternFill('solid', fgColor='FFF9C4')

    akun_waktu_pairs = []
    seen = set()

    for r in rows:
        aw = (safe(r, 'akun', '-'), safe(r, 'waktu', '-'))

        if aw not in seen:
            seen.add(aw)
            akun_waktu_pairs.append(aw)

    row6 = 2

    for akun, waktu in akun_waktu_pairs:
        subset = {
            resi: items
            for resi, items in resi_map2.items()
            if safe(items[0], 'akun', '-') == akun and safe(items[0], 'waktu', '-') == waktu
        }

        if not subset:
            continue

        combo_groups = defaultdict(list)

        for resi, items in subset.items():
            nama_qty = defaultdict(int)

            for it in items:
                nama = clean_text(safe(it, 'nama_produk')) or safe(it, 'sku')
                nama_qty[nama] += int(safe(it, 'qty', 0) or 0)

            parts = [
                f"{nama} x{nama_qty[nama]}"
                for nama in sorted(nama_qty.keys())
            ]

            label = ' | '.join(parts)
            combo_groups[label].append(resi)

        for label, resi_list in sorted(combo_groups.items(), key=lambda x: -len(x[1])):
            total_qty = sum(
                sum(int(safe(r, 'qty', 0) or 0) for r in resi_map2[res])
                for res in resi_list
            )

            contoh = resi_list[0]
            is_multi = '|' in label
            first_qty = sum(int(safe(r, 'qty', 0) or 0) for r in resi_map2[contoh])

            fill = f_multi if is_multi else f_qty if first_qty > 1 else f_single

            vals = [
                akun,
                waktu,
                label,
                len(resi_list),
                total_qty,
                contoh,
            ]

            for col, val in enumerate(vals, 1):
                c = ws6.cell(row=row6, column=col, value=val)
                c.fill = fill
                c.alignment = L if col in [3, 6] else C
                c.border = BD
                c.font = Font(size=10, name='Calibri')

            ws6.cell(row=row6, column=4).font = Font(bold=True, size=12, name='Calibri')
            ws6.cell(row=row6, column=5).font = Font(bold=True, size=12, name='Calibri')

            if is_multi:
                ws6.cell(row=row6, column=4).font = Font(
                    bold=True,
                    size=12,
                    color='CC0000',
                    name='Calibri'
                )

            row6 += 1

        sub_resi = len(subset)
        sub_qty = sum(
            sum(int(safe(r, 'qty', 0) or 0) for r in v)
            for v in subset.values()
        )

        ws6.merge_cells(f'A{row6}:C{row6}')
        c = ws6.cell(row=row6, column=1, value=f'  SUBTOTAL {akun} {waktu} — {sub_resi} resi')
        c.fill = AMBER
        c.font = Font(bold=True, size=10, name='Calibri')
        c.alignment = L
        c.border = BD

        for col, val in [(4, sub_resi), (5, sub_qty)]:
            ct = ws6.cell(row=row6, column=col, value=val)
            ct.fill = AMBER
            ct.font = Font(bold=True, size=12, name='Calibri')
            ct.alignment = C
            ct.border = BD

        row6 += 2

    # =========================
    # SHEET: KODE PENGAMBILAN
    # =========================

    ws7 = wb.create_sheet("Kode Pengambilan (Gojek)")
    ws7.freeze_panes = 'A2'
    ws7.sheet_tab_color = '22C55E'

    for col, width in zip('ABCDE', [22, 14, 14, 22, 30]):
        ws7.column_dimensions[col].width = width

    for col, lbl in enumerate([
        'No. Resi',
        'Kurir',
        'Layanan',
        'Kode Pengambilan',
        'Produk & Qty'
    ], 1):
        c = ws7.cell(row=1, column=col, value=lbl)
        c.fill = HDR
        c.font = HFNT
        c.alignment = C
        c.border = BD

    resi_kode = defaultdict(list)

    for r in rows:
        if safe(r, 'kode_pengambilan'):
            resi_kode[safe(r, 'no_resi')].append(r)

    if resi_kode:
        row7 = 2

        for resi, items in sorted(resi_kode.items()):
            produk = ' | '.join(
                f"{safe(r, 'sku')} x{safe(r, 'qty')}"
                for r in items
            )

            vals = [
                resi,
                safe(items[0], 'courier'),
                safe(items[0], 'layanan'),
                safe(items[0], 'kode_pengambilan'),
                produk,
            ]

            for col, val in enumerate(vals, 1):
                c = ws7.cell(row=row7, column=col, value=val)
                c.fill = PatternFill('solid', fgColor='F0FDF4')
                c.alignment = L if col in [1, 4, 5] else C
                c.border = BD
                c.font = Font(size=10, name='Calibri')

            ws7.cell(row=row7, column=4).font = Font(
                bold=True,
                size=12,
                color='22A85A',
                name='Calibri'
            )

            row7 += 1

    else:
        ws7.merge_cells('A2:E2')
        c = ws7.cell(
            row=2,
            column=1,
            value='Tidak ada kode pengambilan di PDF ini. Biasanya hanya ada untuk layanan Instan/Gojek.'
        )
        c.font = Font(italic=True, color='9CA3AF', size=10, name='Calibri')
        c.alignment = L

    wb.save(output_path)
    return output_path 
