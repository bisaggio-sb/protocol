"""
generate_docx.py – GP2 Protocol Generator
Klonuje szablon Grupa_IND.docx dla każdego meczu i wstawia dane.
"""

import io, csv, re, copy, zipfile, requests
from lxml import etree

W   = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
REL_IMG = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'

def wt(name): return f'{{{W}}}{name}'


# ─── Google Sheets helpers ────────────────────────────────────────────────────

def fetch_sheet_csv(sheet_id, gid):
    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"
    r = requests.get(url, timeout=20)
    r.raise_for_status()
    return list(csv.reader(io.StringIO(r.text)))


def get_group_sheet_ids(sheet_id):
    """
    Pobiera listę (name, gid) dla zakładek Gr. A – Gr. P ze strony HTML arkusza.
    Używa wielu wzorców regex, bo format JSON w HTML Google Sheets bywa różny.
    """
    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/edit"
    r = requests.get(url, timeout=20)
    r.raise_for_status()
    text = r.text

    results = {}

    # Wzorzec 1: ["Gr. A", ..., 123456789]  – GID to OSTATNIA liczba w tablicy
    for m in re.finditer(r'\["(Gr\.\s*[A-Z]+)"(?:[^\[\]]*?),(\d{6,})\]', text):
        name, gid = m.group(1), m.group(2)
        if name not in results:
            results[name] = gid

    # Wzorzec 2: "Gr. A" ... gid:123456789
    if not results:
        for m in re.finditer(r'"(Gr\.\s*[A-Z]+)"[^"]*?"gid"\s*:\s*(\d+)', text):
            name, gid = m.group(1), m.group(2)
            if name not in results:
                results[name] = gid

    # Wzorzec 3: szukaj "Gr. X" i najbliższego 7-9 cyfrowego numeru PO nazwie
    if not results:
        for m in re.finditer(r'"(Gr\.\s*[A-Z]+)"', text):
            name = m.group(1)
            nearby = text[m.end():m.end()+300]
            nums = re.findall(r'\b(\d{7,10})\b', nearby)
            if nums and name not in results:
                results[name] = nums[0]

    # Posortuj A, B, C...
    def sort_key(item):
        letter = item[0].split('.')[-1].strip()
        return letter

    return sorted(results.items(), key=sort_key)


def parse_group_rows(rows):
    if not rows:
        return []
    header_idx = None
    header = []
    for i, row in enumerate(rows):
        norm = [c.strip().lower() for c in row]
        if 'tor' in norm:
            header_idx = i
            header = norm
            break
    if header_idx is None:
        return []

    def col(name):
        try: return header.index(name.lower())
        except ValueError: return None

    ci = {k: col(v) for k, v in {
        'tor': 'tor', 'godz': 'godzina', 'grupa': 'grupa',
        'mecz': 'mecz', 'z1': 'zawodnik 1', 'z2': 'zawodnik 2'
    }.items()}

    matches = []
    for row in rows[header_idx + 1:]:
        if not any(c.strip() for c in row):
            continue
        def g(k):
            idx = ci.get(k)
            if idx is None or idx >= len(row): return ''
            return row[idx].strip()
        tor = g('tor')
        z1  = g('z1')
        if not tor and not z1:
            continue
        matches.append({'tor': tor, 'godz': g('godz'), 'grupa': g('grupa'),
                        'mecz': g('mecz'), 'z1': z1, 'z2': g('z2')})
    return matches


def fetch_all_group_sheets(sheet_id):
    sheet_list = get_group_sheet_ids(sheet_id)
    result = []
    for name, gid in sheet_list:
        try:
            rows = fetch_sheet_csv(sheet_id, gid)
            matches = parse_group_rows(rows)
            result.append((name, matches))
        except Exception as e:
            result.append((name, []))
    return result


# ─── XML helpers ──────────────────────────────────────────────────────────────

def set_cell_text(tc, text, bold=False, size=None, align=None, font='Aptos'):
    paras = tc.findall(wt('p'))
    p = paras[0] if paras else etree.SubElement(tc, wt('p'))

    for r in p.findall(wt('r')):
        p.remove(r)

    if align:
        pPr = p.find(wt('pPr'))
        if pPr is None:
            pPr = etree.SubElement(p, wt('pPr'))
        jc = pPr.find(wt('jc'))
        if jc is None:
            jc = etree.SubElement(pPr, wt('jc'))
        jc.set(f'{{{W}}}val', align)

    if not text:
        return

    r = etree.SubElement(p, wt('r'))
    rPr = etree.SubElement(r, wt('rPr'))
    fonts = etree.SubElement(rPr, wt('rFonts'))
    for attr in ('ascii', 'hAnsi', 'eastAsia', 'cs'):
        fonts.set(f'{{{W}}}{attr}', font)
    for tag in ('b', 'bCs'):
        etree.SubElement(rPr, wt(tag)).set(f'{{{W}}}val', '1' if bold else '0')
    for tag in ('i', 'iCs'):
        etree.SubElement(rPr, wt(tag)).set(f'{{{W}}}val', '0')
    if size:
        for tag in ('sz', 'szCs'):
            etree.SubElement(rPr, wt(tag)).set(f'{{{W}}}val', str(size))
    etree.SubElement(rPr, wt('lang')).set(f'{{{W}}}val', 'pl-PL')

    t = etree.SubElement(r, wt('t'))
    t.text = text
    if text and (text[0] == ' ' or text[-1] == ' '):
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')


def make_page_break():
    p = etree.Element(wt('p'))
    pPr = etree.SubElement(p, wt('pPr'))
    sp = etree.SubElement(pPr, wt('spacing'))
    sp.set(f'{{{W}}}before', '0'); sp.set(f'{{{W}}}after', '0')
    r = etree.SubElement(p, wt('r'))
    br = etree.SubElement(r, wt('br'))
    br.set(f'{{{W}}}type', 'page')
    return p


def make_image_paragraph(rel_id, cx_emu, cy_emu, align='center'):
    WP  = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
    A   = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
    R   = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

    p = etree.Element(wt('p'))
    pPr = etree.SubElement(p, wt('pPr'))
    jc = etree.SubElement(pPr, wt('jc'))
    jc.set(f'{{{W}}}val', align)
    r = etree.SubElement(p, wt('r'))
    inline = etree.SubElement(r, f'{{{WP}}}inline')
    ext = etree.SubElement(inline, f'{{{WP}}}extent')
    ext.set('cx', str(cx_emu)); ext.set('cy', str(cy_emu))
    docPr = etree.SubElement(inline, f'{{{WP}}}docPr')
    docPr.set('id', rel_id); docPr.set('name', 'img')
    graphic = etree.SubElement(inline, f'{{{A}}}graphic')
    gd = etree.SubElement(graphic, f'{{{A}}}graphicData')
    gd.set('uri', PIC)
    pic = etree.SubElement(gd, f'{{{PIC}}}pic')
    nvPicPr = etree.SubElement(pic, f'{{{PIC}}}nvPicPr')
    cNvPr = etree.SubElement(nvPicPr, f'{{{PIC}}}cNvPr')
    cNvPr.set('id', '0'); cNvPr.set('name', 'img')
    etree.SubElement(nvPicPr, f'{{{PIC}}}cNvPicPr')
    blipFill = etree.SubElement(pic, f'{{{PIC}}}blipFill')
    blip = etree.SubElement(blipFill, f'{{{A}}}blip')
    blip.set(f'{{{R}}}embed', rel_id)
    stretch = etree.SubElement(blipFill, f'{{{A}}}stretch')
    etree.SubElement(stretch, f'{{{A}}}fillRect')
    spPr = etree.SubElement(pic, f'{{{PIC}}}spPr')
    xfrm = etree.SubElement(spPr, f'{{{A}}}xfrm')
    off = etree.SubElement(xfrm, f'{{{A}}}off')
    off.set('x','0'); off.set('y','0')
    ext2 = etree.SubElement(xfrm, f'{{{A}}}ext')
    ext2.set('cx', str(cx_emu)); ext2.set('cy', str(cy_emu))
    prstGeom = etree.SubElement(spPr, f'{{{A}}}prstGeom')
    prstGeom.set('prst', 'rect')
    etree.SubElement(prstGeom, f'{{{A}}}avLst')
    return p


def make_qr_bytes(url):
    try:
        import qrcode as _qr
        qr = _qr.QRCode(version=2, box_size=6, border=2)
        qr.add_data(url)
        qr.make(fit=True)
        img = qr.make_image(fill_color='black', back_color='white')
        buf = io.BytesIO()
        img.save(buf, format='PNG')
        return buf.getvalue()
    except ImportError:
        return None


def title_para(text, size=40, bold=False, align='center'):
    p = etree.Element(wt('p'))
    pPr = etree.SubElement(p, wt('pPr'))
    jc = etree.SubElement(pPr, wt('jc'))
    jc.set(f'{{{W}}}val', align)
    r = etree.SubElement(p, wt('r'))
    rPr = etree.SubElement(r, wt('rPr'))
    for attr in ('ascii','hAnsi','eastAsia','cs'):
        etree.SubElement(rPr, wt('rFonts')).set(f'{{{W}}}{attr}', 'Aptos')
    for tag in ('b','bCs'):
        etree.SubElement(rPr, wt(tag)).set(f'{{{W}}}val', '1' if bold else '0')
    for tag in ('sz','szCs'):
        etree.SubElement(rPr, wt(tag)).set(f'{{{W}}}val', str(size))
    t = etree.SubElement(r, wt('t'))
    t.text = text
    return p


# ─── Build docx ───────────────────────────────────────────────────────────────

def build_document(sheet_id, sheets_url, sheets_data, logos=None):
    import os
    template_path = os.path.join(os.path.dirname(__file__), 'Grupa_IND.docx')
    with open(template_path, 'rb') as f:
        tpl_bytes = f.read()

    zin = zipfile.ZipFile(io.BytesIO(tpl_bytes))
    doc_root = etree.fromstring(zin.read('word/document.xml'))
    rels_root = etree.fromstring(zin.read('word/_rels/document.xml.rels'))

    body = doc_root.find(wt('body'))
    sectPr = body.find(wt('sectPr'))
    template_elements = [el for el in body if el.tag != wt('sectPr')]

    # ── Obrazki ──────────────────────────────────────────────────────────────
    next_rid = [200]
    media_files = {}
    logo_rids = {}

    def add_image(key, img_bytes, target_w_cm):
        from PIL import Image as PILImage
        rid = f'rId{next_rid[0]}'
        next_rid[0] += 1
        fname = f'media/logo_{key}.png'
        pil = PILImage.open(io.BytesIO(img_bytes)).convert('RGBA')
        buf = io.BytesIO()
        pil.save(buf, format='PNG')
        media_files[fname] = buf.getvalue()
        cx = int(target_w_cm * 360000)
        cy = int(cx / (pil.width / pil.height))
        rel = etree.SubElement(rels_root, f'{{{REL}}}Relationship')
        rel.set('Id', rid); rel.set('Type', REL_IMG); rel.set('Target', fname)
        return rid, cx, cy

    if logos:
        for key, img_bytes in logos.items():
            if img_bytes:
                w_cm = 16.0 if key == 'banner' else 3.0
                logo_rids[key] = add_image(key, img_bytes, w_cm)

    # ── QR ───────────────────────────────────────────────────────────────────
    qr_rid_info = None
    qr_bytes = make_qr_bytes(sheets_url)
    if qr_bytes:
        rid = f'rId{next_rid[0]}'
        next_rid[0] += 1
        media_files['media/qrcode.png'] = qr_bytes
        rel = etree.SubElement(rels_root, f'{{{REL}}}Relationship')
        rel.set('Id', rid); rel.set('Type', REL_IMG); rel.set('Target', 'media/qrcode.png')
        qr_rid_info = (rid, int(4.5 * 360000), int(4.5 * 360000))

    # ── Nowe body ─────────────────────────────────────────────────────────────
    for el in list(body):
        body.remove(el)

    # Strona tytułowa
    body.append(title_para('GP2 2026', size=52, bold=True))
    body.append(title_para('Protokoły meczowe – faza grupowa', size=28))
    p_sp = etree.Element(wt('p'))
    etree.SubElement(etree.SubElement(p_sp, wt('pPr')), wt('spacing')).set(f'{{{W}}}before','400')
    body.append(p_sp)
    body.append(title_para('Arkusz wyników – zeskanuj QR:', size=22))
    if qr_rid_info:
        body.append(make_image_paragraph(*qr_rid_info, align='center'))
    body.append(title_para(sheets_url, size=14))
    body.append(make_page_break())

    # ── Protokoły ─────────────────────────────────────────────────────────────
    first = True
    for group_name, matches in sheets_data:
        for match in matches:
            if not first:
                body.append(make_page_break())
            first = False

            # Grafiki przed protokołem
            if logo_rids.get('banner'):
                body.append(make_image_paragraph(*logo_rids['banner'], align='center'))
            for key in ('top_left', 'top_right'):
                if logo_rids.get(key):
                    a = 'left' if key == 'top_left' else 'right'
                    body.append(make_image_paragraph(*logo_rids[key], align=a))

            # Klonuj szablon
            for el in template_elements:
                body.append(copy.deepcopy(el))

            # Wypełnij dane — ostatnie len(template_elements) elementów
            count = len(template_elements)
            added = list(body)[-count:]
            tbls = [el for el in added if el.tag == wt('tbl')]

            if tbls:
                tbl1 = tbls[0]
                rows = tbl1.findall(wt('tr'))

                # Wiersz 0: [Tor][val_tor][Godzina][val_godz] ...
                if len(rows) > 0:
                    tcs = rows[0].findall(wt('tc'))
                    # val_tor  = tcs[1]
                    # val_godz = tcs[3] (span=2, więc index 3)
                    # val_grupa = tcs[5] (span=2)
                    # val_mecz  = tcs[7]
                    vals = [(1, match.get('tor','')),
                            (3, match.get('godz','')),
                            (5, match.get('grupa','')),
                            (7, match.get('mecz',''))]
                    for idx, val in vals:
                        if idx < len(tcs):
                            set_cell_text(tcs[idx], val, size=28, align='center')

                # Wiersz 3 = Zawodnik 1, Wiersz 4 = Zawodnik 2
                if len(rows) > 3:
                    tcs3 = rows[3].findall(wt('tc'))
                    if tcs3:
                        set_cell_text(tcs3[0], match.get('z1',''), size=24, align='right')
                if len(rows) > 4:
                    tcs4 = rows[4].findall(wt('tc'))
                    if tcs4:
                        set_cell_text(tcs4[0], match.get('z2',''), size=24, align='right')

            # Logo dolne lewo
            if logo_rids.get('bottom_left'):
                body.append(make_image_paragraph(*logo_rids['bottom_left'], align='left'))

    if sectPr is not None:
        body.append(sectPr)

    # ── Złóż ZIP ──────────────────────────────────────────────────────────────
    doc_out  = etree.tostring(doc_root,  xml_declaration=True, encoding='UTF-8', standalone=True)
    rels_out = etree.tostring(rels_root, xml_declaration=True, encoding='UTF-8', standalone=True)

    zout_buf = io.BytesIO()
    zout = zipfile.ZipFile(zout_buf, 'w', compression=zipfile.ZIP_DEFLATED)
    skip = {'word/document.xml', 'word/_rels/document.xml.rels'} | {f'word/{f}' for f in media_files}
    for item in zin.infolist():
        if item.filename not in skip:
            zout.writestr(item, zin.read(item.filename))
    zout.writestr('word/document.xml', doc_out)
    zout.writestr('word/_rels/document.xml.rels', rels_out)
    for fname, data in media_files.items():
        zout.writestr(f'word/{fname}', data)
    zout.close(); zin.close()
    return zout_buf.getvalue()
