"""
generate_docx.py – GP2 Protocol Generator
"""

import io, csv, re, copy, zipfile, string
import requests
from lxml import etree

W   = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
REL_IMG = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'

def wt(n): return f'{{{W}}}{n}'


# ─── Google Sheets ────────────────────────────────────────────────────────────

def fetch_sheet_by_name(sheet_id, sheet_name):
    url = (f"https://docs.google.com/spreadsheets/d/{sheet_id}"
           f"/export?format=csv&sheet={requests.utils.quote(sheet_name)}")
    r = requests.get(url, timeout=15)
    if r.status_code != 200 or r.text.strip().startswith('<!'):
        return None
    return list(csv.reader(io.StringIO(r.text)))


def parse_group_rows(rows):
    """
    Obsługuje rzeczywistą strukturę zakładki Gr. X:
      #  | Godzina | Tor | (opcjonalne ukryte kol.) | Grupa X (=Z1) | 1.set | 2.set | Z2 | 1.set | 2.set
    Szuka nagłówka 'Tor' i 'Godzina', a Player1 = kolumna ze startkiem 'Gr',
    Player2 = Player1+3 (po dwóch kolumnach punktów).
    """
    if not rows:
        return []

    # Znajdź wiersz nagłówkowy
    header_idx, header = None, []
    for i, row in enumerate(rows):
        norm = [c.strip().lower() for c in row]
        if 'tor' in norm:
            header_idx = i
            header = norm
            break

    if header_idx is None:
        return []

    def ci(name):
        try: return header.index(name)
        except ValueError: return None

    col_tor  = ci('tor')
    col_godz = ci('godzina')

    # Mecz = '#' albo pierwsza kolumna numeryczna
    col_mecz = ci('#')
    if col_mecz is None:
        for name in ('mecz','lp','nr','numer','no','no.'):
            col_mecz = ci(name)
            if col_mecz is not None:
                break
    if col_mecz is None:
        col_mecz = 0

    # Player 1 = kolumna z nagłówkiem startującym od 'gr'
    col_z1, grupa_raw = None, ''
    for i, h in enumerate(header):
        if h.startswith('gr'):
            col_z1 = i
            grupa_raw = rows[header_idx][i].strip()
            break

    if col_z1 is None and col_tor is not None:
        col_z1 = col_tor + 1   # fallback

    # Player 2 = Player1 + 3 kolumny (P1, set1, set2, P2)
    col_z2 = (col_z1 + 3) if col_z1 is not None else None

    # Wyłuskaj literę grupy z nagłówka: "Grupa C" -> "C"
    m = re.search(r'\b([A-P])\b', grupa_raw)
    grupa = m.group(1) if m else ''

    matches = []
    for row in rows[header_idx + 1:]:
        if not any(c.strip() for c in row):
            continue
        def g(c):
            if c is None or c >= len(row): return ''
            return row[c].strip()
        tor = g(col_tor)
        z1  = g(col_z1)
        if not tor and not z1:
            continue
        matches.append({
            'tor':   tor,
            'godz':  g(col_godz),
            'grupa': grupa,
            'mecz':  g(col_mecz),
            'z1':    z1,
            'z2':    g(col_z2),
        })
    return matches


def fetch_all_group_sheets(sheet_id):
    results = []
    for letter in string.ascii_uppercase[:16]:
        name = f"Gr. {letter}"
        try:
            rows = fetch_sheet_by_name(sheet_id, name)
            if rows is None:
                continue
            matches = parse_group_rows(rows)
            if matches:
                results.append((name, matches))
        except Exception:
            continue
    return results


def get_sheet_names_debug(sheet_id):
    info = []
    for letter in string.ascii_uppercase[:16]:
        name = f"Gr. {letter}"
        try:
            rows = fetch_sheet_by_name(sheet_id, name)
            if rows is None:
                info.append(f"❌ {name}: brak zakładki")
                continue
            matches = parse_group_rows(rows)
            # Pokaż też pierwsze wiersze CSV dla diagnozy
            header_preview = rows[0][:8] if rows else []
            info.append(f"✅ {name}: {len(matches)} meczów | nagłówki: {header_preview}")
        except Exception as e:
            info.append(f"⚠️ {name}: błąd – {e}")
    return info


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
    for a in ('ascii','hAnsi','eastAsia','cs'):
        fonts.set(f'{{{W}}}{a}', font)
    for tag in ('b','bCs'):
        etree.SubElement(rPr, wt(tag)).set(f'{{{W}}}val','1' if bold else '0')
    for tag in ('i','iCs'):
        etree.SubElement(rPr, wt(tag)).set(f'{{{W}}}val','0')
    if size:
        for tag in ('sz','szCs'):
            etree.SubElement(rPr, wt(tag)).set(f'{{{W}}}val', str(size))
    etree.SubElement(rPr, wt('lang')).set(f'{{{W}}}val','pl-PL')
    t = etree.SubElement(r, wt('t'))
    t.text = text
    if text and (text[0]==' ' or text[-1]==' '):
        t.set('{http://www.w3.org/XML/1998/namespace}space','preserve')


def make_page_break():
    p = etree.Element(wt('p'))
    pPr = etree.SubElement(p, wt('pPr'))
    sp = etree.SubElement(pPr, wt('spacing'))
    sp.set(f'{{{W}}}before','0'); sp.set(f'{{{W}}}after','0')
    r = etree.SubElement(p, wt('r'))
    br = etree.SubElement(r, wt('br'))
    br.set(f'{{{W}}}type','page')
    return p


def make_image_paragraph(rel_id, cx_emu, cy_emu, align='center'):
    WP  = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
    A   = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
    R   = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    p = etree.Element(wt('p'))
    pPr = etree.SubElement(p, wt('pPr'))
    jc = etree.SubElement(pPr, wt('jc')); jc.set(f'{{{W}}}val', align)
    r = etree.SubElement(p, wt('r'))
    inline = etree.SubElement(r, f'{{{WP}}}inline')
    etree.SubElement(inline, f'{{{WP}}}extent', cx=str(cx_emu), cy=str(cy_emu))
    dPr = etree.SubElement(inline, f'{{{WP}}}docPr')
    dPr.set('id', str(abs(hash(rel_id)) % 9999)); dPr.set('name','img')
    graphic = etree.SubElement(inline, f'{{{A}}}graphic')
    gd = etree.SubElement(graphic, f'{{{A}}}graphicData', uri=PIC)
    pic = etree.SubElement(gd, f'{{{PIC}}}pic')
    nvP = etree.SubElement(pic, f'{{{PIC}}}nvPicPr')
    cNv = etree.SubElement(nvP, f'{{{PIC}}}cNvPr'); cNv.set('id','0'); cNv.set('name','img')
    etree.SubElement(nvP, f'{{{PIC}}}cNvPicPr')
    bf = etree.SubElement(pic, f'{{{PIC}}}blipFill')
    blip = etree.SubElement(bf, f'{{{A}}}blip'); blip.set(f'{{{R}}}embed', rel_id)
    st = etree.SubElement(bf, f'{{{A}}}stretch'); etree.SubElement(st, f'{{{A}}}fillRect')
    spPr = etree.SubElement(pic, f'{{{PIC}}}spPr')
    xfrm = etree.SubElement(spPr, f'{{{A}}}xfrm')
    off = etree.SubElement(xfrm, f'{{{A}}}off'); off.set('x','0'); off.set('y','0')
    ext2 = etree.SubElement(xfrm, f'{{{A}}}ext'); ext2.set('cx', str(cx_emu)); ext2.set('cy', str(cy_emu))
    pg = etree.SubElement(spPr, f'{{{A}}}prstGeom'); pg.set('prst','rect')
    etree.SubElement(pg, f'{{{A}}}avLst')
    return p


def make_qr_bytes(url):
    try:
        import qrcode as _qr
        qr = _qr.QRCode(version=2, box_size=6, border=2)
        qr.add_data(url); qr.make(fit=True)
        img = qr.make_image(fill_color='black', back_color='white')
        buf = io.BytesIO(); img.save(buf, format='PNG')
        return buf.getvalue()
    except ImportError:
        return None


def title_para(text, size=40, bold=False, align='center'):
    p = etree.Element(wt('p'))
    pPr = etree.SubElement(p, wt('pPr'))
    jc = etree.SubElement(pPr, wt('jc')); jc.set(f'{{{W}}}val', align)
    r = etree.SubElement(p, wt('r'))
    rPr = etree.SubElement(r, wt('rPr'))
    for a in ('ascii','hAnsi','eastAsia','cs'):
        etree.SubElement(rPr, wt('rFonts')).set(f'{{{W}}}{a}','Aptos')
    for tag in ('b','bCs'):
        etree.SubElement(rPr, wt(tag)).set(f'{{{W}}}val','1' if bold else '0')
    for tag in ('sz','szCs'):
        etree.SubElement(rPr, wt(tag)).set(f'{{{W}}}val', str(size))
    t = etree.SubElement(r, wt('t')); t.text = text
    return p


# ─── Build docx ───────────────────────────────────────────────────────────────

def build_document(sheet_id, sheets_url, sheets_data, logos=None):
    import os
    tpl_path = os.path.join(os.path.dirname(__file__), 'Grupa_IND.docx')
    with open(tpl_path, 'rb') as f:
        tpl_bytes = f.read()

    zin = zipfile.ZipFile(io.BytesIO(tpl_bytes))
    doc_root = etree.fromstring(zin.read('word/document.xml'))
    rels_root = etree.fromstring(zin.read('word/_rels/document.xml.rels'))

    body = doc_root.find(wt('body'))
    sectPr = body.find(wt('sectPr'))
    template_elements = [el for el in body if el.tag != wt('sectPr')]

    next_rid = [200]
    media_files = {}
    logo_rids = {}

    def add_image(key, img_bytes, target_w_cm):
        from PIL import Image as PILImage
        rid = f'rId{next_rid[0]}'; next_rid[0] += 1
        fname = f'media/logo_{key}.png'
        pil = PILImage.open(io.BytesIO(img_bytes)).convert('RGBA')
        buf = io.BytesIO(); pil.save(buf, format='PNG')
        media_files[fname] = buf.getvalue()
        cx = int(target_w_cm * 360000)
        cy = int(cx / (pil.width / pil.height))
        rel = etree.SubElement(rels_root, f'{{{REL}}}Relationship')
        rel.set('Id', rid); rel.set('Type', REL_IMG); rel.set('Target', fname)
        return rid, cx, cy

    if logos:
        for key, img_bytes in logos.items():
            if img_bytes:
                logo_rids[key] = add_image(key, img_bytes, 16.0 if key=='banner' else 3.0)

    qr_rid_info = None
    qr_bytes = make_qr_bytes(sheets_url)
    if qr_bytes:
        rid = f'rId{next_rid[0]}'; next_rid[0] += 1
        media_files['media/qrcode.png'] = qr_bytes
        rel = etree.SubElement(rels_root, f'{{{REL}}}Relationship')
        rel.set('Id', rid); rel.set('Type', REL_IMG); rel.set('Target','media/qrcode.png')
        qr_rid_info = (rid, int(4.5*360000), int(4.5*360000))

    for el in list(body): body.remove(el)

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

    first = True
    for group_name, matches in sheets_data:
        for match in matches:
            if not first: body.append(make_page_break())
            first = False

            if logo_rids.get('banner'):
                body.append(make_image_paragraph(*logo_rids['banner'], align='center'))
            for key in ('top_left','top_right'):
                if logo_rids.get(key):
                    body.append(make_image_paragraph(*logo_rids[key],
                                align='left' if key=='top_left' else 'right'))

            for el in template_elements:
                body.append(copy.deepcopy(el))

            added = list(body)[-len(template_elements):]
            tbls = [el for el in added if el.tag == wt('tbl')]

            if tbls:
                rows = tbls[0].findall(wt('tr'))
                if len(rows) > 0:
                    tcs = rows[0].findall(wt('tc'))
                    for idx, key in [(1,'tor'),(3,'godz'),(5,'grupa'),(7,'mecz')]:
                        if idx < len(tcs):
                            set_cell_text(tcs[idx], match.get(key,''), size=28, align='center')
                if len(rows) > 3:
                    tcs = rows[3].findall(wt('tc'))
                    if tcs: set_cell_text(tcs[0], match.get('z1',''), size=24, align='right')
                if len(rows) > 4:
                    tcs = rows[4].findall(wt('tc'))
                    if tcs: set_cell_text(tcs[0], match.get('z2',''), size=24, align='right')

            if logo_rids.get('bottom_left'):
                body.append(make_image_paragraph(*logo_rids['bottom_left'], align='left'))

    if sectPr is not None: body.append(sectPr)

    doc_out  = etree.tostring(doc_root,  xml_declaration=True, encoding='UTF-8', standalone=True)
    rels_out = etree.tostring(rels_root, xml_declaration=True, encoding='UTF-8', standalone=True)

    zout_buf = io.BytesIO()
    zout = zipfile.ZipFile(zout_buf, 'w', compression=zipfile.ZIP_DEFLATED)
    skip = {'word/document.xml','word/_rels/document.xml.rels'} | {f'word/{f}' for f in media_files}
    for item in zin.infolist():
        if item.filename not in skip:
            zout.writestr(item, zin.read(item.filename))
    zout.writestr('word/document.xml', doc_out)
    zout.writestr('word/_rels/document.xml.rels', rels_out)
    for fname, data in media_files.items():
        zout.writestr(f'word/{fname}', data)
    zout.close(); zin.close()
    return zout_buf.getvalue()
