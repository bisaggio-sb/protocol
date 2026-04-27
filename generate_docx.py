"""
generate_docx.py – Generator protokołów meczowych Mölkky.

Klonuje szablon Grupa_IND.docx 1:1 dla każdego meczu, podmienia 6 komórek,
dodaje kod QR i grafiki pozycjonowalne w lewym obszarze "Wyniki turnieju".
"""

import io, csv, re, copy, zipfile, string
from urllib.parse import quote
import requests
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
REL_IMG = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'

def wt(n): return f'{{{W}}}{n}'


# ─── Google Sheets fetching ───────────────────────────────────────────────────

def _is_html(text):
    s = text.lstrip()[:200].lower()
    return s.startswith('<!doctype') or s.startswith('<html') or '<head' in s

def fetch_via_gviz(sheet_id, sheet_name):
    url = (f"https://docs.google.com/spreadsheets/d/{sheet_id}"
           f"/gviz/tq?tqx=out:csv&sheet={quote(sheet_name)}")
    r = requests.get(url, timeout=15)
    if r.status_code != 200 or _is_html(r.text):
        return None
    return list(csv.reader(io.StringIO(r.text)))

def fetch_via_gid(sheet_id, gid):
    url = (f"https://docs.google.com/spreadsheets/d/{sheet_id}"
           f"/export?format=csv&gid={gid}")
    r = requests.get(url, timeout=15)
    if r.status_code != 200 or _is_html(r.text):
        return None
    return list(csv.reader(io.StringIO(r.text)))

def fetch_via_export_name(sheet_id, sheet_name):
    url = (f"https://docs.google.com/spreadsheets/d/{sheet_id}"
           f"/export?format=csv&sheet={quote(sheet_name)}")
    r = requests.get(url, timeout=15)
    if r.status_code != 200 or _is_html(r.text):
        return None
    return list(csv.reader(io.StringIO(r.text)))

def get_sheet_gids(sheet_id):
    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/edit"
    try:
        r = requests.get(url, timeout=20)
    except Exception:
        return {}
    text = r.text
    mapping = {}
    for m in re.finditer(r'value="(\d{6,12})"[^>]*>([^<]{1,80})</option>', text):
        gid, name = m.group(1), m.group(2).strip()
        if name and name not in mapping:
            mapping[name] = gid
    if not mapping:
        for m in re.finditer(r'\["([^"\\]{1,60})"(?:\s*,[^,\[\]]*){1,15},(\d{6,12})\]', text):
            name, gid = m.group(1).strip(), m.group(2)
            if name and name not in mapping:
                mapping[name] = gid
    return mapping

def fetch_sheet(sheet_id, sheet_name, gid_map=None):
    rows = fetch_via_gviz(sheet_id, sheet_name)
    if rows: return rows
    if gid_map and sheet_name in gid_map:
        rows = fetch_via_gid(sheet_id, gid_map[sheet_name])
        if rows: return rows
    return fetch_via_export_name(sheet_id, sheet_name)


# ─── Parser zakładki grupy ────────────────────────────────────────────────────

def _is_valid_match_row(tor, godz, z1, z2):
    if not (tor and tor.strip().isdigit()): return False
    if not (godz and re.match(r'^\d{1,2}:\d{2}$', godz.strip())): return False
    if not z1 or not z2: return False
    return True

def parse_group_rows(rows):
    if not rows: return []
    header_idx, header = None, []
    for i, row in enumerate(rows):
        norm = [c.strip().lower() for c in row]
        if 'tor' in norm:
            header_idx = i; header = norm; break
    if header_idx is None: return []
    raw_header = rows[header_idx]
    def ci(name):
        try: return header.index(name)
        except ValueError: return None
    col_tor = ci('tor'); col_godz = ci('godzina')
    col_mecz = ci('#')
    if col_mecz is None:
        for n in ('mecz','lp','nr'):
            col_mecz = ci(n)
            if col_mecz is not None: break
    if col_mecz is None: col_mecz = 0
    col_z1, grupa_raw = None, ''
    for i, h in enumerate(header):
        if h.startswith('gr'):
            col_z1 = i
            grupa_raw = raw_header[i].strip() if i < len(raw_header) else ''
            break
    if col_z1 is None and col_tor is not None:
        col_z1 = col_tor + 1
    col_z2 = col_z1 + 3 if col_z1 is not None else None
    m = re.search(r'\b([A-P])\b', grupa_raw)
    grupa = m.group(1) if m else ''
    matches = []
    for row in rows[header_idx + 1:]:
        if not any(c.strip() for c in row): continue
        def g(c):
            if c is None or c >= len(row): return ''
            return row[c].strip()
        tor = g(col_tor); godz = g(col_godz); z1 = g(col_z1); z2 = g(col_z2)
        if not _is_valid_match_row(tor, godz, z1, z2): continue
        matches.append({'tor':tor,'godz':godz,'grupa':grupa,
                        'mecz':g(col_mecz),'z1':z1,'z2':z2})
    return matches

def fetch_all_group_sheets(sheet_id):
    gid_map = get_sheet_gids(sheet_id)
    results = []
    for letter in string.ascii_uppercase[:16]:
        name = f"Gr. {letter}"
        try:
            rows = fetch_sheet(sheet_id, name, gid_map)
            if rows is None: continue
            matches = parse_group_rows(rows)
            if matches: results.append((name, matches))
        except Exception:
            continue
    return results

def get_sheet_names_debug(sheet_id):
    info = []
    gid_map = get_sheet_gids(sheet_id)
    info.append(f"📋 Mapa GID: {len(gid_map)} zakładek\n")
    for letter in string.ascii_uppercase[:16]:
        name = f"Gr. {letter}"
        method = ""
        rows = fetch_via_gviz(sheet_id, name)
        if rows: method = "gviz"
        else:
            if name in gid_map:
                rows = fetch_via_gid(sheet_id, gid_map[name])
                if rows: method = f"gid={gid_map[name]}"
            if not rows:
                rows = fetch_via_export_name(sheet_id, name)
                if rows: method = "export"
        if not rows:
            info.append(f"❌ {name}: brak"); continue
        matches = parse_group_rows(rows)
        info.append(f"✅ {name} [{method}]: {len(matches)} meczów")
    return info


# ─── XML helpers ──────────────────────────────────────────────────────────────

def _set_cell_value(tc, text, *, bold=True, size=28, align='center'):
    for p in tc.findall(wt('p')):
        tc.remove(p)
    p = etree.SubElement(tc, wt('p'))
    pPr = etree.SubElement(p, wt('pPr'))
    if align:
        jc = etree.SubElement(pPr, wt('jc'))
        jc.set(f'{{{W}}}val', align)
    if not text: return
    r = etree.SubElement(p, wt('r'))
    rPr = etree.SubElement(r, wt('rPr'))
    fonts = etree.SubElement(rPr, wt('rFonts'))
    fonts.set(f'{{{W}}}ascii', 'Aptos')
    fonts.set(f'{{{W}}}hAnsi', 'Aptos')
    fonts.set(f'{{{W}}}eastAsia', 'Aptos')
    fonts.set(f'{{{W}}}cs', 'Aptos')
    if bold:
        for tag in ('b','bCs'):
            etree.SubElement(rPr, wt(tag)).set(f'{{{W}}}val','1')
    for tag in ('sz','szCs'):
        etree.SubElement(rPr, wt(tag)).set(f'{{{W}}}val', str(size))
    etree.SubElement(rPr, wt('lang')).set(f'{{{W}}}val','pl-PL')
    t = etree.SubElement(r, wt('t'))
    t.text = text


def _make_page_break_para():
    p = etree.Element(wt('p'))
    pPr = etree.SubElement(p, wt('pPr'))
    sp = etree.SubElement(pPr, wt('spacing'))
    sp.set(f'{{{W}}}before','0'); sp.set(f'{{{W}}}after','0')
    r = etree.SubElement(p, wt('r'))
    br = etree.SubElement(r, wt('br'))
    br.set(f'{{{W}}}type','page')
    return p


def _fill_protocol(elements, match):
    tbls = [el for el in elements if el.tag == wt('tbl')]
    if not tbls: return
    rows = tbls[0].findall(wt('tr'))
    if len(rows) > 0:
        tcs = rows[0].findall(wt('tc'))
        if len(tcs) > 1: _set_cell_value(tcs[1], match.get('tor',''),  size=28)
        if len(tcs) > 3: _set_cell_value(tcs[3], match.get('godz',''), size=28)
        if len(tcs) > 5: _set_cell_value(tcs[5], match.get('grupa',''),size=28)
        if len(tcs) > 7: _set_cell_value(tcs[7], match.get('mecz',''), size=28)
    if len(rows) > 3:
        tcs = rows[3].findall(wt('tc'))
        if tcs: _set_cell_value(tcs[0], match.get('z1',''), size=24, align='right')
    if len(rows) > 4:
        tcs = rows[4].findall(wt('tc'))
        if tcs: _set_cell_value(tcs[0], match.get('z2',''), size=24, align='right')


# ─── QR code ──────────────────────────────────────────────────────────────────

def make_qr_bytes(url):
    """Generuje PNG kodu QR jako bytes."""
    try:
        import qrcode as _qr
        qr = _qr.QRCode(version=None, box_size=8, border=2,
                        error_correction=_qr.constants.ERROR_CORRECT_M)
        qr.add_data(url); qr.make(fit=True)
        img = qr.make_image(fill_color='black', back_color='white')
        buf = io.BytesIO(); img.save(buf, format='PNG')
        return buf.getvalue()
    except ImportError:
        return None


# ─── Inserting QR + graphics into "Wyniki turnieju" cell ─────────────────────

def _make_image_paragraph(rel_id, cx_emu, cy_emu, align='center'):
    """Tworzy paragraf z obrazem inline - wzór sprawdzony przez python-docx."""
    return etree.fromstring(f'''<w:p xmlns:w="{W}"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
  <w:pPr><w:jc w:val="{align}"/><w:spacing w:before="60" w:after="60"/></w:pPr>
  <w:r><w:drawing>
    <wp:inline distT="0" distB="0" distL="0" distR="0"
      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
      xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
      <wp:extent cx="{int(cx_emu)}" cy="{int(cy_emu)}"/>
      <wp:docPr id="{abs(hash(rel_id))%9000+1000}" name="Picture"/>
      <wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
        <pic:pic>
          <pic:nvPicPr>
            <pic:cNvPr id="{abs(hash(rel_id))%9000+1000}" name="Picture"/>
            <pic:cNvPicPr/>
          </pic:nvPicPr>
          <pic:blipFill>
            <a:blip r:embed="{rel_id}"/>
            <a:stretch><a:fillRect/></a:stretch>
          </pic:blipFill>
          <pic:spPr>
            <a:xfrm><a:off x="0" y="0"/>
              <a:ext cx="{int(cx_emu)}" cy="{int(cy_emu)}"/>
            </a:xfrm>
            <a:prstGeom prst="rect"/>
          </pic:spPr>
        </pic:pic>
      </a:graphicData></a:graphic>
    </wp:inline>
  </w:drawing></w:r>
</w:p>''')


def _make_text_para(text, size=20, bold=True, align='center'):
    """Tworzy paragraf z tekstem."""
    p = etree.Element(wt('p'))
    pPr = etree.SubElement(p, wt('pPr'))
    if align:
        jc = etree.SubElement(pPr, wt('jc'))
        jc.set(f'{{{W}}}val', align)
    sp = etree.SubElement(pPr, wt('spacing'))
    sp.set(f'{{{W}}}before','60'); sp.set(f'{{{W}}}after','60')
    r = etree.SubElement(p, wt('r'))
    rPr = etree.SubElement(r, wt('rPr'))
    fonts = etree.SubElement(rPr, wt('rFonts'))
    for a in ('ascii','hAnsi','eastAsia','cs'):
        fonts.set(f'{{{W}}}{a}', 'Aptos')
    if bold:
        for tag in ('b','bCs'):
            etree.SubElement(rPr, wt(tag)).set(f'{{{W}}}val','1')
    for tag in ('sz','szCs'):
        etree.SubElement(rPr, wt(tag)).set(f'{{{W}}}val', str(size))
    t = etree.SubElement(r, wt('t')); t.text = text
    return p


def _populate_left_area(tbl_elements, items_for_cell):
    """
    Wstawia listę paragrafów (QR + grafiki + napis "Wyniki turnieju")
    do komórki "Wyniki turnieju" (tabela 2, wiersz 1, komórka 0).

    items_for_cell to lista paragrafów XML w kolejności od góry do dołu.
    """
    # Znajdź drugą tabelę (z wynikami) - jest po kilku paragrafach
    tbls = [el for el in tbl_elements if el.tag == wt('tbl')]
    if len(tbls) < 2: return
    score_tbl = tbls[1]
    rows = score_tbl.findall(wt('tr'))
    if len(rows) < 2: return
    # Komórka "Wyniki turnieju" jest w wierszu 1 (drugim), tc[0]
    # (wiersz 0 to "SET 1 / SET 2" header)
    row1 = rows[1]
    tcs = row1.findall(wt('tc'))
    if not tcs: return
    target_cell = tcs[0]

    # Wyczyść istniejące paragrafy w tej komórce
    for p in target_cell.findall(wt('p')):
        target_cell.remove(p)

    # Wstaw nowe paragrafy
    for item in items_for_cell:
        target_cell.append(item)


# ─── Build document ───────────────────────────────────────────────────────────

def build_document(sheet_id, sheets_url, sheets_data, logos=None,
                   tournament_name=None, include_qr=True,
                   image_order=None):
    """
    `logos` to dict {key: bytes} z grafikami.
    `image_order` to lista ['qr', 'logo1', 'logo2', ...] określająca kolejność
    od góry do dołu w lewym obszarze "Wyniki turnieju".
    """
    import os
    tpl_path = os.path.join(os.path.dirname(__file__), 'Grupa_IND.docx')
    with open(tpl_path, 'rb') as f:
        tpl_bytes = f.read()

    zin = zipfile.ZipFile(io.BytesIO(tpl_bytes))
    doc_xml = zin.read('word/document.xml')
    doc_root = etree.fromstring(doc_xml)
    body = doc_root.find(wt('body'))

    # ── Fix marginesy
    sectPr_check = body.find(wt('sectPr'))
    if sectPr_check is not None:
        pgMar = sectPr_check.find(wt('pgMar'))
        if pgMar is not None:
            for side in ('top','bottom','left','right'):
                pgMar.set(f'{{{W}}}{side}', '720')

    # ── Fix fonty etykiet (sz 24 → 18)
    LABEL_TEXTS = {'Tor','Godzina','Grupa','Mecz','#','Punkty','SET 1','SET 2',
                    'Wygrane','sety','Podpis'}
    for r in body.iter(wt('r')):
        ts = r.findall(wt('t'))
        if not ts: continue
        text_content = ''.join((t.text or '') for t in ts).strip()
        if text_content in LABEL_TEXTS:
            sz_el  = r.find(f'{wt("rPr")}/{wt("sz")}')
            szCs_el = r.find(f'{wt("rPr")}/{wt("szCs")}')
            if sz_el is not None and sz_el.get(f'{{{W}}}val') == '24':
                sz_el.set(f'{{{W}}}val', '18')
            if szCs_el is not None and szCs_el.get(f'{{{W}}}val') == '24':
                szCs_el.set(f'{{{W}}}val', '18')

    # ── Fix wyrównanie tabeli 1 z tabelą 2:
    # We wzorcu Oławy obie tabele zaczynają się od lewej krawędzi (brak tblInd)
    # i mają zbliżoną szerokość (~9700 DXA).
    # Naszej tabeli 1 (suma 9090) musimy dać szerokość 9690 (= tabela 2).
    # Robimy to przez wyrównanie SET 1=SET 2 (1080 DXA każda) i rozszerzenie
    # kolumny "Podpis" o 600 DXA (1260 → 1860).
    first_tbl = body.find(wt('tbl'))
    if first_tbl is not None:
        gcs = first_tbl.findall(f'{wt("tblGrid")}/{wt("gridCol")}')
        if len(gcs) == 12:
            gcs[4].set(f'{{{W}}}w', '720')
            gcs[5].set(f'{{{W}}}w', '360')
            gcs[11].set(f'{{{W}}}w', '1860')  # Podpis: 1260 → 1860 (+600)
            for row in first_tbl.findall(wt('tr')):
                for tc in row.findall(wt('tc')):
                    tcPr = tc.find(wt('tcPr'))
                    if tcPr is None: continue
                    tcW = tcPr.find(wt('tcW'))
                    gs  = tcPr.find(wt('gridSpan'))
                    if tcW is None: continue
                    cur_w = int(tcW.get(f'{{{W}}}w','0'))
                    span  = int(gs.get(f'{{{W}}}val','1')) if gs is not None else 1
                    # SET 1 (1050, span=2) → 1080
                    if cur_w == 1050 and span == 2:
                        tcW.set(f'{{{W}}}w', '1080')
                    # Komórka "Podpis" w wierszu nagłówków/zawodnika (cur=1800 lub 2220, span=2/3)
                    # rozszerzamy o 600 DXA bo gridCol[11] urosło o 600
                    elif cur_w == 1800 and span == 2:
                        tcW.set(f'{{{W}}}w', '2400')
                    elif cur_w == 2220 and span == 3:
                        tcW.set(f'{{{W}}}w', '2820')
                    elif cur_w == 1260 and span == 1:
                        tcW.set(f'{{{W}}}w', '1860')
            # NIE dodajemy tblInd - tabela ma się zaczynać od lewej krawędzi

    # ── Wyciągnij sectPr i template
    sectPr = body.find(wt('sectPr'))
    template_elements = [copy.deepcopy(el) for el in body if el.tag != wt('sectPr')]
    while template_elements and template_elements[-1].tag == wt('p'):
        last = template_elements[-1]
        ts = last.findall(f'.//{wt("t")}')
        if not ts or not any((t.text or '').strip() for t in ts):
            template_elements.pop()
        else:
            break

    # ── Przygotuj media (QR + grafiki) i relacje
    next_rid = [200]
    media_files = {}
    rels_root = etree.fromstring(zin.read('word/_rels/document.xml.rels'))

    qr_rid_info = None
    if include_qr and sheets_url:
        qr_bytes = make_qr_bytes(sheets_url)
        if qr_bytes:
            rid = f'rId{next_rid[0]}'; next_rid[0] += 1
            media_files['media/qrcode.png'] = qr_bytes
            rel = etree.SubElement(rels_root, f'{{{REL}}}Relationship')
            rel.set('Id', rid); rel.set('Type', REL_IMG); rel.set('Target','media/qrcode.png')
            qr_rid_info = (rid, int(2.0*360000), int(2.0*360000))

    logo_rids = {}
    if logos:
        from PIL import Image as PILImage
        for key, img_bytes in logos.items():
            if not img_bytes: continue
            rid = f'rId{next_rid[0]}'; next_rid[0] += 1
            fname = f'media/{key}.png'
            pil = PILImage.open(io.BytesIO(img_bytes)).convert('RGBA')
            buf = io.BytesIO(); pil.save(buf, format='PNG')
            media_files[fname] = buf.getvalue()
            target_w_cm = 2.0
            cx = int(target_w_cm * 360000)
            cy = int(cx / (pil.width / pil.height))
            rel = etree.SubElement(rels_root, f'{{{REL}}}Relationship')
            rel.set('Id', rid); rel.set('Type', REL_IMG); rel.set('Target', fname)
            logo_rids[key] = (rid, cx, cy)

    # Wyczyść body
    for el in list(body): body.remove(el)

    # Generuj protokoły
    first = True
    for group_name, matches in sheets_data:
        for match in matches:
            if not first:
                body.append(_make_page_break_para())
            first = False

            cloned = [copy.deepcopy(el) for el in template_elements]
            _fill_protocol(cloned, match)

            # Zbuduj zawartość lewego obszaru "Wyniki turnieju"
            order = image_order if image_order else (
                (['qr'] if qr_rid_info else []) +
                sorted(logo_rids.keys())
            )
            items = []
            for key in order:
                if key == 'qr' and qr_rid_info:
                    items.append(_make_image_paragraph(*qr_rid_info, align='center'))
                elif key in logo_rids:
                    items.append(_make_image_paragraph(*logo_rids[key], align='center'))
            # Dodaj napis "Wyniki turnieju" jako ostatni
            items.append(_make_text_para('Wyniki turnieju', size=20, bold=True))

            _populate_left_area(cloned, items)

            for el in cloned:
                body.append(el)

    if sectPr is not None:
        body.append(sectPr)

    doc_out = etree.tostring(doc_root, xml_declaration=True,
                             encoding='UTF-8', standalone=True)
    rels_out = etree.tostring(rels_root, xml_declaration=True,
                              encoding='UTF-8', standalone=True)

    zout_buf = io.BytesIO()
    zout = zipfile.ZipFile(zout_buf, 'w', compression=zipfile.ZIP_DEFLATED)
    skip = {'word/document.xml','word/_rels/document.xml.rels'} | {f'word/{f}' for f in media_files}
    for item in zin.infolist():
        if item.filename in skip:
            continue
        if item.filename == 'word/fontTable.xml':
            ft = zin.read(item.filename).decode('utf-8')
            for fn in ('Aptos', 'Aptos Display'):
                pat = f'<w:font w:name="{fn}">'
                if pat in ft:
                    after = ft.split(pat,1)[1][:200]
                    if '<w:altName' not in after:
                        ft = ft.replace(pat, pat+'<w:altName w:val="Calibri"/>', 1)
            if 'w:name="Aptos Narrow"' not in ft:
                add = ('<w:font w:name="Aptos Narrow"><w:altName w:val="Calibri"/>'
                       '<w:charset w:val="00"/><w:family w:val="swiss"/>'
                       '<w:pitch w:val="variable"/></w:font>')
                ft = ft.replace('</w:fonts>', add + '</w:fonts>')
            zout.writestr(item, ft.encode('utf-8'))
        elif '[Content_Types].xml' in item.filename:
            ct = zin.read(item.filename).decode('utf-8')
            if 'image/png' not in ct:
                ct = ct.replace('</Types>',
                    '<Default Extension="png" ContentType="image/png"/></Types>')
            zout.writestr(item, ct.encode('utf-8'))
        else:
            zout.writestr(item, zin.read(item.filename))
    zout.writestr('word/document.xml', doc_out)
    zout.writestr('word/_rels/document.xml.rels', rels_out)
    for fname, data in media_files.items():
        zout.writestr(f'word/{fname}', data)
    zout.close(); zin.close()
    return zout_buf.getvalue()
