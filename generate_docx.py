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

def _make_anchored_text(text, posY_emu, posX_emu, width_emu, height_emu, size=18, bold=True):
    """
    Pływające pole tekstowe (text box / shape with text).
    Używamy <w:drawing> z <wp:anchor> i wewnątrz <wps:txbx>.
    """
    uid = abs(hash(text)) % 9000 + 7000
    bold_xml = '<w:b/><w:bCs/>' if bold else ''
    return etree.fromstring(f'''<w:drawing xmlns:w="{W}"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
  <wp:anchor distT="0" distB="0" distL="0" distR="0"
             simplePos="0" relativeHeight="{uid}"
             behindDoc="0" locked="0" layoutInCell="0" allowOverlap="1">
    <wp:simplePos x="0" y="0"/>
    <wp:positionH relativeFrom="column">
      <wp:posOffset>{int(posX_emu)}</wp:posOffset>
    </wp:positionH>
    <wp:positionV relativeFrom="paragraph">
      <wp:posOffset>{int(posY_emu)}</wp:posOffset>
    </wp:positionV>
    <wp:extent cx="{int(width_emu)}" cy="{int(height_emu)}"/>
    <wp:effectExtent l="0" t="0" r="0" b="0"/>
    <wp:wrapNone/>
    <wp:docPr id="{uid}" name="TextBox {uid}"/>
    <wp:cNvGraphicFramePr/>
    <a:graphic>
      <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
        <wps:wsp>
          <wps:cNvSpPr txBox="1"/>
          <wps:spPr>
            <a:xfrm><a:off x="0" y="0"/>
              <a:ext cx="{int(width_emu)}" cy="{int(height_emu)}"/>
            </a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
            <a:noFill/>
            <a:ln><a:noFill/></a:ln>
          </wps:spPr>
          <wps:txbx>
            <w:txbxContent>
              <w:p>
                <w:pPr><w:jc w:val="center"/>
                  <w:spacing w:before="0" w:after="0"/>
                </w:pPr>
                <w:r>
                  <w:rPr>
                    <w:rFonts w:ascii="Aptos" w:hAnsi="Aptos" w:eastAsia="Aptos" w:cs="Aptos"/>
                    {bold_xml}
                    <w:sz w:val="{size}"/>
                    <w:szCs w:val="{size}"/>
                  </w:rPr>
                  <w:t>{text}</w:t>
                </w:r>
              </w:p>
            </w:txbxContent>
          </wps:txbx>
          <wps:bodyPr rot="0" spcFirstLastPara="0" vertOverflow="visible"
                      horzOverflow="visible" wrap="square" lIns="0" tIns="0" rIns="0" bIns="0"
                      anchor="t" anchorCtr="0"/>
        </wps:wsp>
      </a:graphicData>
    </a:graphic>
  </wp:anchor>
</w:drawing>''')


def _make_anchored_image(rel_id, cx_emu, cy_emu, posY_emu, posX_emu=0):
    """
    Tworzy <w:drawing> z pływającym (anchor) obrazem, kotwiczonym w paragrafie.
    posX_emu: pozycja X względem column (lewa krawędź obszaru tekstu)
    posY_emu: pozycja Y względem paragraph (od góry akapitu w którym jest kotwiczony)
    Obraz "behindDoc=0" (przed tekstem), wrap=none (tekst go ignoruje).
    """
    uid = abs(hash(rel_id)) % 9000 + 5000
    return etree.fromstring(f'''<w:drawing xmlns:w="{W}"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
  <wp:anchor distT="0" distB="0" distL="0" distR="0"
             simplePos="0" relativeHeight="{uid}"
             behindDoc="0" locked="0" layoutInCell="0" allowOverlap="1">
    <wp:simplePos x="0" y="0"/>
    <wp:positionH relativeFrom="column">
      <wp:posOffset>{int(posX_emu)}</wp:posOffset>
    </wp:positionH>
    <wp:positionV relativeFrom="paragraph">
      <wp:posOffset>{int(posY_emu)}</wp:posOffset>
    </wp:positionV>
    <wp:extent cx="{int(cx_emu)}" cy="{int(cy_emu)}"/>
    <wp:effectExtent l="0" t="0" r="0" b="0"/>
    <wp:wrapNone/>
    <wp:docPr id="{uid}" name="Picture {uid}"/>
    <wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>
    <a:graphic>
      <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
        <pic:pic>
          <pic:nvPicPr>
            <pic:cNvPr id="{uid}" name="Picture"/>
            <pic:cNvPicPr/>
          </pic:nvPicPr>
          <pic:blipFill>
            <a:blip r:embed="{rel_id}"/>
            <a:stretch><a:fillRect/></a:stretch>
          </pic:blipFill>
          <pic:spPr>
            <a:xfrm>
              <a:off x="0" y="0"/>
              <a:ext cx="{int(cx_emu)}" cy="{int(cy_emu)}"/>
            </a:xfrm>
            <a:prstGeom prst="rect"/>
          </pic:spPr>
        </pic:pic>
      </a:graphicData>
    </a:graphic>
  </wp:anchor>
</w:drawing>''')


def _make_image_paragraph(rel_id, cx_emu, cy_emu, align='center'):
    """Tworzy paragraf z obrazem inline - wzór sprawdzony przez python-docx."""
    return etree.fromstring(f'''<w:p xmlns:w="{W}"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
  <w:pPr><w:jc w:val="{align}"/><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
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
    """Tworzy paragraf z tekstem (zerowy padding)."""
    p = etree.Element(wt('p'))
    pPr = etree.SubElement(p, wt('pPr'))
    if align:
        jc = etree.SubElement(pPr, wt('jc'))
        jc.set(f'{{{W}}}val', align)
    sp = etree.SubElement(pPr, wt('spacing'))
    sp.set(f'{{{W}}}before','0'); sp.set(f'{{{W}}}after','0')
    sp.set(f'{{{W}}}line','240'); sp.set(f'{{{W}}}lineRule','auto')
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


def _populate_left_area(elements, cell_items, anchored_drawings):
    """
    Wstawia napis "Wyniki turnieju" jako paragraf w komórce tabeli 2 (jak we wzorcu),
    a obrazy (QR + grafiki) jako pływające obrazy KOTWICZONE WEWNĄTRZ TEJ KOMÓRKI.
    Komórka pozostaje wysokości pojedynczego wiersza, obrazy "wystają" poza jej granice
    ale layoutInCell=0 sprawia że nie rozciągają komórki.
    """
    tbls = [el for el in elements if el.tag == wt('tbl')]
    if len(tbls) < 2: return

    score_tbl = tbls[1]
    rows = score_tbl.findall(wt('tr'))
    if len(rows) < 2: return

    target_cell = rows[1].findall(wt('tc'))[0]
    # Wyczyść komórkę
    for p in target_cell.findall(wt('p')):
        target_cell.remove(p)

    # Wstaw paragraf z pływającymi obrazami (wszystkie w 1 paragrafie)
    p = etree.SubElement(target_cell, wt('p'))
    pPr = etree.SubElement(p, wt('pPr'))
    sp = etree.SubElement(pPr, wt('spacing'))
    sp.set(f'{{{W}}}before','0'); sp.set(f'{{{W}}}after','0')
    sp.set(f'{{{W}}}line','240'); sp.set(f'{{{W}}}lineRule','auto')

    for drawing in anchored_drawings:
        r = etree.SubElement(p, wt('r'))
        r.append(drawing)

    # Dodaj napis "Wyniki turnieju" do paragrafu w komórce
    for item in cell_items:
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
    # Zwężamy 1. kolumnę o 280 DXA (~0.5 cm) i dodajemy wcięcie tabeli z lewej
    # tblInd=280, żeby etykieta "Tor" miała trochę odstępu od krawędzi.
    # Tabelę 1 dopasowujemy do szerokości tabeli 2 (9690 DXA).
    first_tbl = body.find(wt('tbl'))
    if first_tbl is not None:
        gcs = first_tbl.findall(f'{wt("tblGrid")}/{wt("gridCol")}')
        if len(gcs) == 12:
            # Kolumna z "Tor" — zostawiamy bez zmian (555),
            # ale zwężamy następną pustą o 280 (960→680) i dajemy tblInd 280
            gcs[1].set(f'{{{W}}}w', '680')
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
                    # Pusta kolumna po "Tor" (cur=960, span=1) → 680
                    if cur_w == 960 and span == 1:
                        tcW.set(f'{{{W}}}w', '680')
                    # SET 1 (1050, span=2) → 1080
                    elif cur_w == 1050 and span == 2:
                        tcW.set(f'{{{W}}}w', '1080')
                    # Komórka nazwiska zawodnika (cur=3915, span=4) → 3635 (3915-280)
                    elif cur_w == 3915 and span == 4:
                        tcW.set(f'{{{W}}}w', '3635')
                    # Wiersz 0 spacer (cur=1965, span=2) → 1685 (1965-280)
                    elif cur_w == 1965 and span == 2:
                        tcW.set(f'{{{W}}}w', '1685')
                    # Wiersz 1 spacer (cur=1965, span=2 też) - już objęte powyżej
                    # Komórka "Podpis" w wierszu nagłówków/zawodnika
                    elif cur_w == 1800 and span == 2:
                        tcW.set(f'{{{W}}}w', '2400')
                    elif cur_w == 2220 and span == 3:
                        tcW.set(f'{{{W}}}w', '2820')
                    elif cur_w == 1260 and span == 1:
                        tcW.set(f'{{{W}}}w', '1860')
            # Dodaj tblInd 280 DXA (~0.5 cm) z lewej
            tblPr = first_tbl.find(wt('tblPr'))
            if tblPr is not None:
                tblInd = tblPr.find(wt('tblInd'))
                if tblInd is None:
                    tblInd = etree.SubElement(tblPr, wt('tblInd'))
                tblInd.set(f'{{{W}}}w', '280')
                tblInd.set(f'{{{W}}}type', 'dxa')

    # Wyrównaj etykietę "Tor" do LEWEJ (zamiast do center)
    if first_tbl is not None:
        first_row = first_tbl.find(wt('tr'))
        if first_row is not None:
            tcs = first_row.findall(wt('tc'))
            if tcs:
                for p in tcs[0].findall(wt('p')):
                    pPr = p.find(wt('pPr'))
                    if pPr is None:
                        pPr = etree.Element(wt('pPr'))
                        p.insert(0, pPr)
                    jc = pPr.find(wt('jc'))
                    if jc is None:
                        jc = etree.SubElement(pPr, wt('jc'))
                    jc.set(f'{{{W}}}val', 'left')

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
            qr_rid_info = (rid, int(1.8*360000), int(1.8*360000))

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
            # Ograniczenie: max szerokość 2.4 cm, max wysokość 1.4 cm
            # (proporcje zachowane).
            max_w_cm = 2.4
            max_h_cm = 1.4
            ratio = pil.width / pil.height
            # Najpierw dopasuj do max szerokości
            target_w_cm = max_w_cm
            target_h_cm = target_w_cm / ratio
            # Jeśli za wysoki, dopasuj do max wysokości
            if target_h_cm > max_h_cm:
                target_h_cm = max_h_cm
                target_w_cm = target_h_cm * ratio
            cx = int(target_w_cm * 360000)
            cy = int(target_h_cm * 360000)
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

            order = image_order if image_order else (
                (['qr'] if qr_rid_info else []) +
                sorted(logo_rids.keys())
            )

            anchored = []
            current_y_emu = int(0.1 * 360000)   # zaczynamy ~1mm pod kotwicą
            spacing_emu = int(0.15 * 360000)    # 1.5mm między obrazami
            cell_width_emu = int(5.24 * 360000) # szerokość komórki "Wyniki turnieju"

            label_inserted_after = None  # po którym kluczu wstawić "Wyniki turnieju"

            for i, key in enumerate(order):
                rid_info = None
                if key == 'qr' and qr_rid_info:
                    rid_info = qr_rid_info
                elif key in logo_rids:
                    rid_info = logo_rids[key]
                if rid_info is None:
                    continue
                rid, cx, cy = rid_info
                center_x = (cell_width_emu - cx) // 2
                anchored.append(_make_anchored_image(rid, cx, cy,
                                                     posY_emu=current_y_emu,
                                                     posX_emu=center_x))
                current_y_emu += cy + spacing_emu
                # Po QR (jeśli był jednym z elementów) wstawiamy napis jako pływający text box
                if key == 'qr':
                    anchored.append(_make_anchored_text(
                        'Wyniki turnieju',
                        posY_emu=current_y_emu,
                        posX_emu=0,
                        width_emu=cell_width_emu,
                        height_emu=int(0.5*360000),
                        size=18, bold=True))
                    current_y_emu += int(0.5*360000) + spacing_emu

            # Jeśli nie ma QR, dodaj napis "Wyniki turnieju" na końcu jako floating
            if not (include_qr and qr_rid_info):
                anchored.append(_make_anchored_text(
                    'Wyniki turnieju',
                    posY_emu=current_y_emu,
                    posX_emu=0,
                    width_emu=cell_width_emu,
                    height_emu=int(0.5*360000),
                    size=18, bold=True))

            # Komórka pozostaje pusta (ma tylko paragraf-kotwicę dla pływających obrazów)
            cell_items = []
            _populate_left_area(cloned, cell_items, anchored)

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
