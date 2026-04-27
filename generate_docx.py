"""
generate_docx.py – Generator protokołów meczowych Mölkky.

Strategia: szablon Grupa_IND.docx jest klonowany 1:1 dla każdego meczu.
Wypełniamy tylko 6 komórek w tabeli nagłówkowej:
  Tor, Godzina, Grupa, Mecz #, Zawodnik 1, Zawodnik 2.
Reszta (tabele, formatowanie, fonty, marginesy) pozostaje dokładnie jak w szablonie.
"""

import io, csv, re, copy, zipfile, string
from urllib.parse import quote
import requests
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

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
    if not (tor and tor.strip().isdigit()):
        return False
    if not (godz and re.match(r'^\d{1,2}:\d{2}$', godz.strip())):
        return False
    if not z1 or not z2:
        return False
    return True

def parse_group_rows(rows):
    if not rows:
        return []
    header_idx, header = None, []
    for i, row in enumerate(rows):
        norm = [c.strip().lower() for c in row]
        if 'tor' in norm:
            header_idx = i; header = norm; break
    if header_idx is None:
        return []
    raw_header = rows[header_idx]
    def ci(name):
        try: return header.index(name)
        except ValueError: return None
    col_tor  = ci('tor')
    col_godz = ci('godzina')
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
        if not any(c.strip() for c in row):
            continue
        def g(c):
            if c is None or c >= len(row): return ''
            return row[c].strip()
        tor  = g(col_tor); godz = g(col_godz)
        z1   = g(col_z1);  z2   = g(col_z2)
        mecz = g(col_mecz)
        if not _is_valid_match_row(tor, godz, z1, z2):
            continue
        matches.append({'tor':tor,'godz':godz,'grupa':grupa,
                        'mecz':mecz,'z1':z1,'z2':z2})
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
    info.append(f"📋 Mapa GID: {len(gid_map)} zakładek")
    info.append("")
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


# ─── XML helpers: wstawianie tekstu do komórki ───────────────────────────────

def _set_cell_value(tc, text, *, bold=True, size=28, align='center'):
    """
    Wstawia tekst do KOMÓRKI istniejącego szablonu zachowując wszystko inne.
    Czyści wszystkie paragrafy w komórce, tworzy jeden nowy z tekstem.
    """
    # Usuń wszystkie istniejące paragrafy
    for p in tc.findall(wt('p')):
        tc.remove(p)
    # Dodaj nowy paragraf
    p = etree.SubElement(tc, wt('p'))
    pPr = etree.SubElement(p, wt('pPr'))
    if align:
        jc = etree.SubElement(pPr, wt('jc'))
        jc.set(f'{{{W}}}val', align)
    if not text:
        return
    r = etree.SubElement(p, wt('r'))
    rPr = etree.SubElement(r, wt('rPr'))
    fonts = etree.SubElement(rPr, wt('rFonts'))
    # Aptos jako preferowany, Calibri jako fallback (kompatybilność ze starszym Wordem)
    fonts.set(f'{{{W}}}ascii',     'Aptos')
    fonts.set(f'{{{W}}}hAnsi',     'Aptos')
    fonts.set(f'{{{W}}}eastAsia',  'Aptos')
    fonts.set(f'{{{W}}}cs',        'Aptos')
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
    """
    Wypełnia 1 protokół: znajduje pierwszą tabelę (nagłówkową) wśród
    sklonowanych elementów i wstawia dane do odpowiednich komórek.
    Tabela nagłówkowa ma 5 wierszy. Wiersz 0: Tor|val|Godz|val|Gr|val|Mecz#|val.
    Wiersz 3: Zawodnik 1 (komórka 0). Wiersz 4: Zawodnik 2 (komórka 0).
    """
    tbls = [el for el in elements if el.tag == wt('tbl')]
    if not tbls:
        return
    rows = tbls[0].findall(wt('tr'))
    if len(rows) > 0:
        tcs = rows[0].findall(wt('tc'))
        # tcs[0]=Tor label, tcs[1]=val_tor, tcs[2]=Godz label, tcs[3]=val_godz,
        # tcs[4]=Gr label, tcs[5]=val_gr, tcs[6]=Mecz# label, tcs[7]=val_mecz
        if len(tcs) > 1: _set_cell_value(tcs[1], match.get('tor',''),  size=28)
        if len(tcs) > 3: _set_cell_value(tcs[3], match.get('godz',''), size=28)
        if len(tcs) > 5: _set_cell_value(tcs[5], match.get('grupa',''),size=28)
        if len(tcs) > 7: _set_cell_value(tcs[7], match.get('mecz',''), size=28)
    # Zawodnicy
    if len(rows) > 3:
        tcs = rows[3].findall(wt('tc'))
        if tcs: _set_cell_value(tcs[0], match.get('z1',''), size=24, align='right')
    if len(rows) > 4:
        tcs = rows[4].findall(wt('tc'))
        if tcs: _set_cell_value(tcs[0], match.get('z2',''), size=24, align='right')


# ─── Build document ───────────────────────────────────────────────────────────

def build_document(sheet_id, sheets_url, sheets_data, logos=None,
                   tournament_name=None):
    """
    Klonuje szablon 1:1 raz na każdy mecz, podmienia wartości w 6 komórkach.
    `logos` i `tournament_name` zarezerwowane dla przyszłych wersji.
    """
    import os
    tpl_path = os.path.join(os.path.dirname(__file__), 'Grupa_IND.docx')
    with open(tpl_path, 'rb') as f:
        tpl_bytes = f.read()

    zin = zipfile.ZipFile(io.BytesIO(tpl_bytes))
    doc_xml = zin.read('word/document.xml')
    doc_root = etree.fromstring(doc_xml)
    body = doc_root.find(wt('body'))

    # ── Fix 1: marginesy 720 DXA (1.27cm) z każdej strony
    sectPr_check = body.find(wt('sectPr'))
    if sectPr_check is not None:
        pgMar = sectPr_check.find(wt('pgMar'))
        if pgMar is not None:
            for side in ('top','bottom','left','right'):
                pgMar.set(f'{{{W}}}{side}', '720')

    # ── Fix 2: pomniejsz fonty etykiet z 24 (12pt) → 18 (9pt)
    # żeby "Tor", "Godzina", "Mecz #", "Wygrane sety" mieściły się w 1 linii
    # nawet w przypadku fallback fontu (Calibri/Carlito).
    LABEL_TEXTS = {'Tor','Godzina','Grupa','Mecz','#','Punkty','SET 1','SET 2',
                    'Wygrane','sety','Podpis'}
    for r in body.iter(wt('r')):
        ts = r.findall(wt('t'))
        if not ts:
            continue
        text_content = ''.join((t.text or '') for t in ts).strip()
        if text_content in LABEL_TEXTS:
            sz_el  = r.find(f'{wt("rPr")}/{wt("sz")}')
            szCs_el = r.find(f'{wt("rPr")}/{wt("szCs")}')
            if sz_el is not None and sz_el.get(f'{{{W}}}val') == '24':
                sz_el.set(f'{{{W}}}val', '18')
            if szCs_el is not None and szCs_el.get(f'{{{W}}}val') == '24':
                szCs_el.set(f'{{{W}}}val', '18')

    # ── Fix 3: zrównaj szerokość tabeli 1 do tabeli 2 (9090 → 9690 DXA)
    # Rozszerzamy ostatnią kolumnę "Podpis" (gridCol[11]: 1260 → 1860)
    # i wyrównujemy SET 1 (1050) z SET 2 (1080) → oba 1080.
    first_tbl = body.find(wt('tbl'))
    if first_tbl is not None:
        gcs = first_tbl.findall(f'{wt("tblGrid")}/{wt("gridCol")}')
        if len(gcs) == 12:
            gcs[4].set(f'{{{W}}}w', '720')   # SET 1 części: było 690 → 720
            gcs[5].set(f'{{{W}}}w', '360')   # SET 1 część 2 (bez zmian)
            gcs[11].set(f'{{{W}}}w', '1860') # Podpis: 1260 → 1860 (+600)
            # Update tcW dla komórek: SET 1 (cur 1050→1080), Podpis (cur 2220 → 2820)
            for row in first_tbl.findall(wt('tr')):
                for tc in row.findall(wt('tc')):
                    tcPr = tc.find(wt('tcPr'))
                    if tcPr is None: continue
                    tcW = tcPr.find(wt('tcW'))
                    gs  = tcPr.find(wt('gridSpan'))
                    if tcW is None: continue
                    cur_w = int(tcW.get(f'{{{W}}}w','0'))
                    span  = int(gs.get(f'{{{W}}}val','1')) if gs is not None else 1
                    # SET 1 komórka (1050, span=2) → 1080
                    if cur_w == 1050 and span == 2:
                        tcW.set(f'{{{W}}}w', '1080')
                    # Podpis komórka (2220, span=3) → 2820
                    elif cur_w == 2220 and span == 3:
                        tcW.set(f'{{{W}}}w', '2820')

    # ── Fix 4: dodaj alias fontu Aptos→Calibri w fontTable (Calibri jest
    # dostępny wszędzie i sans-serif, dzięki temu unikamy szeryfowego fallback).
    # Tu nie modyfikujemy fontTable.xml bo sandbox tego nie potrzebuje
    # — ale dodajemy fallback inline w każdym runie:
    # Robimy to później przy zapisie (modyfikujemy fontTable.xml)

    # Wyciągnij sectPr (musi zostać na końcu) i zachowaj template
    sectPr = body.find(wt('sectPr'))
    template_elements = [copy.deepcopy(el) for el in body if el.tag != wt('sectPr')]

    # Pomiń końcowy pusty paragraf jeśli go ma (powoduje pustą stronę)
    while template_elements and template_elements[-1].tag == wt('p'):
        last = template_elements[-1]
        texts = last.findall(f'.//{wt("t")}')
        if not texts or not any((t.text or '').strip() for t in texts):
            template_elements.pop()
        else:
            break

    # Wyczyść body
    for el in list(body):
        body.remove(el)

    # Generuj protokoły
    first = True
    for group_name, matches in sheets_data:
        for match in matches:
            if not first:
                body.append(_make_page_break_para())
            first = False

            # Sklonuj szablon
            cloned = [copy.deepcopy(el) for el in template_elements]
            _fill_protocol(cloned, match)
            for el in cloned:
                body.append(el)

    # Przywróć sectPr na końcu
    if sectPr is not None:
        body.append(sectPr)

    # Zapakuj
    doc_out = etree.tostring(doc_root, xml_declaration=True,
                             encoding='UTF-8', standalone=True)

    zout_buf = io.BytesIO()
    zout = zipfile.ZipFile(zout_buf, 'w', compression=zipfile.ZIP_DEFLATED)
    for item in zin.infolist():
        if item.filename == 'word/document.xml':
            zout.writestr(item, doc_out)
        elif item.filename == 'word/fontTable.xml':
            # Dodaj <w:altName w:val="Calibri"/> dla fontów Aptos
            # żeby Word/LibreOffice używał Calibri zamiast szeryfowego fallback.
            ft_xml = zin.read(item.filename).decode('utf-8')
            for font_name in ('Aptos', 'Aptos Display', 'Aptos Narrow'):
                pattern = f'<w:font w:name="{font_name}">'
                replacement = f'<w:font w:name="{font_name}"><w:altName w:val="Calibri"/>'
                if pattern in ft_xml and '<w:altName' not in ft_xml.split(pattern, 1)[1][:200]:
                    ft_xml = ft_xml.replace(pattern, replacement, 1)
                # Dla Aptos Narrow i Display które nie ma w fontTable, dodaj jako nowe <w:font>
            # Dodaj brakujące fonty:
            for font_name in ('Aptos Narrow',):
                if f'w:name="{font_name}"' not in ft_xml:
                    new_font = (
                        f'<w:font w:name="{font_name}"><w:altName w:val="Calibri"/>'
                        f'<w:charset w:val="00"/><w:family w:val="swiss"/>'
                        f'<w:pitch w:val="variable"/></w:font>'
                    )
                    ft_xml = ft_xml.replace('</w:fonts>', new_font + '</w:fonts>')
            zout.writestr(item, ft_xml.encode('utf-8'))
        else:
            zout.writestr(item, zin.read(item.filename))
    zout.close(); zin.close()
    return zout_buf.getvalue()
