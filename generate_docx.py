"""
generate_docx.py
Pobiera dane z Google Sheets i generuje .docx z protokołami meczowymi.
Odwzorowuje dokładnie XML z szablonu Grupa_IND.docx.
"""

import io
import csv
import re
import copy
import zipfile
import requests
from lxml import etree

# ─── Namespace map (z oryginalnego szablonu) ─────────────────────────────────
NS = {
    'w':   'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r':   'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'wp':  'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a':   'http://schemas.drawingml/2006/main',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
}

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

def wtag(name): return f'{{{W}}}{name}'

# ─── Google Sheets helpers ─────────────────────────────────────────────────────

def fetch_sheet_csv(sheet_id, gid):
    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"
    r = requests.get(url, timeout=20)
    r.raise_for_status()
    return list(csv.reader(io.StringIO(r.text)))


def get_group_sheet_ids(sheet_id):
    """Zwraca listę (name, gid) dla zakładek Gr. A – Gr. P"""
    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/edit"
    r = requests.get(url, timeout=20)
    r.raise_for_status()
    # Szukamy w HTML: ["Gr. A",...,123456]
    found = re.findall(r'\["(Gr\.\s*[A-Z]+)"[^\]]*?,(\d{5,})\]', r.text)
    if not found:
        found = re.findall(r'"(Gr\.\s*[A-Z]+)"[^,]*,[^,]*,[^,]*,(\d{5,})', r.text)
    return found  # [(name, gid_str), ...]


def parse_group_rows(rows):
    """Parsuje CSV zakładki grupy → lista słowników meczów"""
    if not rows:
        return []
    # Znajdź wiersz nagłówkowy z 'Tor'
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

    ci = {
        'tor':   col('tor'),
        'godz':  col('godzina'),
        'grupa': col('grupa'),
        'mecz':  col('mecz'),
        'z1':    col('zawodnik 1'),
        'z2':    col('zawodnik 2'),
    }

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
        matches.append({
            'tor':    tor,
            'godz':   g('godz'),
            'grupa':  g('grupa'),
            'mecz':   g('mecz'),
            'z1':     z1,
            'z2':     g('z2'),
        })
    return matches


def fetch_all_group_sheets(sheet_id):
    """→ [(group_name, [match_dict, ...]), ...]"""
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


# ─── XML builder ──────────────────────────────────────────────────────────────

def _e(tag, attrib=None, text=None):
    el = etree.Element(wtag(tag), attrib={
        (f'{{{NS["w"]}}}{k[2:]}' if k.startswith('w:') else k): v
        for k, v in (attrib or {}).items()
    })
    if text:
        el.text = text
    return el


def make_run(text, bold=False, size=24, font='Aptos', lang='pl-PL'):
    r = etree.Element(wtag('r'))
    rPr = etree.SubElement(r, wtag('rPr'))
    fonts_el = etree.SubElement(rPr, wtag('rFonts'))
    for attr in ('ascii','hAnsi','eastAsia','cs'):
        fonts_el.set(f'{{{W}}}{attr}', font)
    b_val = '1' if bold else '0'
    for tag in ('b','bCs'):
        etree.SubElement(rPr, wtag(tag)).set(f'{{{W}}}val', b_val)
    for tag in ('i','iCs'):
        etree.SubElement(rPr, wtag(tag)).set(f'{{{W}}}val', '0')
    sz = etree.SubElement(rPr, wtag('sz'))
    sz.set(f'{{{W}}}val', str(size))
    szCs = etree.SubElement(rPr, wtag('szCs'))
    szCs.set(f'{{{W}}}val', str(size))
    if lang:
        lg = etree.SubElement(rPr, wtag('lang'))
        lg.set(f'{{{W}}}val', lang)
    t = etree.SubElement(r, wtag('t'))
    t.text = text
    if text and (text[0] == ' ' or text[-1] == ' '):
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    return r


def make_para(children=None, align=None, space_before=None, space_after=None):
    p = etree.Element(wtag('p'))
    pPr = etree.SubElement(p, wtag('pPr'))
    if align:
        jc = etree.SubElement(pPr, wtag('jc'))
        jc.set(f'{{{W}}}val', align)
    if space_before is not None or space_after is not None:
        sp = etree.SubElement(pPr, wtag('spacing'))
        if space_before is not None:
            sp.set(f'{{{W}}}before', str(space_before))
        if space_after is not None:
            sp.set(f'{{{W}}}after', str(space_after))
    for c in (children or []):
        p.append(c)
    return p


def make_cell(width, children=None, span=1, bg=None, borders_nil=False,
              valign=None, margins=None):
    tc = etree.Element(wtag('tc'))
    tcPr = etree.SubElement(tc, wtag('tcPr'))
    w_el = etree.SubElement(tcPr, wtag('tcW'))
    w_el.set(f'{{{W}}}w', str(width))
    w_el.set(f'{{{W}}}type', 'dxa')
    if span > 1:
        gs = etree.SubElement(tcPr, wtag('gridSpan'))
        gs.set(f'{{{W}}}val', str(span))
    if borders_nil:
        bdr = etree.SubElement(tcPr, wtag('tcBorders'))
        for side in ('top','left','bottom','right'):
            s = etree.SubElement(bdr, wtag(side))
            s.set(f'{{{W}}}val', 'nil')
    if bg:
        shd = etree.SubElement(tcPr, wtag('shd'))
        shd.set(f'{{{W}}}val', 'clear')
        shd.set(f'{{{W}}}color', 'auto')
        shd.set(f'{{{W}}}fill', bg)
    if valign:
        va = etree.SubElement(tcPr, wtag('vAlign'))
        va.set(f'{{{W}}}val', valign)
    m = margins or {'left': 105, 'right': 105}
    mar = etree.SubElement(tcPr, wtag('tcMar'))
    for side, val in m.items():
        s = etree.SubElement(mar, wtag(side))
        s.set(f'{{{W}}}w', str(val))
        s.set(f'{{{W}}}type', 'dxa')
    for child in (children or []):
        tc.append(child)
    return tc


def empty_para(size=24):
    return make_para([make_run('', size=size)])


# ─── Table 1: Nagłówek (Tor / Godzina / Grupa / Mecz#) ──────────────────────
# GridCols z oryginału: 555 960 1125 1275 690 360 630 450 825 420 540 1260
# Total: 9090
# Wiersz 1 (labels): Tor(555) | val(960+nil) | Godzina(1125) | val(1275+690 span2 nil) |
#                    Grupa(360+630 span2) | val(450+825 span2 nil) | Mecz#(420+540+1260 span3) | val nil

def make_header_table(tor, godz, grupa, mecz):
    tbl = etree.Element(wtag('tbl'))
    tblPr = etree.SubElement(tbl, wtag('tblPr'))
    ts = etree.SubElement(tblPr, wtag('tblStyle'))
    ts.set(f'{{{W}}}val', 'TableGrid')
    bi = etree.SubElement(tblPr, wtag('bidiVisual'))
    bi.set(f'{{{W}}}val', '0')
    tw = etree.SubElement(tblPr, wtag('tblW'))
    tw.set(f'{{{W}}}w', '0')
    tw.set(f'{{{W}}}type', 'auto')
    # No outer borders on header table
    tblBdr = etree.SubElement(tblPr, wtag('tblBorders'))
    for side in ('top','left','bottom','right'):
        s = etree.SubElement(tblBdr, wtag(side))
        s.set(f'{{{W}}}val', 'single')
        s.set(f'{{{W}}}sz', '6')

    tblGrid = etree.SubElement(tbl, wtag('tblGrid'))
    for w in [555,960,1125,1275,690,360,630,450,825,420,540,1260]:
        gc = etree.SubElement(tblGrid, wtag('gridCol'))
        gc.set(f'{{{W}}}w', str(w))

    def label_row():
        tr = etree.Element(wtag('tr'))
        trPr = etree.SubElement(tr, wtag('trPr'))
        trH = etree.SubElement(trPr, wtag('trHeight'))
        trH.set(f'{{{W}}}val', '300')

        # Tor
        tr.append(make_cell(555, borders_nil=True, valign='bottom', children=[
            make_para([make_run('Tor', size=24)], align='center')
        ]))
        # val tor (puste)
        tr.append(make_cell(960, borders_nil=True, valign='bottom', children=[empty_para(28)]))
        # Godzina
        tr.append(make_cell(1125, borders_nil=True, valign='bottom', children=[
            make_para([make_run('Godzina', size=24)], align='center')
        ]))
        # val godzina (span 2)
        tr.append(make_cell(1275+690, span=2, borders_nil=True, valign='bottom', children=[empty_para(28)]))
        # Grupa (span 2)
        tr.append(make_cell(360+630, span=2, borders_nil=True, valign='bottom', children=[
            make_para([make_run('Grupa', size=24)], align='center')
        ]))
        # val grupa (span 2)
        tr.append(make_cell(450+825, span=2, borders_nil=True, valign='bottom', children=[empty_para(28)]))
        # Mecz # (span 3)
        mecz_r = make_run('Mecz', size=24)
        hash_r = make_run(' #', bold=True, size=24)
        tr.append(make_cell(420+540+1260, span=3, borders_nil=True, valign='bottom', children=[
            make_para([mecz_r, hash_r], align='center')
        ]))
        return tr

    def value_row():
        tr = etree.Element(wtag('tr'))
        trPr = etree.SubElement(tr, wtag('trPr'))
        trH = etree.SubElement(trPr, wtag('trHeight'))
        trH.set(f'{{{W}}}val', '425')

        vals = [
            (555,  tor,   False, 28),
            (960,  '',    False, 28),
            (1125, godz,  False, 28),
        ]
        for w, txt, bold, sz in vals:
            tr.append(make_cell(w, borders_nil=True, valign='bottom', children=[
                make_para([make_run(txt, size=sz, bold=bold)], align='center')
            ]))
        # val godzina span2
        tr.append(make_cell(1275+690, span=2, borders_nil=True, valign='bottom', children=[empty_para(28)]))
        # Grupa val span2
        tr.append(make_cell(360+630, span=2, borders_nil=True, valign='bottom', children=[
            make_para([make_run(grupa, size=28, bold=True)], align='center')
        ]))
        # val span2
        tr.append(make_cell(450+825, span=2, borders_nil=True, valign='bottom', children=[empty_para(28)]))
        # Mecz val span3
        tr.append(make_cell(420+540+1260, span=3, borders_nil=True, valign='bottom', children=[
            make_para([make_run(mecz, size=36, bold=True)], align='center')
        ]))
        return tr

    tbl.append(label_row())
    tbl.append(value_row())
    return tbl


# ─── Table 2: Zawodnicy + wyniki setów ───────────────────────────────────────
# GridCols z oryginału row 3-5: szerokości komórek (zawodnik | set1 | set2 | sety | podpis)
# Z XML: tbl1 rows 3,4,5 → 12 kolumn
# Praktycznie: zawodnik(555+960+1125) span3 | set1(1275+690) span2 | set2(360+630) span2 | sety(450+825) span2 | podpis(420+540+1260) span3

ZAWODNIK_W = 555+960+1125   # 2640
SET1_W = 1275+690           # 1965
SET2_W = 360+630            # 990
SETY_W = 450+825            # 1275
PODPIS_W = 420+540+1260     # 2220

BG_GRAY = 'D9D9D9'
BG_BLUE = 'DDEEFF'

def make_players_table(z1, z2):
    tbl = etree.Element(wtag('tbl'))
    tblPr = etree.SubElement(tbl, wtag('tblPr'))
    ts = etree.SubElement(tblPr, wtag('tblStyle'))
    ts.set(f'{{{W}}}val', 'TableGrid')
    bi = etree.SubElement(tblPr, wtag('bidiVisual'))
    bi.set(f'{{{W}}}val', '0')
    tw = etree.SubElement(tblPr, wtag('tblW'))
    tw.set(f'{{{W}}}w', '0')
    tw.set(f'{{{W}}}type', 'auto')

    tblGrid = etree.SubElement(tbl, wtag('tblGrid'))
    for w in [555,960,1125,1275,690,360,630,450,825,420,540,1260]:
        gc = etree.SubElement(tblGrid, wtag('gridCol'))
        gc.set(f'{{{W}}}w', str(w))

    def hdr_row():
        tr = etree.Element(wtag('tr'))
        # pusta lewa
        tr.append(make_cell(ZAWODNIK_W, span=3, bg=BG_GRAY, children=[empty_para()]))
        # Punkty SET 1
        tr.append(make_cell(SET1_W, span=2, bg=BG_GRAY, children=[
            make_para([make_run('Punkty', size=18, bold=True),
                       make_run(' SET 1', size=18, bold=True)], align='center')
        ]))
        # Punkty SET 2
        tr.append(make_cell(SET2_W, span=2, bg=BG_GRAY, children=[
            make_para([make_run('Punkty', size=18, bold=True),
                       make_run(' SET 2', size=18, bold=True)], align='center')
        ]))
        # Wygrane sety
        tr.append(make_cell(SETY_W, span=2, bg=BG_GRAY, children=[
            make_para([make_run('Wygrane', size=18, bold=True),
                       make_run(' sety', size=18, bold=True)], align='center')
        ]))
        # Podpis
        tr.append(make_cell(PODPIS_W, span=3, bg=BG_GRAY, children=[
            make_para([make_run('Podpis', size=18, bold=True)], align='center')
        ]))
        return tr

    def player_row(name, bg=None):
        tr = etree.Element(wtag('tr'))
        trPr = etree.SubElement(tr, wtag('trPr'))
        trH = etree.SubElement(trPr, wtag('trHeight'))
        trH.set(f'{{{W}}}val', '500')
        tr.append(make_cell(ZAWODNIK_W, span=3, bg=bg, children=[
            make_para([make_run(name, size=24, bold=True)])
        ]))
        for w, sp in [(SET1_W,2),(SET2_W,2),(SETY_W,2),(PODPIS_W,3)]:
            tr.append(make_cell(w, span=sp, children=[empty_para()]))
        return tr

    tbl.append(hdr_row())
    tbl.append(player_row(z1, bg=BG_BLUE))
    tbl.append(player_row(z2))
    return tbl


# ─── Table 3: Wyniki turnieju ─────────────────────────────────────────────────
# GridCols z oryginału: 2970 840 735 735 735 735 735 735 735 735
# Total: 9690
# Header row: Wyniki turnieju(2970) | IMIONA(840) | SET1(735) | SUMA(735) | SET2(735) | SUMA ...
# 18 empty rows + WYNIK row

SCORE_COLS = [2970, 840, 735, 735, 735, 735, 735, 735, 735, 735]
SCORE_ROWS = 18

def make_score_table():
    tbl = etree.Element(wtag('tbl'))
    tblPr = etree.SubElement(tbl, wtag('tblPr'))
    ts = etree.SubElement(tblPr, wtag('tblStyle'))
    ts.set(f'{{{W}}}val', 'TableGrid')
    bi = etree.SubElement(tblPr, wtag('bidiVisual'))
    bi.set(f'{{{W}}}val', '0')
    tw = etree.SubElement(tblPr, wtag('tblW'))
    tw.set(f'{{{W}}}w', '0')
    tw.set(f'{{{W}}}type', 'auto')

    tblGrid = etree.SubElement(tbl, wtag('tblGrid'))
    for w in SCORE_COLS:
        gc = etree.SubElement(tblGrid, wtag('gridCol'))
        gc.set(f'{{{W}}}w', str(w))

    # Row 0: Superheader – "SET 1" and "SET 2" spanning cols
    r0 = etree.Element(wtag('tr'))
    r0.append(make_cell(SCORE_COLS[0]+SCORE_COLS[1], span=2, children=[empty_para()]))
    # SET 1 spans cols 2-3 (735+735=1470)
    r0.append(make_cell(735+735, span=2, bg=BG_GRAY, children=[
        make_para([make_run('SET 1', size=20, bold=True)], align='center')
    ]))
    r0.append(make_cell(735+735, span=2, bg=BG_GRAY, children=[
        make_para([make_run('SET 2', size=20, bold=True)], align='center')
    ]))
    # remaining 4 cells
    for _ in range(4):
        r0.append(make_cell(735, children=[empty_para()]))
    tbl.append(r0)

    # Row 1: Headers
    r1 = etree.Element(wtag('tr'))
    r1.append(make_cell(SCORE_COLS[0], bg=BG_GRAY, children=[
        make_para([make_run('Wyniki turnieju', size=20, bold=True)], align='center')
    ]))
    # IMIONA – spaced letters
    r1.append(make_cell(SCORE_COLS[1], bg=BG_GRAY, children=[
        make_para([make_run('I M I O N A', size=20, bold=True)], align='center')
    ]))
    for lbl in ['', 'S U M A', '', 'S U M A', '', 'S U M A', '', 'S U M A']:
        r1.append(make_cell(SCORE_COLS[2], bg=BG_GRAY if lbl else None, children=[
            make_para([make_run(lbl, size=18, bold=True)], align='center')
        ]))
    tbl.append(r1)

    # 18 empty rows
    for _ in range(SCORE_ROWS):
        tr = etree.Element(wtag('tr'))
        trPr = etree.SubElement(tr, wtag('trPr'))
        trH = etree.SubElement(trPr, wtag('trHeight'))
        trH.set(f'{{{W}}}val', '300')
        for w in SCORE_COLS:
            tr.append(make_cell(w, children=[empty_para()]))
        tbl.append(tr)

    # WYNIK row
    rW = etree.Element(wtag('tr'))
    rW.append(make_cell(SCORE_COLS[0], children=[empty_para()]))
    rW.append(make_cell(SCORE_COLS[1], bg=BG_GRAY, children=[
        make_para([make_run('WYNIK', size=20, bold=True)], align='center')
    ]))
    for w in SCORE_COLS[2:]:
        rW.append(make_cell(w, children=[empty_para()]))
    tbl.append(rW)

    return tbl


# ─── Image helper (logo/banner) ───────────────────────────────────────────────

def make_image_para(rel_id, cx_emu, cy_emu, align='left'):
    """Buduje paragraf z inline obrazkiem"""
    p = etree.Element(wtag('p'))
    pPr = etree.SubElement(p, wtag('pPr'))
    if align == 'center':
        jc = etree.SubElement(pPr, wtag('jc'))
        jc.set(f'{{{W}}}val', 'center')
    r = etree.SubElement(p, wtag('r'))
    drawing = etree.SubElement(r, '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}inline')
    ext = etree.SubElement(drawing, '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}extent')
    ext.set('cx', str(cx_emu))
    ext.set('cy', str(cy_emu))
    graphic = etree.SubElement(drawing,
        '{http://schemas.openxmlformats.org/drawingml/2006/main}graphic')
    gd = etree.SubElement(graphic,
        '{http://schemas.openxmlformats.org/drawingml/2006/main}graphicData')
    gd.set('uri', 'http://schemas.openxmlformats.org/drawingml/2006/picture')
    pic = etree.SubElement(gd,
        '{http://schemas.openxmlformats.org/drawingml/2006/picture}pic')
    blipFill = etree.SubElement(pic,
        '{http://schemas.openxmlformats.org/drawingml/2006/picture}blipFill')
    blip = etree.SubElement(blipFill,
        '{http://schemas.openxmlformats.org/drawingml/2006/main}blip')
    blip.set('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed', rel_id)
    return p


# ─── Build one protocol as list of XML body children ─────────────────────────

def make_protocol_elements(match, logos_rel=None):
    """Zwraca listę elementów XML (tabele + paragrafy) dla jednego protokołu"""
    els = []

    # Opcjonalne logo/banner na górze
    if logos_rel and 'banner' in logos_rel:
        rel_id, cx, cy = logos_rel['banner']
        els.append(make_image_para(rel_id, cx, cy, align='center'))

    # Logo górne lewo + prawo obok siebie
    # (uproszczone – dodajemy je kolejno, bo Word nie wspiera łatwo side-by-side inline)
    for key in ('top_left', 'top_right'):
        if logos_rel and key in logos_rel:
            rel_id, cx, cy = logos_rel[key]
            align = 'left' if key == 'top_left' else 'right'
            els.append(make_image_para(rel_id, cx, cy, align=align))

    # Tabela nagłówkowa
    els.append(make_header_table(
        match['tor'], match['godz'], match['grupa'], match['mecz']
    ))

    # Paragraf pusty między tabelami
    els.append(make_para([], space_before=0, space_after=40))

    # Tabela zawodników
    els.append(make_players_table(match['z1'], match['z2']))

    # Zasady
    els.append(make_para(
        [make_run('Każdy zawodnik zaczyna po jednym secie (w dowolnej kolejności)', size=16, lang='pl-PL')],
        space_before=40, space_after=0
    ))
    els.append(make_para(
        [make_run('Set przegrany przez 3 kolejne chybienia oznacza wynik 0:50', size=16, lang='pl-PL')],
        space_before=0, space_after=40
    ))

    # Tabela wyników
    els.append(make_score_table())

    # Logo dolne lewo
    if logos_rel and 'bottom_left' in logos_rel:
        rel_id, cx, cy = logos_rel['bottom_left']
        els.append(make_image_para(rel_id, cx, cy, align='left'))

    return els


def make_page_break():
    p = etree.Element(wtag('p'))
    pPr = etree.SubElement(p, wtag('pPr'))
    sp = etree.SubElement(pPr, wtag('spacing'))
    sp.set(f'{{{W}}}before', '0')
    sp.set(f'{{{W}}}after', '0')
    r = etree.SubElement(p, wtag('r'))
    br = etree.SubElement(r, wtag('br'))
    br.set(f'{{{W}}}type', 'page')
    return p


# ─── Build full document ──────────────────────────────────────────────────────

def build_document(sheet_id, sheets_url, sheets_data, logos=None):
    """
    sheet_id: str
    sheets_url: str
    sheets_data: [(group_name, [match_dict, ...]), ...]
    logos: dict {key: bytes} – None lub {top_left, top_right, bottom_left, banner}
    Returns: bytes (docx)
    """
    from lxml import etree
    import io, zipfile, copy

    # ── Wczytaj szablon ──────────────────────────────────────────────────────
    import importlib.resources, os
    template_path = os.path.join(os.path.dirname(__file__), 'Grupa_IND.docx')

    with open(template_path, 'rb') as f:
        template_bytes = f.read()

    zin = zipfile.ZipFile(io.BytesIO(template_bytes))
    zout_buf = io.BytesIO()
    zout = zipfile.ZipFile(zout_buf, 'w', compression=zipfile.ZIP_DEFLATED)

    # Pliki z szablonu które kopiujemy bez zmian (poza document.xml i rels)
    skip = {'word/document.xml', 'word/_rels/document.xml.rels'}
    for item in zin.infolist():
        if item.filename not in skip:
            zout.writestr(item, zin.read(item.filename))

    # ── Parsuj document.xml ──────────────────────────────────────────────────
    doc_xml = zin.read('word/document.xml')
    doc_tree = etree.fromstring(doc_xml)
    body = doc_tree.find(f'{{{W}}}body')

    # Usuń wszystkie dzieci body poza sectPr
    sect_pr = body.find(f'{{{W}}}sectPr')
    for child in list(body):
        body.remove(child)

    # ── Rels ─────────────────────────────────────────────────────────────────
    rels_xml = zin.read('word/_rels/document.xml.rels')
    rels_tree = etree.fromstring(rels_xml)
    REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'

    # Dodaj obrazki do docx i rels
    logos_rel = {}
    rel_id_counter = 100  # zaczynamy od rId100 żeby nie kolidować z istniejącymi

    if logos:
        from PIL import Image as PILImage
        for key, img_bytes in logos.items():
            if not img_bytes:
                continue
            rel_id_counter += 1
            rid = f'rId{rel_id_counter}'
            ext = 'png'
            media_name = f'media/logo_{key}.png'

            # Konwertuj do PNG
            pil_img = PILImage.open(io.BytesIO(img_bytes)).convert('RGBA')

            # Rozmiary w EMU (914400 EMU = 1 cal)
            # Banner: pełna szerokość ~16 cm, logo: max 3 cm wysokości
            if key == 'banner':
                target_w_cm = 16.0
            else:
                target_w_cm = 3.0
            target_w_emu = int(target_w_cm * 360000)
            ratio = pil_img.width / pil_img.height
            target_h_emu = int(target_w_emu / ratio)

            png_buf = io.BytesIO()
            pil_img.save(png_buf, format='PNG')
            png_buf.seek(0)

            zout.writestr(f'word/{media_name}', png_buf.read())

            # Relationship
            rel_el = etree.SubElement(rels_tree, f'{{{REL_NS}}}Relationship')
            rel_el.set('Id', rid)
            rel_el.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image')
            rel_el.set('Target', media_name)

            logos_rel[key] = (rid, target_w_emu, target_h_emu)

    # ── QR code ──────────────────────────────────────────────────────────────
    try:
        import qrcode as _qrcode
        qr = _qrcode.QRCode(version=2, box_size=6, border=2)
        qr.add_data(sheets_url)
        qr.make(fit=True)
        qr_img = qr.make_image(fill_color='black', back_color='white')
        qr_buf = io.BytesIO()
        qr_img.save(qr_buf, format='PNG')
        qr_buf.seek(0)
        qr_bytes = qr_buf.read()
        has_qr = True
    except ImportError:
        has_qr = False

    if has_qr:
        qr_rid_n = rel_id_counter + 1
        qr_rid = f'rId{qr_rid_n}'
        zout.writestr('word/media/qrcode.png', qr_bytes)
        rel_el = etree.SubElement(rels_tree, f'{{{REL_NS}}}Relationship')
        rel_el.set('Id', qr_rid)
        rel_el.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image')
        rel_el.set('Target', 'media/qrcode.png')
        qr_emu = int(4.0 * 360000)  # 4 cm

    # ── Strona tytułowa ──────────────────────────────────────────────────────
    def title_para(text, size=40, bold=False, align='center', color=None):
        p = etree.Element(wtag('p'))
        pPr = etree.SubElement(p, wtag('pPr'))
        jc = etree.SubElement(pPr, wtag('jc'))
        jc.set(f'{{{W}}}val', align)
        r = make_run(text, size=size, bold=bold)
        if color:
            rpr = r.find(wtag('rPr'))
            col_el = etree.SubElement(rpr, wtag('color'))
            col_el.set(f'{{{W}}}val', color)
        p.append(r)
        return p

    body.append(title_para('GP2 2026', size=52, bold=True))
    body.append(title_para('Protokoły meczowe – faza grupowa', size=28))
    body.append(make_para([], space_before=400, space_after=200))
    body.append(title_para('Arkusz wyników:', size=22))

    if has_qr:
        body.append(make_image_para(qr_rid, qr_emu, qr_emu, align='center'))

    body.append(title_para(sheets_url, size=16, color='0000CC'))
    body.append(make_page_break())

    # ── Protokoły ─────────────────────────────────────────────────────────────
    first = True
    for group_name, matches in sheets_data:
        if not matches:
            continue
        for i, match in enumerate(matches):
            if not first:
                body.append(make_page_break())
            first = False
            for el in make_protocol_elements(match, logos_rel):
                body.append(el)

    # Przywróć sectPr
    if sect_pr is not None:
        body.append(sect_pr)

    # ── Serializuj ───────────────────────────────────────────────────────────
    doc_out = etree.tostring(doc_tree, xml_declaration=True, encoding='UTF-8', standalone=True)
    zout.writestr('word/document.xml', doc_out)

    rels_out = etree.tostring(rels_tree, xml_declaration=True, encoding='UTF-8', standalone=True)
    zout.writestr('word/_rels/document.xml.rels', rels_out)

    zout.close()
    return zout_buf.getvalue()
