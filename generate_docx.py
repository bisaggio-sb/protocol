"""
generate_docx.py – Generator protokołów meczowych Mölkky.

CZWORKA layout (v cz4):
- Anchor X relativeFrom="page" + LEFT_MARGIN_EMU offset (1.27 cm) — jednoznaczne.
- QR posX=0 (col) → page-X=1.27 cm = lewa krawędź tabeli.
- Last graphic posX+w=18.46 (col) → page-X=19.73 cm = prawa krawędź tabeli.
- "Wyniki" body paragraph: ind left=0 right=9559, jc center → centrowane w 1.6 cm boxie nad QR.
- col_width_cm clamp = 18.46 (= TARGET_WIDTH) — nie skaluje grafik które już mieszczą się w tabeli.
"""

import io, os, csv, re, copy, zipfile, string
from urllib.parse import quote
import requests
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
REL_IMG = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'

def wt(n): return f'{{{W}}}{n}'


# ─── Placeholder detection (IND) ──────────────────────────────────────────────
# „Gracz N" w grupowej = wolne miejsce na drabinkę przed losowaniem (np. „Gracz 44").
# „bye" w drabince = wolny los przy nieparzystej liczbie / wysokim rozstawieniu.
# Mecze z placeholderem po dowolnej stronie pomijamy gdy `skip_placeholders=True`.

_PLACEHOLDER_GRACZ_RE = re.compile(r'^\s*gracz\s+\d+\s*$', re.IGNORECASE)

def _is_placeholder_name(name):
    if not name:
        return False
    s = str(name).strip()
    if s.lower() == 'bye':
        return True
    if s.upper() == 'X':  # sentinel pustej drużyny/zawodnika (np. DWÓJKA bez kompletu)
        return True
    if _PLACEHOLDER_GRACZ_RE.match(s):
        return True
    return False


# ─── Google Sheets fetching ───────────────────────────────────────────────────

def _is_html(text):
    s = text.lstrip()[:200].lower()
    return s.startswith('<!doctype') or s.startswith('<html') or '<head' in s

def fetch_via_gviz(sheet_id, sheet_name):
    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={quote(sheet_name)}"
    r = requests.get(url, timeout=15)
    if r.status_code != 200 or _is_html(r.text):
        return None
    return list(csv.reader(io.StringIO(r.text)))

def fetch_via_gid(sheet_id, gid):
    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"
    r = requests.get(url, timeout=15)
    if r.status_code != 200 or _is_html(r.text):
        return None
    return list(csv.reader(io.StringIO(r.text)))

def fetch_via_export_name(sheet_id, sheet_name):
    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&sheet={quote(sheet_name)}"
    r = requests.get(url, timeout=15)
    if r.status_code != 200 or _is_html(r.text):
        return None
    return list(csv.reader(io.StringIO(r.text)))

def get_sheet_gids(sheet_id):
    """Skraping HTML arkusza by wyciągnąć listę nazwa→gid wszystkich zakładek.
    Google zmienia HTML z czasem, więc używamy kilku wzorców. Browser-like UA
    daje stabilniejszą odpowiedź niż domyślny requests UA."""
    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/edit"
    headers = {
        'User-Agent': ('Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 '
                       '(KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36'),
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'en-US,en;q=0.5',
    }
    try:
        r = requests.get(url, headers=headers, timeout=20)
    except Exception:
        return {}
    text = r.text
    mapping = {}
    # Wzór 1: <option value="GID">Nazwa</option> (legacy)
    for m in re.finditer(r'value="(\d{6,12})"[^>]*>([^<]{1,80})</option>', text):
        gid, name = m.group(1), m.group(2).strip()
        if name and name not in mapping:
            mapping[name] = gid
    # Wzór 2: bootstrap data array ["Nazwa", ..., GID]
    if not mapping:
        for m in re.finditer(r'\["([^"\\]{1,60})"(?:\s*,[^,\[\]]*){1,15},(\d{6,12})\]', text):
            name, gid = m.group(1).strip(), m.group(2)
            if name and name not in mapping:
                mapping[name] = gid
    # Wzór 3: nowy format Google 2024+ — JSON-like z "name" i "sheetId"
    if not mapping:
        for m in re.finditer(r'"name"\s*:\s*"([^"\\]{1,80})"[^}]*?"sheetId"\s*:\s*(\d{1,12})', text):
            name, gid = m.group(1).strip(), m.group(2)
            if name and name not in mapping:
                mapping[name] = gid
    # Wzór 4: HTML tab names w title atrybutach
    if not mapping:
        for m in re.finditer(r'title="([^"]{1,80})"[^>]*data-id="(\d{6,12})"', text):
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
        tor=g(col_tor); godz=g(col_godz); z1=g(col_z1); z2=g(col_z2)
        # Normalizacja formatu (arkusz/gviz bywa różny): tor "1.0"→"1",
        # godzina "09:30:00"→"09:30". Bez tego komórki sformatowane jako liczba/czas
        # z sekundami wywalały walidację (0 meczów) lub dawały "Tor 1.0" / "09:30:00".
        if re.match(r'^\d+\.0+$', tor):
            tor = tor.split('.')[0]
        _gm = re.match(r'^(\d{1,2}:\d{2})(?::\d{2})?$', godz)
        if _gm:
            godz = _gm.group(1)
        if not _is_valid_match_row(tor,godz,z1,z2): continue
        mecz_raw = g(col_mecz)
        # Mecz # bywa float ("1.0") gdy arkusz formatuje kolumnę jako liczbę.
        if re.match(r'^\d+\.0+$', mecz_raw):
            mecz_raw = mecz_raw.split('.')[0]
        matches.append({'tor':tor,'godz':godz,'grupa':grupa,
                        'mecz':mecz_raw,'z1':z1,'z2':z2})
    return matches

def fetch_all_group_sheets(sheet_id, progress_cb=None):
    """Skanuje zakładki Gr. A..Z faktycznie istniejące w arkuszu. Zwraca tylko te z meczami.

    Optymalizacja: gdy gid_map jest dostępne, iterujemy tylko po nazwach które
    naprawdę istnieją (zazwyczaj 4-10 grup zamiast 26 zbędnych HTTP).
    progress_cb: callable(done, total, label) — wywoływany dla każdej zakładki."""
    gid_map = get_sheet_gids(sheet_id)
    if gid_map:
        candidates = [f"Gr. {L}" for L in string.ascii_uppercase if f"Gr. {L}" in gid_map]
        if not candidates:
            candidates = [f"Gr. {L}" for L in string.ascii_uppercase]
    else:
        candidates = [f"Gr. {L}" for L in string.ascii_uppercase]
    results = []
    total = max(1, len(candidates))
    is_fallback = not gid_map
    empty_streak = 0
    for i, name in enumerate(candidates):
        if progress_cb:
            try:
                # Fallback A-Z: nie pokazuj nazwy (większość pusta) — tylko licznik znalezionych
                label = f"znaleziono {len(results)}" if is_fallback else name
                progress_cb(i + 1, total, label)
            except Exception: pass
        try:
            rows = fetch_sheet(sheet_id, name, gid_map)
            if rows is None:
                empty_streak += 1
            else:
                matches = parse_group_rows(rows)
                if matches:
                    results.append((name, matches))
                    empty_streak = 0
                else:
                    empty_streak += 1
        except Exception:
            empty_streak += 1
            continue
        # Early stop: po 3 pustych pod rząd w fallback mode (gdy mamy już 1+ wynik)
        if is_fallback and results and empty_streak >= 3:
            break
    return results


# ─────────────────────────────────────────────────────────────────────
# Parser zakładki "Drabinka" - dla meczów fazy pucharowej
# ─────────────────────────────────────────────────────────────────────

# Rozpoznawane nazwy faz — funkcja detect_phase rozpoznaje dynamicznie:
#   • "1/N finału" / "1/N FINAŁU"
#   • "PÓŁFINAŁY" / "półfinał"
#   • "FINAŁ"
#   • "MECZ O N. MIEJSCE" (gdzie N = 3, 5, 7, ..., 31)
#   • "MIEJSCA X-Y" (dowolny zakres)
# Zwraca (phase_key, phase_full_name) lub (None, None).
# phase_key jest unikatowym kluczem fazy (do dopasowania target_phase).

# ─── Rozpiski meczowe per zawodnik (ukryte kolumny AC-AI w Gr. *) ─────────────
# ─── Rozpiski meczowe per zawodnik ────────────────────────────────────────
# Derywujemy z meczów grupowych (parse_group_rows) — każdy mecz ląduje w
# rozpiskach OBOJU graczy (z1 i z2) bez modyfikacji oryginalnej kolejności
# z arkusza. Wcześniej parsowaliśmy z ukrytych kolumn AC-AI, ale to były
# duplicate dane — wszystko jest w widocznych kolumnach Tor/Godzina/Gracz1/Gracz2.

def matches_to_player_schedules(matches, group_letter):
    """Z listy meczów grupowych (parse_group_rows) tworzy listę rozpisek per zawodnik.

    Output: list[{name, group, matches: [{godzina, tor, z1, z2}]}]
    Zawodnicy posortowani alfabetycznie. Mecze w kolejności jak w arkuszu (= chronologicznie/po torach).
    z1/z2 ZACHOWANE jak w arkuszu (nie przesuwamy własnego zawodnika na lewo)."""
    by_player = {}  # name -> list of original matches
    for m in matches:
        for who in (m.get('z1', ''), m.get('z2', '')):
            who = (who or '').strip()
            if not who: continue
            if _is_placeholder_name(who): continue  # nie rób rozpiski dla 'X'/'bye'/'Gracz N'
            by_player.setdefault(who, []).append(m)
    out = []
    for name in sorted(by_player.keys(), key=lambda s: s.lower()):
        items = []
        for mm in by_player[name]:
            items.append({
                'godzina': mm.get('godz', ''),
                'tor': mm.get('tor', ''),
                'z1': mm.get('z1', ''),
                'z2': mm.get('z2', ''),
            })
        out.append({'name': name, 'group': group_letter, 'matches': items})
    return out


def fetch_all_player_schedules(sheet_id, progress_cb=None):
    """Pobiera rozpiski z wszystkich zakładek Gr. * w arkuszu (przez gviz).
    Reuse parse_group_rows + matches_to_player_schedules.
    progress_cb(done, total, label): postęp dla UI.
    Gdy gid_map dostępna — iterujemy tylko po istniejących zakładkach.
    Gdy nie — probujemy A..Z ale zatrzymujemy się po 2 z rzędu pustych,
    żeby nie pokazywać A..Z gdy turniej ma np. tylko grupy A-D."""
    gid_map = get_sheet_gids(sheet_id)
    if gid_map:
        candidates = [f"Gr. {L}" for L in string.ascii_uppercase if f"Gr. {L}" in gid_map]
        # Z gid_map znamy realną liczbę grup → progres pełny.
        total_for_ui = max(1, len(candidates))
    else:
        candidates = [f"Gr. {L}" for L in string.ascii_uppercase]
        # Bez gid_map total jest nieznany — sygnalizujemy 0 by UI ukrył „/N".
        total_for_ui = 0
    all_schedules = []
    consecutive_empty = 0
    for i, name in enumerate(candidates):
        if progress_cb:
            try: progress_cb(i + 1, total_for_ui, name)
            except Exception: pass
        try:
            rows = fetch_sheet(sheet_id, name, gid_map=gid_map)
        except Exception:
            rows = None
        if not rows:
            # Bez gid_map nie wiemy ile jest grup — early stop po 2 z rzędu.
            if not gid_map:
                consecutive_empty += 1
                if consecutive_empty >= 2 and all_schedules:
                    break
            continue
        matches = parse_group_rows(rows)
        if not matches:
            if not gid_map:
                consecutive_empty += 1
                if consecutive_empty >= 2 and all_schedules:
                    break
            continue
        consecutive_empty = 0
        # grupa: bierzemy z pierwszego meczu (parse_group_rows wypełnia z headera 'Grupa X')
        letter = matches[0].get('grupa') or name.rsplit(' ', 1)[-1]
        all_schedules.extend(matches_to_player_schedules(matches, letter))
    return all_schedules


# ─── Builder docx z kartami zawodników (do druku + wycinania) ────────────────
# Layout: A4 portrait, marginesy 1cm, grid 2 kolumny × 4 wiersze = 8 kart/stronę.
# Każda karta = komórka outer-table; w środku tytuł + małe info + tabelka 5 meczów.

def _emu_cm(cm): return int(cm * 360000)
def _dxa_cm(cm): return int(cm * 567)  # 1cm = 567 dxa (twentieth of point)

def _new_para(text='', *, bold=False, italic=False, size_pt=None, align=None,
              after_pt=0, font='Calibri', color=None):
    """Tworzy <w:p> z jednym runem."""
    p = etree.Element(wt('p'))
    pPr = etree.SubElement(p, wt('pPr'))
    if align:
        jc = etree.SubElement(pPr, wt('jc')); jc.set(wt('val'), align)
    sp = etree.SubElement(pPr, wt('spacing'))
    sp.set(wt('before'), '0'); sp.set(wt('after'), str(after_pt * 20))
    sp.set(wt('line'), '240'); sp.set(wt('lineRule'), 'auto')
    if text:
        r = etree.SubElement(p, wt('r'))
        rPr = etree.SubElement(r, wt('rPr'))
        fonts = etree.SubElement(rPr, wt('rFonts'))
        for a in ('ascii', 'hAnsi', 'eastAsia', 'cs'): fonts.set(wt(a), font)
        if bold: etree.SubElement(rPr, wt('b'))
        if italic:
            etree.SubElement(rPr, wt('i')); etree.SubElement(rPr, wt('iCs'))
        if color:
            c = etree.SubElement(rPr, wt('color')); c.set(wt('val'), color)
        if size_pt:
            sz_val = str(int(round(size_pt * 2)))
            sz = etree.SubElement(rPr, wt('sz')); sz.set(wt('val'), sz_val)
            szCs = etree.SubElement(rPr, wt('szCs')); szCs.set(wt('val'), sz_val)
        t = etree.SubElement(r, wt('t'))
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        t.text = text
    return p


def _make_border(val='single', sz='4', color='000000'):
    el = etree.Element(wt(''))
    el.set(wt('val'), val); el.set(wt('sz'), sz); el.set(wt('space'), '0'); el.set(wt('color'), color)
    return el


def _set_cell_borders(tc, *, all_val='single', sz='4', color='000000'):
    tcPr = tc.find(wt('tcPr'))
    if tcPr is None:
        tcPr = etree.Element(wt('tcPr')); tc.insert(0, tcPr)
    # usuń stare borders jeśli są
    old = tcPr.find(wt('tcBorders'))
    if old is not None: tcPr.remove(old)
    bdr = etree.SubElement(tcPr, wt('tcBorders'))
    for name in ('top', 'left', 'bottom', 'right'):
        b = etree.SubElement(bdr, wt(name))
        b.set(wt('val'), all_val); b.set(wt('sz'), sz); b.set(wt('space'), '0'); b.set(wt('color'), color)


def _set_cell_margins(tc, top=80, left=120, bottom=80, right=120):
    tcPr = tc.find(wt('tcPr'))
    if tcPr is None:
        tcPr = etree.Element(wt('tcPr')); tc.insert(0, tcPr)
    old = tcPr.find(wt('tcMar'))
    if old is not None: tcPr.remove(old)
    mar = etree.SubElement(tcPr, wt('tcMar'))
    for name, val in (('top', top), ('left', left), ('bottom', bottom), ('right', right)):
        m = etree.SubElement(mar, wt(name))
        m.set(wt('w'), str(val)); m.set(wt('type'), 'dxa')


def build_player_schedules_doc(schedules, *, tournament_name=None,
                                tournament_date=None,
                                cols=2, rows_per_page=5,
                                is_team=False):
    """Buduje docx z gridem kart zawodników (do druku + wycinania).
    schedules: lista z parse_player_schedules / fetch_all_player_schedules.
    Returns: bytes (docx).

    Używamy python-docx jako bazy (skeleton docx z poprawnymi stylami) i
    wstrzykujemy własną tabelę przez lxml. Minimalny docx-from-scratch nie
    renderuje multi-column tabel w LibreOffice (cards stackowały się jednokolumnowo).
    """
    from docx import Document
    from docx.shared import Cm, Pt

    PAGE_W_CM = 21.0
    PAGE_H_CM = 29.7
    MARG_CM = 1.0
    usable_w_dxa = _dxa_cm(PAGE_W_CM - 2 * MARG_CM)
    card_w_dxa = usable_w_dxa // cols

    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(PAGE_W_CM)
    section.page_height = Cm(PAGE_H_CM)
    section.left_margin = Cm(MARG_CM); section.right_margin = Cm(MARG_CM)
    section.top_margin = Cm(MARG_CM); section.bottom_margin = Cm(MARG_CM)

    n_rows = (len(schedules) + cols - 1) // cols
    if n_rows == 0:
        return _save_doc_to_bytes(doc)

    # Tabela cols × n_rows na karty. python-docx tworzy strukturę, my potem
    # ustawiamy szerokości i wstrzykujemy bogate komórki.
    table = doc.add_table(rows=n_rows, cols=cols)
    table.autofit = False

    tbl_el = table._tbl
    tblPr = tbl_el.find(wt('tblPr'))
    if tblPr is None:
        tblPr = etree.SubElement(tbl_el, wt('tblPr'))
    # tblW + tblLayout=fixed + jc=center (centruje grid kart na stronie)
    for child in list(tblPr):
        if child.tag in (wt('tblW'), wt('tblLayout'), wt('jc')):
            tblPr.remove(child)
    tblW = etree.SubElement(tblPr, wt('tblW'))
    tblW.set(wt('w'), str(usable_w_dxa)); tblW.set(wt('type'), 'dxa')
    jc = etree.SubElement(tblPr, wt('jc')); jc.set(wt('val'), 'center')
    tblLayout = etree.SubElement(tblPr, wt('tblLayout'))
    tblLayout.set(wt('type'), 'fixed')

    # tblGrid z naszymi szerokościami
    grid = tbl_el.find(wt('tblGrid'))
    if grid is not None:
        tbl_el.remove(grid)
    grid = etree.Element(wt('tblGrid'))
    for _ in range(cols):
        gc = etree.SubElement(grid, wt('gridCol')); gc.set(wt('w'), str(card_w_dxa))
    tblPr.addnext(grid)

    # Wypełnianie komórek
    for ri in range(n_rows):
        # cantSplit per row — karta nie łamie się między stronami
        tr = table.rows[ri]._tr
        trPr = tr.find(wt('trPr'))
        if trPr is None:
            trPr = etree.Element(wt('trPr'))
            tr.insert(0, trPr)
        if trPr.find(wt('cantSplit')) is None:
            etree.SubElement(trPr, wt('cantSplit'))
        for ci in range(cols):
            idx = ri * cols + ci
            tc = table.rows[ri].cells[ci]._tc
            # Usuń domyślny pusty paragraf z python-docx
            for p in list(tc.findall(wt('p'))):
                tc.remove(p)
            # Ustaw szerokość tc + border + marginesy
            tcPr = tc.find(wt('tcPr'))
            if tcPr is None:
                tcPr = etree.Element(wt('tcPr')); tc.insert(0, tcPr)
            old_tcW = tcPr.find(wt('tcW'))
            if old_tcW is not None: tcPr.remove(old_tcW)
            tcW = etree.SubElement(tcPr, wt('tcW'))
            tcW.set(wt('w'), str(card_w_dxa)); tcW.set(wt('type'), 'dxa')
            if idx < len(schedules):
                _fill_player_card_tc(tc, schedules[idx], tournament_name,
                                     tournament_date, card_w_dxa,
                                     is_team=is_team)
            else:
                # Pusta komórka — ramka kropkowana (linia cięcia placeholder)
                _set_cell_borders(tc, all_val='dashed', sz='4', color='AAAAAA')
                tc.append(_new_para(''))

    return _save_doc_to_bytes(doc)


def _save_doc_to_bytes(doc):
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def _fill_player_card_tc(tc, player, tournament_name, tournament_date, card_w_dxa, *, is_team=False):
    """Wypełnia istniejący <w:tc> zawartością karty zawodnika.
    Layout: header (imię) → subtitle (Grupa/turniej/data) → tabela 4 kolumny
    (godzina | tor | gracz 1 | gracz 2). Własne nazwisko w wierszach pogrubione
    (jak na wzorcu z arkusza)."""
    _set_cell_borders(tc, all_val='single', sz='12', color='000000')
    _set_cell_margins(tc, top=80, left=140, bottom=80, right=140)

    # Header: imię + nazwisko
    tc.append(_new_para(player['name'], bold=True, size_pt=12, align='left', after_pt=1))
    # Subtitle: Grupa · turniej · data
    sub_parts = [f"Grupa {player['group']}"]
    if tournament_name: sub_parts.append(tournament_name)
    if tournament_date: sub_parts.append(tournament_date)
    tc.append(_new_para(' · '.join(sub_parts), size_pt=7, align='left', after_pt=2,
                        italic=True, color='707070'))

    # Inner table 4 col × (1 + N) rows
    inner_w = card_w_dxa - 280  # tcMar L+R
    # godz. | tor | gracz 1 | gracz 2 — gracze szerokie; "godz." mieści się w 1 linii
    cw = [int(inner_w * 0.12), int(inner_w * 0.07), 0, 0]
    cw[2] = (inner_w - cw[0] - cw[1]) // 2
    cw[3] = inner_w - cw[0] - cw[1] - cw[2]

    inner_tbl = etree.Element(wt('tbl'))
    ip = etree.SubElement(inner_tbl, wt('tblPr'))
    iw = etree.SubElement(ip, wt('tblW'))
    iw.set(wt('w'), str(inner_w)); iw.set(wt('type'), 'dxa')
    ijc = etree.SubElement(ip, wt('jc')); ijc.set(wt('val'), 'center')
    bdr = etree.SubElement(ip, wt('tblBorders'))
    for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        b = etree.SubElement(bdr, wt(side))
        b.set(wt('val'), 'single'); b.set(wt('sz'), '4'); b.set(wt('color'), '707070')
    ilayout = etree.SubElement(ip, wt('tblLayout'))
    ilayout.set(wt('type'), 'fixed')
    igrid = etree.SubElement(inner_tbl, wt('tblGrid'))
    for w in cw:
        gc = etree.SubElement(igrid, wt('gridCol')); gc.set(wt('w'), str(w))

    def _add_row(godz, tor, p1, p2, *, header=False, shade=False, own_name=None):
        tr = etree.SubElement(inner_tbl, wt('tr'))
        trPr = etree.SubElement(tr, wt('trPr'))
        etree.SubElement(trPr, wt('cantSplit'))
        if header:
            etree.SubElement(trPr, wt('tblHeader'))
        for i, (val, w) in enumerate(zip((godz, tor, p1, p2), cw)):
            cell = etree.SubElement(tr, wt('tc'))
            cPr = etree.SubElement(cell, wt('tcPr'))
            wEl = etree.SubElement(cPr, wt('tcW'))
            wEl.set(wt('w'), str(w)); wEl.set(wt('type'), 'dxa')
            vA = etree.SubElement(cPr, wt('vAlign')); vA.set(wt('val'), 'center')
            if header or shade:
                sh = etree.SubElement(cPr, wt('shd'))
                sh.set(wt('val'), 'clear'); sh.set(wt('color'), 'auto')
                sh.set(wt('fill'), 'EEEEEE' if header else 'F7F7F7')
            mar = etree.SubElement(cPr, wt('tcMar'))
            for nm, mv in (('top', 10), ('left', 50), ('bottom', 10), ('right', 50)):
                me = etree.SubElement(mar, wt(nm))
                me.set(wt('w'), str(mv)); me.set(wt('type'), 'dxa')
            # Kolumny gracz 1 / gracz 2 do lewej; godzina/tor wycentrowane
            align = 'left' if i >= 2 else 'center'
            # Pogrubienie własnego nazwiska (kolumny 2-3); header zawsze italic
            cell_bold = False
            if header:
                cell_italic = True
            else:
                cell_italic = False
                if own_name and i >= 2 and val == own_name:
                    cell_bold = True
            # Auto-shrink dla długich nazw w kol. gracz/drużyna 1/2 — kiedy bold
            # robi że "Stowarzyszenie Aktywny Orlik" rozjeżdża się na 2 wiersze.
            # Heurystyka ~75 dxa/char przy sz=16 bold (Calibri/Carlito) — tylko
            # gdy NAPRAWDĘ się nie mieści. Wcześniejsze 110 przesadnie zmniejszało
            # nawet "Warsaw Adventure Team" (21 znaków) który spokojnie mieści się
            # w 1 linii bez skalowania.
            cell_sz_pt = 7 if header else 8
            if not header and i >= 2 and val:
                usable_dxa = w - 100  # cell width minus tcMar L+R
                est_dxa = len(val) * 75
                if est_dxa > usable_dxa:
                    shrunk_half_pt = max(12, int(16 * usable_dxa / est_dxa))
                    cell_sz_pt = shrunk_half_pt / 2.0
            cell.append(_new_para(val, bold=(header or cell_bold),
                                  italic=cell_italic,
                                  size_pt=cell_sz_pt,
                                  align=align, after_pt=0))
        return tr

    side_label = 'drużyna' if is_team else 'gracz'
    _add_row('godz.', 'tor', f'{side_label} 1', f'{side_label} 2', header=True)
    for idx, m in enumerate(player['matches']):
        _add_row(m.get('godzina', '') or '—',
                 (m.get('tor') or '—'),
                 m.get('z1') or '—',
                 m.get('z2') or '—',
                 shade=(idx % 2 == 1),
                 own_name=player['name'])

    tc.append(inner_tbl)
    # Pusty paragraf po tabeli (wymaganie OOXML: tc musi kończyć się <w:p>)
    tc.append(_new_para(''))


def pluralize(n, sing, few, many):
    """Polish plural: 1 → sing, 2-4 (oprócz 12-14) → few, reszta → many.
    Przykład: 1 grupa / 2 grupy / 5 grup; 1 mecz / 2 mecze / 5 meczów; 1 faza / 2 fazy / 5 faz."""
    last = n % 10
    last2 = n % 100
    if n == 1:
        return sing
    elif last in (2, 3, 4) and last2 not in (12, 13, 14):
        return few
    else:
        return many


def detect_phase(cell_text):
    """Rozpoznaje fazę turnieju z tekstu komórki. Zwraca (key, full_name) lub (None, None)."""
    s = cell_text.strip().lower()
    if not s:
        return None, None
    # "MECZ O N. MIEJSCE" / "Mecz o 3 miejsce" / "MIEJSCE 3" — single match for Nth place
    m = re.search(r'mecz\s+o\s+(\d+)\.?\s*miejsc', s)
    if m:
        n = m.group(1)
        return f'mecz o {n}', f'MECZ O {n}. MIEJSCE'
    # Plain "3. MIEJSCE" / "5. MIEJSCE" itp. (bez "mecz o")
    m = re.search(r'^(\d+)\.\s*miejsc', s)
    if m:
        n = m.group(1)
        return f'mecz o {n}', f'MECZ O {n}. MIEJSCE'
    # "MIEJSCA X-Y" (z lub bez czasu w nawiasie)
    m = re.search(r'miejsca\s+(\d+)\s*-\s*(\d+)', s)
    if m:
        a, b = m.group(1), m.group(2)
        return f'miejsca {a}-{b}', f'MIEJSCA {a}-{b}'
    # "1/N finału"
    m = re.search(r'1/(\d+)\b', s)
    if m:
        n = m.group(1)
        if n == '2':
            return '1/2', '1/2 FINAŁU'
        return f'1/{n}', f'1/{n} FINAŁU'
    # "PÓŁFINAŁY" / "półfinał"
    if 'półfinał' in s:
        return '1/2', '1/2 FINAŁU'
    # "FINAŁ" (nie "półfinał" - sprawdzone wyżej)
    if 'finał' in s:
        return 'finał', 'FINAŁ'
    return None, None


# Wstecznej kompatybilności: lista znanych phase keys do iteracji w "Sprawdź zakładki".
# (Tylko predefiniowane — dynamiczne fazy z arkusza wykrywa detect_phase.)
PHASE_NAMES = {
    '1/64': '1/64 FINAŁU',
    '1/32': '1/32 FINAŁU',
    '1/16': '1/16 FINAŁU',
    '1/8':  '1/8 FINAŁU',
    '1/4':  '1/4 FINAŁU',
    '1/2':  '1/2 FINAŁU',
    'półfinał': '1/2 FINAŁU',
    'mecz o 3': 'MECZ O 3. MIEJSCE',
    'miejsca 5-8':   'MIEJSCA 5-8',
    'miejsca 9-16':  'MIEJSCA 9-16',
    'miejsca 17-24': 'MIEJSCA 17-24',
    'miejsca 17-32': 'MIEJSCA 17-32',
    'miejsca 25-32': 'MIEJSCA 25-32',
    'finał': 'FINAŁ',
}


def parse_drabinka_rows(rows, target_phase=None):
    """
    Parsuje zakładkę Drabinka i zwraca mecze z konkretnej fazy.
    
    Drabinka może mieć WIELE faz w jednym arkuszu, w RÓŻNYCH KOLUMNACH:
      np. D = "1/8 FINAŁU"  | N = "1/4 FINAŁU"  | Y = "PÓŁFINAŁY"  | AI = "FINAŁ"
    Plus podsekcje w tej samej kolumnie ale niżej:
      np. AI: row 1 "FINAŁ" + row 5 "3. MIEJSCE"
      np. D:  row 1 "1/8 FINAŁU" + row 20 "MIEJSCA 17-24"
    
    Każdy blok fazy ma strukturę: [Tor][nazwa fazy][Set 1]...[Set 5][SETY]
    Player names są w kolumnie z nazwą fazy. Tor jest LEWO od niej (sąsiednia
    lub +1 jeśli między nimi jest "Grupa").
    
    target_phase: np. "Pucharowa 1/4 finału" lub "1/4 FINAŁU" lub "Mecz o 3. miejsce".
    
    Zwraca: (phase_full_name, time, [matches])
        matches: [{tor, godz, grupa, mecz, z1, z2}]
    """
    if not rows:
        return None, None, []
    
    # Słowa-stop dla section headers (używane też w walidacji nazwisk)
    stop_keywords = ['miejsc', 'finał', 'finałow', 'mecz o', 'play-off',
                     'ranking', 'klasyfikacja', '1/64', '1/32', '1/16',
                     '1/8', '1/4', '1/2', 'półfinał']
    
    # ── KROK 1: skanowanie WSZYSTKICH komórek pod kątem nazw faz ──
    # Każde znalezienie = potencjalny blok fazy. Używamy detect_phase()
    # (regex-based) — łapie też dynamiczne wartości "MIEJSCA X-Y" i
    # "MECZ O N. MIEJSCE" których nie ma w predefiniowanym PHASE_NAMES.
    phase_blocks = []
    for r_idx, row in enumerate(rows):
        for c_idx, cell in enumerate(row):
            phase_key, phase_full = detect_phase(cell)
            if phase_key is None:
                continue
            # Precyzyjna detekcja Tor/Grupa. Układ bloku fazy to:
            #   "Tor Grupa Phase" (faza grupowa-style) LUB "Tor Phase" (drabinka bez grup).
            # Kolumna DOKŁADNIE przed phase to albo Grupa (wtedy Tor o 1 dalej w lewo)
            # albo Tor (wtedy brak Grupa). Sprawdzamy to wprost zamiast szerokiego skanu,
            # bo szeroki skan łapał etykiety z INNYCH bloków fazy w tym samym wierszu
            # (np. "Grupa" z 1/32 przy parsowaniu MIEJSCA 33-48).
            def _cell_label(r_arr, ci):
                if 0 <= ci < len(r_arr):
                    return r_arr[ci].strip().lower().rstrip('.:')
                return ''
            col_tor = None
            col_grupa = None
            # Szukamy etykiet "Tor"/"Grupa" w wierszu nagłówka fazy (i ±2 sąsiednich),
            # skanując W LEWO od kolumny fazy do col 0 albo do innej fazy.
            # Może być WIELE kolumn etykietowanych "Tor" (np. arkusz ma seed IDs w
            # kolumnie A pod etykietą "Tor" + prawdziwe numery torów w kolumnie B).
            # Wybieramy tę, w której DATA wyglądają jak numery torów (1-99).
            def _looks_like_tor_col(ci, r_after):
                """Sprawdza czy w kolumnie ci data pod nagłówkiem (od r_after+1)
                wygląda jak numery torów: cyfry / puste (merged) / NIE-seed-id."""
                n_ok = 0
                n_bad = 0
                for ri in range(r_after + 1, min(r_after + 10, len(rows))):
                    if ci >= len(rows[ri]):
                        continue
                    v = str(rows[ri][ci]).strip()
                    if not v:
                        continue
                    # Numer toru: czysta liczba 1-99 lub "1.0" (gviz float)
                    if re.match(r'^\d{1,2}(\.0+)?$', v):
                        n_ok += 1
                    # Seed ID: litera + cyfra (A1, B4, AB12 itp.) — to NIE Tor
                    elif re.match(r'^[A-Z]+\d+$', v):
                        n_bad += 1
                # Akceptujemy gdy więcej numerów niż seed IDs (lub same numery).
                return n_ok > n_bad
            for dr in (0, -1, 1, -2, 2):
                r2 = r_idx + dr
                if not (0 <= r2 < len(rows)):
                    continue
                hr = rows[r2]
                _ct = _cg = None
                _tor_candidates = []  # kolumny z etykietą 'tor' (w kolejności od najbliższej do fazy)
                for ci in range(c_idx - 1, -1, -1):
                    # Stop gdy trafimy inną fazę (jej kolumna oznacza początek innego bloku).
                    if 0 <= ci < len(hr):
                        cell_v = hr[ci]
                        ph_key, _ = detect_phase(cell_v) if isinstance(cell_v, str) else (None, None)
                        if ph_key is not None and ci != c_idx:
                            break
                    label = _cell_label(hr, ci)
                    if label == 'tor':
                        _tor_candidates.append(ci)
                    elif label == 'grupa' and _cg is None:
                        _cg = ci
                # Wybór candidate: pierwszy z numerycznymi data (preferuje kolumnę
                # z liczbami nad kolumną z seed IDs).
                for ci in _tor_candidates:
                    if _looks_like_tor_col(ci, r_idx):
                        _ct = ci
                        break
                if _ct is None and _tor_candidates:
                    _ct = _tor_candidates[0]  # fallback: pierwsza znaleziona
                # Akceptujemy tylko jeśli faktycznie coś znaleźliśmy w tym wierszu.
                if _ct is not None or _cg is not None:
                    col_tor = _ct
                    col_grupa = _cg
                    break
            m = re.search(r'\((\d{1,2}:\d{2})\)', cell)
            phase_time = m.group(1) if m else None
            phase_blocks.append({
                'header_row': r_idx,
                'col': c_idx,
                'phase_full': phase_full,
                'phase_key': phase_key,
                'time': phase_time,
                'col_tor': col_tor,
                'col_grupa': col_grupa,
            })
    
    if not phase_blocks:
        return None, None, []
    
    # ── KROK 2: wybór bloku(ów) pasujących do target_phase ──
    # Używamy detect_phase() do parsowania target_phase (np. "Pucharowa 1/4 finału",
    # "Mecz o 9. miejsce", "Miejsca 17-20"). Potem szukamy WSZYSTKICH bloków których
    # phase_key == target_key, ALBO których phase_full pasuje (np. 'półfinał' z arkusza
    # i '1/2' z UI oba mapują na '1/2 FINAŁU').
    #
    # WAŻNE: jedna faza może być ROZBITA na KILKA bloków nagłówkowych ułożonych
    # pionowo w tej samej kolumnie (np. "MIEJSCA 5-8" = 2 półfinały drabinki
    # o miejsca 5-8, każdy z własnym nagłówkiem "MIEJSCA 5-8 (16:45)" i pustym
    # wierszem między nimi). Wcześniej braliśmy TYLKO pierwszy blok i czytaliśmy
    # do nagłówka drugiego (stop-keyword) → gubiliśmy połowę meczów. Teraz
    # zbieramy mecze ze WSZYSTKICH bloków tej fazy i sklejamy.
    chosen_blocks = []
    target_key = None
    target_full = None
    if target_phase:
        target_key, target_full = detect_phase(target_phase)
        if target_key:
            chosen_blocks = [b for b in phase_blocks if b['phase_key'] == target_key]
            if not chosen_blocks:
                chosen_blocks = [b for b in phase_blocks if b['phase_full'] == target_full]
            # Jeśli target_phase ma sens ale brak takiego bloku w arkuszu —
            # zwracamy pustą listę z poprawnym phase_full_name. User dostanie
            # czytelny komunikat "Nie znaleziono fazy".
            if not chosen_blocks:
                return target_full, None, []

    # target_phase=None lub niezrozumiały → fallback do pierwszego bloku
    if not chosen_blocks:
        chosen_blocks = [phase_blocks[0]]

    # Bloki w kolejności wystąpienia (scan jest row-major, więc już posortowane).
    chosen_blocks = sorted(chosen_blocks, key=lambda b: (b['header_row'], b['col']))

    phase_full_name = chosen_blocks[0]['phase_full']
    phase_time = chosen_blocks[0]['time']

    def g(row, c):
        if c is None or c >= len(row): return ''
        return row[c].strip()

    # Słowa-stop dla section headers (pod-nagłówek = koniec bieżącego bloku)
    # — zdefiniowane wyżej jako stop_keywords.

    # ── KROK 3: czytanie meczów ze WSZYSTKICH wybranych bloków ──
    matches = []
    match_num = 1
    # Dedup par zawodników w obrębie fazy. W drabince (knockout) każda para gra
    # raz na fazę, więc identyczna para = duplikat (np. arkusz z przypadkowo
    # zduplikowanym nagłówkiem+treścią). Klucz nieuporządkowany (kolejność z1/z2
    # bez znaczenia). To uodparnia na błędy w przygotowaniu arkusza — gdy split
    # bloków to dwa RÓŻNE mecze (legit półfinały), pary się różnią i nic nie ginie.
    seen_pairs = set()
    for chosen in chosen_blocks:
        col_player = chosen['col']
        col_grupa = chosen['col_grupa']
        # AGGRESSIVE: jeśli col_grupa znalezione, ZAWSZE używamy col_grupa-1 jako tor.
        # W arkuszach Mölkky układ to ZAWSZE "Tor Grupa Phase" — deterministyczna prawda.
        if col_grupa is not None and col_grupa > 0:
            col_tor = col_grupa - 1
        else:
            col_tor = chosen['col_tor']
        # Brak etykiety 'Tor' → gviz typuje kolumnę torów jako LICZBOWĄ i GUBI nagłówek.
        # Tor jest tuż na lewo od kolumny gracza, ALE bezpośrednio przed player może
        # być kolumna z seed ID (np. "A1"/"B4") — wtedy Tor jest col_player-2.
        # Fix: scan leftward, użyj numeric-vs-seed heurystyki by odrzucić seed IDs.
        if col_tor is None:
            header_idx_tmp = chosen['header_row']
            for offset in (1, 2, 3):
                ci = col_player - offset
                if ci < 0:
                    break
                n_ok = 0
                n_bad = 0
                for ri in range(header_idx_tmp + 1, min(header_idx_tmp + 10, len(rows))):
                    if ci >= len(rows[ri]):
                        continue
                    v = rows[ri][ci].strip()
                    if not v:
                        continue
                    if re.match(r'^\d{1,2}(\.0+)?$', v):
                        n_ok += 1
                    elif re.match(r'^[A-Z]+\d+$', v):
                        n_bad += 1
                if n_ok > n_bad and n_ok > 0:
                    col_tor = ci
                    break
            if col_tor is None:
                col_tor = max(0, col_player - 1)

        header_idx = chosen['header_row']
        data_rows = rows[header_idx + 1:]

        i = 0
        empty_streak = 0
        last_known_tor = ''  # cache dla scalonych komórek tora
        while i < len(data_rows):
            row1 = data_rows[i]
            # Patrzymy TYLKO na kolumnę zawodnika i tora — inne fazy w tym samym
            # wierszu nie mają znaczenia (są w innych kolumnach).
            z1_raw = g(row1, col_player)
            z1_lower = z1_raw.lower()

            # Pusta komórka zawodnika? Sprawdź następny wiersz lub stop.
            if not z1_raw:
                empty_streak += 1
                if empty_streak >= 5:
                    break  # 5+ pustych = koniec bloku
                i += 1
                continue
            empty_streak = 0

            # Czy ta komórka to nagłówek nowej sekcji (np. drugi "MIEJSCA 5-8"
            # albo "MIEJSCA 17-24" pod 1/8)? To koniec bieżącego bloku — kolejny
            # blok tej samej fazy zostanie obsłużony w następnej iteracji pętli.
            if any(kw in z1_lower for kw in stop_keywords):
                break

            # Drugi zawodnik w następnym wierszu
            row2 = data_rows[i+1] if i+1 < len(data_rows) else []
            z2_raw = g(row2, col_player)
            z2_lower = z2_raw.lower()

            # Tor: hierarchia źródeł
            tor_raw = g(row1, col_tor) or g(row2, col_tor)
            if tor_raw:
                tor = tor_raw.strip()
                if re.match(r'^\d+\.0+$', tor):
                    tor = tor.split('.')[0]
                if 'TBA' in tor.upper() or tor == '?':
                    tor = ''
                elif tor.isdigit():
                    last_known_tor = tor
            elif last_known_tor and last_known_tor.isdigit():
                tor = str(int(last_known_tor) + 1)
                last_known_tor = tor
            else:
                tor = ''

            # Grupa: czytamy z col_grupa jeśli istnieje
            grupa_v = ''
            if col_grupa is not None:
                grupa_v = (g(row1, col_grupa) or g(row2, col_grupa)).strip()
            # Walidacja: oba nazwiska niepuste, sensowna długość, brak stop-keywords
            if z1_raw and z2_raw and len(z1_raw) >= 3 and len(z2_raw) >= 3:
                has_stop = any(kw in z1_lower or kw in z2_lower for kw in stop_keywords)
                pair_key = frozenset((z1_raw.strip().lower(), z2_raw.strip().lower()))
                is_dup = pair_key in seen_pairs
                if not has_stop and not is_dup \
                   and any(c.isupper() for c in z1_raw) and any(c.isupper() for c in z2_raw):
                    seen_pairs.add(pair_key)
                    matches.append({
                        'tor': tor,
                        'godz': phase_time or '',
                        'grupa': grupa_v,
                        'mecz': str(match_num),
                        'z1': z1_raw,
                        'z2': z2_raw,
                    })
                    match_num += 1

            i += 2  # przeskakujemy 2 wiersze (mecz)

    # Do hard-capu używamy phase_key pierwszego bloku.
    chosen = chosen_blocks[0]

    # ── HARD CAP ── parser czasem łapie więcej meczów niż powinno (np. brak
    # nagłówka "MECZE O MIEJSCA" oddzielającego 1/8 finału od meczów o miejsca
    # 17-32). Liczba meczów w fazie jest deterministyczna:
    #   1/N finału = N/2 par, MIEJSCA X-Y = (Y-X+1)/2 par, MECZ O N = 1.
    expected = None
    src_full = phase_full_name or ''
    src_key = chosen.get('phase_key', '') if chosen else ''
    # 1/N finału — N meczów (np. 1/8 = 8 meczów, 1/4 = 4, 1/2 = 2)
    m = re.match(r'1/(\d+)', src_key)
    if m:
        expected = int(m.group(1))
    # Mecz o N. miejsce
    elif src_key.startswith('mecz o '):
        expected = 1
    # MIEJSCA X-Y
    elif src_key.startswith('miejsca '):
        rng = re.match(r'miejsca (\d+)-(\d+)', src_key)
        if rng:
            x, y = int(rng.group(1)), int(rng.group(2))
            expected = max(1, (y - x + 1) // 2)
    # FINAŁ
    elif src_key == 'finał':
        expected = 1
    if expected is not None and len(matches) > expected:
        matches = matches[:expected]
    
    return phase_full_name, phase_time, matches


def fetch_drabinka_phase(sheet_id, target_phase):
    """
    Pobiera mecze z zakładki "Drabinka" dla wybranej fazy.
    
    target_phase: np. "Pucharowa 1/32 (best of 3)" lub "1/32 FINAŁU"
    
    Zwraca: (phase_name, [matches]) lub (None, [])
    """
    gid_map = get_sheet_gids(sheet_id)
    
    # Próbujemy różne warianty nazwy zakładki
    for tab_name in ('Drabinka', 'drabinka', 'DRABINKA'):
        try:
            rows = fetch_sheet(sheet_id, tab_name, gid_map)
            if rows is None: continue
            phase_name, phase_time, matches = parse_drabinka_rows(rows, target_phase)
            if matches:
                return phase_name, matches
        except Exception:
            continue
    
    return None, []


def detect_drabinka_phases(sheet_id, progress_cb=None):
    """
    Wykrywa WSZYSTKIE fazy zaprezentowane w zakładce Drabinka oraz fazę grupową.
    Zwraca dict:
        {
            'has_grupowa': bool,
            'group_count': int,
            'group_total_matches': int,
            'glowna': [{'key', 'full', 'time', 'n_matches'}, ...],
            'b':      [{'key', 'full', 'time', 'n_matches'}, ...],
            'all_times': sorted unique list of times across all phases,
        }

    progress_cb: optional callable(pct: int, msg: str) — invoked at milestones for UI progress.
    """
    def _p(pct, msg):
        if progress_cb:
            try:
                progress_cb(int(pct), msg)
            except Exception:
                pass

    result = {
        'has_grupowa': False,
        'group_count': 0,
        'group_total_matches': 0,
        'glowna': [],
        'b': [],
        'all_times': [],
    }
    _p(2, "🔌 Łączę się z arkuszem…")
    gid_map = get_sheet_gids(sheet_id)
    if gid_map:
        _p(8, f"📋 Pobrałem listę zakładek ({len(gid_map)})")
    else:
        _p(8, "⚠️ Lista zakładek niedostępna — skanuję A-Z")

    # Faza grupowa - zakładki Gr. A, Gr. B, ...
    # Optymalizacja: filtrujemy LETTERS do tych dla których nazwa zakładki
    # FAKTYCZNIE istnieje w gid_map. Zamiast 26 × 3 = 78 prób HTTP (większość
    # nieudanych), robimy tylko realne zapytania dla istniejących zakładek.
    # Fallback do pełnego skanu jeśli gid_map jest puste (np. arkusz nieskanowany).
    LETTERS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    candidate_tabs = []  # [(letter, tab_name)]
    if gid_map:
        for letter in LETTERS:
            for prefix in ('Gr. ', 'Gr.', 'Grupa '):
                tab_name = f'{prefix}{letter}'
                if tab_name in gid_map:
                    candidate_tabs.append((letter, tab_name))
                    break
    else:
        # Fallback: bez gid_map skanujemy wszystkie litery (jak poprzednio)
        candidate_tabs = [(L, f'Gr. {L}') for L in LETTERS]

    group_count = 0
    group_matches_total = 0
    total = max(1, len(candidate_tabs))
    is_fallback_scan = not gid_map
    EMPTY_STOP = 3
    empty_streak = 0
    for li, (letter, tab_name) in enumerate(candidate_tabs):
        # W trybie fallback (A-Z scan) NIE pokazujemy litery — bo zwykle 80%
        # to puste sprawdzenia. Pokazujemy tylko liczbę faktycznie znalezionych.
        if is_fallback_scan:
            _p(10 + int(35 * (li + 1) / total),
               f"🔍 Skanuję grupy… (znaleziono: {group_count})")
        else:
            _p(10 + int(35 * (li + 1) / total),
               f"📂 Wczytuję {tab_name} ({li + 1}/{total})…")
        try:
            rows = fetch_sheet(sheet_id, tab_name, gid_map)
            if rows:
                matches = parse_group_rows(rows)
                if matches:
                    group_count += 1
                    group_matches_total += len(matches)
                    empty_streak = 0
                    continue
        except Exception:
            pass
        empty_streak += 1
        if is_fallback_scan and group_count > 0 and empty_streak >= EMPTY_STOP:
            break
    if group_count > 0:
        result['has_grupowa'] = True
        result['group_count'] = group_count
        result['group_total_matches'] = group_matches_total

    _p(48, "🏆 Wczytuję zakładkę Drabinka…")
    # Drabinka — zbieramy wszystkie wykryte fazy
    drabinka_rows = None
    for tab_name in ('Drabinka', 'drabinka', 'DRABINKA'):
        try:
            r = fetch_sheet(sheet_id, tab_name, gid_map)
            if r:
                drabinka_rows = r
                break
        except Exception:
            continue
    
    if drabinka_rows:
        _p(62, "🔎 Skanuję nagłówki faz…")
        # PASS 1 — single full scan: znajdź wszystkie phase block markers
        phase_block_markers = []  # [{'header_row', 'col', 'key', 'full', 'time'}]
        for r_idx, row in enumerate(drabinka_rows):
            for c_idx, cell in enumerate(row):
                pkey, pfull = detect_phase(cell)
                if pkey is None:
                    continue
                ptime = None
                m = re.search(r'\((\d{1,2}:\d{2})\)', cell)
                if m: ptime = m.group(1)
                # Tor w lewo (max 4 col)
                col_tor = None
                for offset in range(1, 5):
                    if c_idx - offset >= 0:
                        t = row[c_idx - offset].strip().lower()
                        if t == 'tor':
                            col_tor = c_idx - offset
                            break
                phase_block_markers.append({
                    'header_row': r_idx, 'col': c_idx, 'col_tor': col_tor,
                    'key': pkey, 'full': pfull, 'time': ptime,
                })

        _p(75, f"📊 Liczę mecze w {len(phase_block_markers)} fazach…")
        # PASS 2 — dla każdego markera policz mecze schodząc w dół jego kolumny.
        # Dużo szybsze niż wcześniejsze N×parse_drabinka_rows.
        STOP_KWS = ['miejsc', 'finał', 'mecz o', '1/64', '1/32',
                    '1/16', '1/8', '1/4', '1/2', 'półfinał']
        # WAŻNE: jedna faza może być rozbita na KILKA bloków nagłówkowych w tej
        # samej kolumnie (np. "MIEJSCA 5-8" = 2 półfinały, każdy z własnym
        # nagłówkiem). Wcześniejszy `seen_keys` dedup liczył TYLKO pierwszy blok
        # → UI pokazywało np. „1 mecz" zamiast „2 mecze". Teraz SUMUJEMY mecze
        # ze wszystkich bloków tej samej fazy (po phase_key).
        agg_entries = {}  # key -> entry dict (ref współdzielony z result['b']/['glowna'])
        pairs_by_key = {}  # key -> set nieuporządkowanych par (dedup, spójne z parserem)
        for marker in phase_block_markers:
            key = marker['key']
            col = marker['col']
            i = marker['header_row'] + 1
            n_matches = 0
            empty_streak = 0
            seen_pairs = pairs_by_key.setdefault(key, set())
            while i < len(drabinka_rows):
                row1 = drabinka_rows[i]
                z1 = row1[col].strip() if col < len(row1) else ''
                if not z1:
                    empty_streak += 1
                    if empty_streak >= 5: break
                    i += 1
                    continue
                empty_streak = 0
                # Stop na pod-nagłówek nowej sekcji
                z1_lower = z1.lower()
                if any(kw in z1_lower for kw in STOP_KWS):
                    break
                # Para zawodników
                row2 = drabinka_rows[i+1] if i+1 < len(drabinka_rows) else []
                z2 = row2[col].strip() if col < len(row2) else ''
                if z1 and z2 and len(z1) >= 3 and len(z2) >= 3 \
                   and any(c.isupper() for c in z1) and any(c.isupper() for c in z2):
                    # Dedup po nieuporządkowanej parze — spójnie z parse_drabinka_rows,
                    # żeby liczba w UI = liczba realnie wygenerowanych protokołów.
                    pair_key = frozenset((z1.lower(), z2.lower()))
                    if pair_key not in seen_pairs:
                        seen_pairs.add(pair_key)
                        n_matches += 1
                i += 2

            if n_matches > 0:
                if key in agg_entries:
                    # Kolejny blok tej samej fazy — dodaj mecze do istniejącego wpisu.
                    agg_entries[key]['n_matches'] += n_matches
                    # Czas: jeśli pierwszy blok nie miał czasu, weź z tego.
                    if not agg_entries[key]['time'] and marker['time']:
                        agg_entries[key]['time'] = marker['time']
                else:
                    entry = {
                        'key': key,
                        'full': marker['full'],
                        'time': marker['time'] or '',
                        'n_matches': n_matches,
                    }
                    agg_entries[key] = entry
                    if key.startswith('miejsca ') or key.startswith('mecz o '):
                        result['b'].append(entry)
                    else:
                        result['glowna'].append(entry)
    
    # Wszystkie wykryte godziny (do filtra czasowego)
    all_times = set()
    for entry in result['glowna'] + result['b']:
        if entry['time']:
            all_times.add(entry['time'])
    result['all_times'] = sorted(all_times)

    _p(100, "✅ Gotowe")
    return result


# ─────────────────────────────────────────────────────────────────────


def get_sheet_names_debug(sheet_id, detected=None):
    """Czytelny debug oparty o detect_drabinka_phases: 1 scan, brak duplikatów,
    grupowanie faz drabinki po godzinie.

    `detected`: opcjonalny pre-policzony wynik detect_drabinka_phases.
                Jeśli podany — unikamy ponownego skanowania arkusza.
    """
    info = []
    if detected is None:
        detected = detect_drabinka_phases(sheet_id)
    
    if not detected['has_grupowa'] and not detected['glowna'] and not detected['b']:
        info.append("❌ Nie znaleziono żadnych grup z meczami ani zakładki Drabinka.")
        info.append("   Sprawdź czy arkusz jest publiczny.")
        return info
    
    # Faza grupowa
    if detected['has_grupowa']:
        ng = detected['group_count']
        nm = detected['group_total_matches']
        info.append(f"✅ Faza grupowa: {ng} {pluralize(ng, 'grupa', 'grupy', 'grup')}, "
                    f"{nm} {pluralize(nm, 'mecz', 'mecze', 'meczów')}")
    
    # Drabinka — grupowanie po godzinie
    all_phases = detected['glowna'] + detected['b']
    if all_phases:
        if info: info.append("")
        n_total = len(all_phases)
        info.append(f"✅ Drabinka: {n_total} {pluralize(n_total, 'faza', 'fazy', 'faz')} "
                    "(pogrupowane po godzinie startu):")
        
        # Grupuj po godzinie
        by_time = {}
        no_time = []
        for p in all_phases:
            t = p['time']
            if t:
                by_time.setdefault(t, []).append(p)
            else:
                no_time.append(p)
        
        for t in sorted(by_time.keys()):
            # Sortuj fazy NUMERYCZNIE (nie alfabetycznie):
            # • "1/N FINAŁU" → po N (mniejsze N = późniejsza runda, ważniejsza)
            # • "MIEJSCA X-Y" → po X
            # • "MECZ O N. MIEJSCE" → po N (9 przed 11, nie 11 przed 9)
            def _phase_sort_key(p):
                full_lower = p['full'].lower()
                m = re.search(r'1/(\d+)', full_lower)
                if m: return (0, int(m.group(1)))
                if 'finał' in full_lower and 'półfinał' not in full_lower:
                    return (0, 1)  # FINAŁ traktowany jak 1/1
                m = re.search(r'miejsca\s+(\d+)', full_lower)
                if m: return (1, int(m.group(1)))
                m = re.search(r'mecz\s+o\s+(\d+)', full_lower)
                if m: return (2, int(m.group(1)))
                return (3, 0)
            phases_at_t = sorted(by_time[t], key=_phase_sort_key)
            info.append(f"  🕐 {t}:")
            for p in phases_at_t:
                n_m = p['n_matches']
                info.append(f"     • {p['full']}: {n_m} {pluralize(n_m, 'mecz', 'mecze', 'meczów')}")
        if no_time:
            info.append("  🕐 (bez wykrytej godziny):")
            for p in sorted(no_time, key=lambda x: x['full']):
                n_m = p['n_matches']
                info.append(f"     • {p['full']}: {n_m} {pluralize(n_m, 'mecz', 'mecze', 'meczów')}")
    
    return info


# ─── XML helpers: tekst w komórce ─────────────────────────────────────────────

def _set_cell_value(tc, text, *, bold=True, size=28, align='center'):
    # Auto-shrink: długie nazwiska/nazwy zespołów zmniejszamy żeby zmieściły się
    # w jednym wierszu. Gdy znamy szerokość komórki, dobieramy największy rozmiar
    # przy którym tekst się mieści (lepiej niż stałe progi: szersza komórka =
    # większy font). Fallback (bez tcW): oryginalne progi.
    if text and size >= 24:
        n = len(text)
        cell_w = 0
        tcPr0 = tc.find(wt('tcPr'))
        if tcPr0 is not None:
            tcW_el = tcPr0.find(wt('tcW'))
            if tcW_el is not None:
                try:
                    cell_w = int(tcW_el.get(f'{{{W}}}w', '0'))
                except (TypeError, ValueError):
                    cell_w = 0
        if cell_w > 400:
            # Calibri PL (Carlito w LO): avg char ≈ sz_twentieths × 5.0 dxa (empirycznie
            # zmierzone na „Stowarzyszenie Aktywny Orlik I" w komórce 3060 dxa — przy 4.2
            # sz=20 jeszcze zawijał ostatni znak). Margines komórki + rezerwa ~200 dxa.
            usable = cell_w - 200
            # Pełen zakres: 28→14. Min 14 (7pt) jest jeszcze komfortowo czytelne,
            # mieści ~35 znaków w komórce 2800 dxa (DWÓJKA Bo7 nazwa drużyny: 3009
            # dxa). User: „wolałbym ciut zmniejszoną czcionkę niż kolejną linię
            # i rozjechanie w tabelce". Bardzo długie (>40 znaków) wciąż zawiną,
            # ale poniżej 14pt tekst robi się nieczytelny — granica praktyczna.
            chosen = None
            for try_sz in (28, 26, 24, 22, 20, 18, 16, 14):
                if try_sz > size:
                    continue
                if n * try_sz * 5.0 <= usable:
                    chosen = try_sz
                    break
            size = chosen if chosen is not None else 14
        else:
            if n > 34:
                size = 14
            elif n > 30:
                size = 16
            elif n > 28:
                size = 18
            elif n > 24:
                size = 20
            elif n > 19:
                size = 22
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
    for a in ('ascii','hAnsi','eastAsia','cs'):
        # Calibri (nie Aptos) — Aptos nie jest znany LibreOffice (Streamlit Cloud) i
        # podstawiany jest fontem SZERYFOWYM, przez co nazwiska/wartości wychodziły
        # szeryfowe. Calibri ma zamiennik Carlito (bezszeryfowy) → spójnie z etykietami.
        fonts.set(f'{{{W}}}{a}', 'Calibri')
    if bold:
        for tag in ('b','bCs'):
            etree.SubElement(rPr, wt(tag)).set(f'{{{W}}}val','1')
    for tag in ('sz','szCs'):
        etree.SubElement(rPr, wt(tag)).set(f'{{{W}}}val', str(size))
    etree.SubElement(rPr, wt('lang')).set(f'{{{W}}}val','pl-PL')
    t = etree.SubElement(r, wt('t'))
    t.text = text


def _set_cell_label(tc, text):
    """Zmień TEKST etykiety w komórce, ZACHOWUJĄC oryginalne formatowanie
    (font, size, bold) z szablonu. Używane do podmiany 'Grupa' → 'Faza'
    żeby etykieta wyglądała identycznie jak 'Tor', 'Godzina', 'Mecz #'."""
    # Znajdź pierwszy run z tekstem
    for p in tc.findall(wt('p')):
        for r in p.findall(wt('r')):
            ts = r.findall(wt('t'))
            if ts:
                # Wyczyść wszystkie t i ustaw nowy tekst w pierwszym
                for i, t in enumerate(ts):
                    if i == 0:
                        t.text = text
                    else:
                        r.remove(t)
                # Pozostałe runs w tym paragrafie mogą zawierać dodatkowy tekst
                # (np. spację między 'Mecz' a '#') - usuwamy je
                runs_to_remove = []
                found_first = False
                for r2 in p.findall(wt('r')):
                    if r2 is r:
                        found_first = True
                        continue
                    if found_first:
                        # Usuwamy wszystko po pierwszym runie z tekstem
                        ts2 = r2.findall(wt('t'))
                        if ts2:
                            runs_to_remove.append(r2)
                for r2 in runs_to_remove:
                    p.remove(r2)
                return
    # Fallback: komórka nie miała żadnego tekstu - dodaj nowy paragraf jak _set_cell_value
    # ale BEZ bold (etykiety z szablonu mają b=0)
    _set_cell_value(tc, text, bold=False, size=24)


def _fix_pkt_set_cells(body, narrow='16', wide='20'):
    """Przebudowuje komórki 'Pkt SET N' / 'Punkty SET N' na CZYSTE 2 linie:
    'Pkt'/'Punkty' + <br/> + 'SET N'. Szablony mają 2× <w:br/> + litery jako
    osobne runy, co dawało 3 linie. Font Calibri, rozmiar zależny od szerokości
    komórki (wąskie <800 dxa → narrow, szersze → wide)."""
    import re as _re
    XMLSPACE = '{http://www.w3.org/XML/1998/namespace}space'
    for tbl in body.findall(wt('tbl')):
        for tr in tbl.findall(wt('tr')):
            for tc in tr.findall(wt('tc')):
                full = ''.join((t.text or '') for t in tc.iter(wt('t')))
                m = _re.search(r'(Punkty|Pkt)\.?\s*SET\s*(\d)', full.replace('\n', ''))
                if not m:
                    continue
                # NOWE szablony mają już CZYSTĄ 1-linijną postać „Pkt. SET N" bez <w:br/>.
                # Tylko STARE (z 2× <w:br/> + literami w osobnych runach) trzeba przebudować.
                # Bez tej kontroli psujemy nowe szablony — runy „." i „ " znikały,
                # zostawało „PktSET 1" jako jedno słowo i LO wrapowało na 2 linie.
                has_break = any(True for _ in tc.iter(wt('br')))
                if not has_break:
                    continue
                prefix, setnum = m.group(1), m.group(2)
                tcW = tc.find(f'{wt("tcPr")}/{wt("tcW")}')
                cw = int(tcW.get(f'{{{W}}}w', '990')) if tcW is not None else 990
                sz = narrow if cw < 800 else wide
                ps = tc.findall(wt('p'))
                if not ps:
                    continue
                p = ps[0]
                for extra in ps[1:]:
                    tc.remove(extra)
                for r in p.findall(wt('r')):
                    p.remove(r)
                pPr = p.find(wt('pPr'))
                if pPr is None:
                    pPr = etree.Element(wt('pPr')); p.insert(0, pPr)
                jc = pPr.find(wt('jc'))
                if jc is None:
                    jc = etree.SubElement(pPr, wt('jc'))
                jc.set(f'{{{W}}}val', 'center')

                def _mkrun(txt, with_break):
                    r = etree.SubElement(p, wt('r'))
                    rPr = etree.SubElement(r, wt('rPr'))
                    fonts = etree.SubElement(rPr, wt('rFonts'))
                    for a in ('ascii', 'hAnsi', 'eastAsia', 'cs'):
                        fonts.set(f'{{{W}}}{a}', 'Calibri')
                    for tag in ('b', 'bCs'):
                        etree.SubElement(rPr, wt(tag)).set(f'{{{W}}}val', '1')
                    for tag in ('sz', 'szCs'):
                        etree.SubElement(rPr, wt(tag)).set(f'{{{W}}}val', sz)
                    if with_break:
                        etree.SubElement(r, wt('br'))
                    t = etree.SubElement(r, wt('t'))
                    t.text = txt
                    t.set(XMLSPACE, 'preserve')

                _mkrun(prefix, False)
                _mkrun(f'SET {setnum}', True)
                tcPr = tc.find(wt('tcPr'))
                if tcPr is not None:
                    vA = tcPr.find(wt('vAlign'))
                    if vA is None:
                        vA = etree.SubElement(tcPr, wt('vAlign'))
                    vA.set(f'{{{W}}}val', 'center')


def _force_calibri_score_labels(body):
    """Wymusza Calibri na etykietach w tabelach wynikowych (IMIONA / SUMA / WYNIK /
    PKT). Litery są w osobnych runach z fontem 'Aptos Narrow' (brak w LibreOffice
    → krzywy fallback). Dotyczy tabel z >10 wierszami (score tables)."""
    for tbl in body.findall(wt('tbl')):
        rows = tbl.findall(wt('tr'))
        if len(rows) <= 10:
            continue
        # 1) Calibri na WSZYSTKICH runach z tekstem w tabeli wynikowej.
        for r in list(tbl.iter(wt('r'))):
            if not r.findall(wt('t')):
                continue
            rPr = r.find(wt('rPr'))
            if rPr is None:
                rPr = etree.Element(wt('rPr')); r.insert(0, rPr)
            fonts = rPr.find(wt('rFonts'))
            if fonts is None:
                fonts = etree.SubElement(rPr, wt('rFonts'))
            for a in ('ascii', 'hAnsi', 'eastAsia', 'cs'):
                fonts.set(f'{{{W}}}{a}', 'Calibri')
        # 2) Ostatni wiersz (WYNIK/PKT) → sz=16, by etykieta się nie łamała.
        #    Iterujemy BEZPOŚREDNIO po runach ostatniego wiersza. NIE używamy id():
        #    lxml tworzy nowe proxy-obiekty przy każdej iteracji, a id() zwolnionego
        #    proxy potrafi kolidować z innym runem → losowo psuło rozmiar etykiet
        #    (np. SET 1 / SET 3 w nagłówku tabeli stawały się 16 zamiast 28).
        for r in list(rows[-1].iter(wt('r'))):
            if not r.findall(wt('t')):
                continue
            rPr = r.find(wt('rPr'))
            if rPr is None:
                rPr = etree.Element(wt('rPr')); r.insert(0, rPr)
            for tag in ('sz', 'szCs'):
                el = rPr.find(wt(tag))
                if el is None:
                    el = etree.SubElement(rPr, wt(tag))
                el.set(f'{{{W}}}val', '16')


def _make_page_break_para():
    p = etree.Element(wt('p'))
    pPr = etree.SubElement(p, wt('pPr'))
    sp = etree.SubElement(pPr, wt('spacing'))
    sp.set(f'{{{W}}}before','0'); sp.set(f'{{{W}}}after','0')
    r = etree.SubElement(p, wt('r'))
    br = etree.SubElement(r, wt('br'))
    br.set(f'{{{W}}}type','page')
    return p


def _fill_protocol(elements, match, hide_grupa_mecz=False, phase_label=None,
                   template_type='IND'):
    """`hide_grupa_mecz` to historyczna nazwa — obecnie kontroluje tylko czy etykieta
    'Grupa' i jej wartość są ukrywane (dla pucharowej). 'Mecz #' i jego numer
    są ZAWSZE pokazywane (też w pucharowej — numerujemy mecze sekwencyjnie)."""
    tbls = [el for el in elements if el.tag == wt('tbl')]
    if not tbls: return
    rows = tbls[0].findall(wt('tr'))

    # ── Bo3 IND: pucharowa 1v1, header layout jak TROJKA_Bo3 ────────────
    # R0: [Tor 3255] [Godz. 990] [empty 990] [empty 990] [empty 1275] [Mecz # 1860]
    # R1: [empty] [PunktySET 1] [PunktySET 2] [PunktySET 3] [Wygrane sety] [Podpis]
    # R2: zawodnik A (z1)
    # R3: zawodnik B (z2) — w pucharowej 1v1 OBA mają być wypełnione
    if template_type == 'IND_Bo3':
        if len(rows) >= 1:
            tcs = rows[0].findall(wt('tc'))
            tor_val = match.get('tor', '').strip()
            if len(tcs) > 0 and tor_val:
                _set_cell_label(tcs[0], f'Tor  {tor_val}')
            godz_val = match.get('godz', '').strip()
            if len(tcs) > 2 and godz_val:
                _set_cell_value(tcs[2], godz_val, size=22, bold=True, align='left')
            mecz_val = match.get('mecz', '').strip()
            if len(tcs) > 5:
                if mecz_val:
                    _set_cell_label(tcs[5], f'Mecz #  {mecz_val}')
                else:
                    _set_cell_label(tcs[5], '')
        # z1 → R2.tc[0], z2 → R3.tc[0]
        if len(rows) > 2:
            tcs = rows[2].findall(wt('tc'))
            if tcs: _set_cell_value(tcs[0], match.get('z1', ''), size=24, align='right')
        if len(rows) > 3:
            tcs = rows[3].findall(wt('tc'))
            if tcs: _set_cell_value(tcs[0], match.get('z2', ''), size=24, align='right')
        return  # IND_Bo3 ma własną logikę, nie kontynuujemy

    # ── Bo5 IND: pucharowa 1v1, 2 strony (twardy page break w szablonie) ──
    # Szablon ma 4 tabele: all_tbls[0]=nagłówek str.1, [1]=wynik SET 1-3,
    # [2]=nagłówek str.2 (identyczny: Tor/Godz/Mecz#), [3]=wynik (SET 4)(SET 5).
    # Wiersze nagłówka: r0 [Tor(0) _ Godz.(2) godzVal(3) _ Mecz#(5)], r2/r3 = imiona.
    if template_type == 'IND_Bo5':
        tor_val = match.get('tor', '').strip()
        godz_val = match.get('godz', '').strip()
        mecz_val = match.get('mecz', '').strip()
        z1 = match.get('z1', '')
        z2 = match.get('z2', '')
        all_tbls = [el for el in elements if el.tag == wt('tbl')]

        def _fill_bo5_header(tbl):
            hrows = tbl.findall(wt('tr'))
            if not hrows:
                return
            tcs = hrows[0].findall(wt('tc'))
            if len(tcs) > 0 and tor_val:
                _set_cell_label(tcs[0], f'Tor  {tor_val}')
            if len(tcs) > 3 and godz_val:
                _set_cell_value(tcs[3], godz_val, size=24, bold=True, align='left')
            if len(tcs) > 5:
                _set_cell_label(tcs[5], f'Mecz #  {mecz_val}' if mecz_val else '')
            if len(hrows) > 2:
                tcs = hrows[2].findall(wt('tc'))
                if tcs: _set_cell_value(tcs[0], z1, size=24, align='right')
            if len(hrows) > 3:
                tcs = hrows[3].findall(wt('tc'))
                if tcs: _set_cell_value(tcs[0], z2, size=24, align='right')

        if len(all_tbls) >= 1:
            _fill_bo5_header(all_tbls[0])   # nagłówek str.1
        if len(all_tbls) >= 3:
            _fill_bo5_header(all_tbls[2])   # nagłówek str.2
        return  # IND_Bo5 ma własną logikę

    # ── Bo7 IND: pucharowa 1v1, 2 strony (twardy page break w szablonie) ──
    # Identyczna struktura 4-tabelowa jak Bo5: all_tbls[0]=nagłówek str.1,
    # [1]=wynik SET 1-4, [2]=nagłówek str.2, [3]=wynik (SET 5)(SET 6)(SET 7).
    # RÓŻNICA vs Bo5: user USUNĄŁ etykietę „Mecz #" (Bo7 to praktycznie tylko
    # finał — jeden mecz, numer zbędny). Nagłówek r0 ma 10 komórek:
    # tcs[0]=Tor(label), tcs[2]=Godz(label), tcs[3]=godz-val, reszta filler.
    # Imiona w hrows[2]/hrows[3].tc[0] (jak Bo5).
    if template_type == 'IND_Bo7':
        tor_val = match.get('tor', '').strip()
        godz_val = match.get('godz', '').strip()
        mecz_val = match.get('mecz', '').strip()
        z1 = match.get('z1', '')
        z2 = match.get('z2', '')
        all_tbls = [el for el in elements if el.tag == wt('tbl')]

        def _fill_bo7_header(tbl):
            hrows = tbl.findall(wt('tr'))
            if not hrows:
                return
            tcs = hrows[0].findall(wt('tc'))
            if len(tcs) > 0 and tor_val:
                _set_cell_label(tcs[0], f'Tor  {tor_val}')
            if len(tcs) > 3 and godz_val:
                _set_cell_value(tcs[3], godz_val, size=24, bold=True, align='left')
            # „Mecz #": numer dopisany do etykiety w tcs[4] (jak w TRÓJCE/CZWÓRCE) —
            # komórka jest vAlign=center, więc „Mecz #  1" jest pionowo wyśrodkowane
            # (oddzielna komórka wartości tcs[5] nie ma vAlign → wartość lądowała u góry).
            # Dla faz z jednym meczem (Finał, Mecz o N. miejsce) build_document zeruje
            # match['mecz'] → czyścimy całą etykietę, żeby nie wisiała samotnie.
            if len(tcs) > 4:
                _set_cell_label(tcs[4], f'Mecz #  {mecz_val}' if mecz_val else '')
            if len(hrows) > 2:
                tcs = hrows[2].findall(wt('tc'))
                if tcs: _set_cell_value(tcs[0], z1, size=24, align='right')
            if len(hrows) > 3:
                tcs = hrows[3].findall(wt('tc'))
                if tcs: _set_cell_value(tcs[0], z2, size=24, align='right')

        if len(all_tbls) >= 1:
            _fill_bo7_header(all_tbls[0])   # nagłówek str.1
        if len(all_tbls) >= 3:
            _fill_bo7_header(all_tbls[2])   # nagłówek str.2
        return  # IND_Bo7 ma własną logikę

    # ── Bo3/Bo5 CZWÓRKA + DWÓJKA Bo7: landscape, osobne komórki wartości
    #    (tc[1]=tor, tc[3]=godz, tc[5]=mecz) + nazwy drużyn w r1.tc[1] / r2.tc[1].
    #    Header w T0 (i T2 dla Bo5/Bo7). DWÓJKA Bo7 ma identyczną strukturę co
    #    CZWÓRKA Bo5 (landscape) — różni się tylko liczbą zawodników (2 vs 4). ──
    if template_type in ('CZWORKA_Bo3', 'CZWORKA_Bo5', 'DWOJKA_Bo7'):
        tor_val = match.get('tor', '').strip()
        godz_val = match.get('godz', '').strip()
        mecz_val = match.get('mecz', '').strip()
        z1 = match.get('z1', '')
        z2 = match.get('z2', '')
        all_tbls = [el for el in elements if el.tag == wt('tbl')]
        # Header tables: T0 (idx 0) str.1; T2 (idx 2) str.2 — tylko Bo5.
        for t_idx in (0, 2):
            if len(all_tbls) <= t_idx:
                continue
            trows = all_tbls[t_idx].findall(wt('tr'))
            if not trows:
                continue
            r0 = trows[0].findall(wt('tc'))
            # Wartości po lewej — komórki są szerokie (zwł. tc[5] dla Mecz#: 2520 dxa),
            # przy center wartość ląduje daleko od etykiety. Left = tuż obok labela.
            if len(r0) > 1 and tor_val:
                _set_cell_value(r0[1], tor_val, size=24, bold=True, align='left')
            if len(r0) > 3 and godz_val:
                _set_cell_value(r0[3], godz_val, size=24, bold=True, align='left')
            if len(r0) > 5:
                _set_cell_value(r0[5], mecz_val, size=24, bold=True, align='left')

            # Lock wysokości wierszy z nazwami zespołów (hRule=exact) — różne wersje
            # LibreOffice/Carlito renderują tekst minimalnie inaczej i bez locka długi
            # tekst zawijał się w produkcji, rozdmuchując wiersz (R2 stawał się
            # dwukrotnie wyższy od R1 — widać na PDF wygenerowanym ze Streamlita).
            def _lock_row(tr):
                # Tylko hRule=exact (zapobiega expansji R2 gdy LO/Carlito wrapsuje
                # tekst). Wartość zostawiamy z szablonu (R2=215) — wymuszanie większej
                # rozszerza TBL2 i jej R0 wycieka na koniec poprzedniej strony.
                trPr = tr.find(wt('trPr'))
                if trPr is None:
                    trPr = etree.Element(wt('trPr'))
                    tr.insert(0, trPr)
                trH = trPr.find(wt('trHeight'))
                if trH is None:
                    trH = etree.SubElement(trPr, wt('trHeight'))
                    trH.set(f'{{{W}}}val', '215')
                trH.set(f'{{{W}}}hRule', 'exact')

            if len(trows) > 1:
                r1 = trows[1].findall(wt('tc'))
                if len(r1) > 1:
                    _set_cell_value(r1[1], z1, size=24, align='center')
                # R1 NIE lockujemy — zawiera komórkę z instrukcjami (4 linie tekstu)
                # która musi auto-grow do mieszczenia całości.
            if len(trows) > 2:
                r2 = trows[2].findall(wt('tc'))
                if len(r2) > 1:
                    _set_cell_value(r2[1], z2, size=24, align='center')
                _lock_row(trows[2])  # tylko R2 (długa nazwa zespołu) — w nim było puchnięcie
        return  # CZWÓRKA Bo3/Bo5 ma własną logikę

    # ── DWÓJKA_Bo3: layout r0 = [Tor | Tor-val | Godzina | Godz-val(span2) |
    #    filler×3 (span2 each) | Mecz # | Mecz-val(span2)]. Drużyny w r3/r4 c0.
    if template_type == 'DWOJKA_Bo3':
        if len(rows) >= 1:
            tcs = rows[0].findall(wt('tc'))
            tor_val = match.get('tor', '').strip()
            godz_val = match.get('godz', '').strip()
            mecz_val = match.get('mecz', '').strip()
            if len(tcs) > 1 and tor_val:
                _set_cell_value(tcs[1], tor_val, size=22, bold=True)
            if len(tcs) > 3 and godz_val:
                _set_cell_value(tcs[3], godz_val, size=22, bold=True)
            if len(tcs) > 8 and mecz_val:
                _set_cell_value(tcs[8], mecz_val, size=22, bold=True)
        # Drużyny w r3.c0 i r4.c0
        if len(rows) > 3:
            tcs = rows[3].findall(wt('tc'))
            if tcs: _set_cell_value(tcs[0], match.get('z1', ''), size=24, align='right')
        if len(rows) > 4:
            tcs = rows[4].findall(wt('tc'))
            if tcs: _set_cell_value(tcs[0], match.get('z2', ''), size=24, align='right')
        return

    # ── DWÓJKA_Bo5: layout TROJKA_Bo5-style ale z 8 komórkami w r0:
    #    c0=Tor (label+val w jednej szerokiej komórce 2991dxa = 5.27 cm)
    #    c1=Godz. (label, 850dxa)  c2=godz-value (850dxa)
    #    c3..c5 i c6: pustego spacingu
    #    c7=Mecz # (1890dxa, label+val w jednej komórce)
    # Drużyny: r3.c0 (z1), r4.c0 (z2). T2 (str.2 header) ma identyczną strukturę,
    # wypełniany analogicznie. r1 ma 8 etykiet Pkt SET 1..5 + Wygrane sety + Podpis
    # — bez wartości do wstawiania.
    if template_type == 'DWOJKA_Bo5':
        tor_val = match.get('tor', '').strip()
        godz_val = match.get('godz', '').strip()
        mecz_val = match.get('mecz', '').strip()
        z1 = match.get('z1', '')
        z2 = match.get('z2', '')
        # Tor/Mecz # wstawiamy przez _set_cell_label (zachowuje szablonowy font),
        # ale w innych szablonach (CZWÓRKA/DWÓJKA Bo3) te wartości są POGRUBIONE —
        # template DWÓJKA Bo5 ma label nie-bold, więc forsujemy bold na całej
        # komórce (label + wartość), żeby było spójnie. (Godz. value już bold.)
        def _bold_cell(tc):
            for run in tc.iter(wt('r')):
                if not run.findall(wt('t')):
                    continue
                rPr = run.find(wt('rPr'))
                if rPr is None:
                    rPr = etree.Element(wt('rPr')); run.insert(0, rPr)
                for btag in ('b', 'bCs'):
                    if rPr.find(wt(btag)) is None:
                        etree.SubElement(rPr, wt(btag))
        # Strona 1 (T1)
        if len(rows) >= 1:
            tcs = rows[0].findall(wt('tc'))
            if len(tcs) > 0 and tor_val:
                _set_cell_label(tcs[0], f'Tor  {tor_val}')
                _bold_cell(tcs[0])
            if len(tcs) > 2 and godz_val:
                _set_cell_value(tcs[2], godz_val, size=22, bold=True, align='left')
            if len(tcs) >= 8:
                if mecz_val:
                    _set_cell_label(tcs[-1], f'Mecz #  {mecz_val}')
                    _bold_cell(tcs[-1])
                else:
                    _set_cell_label(tcs[-1], '')
        if len(rows) > 2:
            tcs = rows[2].findall(wt('tc'))
            if tcs: _set_cell_value(tcs[0], z1, size=24, align='right')
        if len(rows) > 3:
            tcs = rows[3].findall(wt('tc'))
            if tcs: _set_cell_value(tcs[0], z2, size=24, align='right')
        # Strona 2 (T3, czyli tbls[2]) — identyczny layout
        all_tbls = [el for el in elements if el.tag == wt('tbl')]
        if len(all_tbls) >= 3:
            t3 = all_tbls[2]
            t3_rows = t3.findall(wt('tr'))
            if t3_rows:
                tcs = t3_rows[0].findall(wt('tc'))
                if len(tcs) > 0 and tor_val:
                    _set_cell_label(tcs[0], f'Tor  {tor_val}')
                    _bold_cell(tcs[0])
                if len(tcs) > 2 and godz_val:
                    _set_cell_value(tcs[2], godz_val, size=22, bold=True, align='left')
                if len(tcs) >= 8:
                    if mecz_val:
                        _set_cell_label(tcs[-1], f'Mecz #  {mecz_val}')
                        _bold_cell(tcs[-1])
                    else:
                        _set_cell_label(tcs[-1], '')
            if len(t3_rows) > 2:
                tcs = t3_rows[2].findall(wt('tc'))
                if tcs: _set_cell_value(tcs[0], z1, size=24, align='right')
            if len(t3_rows) > 3:
                tcs = t3_rows[3].findall(wt('tc'))
                if tcs: _set_cell_value(tcs[0], z2, size=24, align='right')
        return

    if template_type in ('TROJKA_Bo3',):
        # Bo3 template R1 cells (TROJKA i CZWORKA mają identyczny layout header):
        #   tc[0] "Tor" 5.24 cm — label + value w jednej komórce
        #   tc[1] "Godz." 1.87 cm — TYLKO label (za wąska na "Godz. 13:00")
        #   tc[2] pusta 1.87 cm — wartość godziny
        #   tc[3] pusta 1.87 cm — pozostawiamy pustą (spacing)
        #   tc[4] pusta 2.20 cm — pozostawiamy pustą (spacing)
        #   tc[5] "Runda" 3.24 cm — USUWAMY label "Runda" (faza już w nagłówku
        #         strony w prawym górnym rogu) i wstawiamy tu numer meczu "Mecz X"
        if len(rows) >= 1:
            tcs = rows[0].findall(wt('tc'))
            tor_val = match.get('tor', '').strip()
            if len(tcs) > 0 and tor_val:
                _set_cell_label(tcs[0], f'Tor  {tor_val}')
            godz_val = match.get('godz', '').strip()
            if len(tcs) > 2 and godz_val:
                _set_cell_value(tcs[2], godz_val, size=22, bold=True, align='left')
            mecz_val = match.get('mecz', '').strip()
            if len(tcs) > 5:
                if mecz_val:
                    # Podmieniamy "Runda" na "Mecz X"
                    _set_cell_label(tcs[5], f'Mecz #  {mecz_val}')
                else:
                    # Bez numeru meczu - czyścimy całkowicie (faza w nagłówku)
                    _set_cell_label(tcs[5], '')
        # Team A → R3.tc[0] (rows[2]), Team B → R4.tc[0] (rows[3])
        if len(rows) > 2:
            tcs = rows[2].findall(wt('tc'))
            if tcs: _set_cell_value(tcs[0], match.get('z1', ''), size=24, align='right')
        if len(rows) > 3:
            tcs = rows[3].findall(wt('tc'))
            if tcs: _set_cell_value(tcs[0], match.get('z2', ''), size=24, align='right')
        return  # Bo3 ma własną logikę, nie kontynuujemy ze standardową
    
    if template_type == 'TROJKA_Bo5':
        # Bo5 ma DWIE strony do wypełnienia (oddzielone pustą stroną):
        #  Strona 1 (Table 1): Tor / Godz. / Runda + Punkty SET 1-3 + Wygrane + Podpis
        #  Strona 3 (Table 3): Tor / Runda + Pkt SET 1-5 + Wygrane + Podpis  (NO Godz.!)
        # Logika identyczna z Bo3 dla strony 1, dodatkowo wypełniamy stronę 3.
        tor_val = match.get('tor', '').strip()
        godz_val = match.get('godz', '').strip()
        mecz_val = match.get('mecz', '').strip()
        z1 = match.get('z1', '')
        z2 = match.get('z2', '')
        
        # ── Strona 1 (Table 1, R1) — taka sama logika jak w Bo3 ──
        if len(rows) >= 1:
            tcs = rows[0].findall(wt('tc'))
            if len(tcs) > 0 and tor_val:
                _set_cell_label(tcs[0], f'Tor  {tor_val}')
            if len(tcs) > 2 and godz_val:
                _set_cell_value(tcs[2], godz_val, size=22, bold=True, align='left')
            # „Mecz #": szablon TRÓJKA_Bo5 ma etykietę w OSTATNIEJ komórce (c7)
            # — wcześniej fill pisał do c5 (pustej) i renderowały się DWA „Mecz #".
            if len(tcs) > 5:
                target = tcs[-1]
                if mecz_val:
                    _set_cell_label(target, f'Mecz #  {mecz_val}')
                else:
                    _set_cell_label(target, '')
        # Drużyny w T1.R3.tc[0] i T1.R4.tc[0]
        if len(rows) > 2:
            tcs = rows[2].findall(wt('tc'))
            if tcs: _set_cell_value(tcs[0], z1, size=24, align='right')
        if len(rows) > 3:
            tcs = rows[3].findall(wt('tc'))
            if tcs: _set_cell_value(tcs[0], z2, size=24, align='right')
        
        # ── Strona 3 (Table 3) — analogiczna struktura do T1 ──
        # Aktualny szablon TRÓJKA_Bo5 ma „Godz." też na str.2 (wcześniejsze
        # wersje nie miały — stąd stary komentarz „NO Godz."). Wypełniamy.
        all_tbls = [el for el in elements if el.tag == wt('tbl')]
        if len(all_tbls) >= 3:
            t3 = all_tbls[2]
            t3_rows = t3.findall(wt('tr'))
            if t3_rows:
                tcs = t3_rows[0].findall(wt('tc'))
                if len(tcs) > 0 and tor_val:
                    _set_cell_label(tcs[0], f'Tor  {tor_val}')
                if len(tcs) > 2 and godz_val:
                    _set_cell_value(tcs[2], godz_val, size=22, bold=True, align='left')
                # „Mecz #" w ostatniej komórce r0 (analogicznie do str.1).
                if len(tcs) >= 7:
                    if mecz_val:
                        _set_cell_label(tcs[-1], f'Mecz #  {mecz_val}')
                    else:
                        _set_cell_label(tcs[-1], '')
            # Drużyny powtarzane w T3.R3.tc[0] i T3.R4.tc[0]
            if len(t3_rows) > 2:
                tcs = t3_rows[2].findall(wt('tc'))
                if tcs: _set_cell_value(tcs[0], z1, size=24, align='right')
            if len(t3_rows) > 3:
                tcs = t3_rows[3].findall(wt('tc'))
                if tcs: _set_cell_value(tcs[0], z2, size=24, align='right')
        return  # Bo5 ma własną logikę
    
    # ── Standard (IND, TROJKA Bo2): zachowane bez zmian ───────────────────
    if len(rows) > 0:
        tcs = rows[0].findall(wt('tc'))
        if len(tcs) > 1: _set_cell_value(tcs[1], match.get('tor',''),  size=28)
        if len(tcs) > 3: _set_cell_value(tcs[3], match.get('godz',''), size=28)
        # Mecz # ZAWSZE
        if len(tcs) > 7: _set_cell_value(tcs[7], match.get('mecz',''), size=28)
        if hide_grupa_mecz:
            # Faza pucharowa: ukryj etykietę "Grupa" i jej wartość
            # (faza jest pokazywana w prawym górnym rogu obok nazwy/daty turnieju)
            if len(tcs) > 4: _set_cell_label(tcs[4], '')
            if len(tcs) > 5: _set_cell_value(tcs[5], '', size=28)
        else:
            if len(tcs) > 5: _set_cell_value(tcs[5], match.get('grupa',''),size=28)
    if len(rows) > 3:
        tcs = rows[3].findall(wt('tc'))
        if tcs: _set_cell_value(tcs[0], match.get('z1',''), size=24, align='right')
    if len(rows) > 4:
        tcs = rows[4].findall(wt('tc'))
        if tcs: _set_cell_value(tcs[0], match.get('z2',''), size=24, align='right')


def _make_inline_image_drawing(rel_id, cx_emu, cy_emu):
    """Inline obraz (nie anchored) — siedzi w paragrafie, layout przez flow.
    Używany dla CZWORKA strip table: każda komórka tabeli ma 1 inline image,
    pozycjonowanie przez komórki tabeli (deterministyczne)."""
    uid = _next_uid()
    return etree.fromstring(f'''<w:drawing xmlns:w="{W}"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
  <wp:inline distT="0" distB="0" distL="0" distR="0">
    <wp:extent cx="{int(cx_emu)}" cy="{int(cy_emu)}"/>
    <wp:effectExtent l="0" t="0" r="0" b="0"/>
    <wp:docPr id="{uid}" name="Picture {uid}"/>
    <wp:cNvGraphicFramePr>
      <a:graphicFrameLocks noChangeAspect="1"/>
    </wp:cNvGraphicFramePr>
    <a:graphic>
      <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
        <pic:pic>
          <pic:nvPicPr>
            <pic:cNvPr id="{uid}" name=""/>
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
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          </pic:spPr>
        </pic:pic>
      </a:graphicData>
    </a:graphic>
  </wp:inline>
</w:drawing>''')


def _make_czworka_strip_table(image_specs):
    """Buduje tabelę-strip dla CZWORKA: jeden wiersz, ALTERNATING cele:
    [img_cell][spacer_cell][img_cell][spacer_cell]...[img_cell]
    Każdy img_cell zawiera 1 inline image. Total table width = 10466 dxa (18.46 cm).

    image_specs: [(rel_id, w_cm, h_cm)] w kolejności od lewej do prawej.
    Bez borderów, bez paddingu. Layout deterministyczny przez table widths."""
    TABLE_W_DXA = 10466  # 18.46 cm
    TABLE_W_CM = 18.46
    DXA_PER_CM = 567.0

    n = len(image_specs)
    if n == 0:
        return None
    total_img_w_cm = sum(w for _, w, _ in image_specs)
    if n == 1:
        gap_cm = 0
    else:
        gap_cm = (TABLE_W_CM - total_img_w_cm) / (n - 1)
        if gap_cm < 0:
            # Skaluj wszystkie do pełnej szerokości bez gapów
            scale = TABLE_W_CM / total_img_w_cm
            image_specs = [(rid, w * scale, h * scale) for rid, w, h in image_specs]
            total_img_w_cm = TABLE_W_CM
            gap_cm = 0

    # Build tblGrid: [img1_w, gap, img2_w, gap, ..., imgN_w] in dxa
    grid_widths_dxa = []
    for i, (rid, w_cm, h_cm) in enumerate(image_specs):
        grid_widths_dxa.append(int(w_cm * DXA_PER_CM))
        if i < n - 1:
            grid_widths_dxa.append(int(gap_cm * DXA_PER_CM))

    # Adjust last cell to make total exactly TABLE_W_DXA (rounding)
    diff = TABLE_W_DXA - sum(grid_widths_dxa)
    if diff != 0:
        grid_widths_dxa[-1] += diff

    tbl = etree.Element(wt('tbl'))
    tblPr = etree.SubElement(tbl, wt('tblPr'))
    tblW = etree.SubElement(tblPr, wt('tblW'))
    tblW.set(f'{{{W}}}w', str(TABLE_W_DXA))
    tblW.set(f'{{{W}}}type', 'dxa')
    # No borders
    tblBorders = etree.SubElement(tblPr, wt('tblBorders'))
    for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        b = etree.SubElement(tblBorders, wt(side))
        b.set(f'{{{W}}}val', 'nil')
    # No cell margins (so image fills cell exactly)
    tblCellMar = etree.SubElement(tblPr, wt('tblCellMar'))
    for side in ('top', 'left', 'bottom', 'right'):
        cm = etree.SubElement(tblCellMar, wt(side))
        cm.set(f'{{{W}}}w', '0')
        cm.set(f'{{{W}}}type', 'dxa')

    grid = etree.SubElement(tbl, wt('tblGrid'))
    for w_dxa in grid_widths_dxa:
        gc = etree.SubElement(grid, wt('gridCol'))
        gc.set(f'{{{W}}}w', str(w_dxa))

    tr = etree.SubElement(tbl, wt('tr'))
    # Row height = max image height (exact). Bez tego cell paragraph z line=240
    # dodaje ~0.5 cm wysokości i strip pcha się na 2 stronę.
    max_h_cm = max(h for _, _, h in image_specs)
    row_h_dxa = int(max_h_cm * DXA_PER_CM)
    trPr = etree.SubElement(tr, wt('trPr'))
    trHeight = etree.SubElement(trPr, wt('trHeight'))
    trHeight.set(f'{{{W}}}val', str(row_h_dxa))
    trHeight.set(f'{{{W}}}hRule', 'atLeast')

    # Iterate cells in same order as grid, alternating image/spacer
    img_idx = 0
    for i, w_dxa in enumerate(grid_widths_dxa):
        tc = etree.SubElement(tr, wt('tc'))
        tcPr = etree.SubElement(tc, wt('tcPr'))
        tcW = etree.SubElement(tcPr, wt('tcW'))
        tcW.set(f'{{{W}}}w', str(w_dxa))
        tcW.set(f'{{{W}}}type', 'dxa')
        vAlign = etree.SubElement(tcPr, wt('vAlign'))
        vAlign.set(f'{{{W}}}val', 'center')

        p = etree.SubElement(tc, wt('p'))
        pPr = etree.SubElement(p, wt('pPr'))
        sp = etree.SubElement(pPr, wt('spacing'))
        sp.set(f'{{{W}}}before', '0')
        sp.set(f'{{{W}}}after', '0')
        # lineRule=auto — paragraph line dopasowuje się do wysokości obrazu (nie clipuje).
        # Wcześniej line=40 exact przycinało image od góry bo image jest większy niż line.
        sp.set(f'{{{W}}}line', '240')
        sp.set(f'{{{W}}}lineRule', 'auto')

        is_image_cell = (i % 2 == 0)
        if is_image_cell:
            rid, w_cm, h_cm = image_specs[img_idx]
            img_idx += 1
            # Center image horizontally (cell width == image width, but explicit center for safety)
            jc = etree.SubElement(pPr, wt('jc'))
            jc.set(f'{{{W}}}val', 'left')  # cell w == image w, so any align is fine
            r = etree.SubElement(p, wt('r'))
            cx_emu = int(w_cm * 360000)
            cy_emu = int(h_cm * 360000)
            r.append(_make_inline_image_drawing(rid, cx_emu, cy_emu))
    return tbl


# ─── QR code ──────────────────────────────────────────────────────────────────

def make_qr_bytes(url):
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


# ─── Anchored image (wzorzec z Oławy) ─────────────────────────────────────────

_anchor_uid = [1000]
def _next_uid():
    _anchor_uid[0] += 1
    return _anchor_uid[0]


def _make_anchored_image_drawing(rel_id, cx_emu, cy_emu, posY_emu, posX_emu=0,
                                  posX_relative_to='margin'):
    """Pływający obraz z layoutInCell=0 (anchor-based)."""
    uid = _next_uid()
    rel_h_val = posX_relative_to if posX_relative_to in ('margin','page','column') else 'margin'
    return etree.fromstring(f'''<w:drawing xmlns:w="{W}"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
  <wp:anchor distT="0" distB="0" distL="0" distR="0"
             simplePos="0" relativeHeight="{251659264 + uid}"
             behindDoc="0" locked="0" layoutInCell="0" allowOverlap="1"
             wp14:anchorId="0000{uid:04X}" wp14:editId="0000{uid:04X}">
    <wp:simplePos x="0" y="0"/>
    <wp:positionH relativeFrom="{rel_h_val}">
      <wp:posOffset>{int(posX_emu)}</wp:posOffset>
    </wp:positionH>
    <wp:positionV relativeFrom="paragraph">
      <wp:posOffset>{int(posY_emu)}</wp:posOffset>
    </wp:positionV>
    <wp:extent cx="{int(cx_emu)}" cy="{int(cy_emu)}"/>
    <wp:effectExtent l="0" t="0" r="0" b="0"/>
    <wp:wrapNone/>
    <wp:docPr id="{uid}" name="Picture {uid}"/>
    <wp:cNvGraphicFramePr>
      <a:graphicFrameLocks noChangeAspect="1"/>
    </wp:cNvGraphicFramePr>
    <a:graphic>
      <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
        <pic:pic>
          <pic:nvPicPr>
            <pic:cNvPr id="{uid}" name=""/>
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
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          </pic:spPr>
        </pic:pic>
      </a:graphicData>
    </a:graphic>
  </wp:anchor>
</w:drawing>''')


def _populate_left_area(elements, anchored_drawings, with_label, label_y_cm=0):
    """
    Wstawia pływające obrazy w komórce + napis "Wyniki turnieju".
    
    Komórka ma swój paragraf-kotwicę dla pływających obrazów.
    Napis "Wyniki turnieju" jest renderowany inline z paddingiem górnym
    który ustawia go w odpowiedniej pozycji Y (np. pod QR).
    """
    tbls = [el for el in elements if el.tag == wt('tbl')]
    if len(tbls) < 2: return
    score_tbl = tbls[1]
    rows = score_tbl.findall(wt('tr'))
    if len(rows) < 2: return
    target_cell = rows[1].findall(wt('tc'))[0]

    # Wyczyść istniejące paragrafy
    for p in target_cell.findall(wt('p')):
        target_cell.remove(p)

    # Pierwszy paragraf - kotwiczy pływające obrazy.
    # Wysokość minimalna (line=40 exact) żeby paragraf nie zajmował dużo miejsca
    p1 = etree.SubElement(target_cell, wt('p'))
    pPr1 = etree.SubElement(p1, wt('pPr'))
    sp1 = etree.SubElement(pPr1, wt('spacing'))
    sp1.set(f'{{{W}}}before', '0')
    sp1.set(f'{{{W}}}after', '0')
    sp1.set(f'{{{W}}}line', '40')
    sp1.set(f'{{{W}}}lineRule', 'exact')
    if anchored_drawings:
        r = etree.SubElement(p1, wt('r'))
        for drawing in anchored_drawings:
            r.append(drawing)

    # Drugi paragraf - napis "Wyniki turnieju".
    # Padding górny przesuwa go do pozycji label_y_cm
    if with_label:
        p2 = etree.SubElement(target_cell, wt('p'))
        pPr2 = etree.SubElement(p2, wt('pPr'))
        jc = etree.SubElement(pPr2, wt('jc'))
        jc.set(f'{{{W}}}val', 'center')
        sp2 = etree.SubElement(pPr2, wt('spacing'))
        # Padding górny w twipach (1 cm = 567 twip)
        before_twip = max(0, int(label_y_cm * 567))
        sp2.set(f'{{{W}}}before', str(before_twip))
        sp2.set(f'{{{W}}}after', '0')
        r2 = etree.SubElement(p2, wt('r'))
        rPr = etree.SubElement(r2, wt('rPr'))
        fonts = etree.SubElement(rPr, wt('rFonts'))
        for a in ('ascii','hAnsi','eastAsia','cs'):
            fonts.set(f'{{{W}}}{a}', 'Aptos')
        for tag in ('b','bCs'):
            etree.SubElement(rPr, wt(tag)).set(f'{{{W}}}val', '1')
        for tag in ('sz','szCs'):
            etree.SubElement(rPr, wt(tag)).set(f'{{{W}}}val', '20')
        t = etree.SubElement(r2, wt('t'))
        t.text = with_label


# ─── Build document ───────────────────────────────────────────────────────────

def build_blank_document(num_pages=1, logos=None, tournament_name=None,
                         tournament_date=None, tournament_phase_text=None,
                         include_qr=False,
                         include_pfm_logo=True, sheets_url='',
                         image_order=None, image_positions=None,
                         template_type='IND'):
    """
    Generuje pusty formularz protokołu (bez wypełnionych danych z arkusza).
    Tworzy `num_pages` identycznych pustych protokołów.
    """
    # Tworzymy listę pustych meczów
    blank_match = {'tor':'','godz':'','grupa':'','mecz':'','z1':'','z2':''}
    blank_data = [('', [blank_match for _ in range(num_pages)])]
    return build_document(
        sheet_id='', sheets_url=sheets_url, sheets_data=blank_data,
        logos=logos,
        tournament_name=tournament_name, tournament_date=tournament_date,
        tournament_phase_text=tournament_phase_text,
        include_qr=include_qr and bool(sheets_url),  # QR tylko jeśli jest URL
        include_pfm_logo=include_pfm_logo,
        image_order=image_order, image_positions=image_positions,
        template_type=template_type
    )


def build_document(sheet_id, sheets_url, sheets_data, logos=None,
                   tournament_name=None, tournament_date=None,
                   tournament_phase_text=None,
                   include_qr=True, include_pfm_logo=True,
                   image_order=None, image_positions=None,
                   hide_grupa_mecz=False, phase_label=None,
                   template_type='IND', progress_cb=None,
                   skip_placeholders=True):
    """
    `tournament_date`: string (np. "10.05.2026") wyświetlany w nagłówku obok nazwy.
    `tournament_phase_text`: tekst fazy turnieju (np. "Grupowa", "1/16 finału") -
                             dodawany w prawym górnym rogu obok nazwy i daty.
    `include_pfm_logo`: czy dodać domyślne logo PFM.
    image_positions: dict {key: (x_cm, y_cm, width_cm)} dla każdej grafiki/QR.
    Jeśli None, używamy domyślnego ułożenia jedna pod drugą.
    hide_grupa_mecz: True dla fazy pucharowej - ukrywa pola "Grupa" i "Mecz #",
                    zamiast tego pokazuje fazę (np. "1/32 finału").
    phase_label: nazwa fazy do wyświetlenia gdy hide_grupa_mecz=True.
    template_type: 'IND' (indywidualny) | 'TROJKA' (3-osobowy) | 'CZWORKA' (4-osobowy)
                   | 'DRUZYNA' (2-osobowy). Wybiera odpowiedni szablon docx.
    """
    import os
    _anchor_uid[0] = 1000

    # Filtruj placeholderowe mecze („Gracz N" w grupowej, „bye" w drabince) dla
    # IND. Tylko indywidualny — drużynówki nie mają tego problemu (zespół ma
    # zwykle prawdziwą nazwę).
    if skip_placeholders and template_type in ('IND', 'IND_Bo3', 'IND_Bo5', 'IND_Bo7'):
        filtered = []
        for entry in sheets_data:
            if len(entry) == 3:
                gname, matches, phase_override = entry
            else:
                gname, matches = entry
                phase_override = None
            kept = [m for m in matches
                    if not (_is_placeholder_name(m.get('z1'))
                            or _is_placeholder_name(m.get('z2')))]
            if not kept:
                continue
            filtered.append((gname, kept, phase_override) if phase_override is not None else (gname, kept))
        sheets_data = filtered

    # Nazwy plików: konwencja PFM SharePoint z UNDERSCOREM (IND_Grupa.docx itp.)
    template_files = {
        'IND': 'IND_Grupa.docx',
        'IND_Bo3': 'IND_Bo3.docx',           # pucharowa indywidualna Best of 3
        'IND_Bo5': 'IND_Bo5.docx',           # pucharowa indywidualna Best of 5
        'IND_Bo7': 'IND_Bo7.docx',           # pucharowa indywidualna Best of 7 (finały)
        'DWOJKA': 'DWÓJKA_Grupa.docx',       # 2-osobowa drużyna grupowa (własny layout: DRUŻYNY pion, 4 SUMA/SET)
        'DWOJKA_Bo3': 'DWÓJKA_Bo3.docx',     # 2-osobowa pucharowa Best of 3
        'DWOJKA_Bo5': 'DWÓJKA_Bo5.docx',     # 2-osobowa pucharowa Best of 5
        'DWOJKA_Bo7': 'DWÓJKA_Bo7.docx',     # 2-osobowa pucharowa Best of 7 (landscape, jak CZWÓRKA)
        'TROJKA': 'TRÓJKA_Grupa.docx',
        'TROJKA_Bo3': 'TRÓJKA_Bo3.docx',
        'TROJKA_Bo5': 'TRÓJKA_Bo5.docx',
        'CZWORKA': 'CZWÓRKA_Grupa.docx',     # 4-osobowa drużyna grupowa
        'CZWORKA_Bo3': 'CZWÓRKA_Bo3.docx',   # 4-osobowa pucharowa Best of 3
        'CZWORKA_Bo5': 'CZWÓRKA_Bo5.docx',   # 4-osobowa pucharowa Best of 5
    }
    tpl_filename = template_files.get(template_type, 'IND_Grupa.docx')
    tpl_path = os.path.join(os.path.dirname(__file__), tpl_filename)
    with open(tpl_path, 'rb') as f:
        tpl_bytes = f.read()

    zin = zipfile.ZipFile(io.BytesIO(tpl_bytes))
    doc_xml = zin.read('word/document.xml')
    doc_root = etree.fromstring(doc_xml)
    body = doc_root.find(wt('body'))

    # ── Marginesy
    # Indywidualny: 720 dxa = 1.27 cm wszędzie.
    # Trójka: 720 dxa (1.27 cm) + zsynchronizowanie tabeli 1 z tabelą 2 (oba 9251 dxa
    # od lewego brzegu marginesu) — bez tego tabela 1 'auto' rozciąga się na całą
    # użyteczną szerokość a tabela 2 zostaje fixed 9251 → rozjazd.
    # CZWORKA: top=360 (0.63 cm) — zmniejszone z 540, by oszczędzić miejsce.
    # bottom=540 (0.95 cm) — większy niż w oryginale (180 dxa = 0.32 cm) bo wiele drukarek
    # ma hardware margin ~0.5-1 cm; przy 0.32 cm dolna część grafik była obcinana.
    # Łącznie zysk względem oryginału (top 540 + bottom 180 = 720 dxa) jest niewielki
    # (top 360 + bottom 540 = 900 dxa = +180 dxa), ALE strip jest przeniesiony wyżej
    # (Y_GRAPHICS=18.5 zamiast 20.5) więc faktycznie zostaje go więcej miejsca do końca strony.
    sectPr_check = body.find(wt('sectPr'))
    if sectPr_check is not None:
        pgMar = sectPr_check.find(wt('pgMar'))
        if pgMar is not None:
            if template_type == 'CZWORKA':
                pgMar.set(f'{{{W}}}top', '360')
                pgMar.set(f'{{{W}}}bottom', '280')
                pgMar.set(f'{{{W}}}left', '720')
                pgMar.set(f'{{{W}}}right', '720')
            elif template_type in ('CZWORKA_Bo3', 'CZWORKA_Bo5', 'DWOJKA_Bo7'):
                # Pucharowa czwórka + DWÓJKA Bo7 — szablony landscape z własnymi
                # marginesami (NIE wymuszamy 720, zepsułoby szerokie tabele 27 cm).
                # NIE wymuszamy 720 (zepsułoby szerokie tabele 27 cm w landscape).
                # ALE: bottom margin szablonu jest ~26 dxa (zero), więc dodany przez
                # build_document paragraf nagłówka turnieju (~280 dxa) pcha ostatnie
                # wiersze tabeli wyników na nową stronę. Kompensujemy zmniejszając
                # top margin o wysokość paragrafu nagłówka — content zostaje na swoim
                # miejscu względem strony, a header sadowi się w zwolnionej przestrzeni.
                has_header = bool(
                    tournament_name or tournament_date or tournament_phase_text or
                    any(len(g) == 3 and g[2] for g in sheets_data)
                )
                if has_header:
                    pgMar.set(f'{{{W}}}top', '260')
            else:
                for side in ('top','bottom','left','right'):
                    pgMar.set(f'{{{W}}}{side}', '720')

    # ── Fonty etykiet: pomniejsz zbyt duże (24→20 dla głównych etykiet,
    # zachowaj 24 dla nagłówków SET 1/SET 2/Wyniki turnieju w tabeli wyników).
    # Tor/Godzina/Grupa/Mecz# (sz=24) → sz=22 (czytelne, mieszczą się w linii)
    # Punkty SET 1/SET 2/Wygrane sety/Podpis (sz=24) → sz=20
    LABELS_BIGGER = {'Tor','Godzina','Grupa','Mecz','#'}
    LABELS_HEADER = {'Punkty','SET 1','SET 2','SET 3','Wygrane','sety','Podpis'}
    # ── Operacje SPECYFICZNE DLA TRÓJKOWEGO szablonu:
    # Wymuszenie Calibri jako fontu etykiet (Tor/Godzina/Grupa/Mecz#/Punkty SET 1/2/
    # Wygrane sety/Podpis). Bez tego LibreOffice (i Word bez Aptos) używa fallback
    # który jest znacznie szerszy i wszystko rozjeżdża się na 2 wiersze.
    # Zachowujemy oryginalne size (24) - Calibri w tym rozmiarze mieści się normalnie.
    if template_type in ('TROJKA', 'TROJKA_Bo3', 'TROJKA_Bo5', 'DWOJKA', 'DWOJKA_Bo3', 'DWOJKA_Bo5'):
        # Bo2 i Bo3/Bo5 mają wspólny zestaw etykiet, ale Bo3/Bo5 NIE potrzebują
        # normalizacji fontu dla 'SET 1/2/3/4/5' / '(SET N)' bo w nowym szablonie
        # te etykiety są inaczej zbudowane (różne run-e) i normalizacja powoduje
        # niespójność (np. SET 2 robi się grubsze niż SET 1/(SET 3)).
        if template_type in ('TROJKA_Bo3', 'TROJKA_Bo5', 'DWOJKA_Bo3', 'DWOJKA_Bo5'):
            TROJKA_LABELS = {'Tor','Godz.','Godzina','Grupa','Mecz','#','Runda',
                             'PunktySET 1','PunktySET 2','PunktySET 3',
                             'PunktySET 4','PunktySET 5',
                             'Punkty','SET 1','SET 2','SET 3','SET 4','SET 5','(SET 3)','SET','3',
                             'PktSET 1','PktSET 2','PktSET 3','PktSET 4','PktSET 5',
                             'Pkt','Wygrane','sety','Podpis','Wygranesety','DRUŻYNY'}
        else:
            TROJKA_LABELS = {'Tor','Godzina','Godz.','Grupa','Mecz','#','Runda',
                             'Punkty','SET 1','SET 2','SET 3','(SET 3)',
                             'Wygrane','sety','Podpis',
                             'PunktySET 1','PunktySET 2','PunktySET 3','Wygranesety',
                             'Wyniki turnieju','Wyniki turnieju:'}
        for r in body.iter(wt('r')):
            ts = r.findall(wt('t'))
            if not ts: continue
            text_content = ''.join((t.text or '') for t in ts).strip()
            # Match po exact label LUB prefix "Mecz #" (po fill _set_cell_label
            # zmienia "Runda" → "Mecz #  1" w jednym runie — single-text nie pasuje
            # do żadnego stringa z TROJKA_LABELS → bez tego LO bierze Aptos Narrow
            # fallback → szeryf na "Mecz # 1" w prawym górnym).
            if text_content in TROJKA_LABELS or text_content.startswith('Mecz #'):
                rPr = r.find(wt('rPr'))
                if rPr is None:
                    rPr = etree.Element(wt('rPr'))
                    r.insert(0, rPr)
                fonts = rPr.find(wt('rFonts'))
                if fonts is None:
                    fonts = etree.SubElement(rPr, wt('rFonts'))
                for a in ('ascii','hAnsi','eastAsia','cs'):
                    fonts.set(f'{{{W}}}{a}', 'Calibri')

        # ── DWÓJKA Bo3/Bo5: nagłówek r0 (Tor / Godz. / Mecz #) — wymuś Calibri +
        # jawny rozmiar na KAŻDYM runie. Powód: „Godz." jest rozbite na 2 runy
        # „Godz"(sz22) + „."(BEZ sz → dziedziczy duży docDefault ~24pt). Match
        # per-run w TROJKA_LABELS nie łapie ani „Godz" (nie ma w secie) ani „."
        # → kropka renderowana ogromna w Aptos, „Godz." zawija na 2 linie
        # („Godz"/„."), a wartość 13:30 w osobnej komórce ląduje wyżej. Cell-level
        # force (jak w CZWÓRCE) naprawia rozbite labele niezależnie od podziału runów.
        # Rozszerzone na TRÓJKĘ Bo3/Bo5: nagłówek str.2 ma węższą komórkę „Mecz #"
        # (1093 vs 1314 na str.1) → „Mecz #  1" przy docDefault ~24 zawijało na 2
        # linie („Mecz #"/„1" — user: „Mecz # rozjechane"). Wymuszamy JEDNOLITY
        # sz=22 na WSZYSTKICH runach r0 (Tor/Godz./Mecz# + wartości) — mieści się
        # na obu stronach i daje spójny rozmiar etykiet (user: „powinno być
        # jednolicie w każdym wariancie wszędzie").
        if template_type in ('DWOJKA_Bo3', 'DWOJKA_Bo5', 'TROJKA_Bo3', 'TROJKA_Bo5'):
            for tbl in body.findall(wt('tbl')):
                first_row = tbl.find(wt('tr'))
                if first_row is None:
                    continue
                rtxt = ''.join((t.text or '') for t in first_row.iter(wt('t')))
                if 'Tor' not in rtxt or 'Mecz' not in rtxt:
                    continue
                for tc in first_row.findall(wt('tc')):
                    for run in tc.iter(wt('r')):
                        if not run.findall(wt('t')):
                            continue
                        rPr = run.find(wt('rPr'))
                        if rPr is None:
                            rPr = etree.Element(wt('rPr')); run.insert(0, rPr)
                        fonts = rPr.find(wt('rFonts'))
                        if fonts is None:
                            fonts = etree.SubElement(rPr, wt('rFonts'))
                        for a in ('ascii', 'hAnsi', 'eastAsia', 'cs'):
                            fonts.set(f'{{{W}}}{a}', 'Calibri')
                        # Jednolity sz=22 (11pt) — FORSUJEMY na każdym runie (także
                        # tych z własnym sz, np. godz=24), żeby Tor/Godz/Mecz miały
                        # ten sam rozmiar i Mecz# nie zawijał w wąskiej komórce str.2.
                        for tag in ('sz', 'szCs'):
                            el = rPr.find(wt(tag))
                            if el is None:
                                el = etree.SubElement(rPr, wt(tag))
                            el.set(f'{{{W}}}val', '22')

        # ── Bo3/Bo5: rozmiar etykiet nagłówka kolumn (Pkt/Punkty SET, Wygrane sety,
        # Podpis). TRÓJKA_Bo3 ma te runy BEZ jawnego sz (dziedziczą duży docDefault)
        # → „Wygrane" (W+g+y szerokie) nie mieści się w wąskiej komórce 1174 dxa
        # i łamie się brzydko na „Wygran"/„e". TRÓJKA_Bo5 ma już sz=20 (renderuje OK),
        # więc wymuszenie 20 jest tam idempotentne. Skala 20 = 10pt, mieści „Wygrane"
        # nawet w 1039 dxa (Bo5). Scope: TYLKO tabele nagłówkowe (≤6 wierszy) — żeby
        # NIE ruszyć wielkich etykiet „SET 1" (sz=28) w tabeli wynikowej.
        if template_type in ('TROJKA_Bo3', 'TROJKA_Bo5', 'DWOJKA_Bo3', 'DWOJKA_Bo5'):
            for tbl in body.findall(wt('tbl')):
                if len(tbl.findall(wt('tr'))) > 6:
                    continue
                for tc in tbl.iter(wt('tc')):
                    ctext = ''.join((t.text or '') for t in tc.iter(wt('t')))
                    cstrip = ctext.replace(' ', '')
                    is_pktset = (('Pkt' in ctext or 'Punkt' in ctext) and 'SET' in ctext)
                    is_wygr = ('Wygr' in ctext and 'set' in ctext.lower())
                    is_podpis = (cstrip == 'Podpis')
                    if not (is_pktset or is_wygr or is_podpis):
                        continue
                    # Adaptacyjny rozmiar: „Wygrane" (7 znaków single-word, nie
                    # wieloliniowe jak „Pkt"+„SET N") w wąskiej komórce zawija na
                    # 3 linie („Wygra/ne/sety"). W TROJKA_Bo5 redystrybucja 300 dxa
                    # zostawia Wygr ~850 dxa — za mało dla sz=20. Drop do sz=16
                    # poniżej 900 dxa (mieści „Wygrane" w jednej linii).
                    # „Wygr. sety" (skrócone, z kropką) mieści się przy sz=20
                    # w wąskiej komórce 850 dxa — tak samo jak kolumny Pkt SET.
                    # (Poprzednio szablon miał „Wygrane sety" 7-znakowe, wymagało
                    # sz=16; user zmienił label na „Wygr. sety" na obu stronach.)
                    sz_val = '20'
                    for run in tc.iter(wt('r')):
                        if not run.findall(wt('t')):
                            continue
                        if not ''.join((t.text or '') for t in run.findall(wt('t'))).strip():
                            continue
                        rPr = run.find(wt('rPr'))
                        if rPr is None:
                            rPr = etree.Element(wt('rPr')); run.insert(0, rPr)
                        fonts = rPr.find(wt('rFonts'))
                        if fonts is None:
                            fonts = etree.SubElement(rPr, wt('rFonts'))
                        for a in ('ascii', 'hAnsi', 'eastAsia', 'cs'):
                            fonts.set(f'{{{W}}}{a}', 'Calibri')
                        for tag in ('sz', 'szCs'):
                            el = rPr.find(wt(tag))
                            if el is None:
                                el = etree.SubElement(rPr, wt(tag))
                            el.set(f'{{{W}}}val', sz_val)
                        for btag in ('b', 'bCs'):
                            if rPr.find(wt(btag)) is None:
                                etree.SubElement(rPr, wt(btag))

        # Zmniejszenie odstępów po tabeli wynikowej żeby mieściło się na 1 stronie.
        # Ostatni paragraf przed sectPr ma duży before - usuwamy.
        # POMIJAMY DLA Bo5 — Bo5 jest 2-stronicowe, potrzebuje zachowania
        # naturalnych odstępów dla page-break między stroną 1 (SET 1-3)
        # a stroną 2 (SET 4-5).
        if template_type != 'TROJKA_Bo5':
            for p in body.iter(wt('p')):
                pPr = p.find(wt('pPr'))
                if pPr is not None:
                    spacing = pPr.find(wt('spacing'))
                    if spacing is not None:
                        # Zmniejsz before/after do 0
                        for attr in ('before','after','beforeLines','afterLines'):
                            a = f'{{{W}}}{attr}'
                            if a in spacing.attrib:
                                spacing.set(a, '0')
        
        # ── Rozszerzenie obu tabel do PEŁNEJ szerokości użytecznej strony ──
        # Marginesy 720 dxa po obu stronach → useable area = 11906 - 1440 = 10466 dxa = 18.46 cm.
        # 
        # Dla tabeli 2 (wynikowej): zwiększenie idzie GŁÓWNIE do lewej kolumny (R1.tc[0])
        # z oryginalnej szerokości (różne dla TROJKA vs Bo3) → 2700 dxa = 4.76 cm.
        # Daje to obszar na grafiki (QR + loga). Pozostałe kolumny dostają to co zostanie.
        TARGET_WIDTH = 10466
        
        # Parametry zależne od szablonu
        if template_type == 'TROJKA_Bo3':
            # TROJKA_Bo3.docx (nowa wersja): tabela 1 = 9239 dxa (6 kol),
            # tabela 2 = 9439 dxa (25 kol, col0=391). Struktura trójkowa
            # (wiersze 1-18 + WYNIK), 3 sety, SUMA per zawodnik per set.
            ORIG_T1_TOTAL = 9239
            ORIG_T2_TOTAL = 9439
            ORIG_LEFT_COL_DXA = 391
        elif template_type == 'TROJKA_Bo5':
            # TROJKA_Bo5.docx (aktualna wersja): T0=9450, T1=9439, T2=9450, T3=6425.
            ORIG_T1_TOTAL = 9450
            ORIG_T2_TOTAL = 9439
            ORIG_T3_TOTAL = 9450
            ORIG_T4_TOTAL = 6425
            ORIG_LEFT_COL_DXA = 393  # T2.col0
        else:
            # TROJKA_Grupa.docx: tabela 1 = 9026 dxa, tabela 2 = 9251 dxa, col0=1186
            ORIG_T1_TOTAL = 9026
            ORIG_T2_TOTAL = 9251
            ORIG_LEFT_COL_DXA = 1186
        
        tbls = body.findall(wt('tbl'))
        if len(tbls) >= 2:
            t1, t2 = tbls[0], tbls[1]

            # ── Wymiary DYNAMICZNE dla wariantów grupowych (TROJKA/DWOJKA) ──
            # Szablony grupowe różnią się szerokościami (DWÓJKA Grupa ma inny
            # layout niż TRÓJKA Grupa). Hardcode 9026/9251/1186 pasował tylko
            # do TRÓJKI → DWÓJKA dostawała ZŁY ORIG_LEFT_COL_DXA, przez co
            # warunek dopasowania komórki lewej (poniżej) nie trafiał i tcW
            # col0 NIE było poszerzane do 2700, mimo że gridCol BYŁO. Efekt:
            # PDF (LibreOffice czyta gridCol) OK, ale DOCX (Word czyta tcW per
            # komórka) miał wąską kolumnę → grafiki nachodziły na tabelę.
            # Czytamy realne wymiary z szablonu — działa dla obu.
            if template_type in ('TROJKA', 'DWOJKA', 'DWOJKA_Bo3', 'DWOJKA_Bo5'):
                g1 = t1.find(wt('tblGrid'))
                if g1 is not None:
                    _w1 = [int(gc.get(f'{{{W}}}w')) for gc in g1.findall(wt('gridCol')) if gc.get(f'{{{W}}}w')]
                    if _w1:
                        ORIG_T1_TOTAL = sum(_w1)
                g2 = t2.find(wt('tblGrid'))
                if g2 is not None:
                    _w2 = [int(gc.get(f'{{{W}}}w')) for gc in g2.findall(wt('gridCol')) if gc.get(f'{{{W}}}w')]
                    if _w2:
                        ORIG_T2_TOTAL = sum(_w2)
                        ORIG_LEFT_COL_DXA = _w2[0]

            # Tabela 1: skaluj wszystkie kolumny proporcjonalnie
            scale_t1 = TARGET_WIDTH / ORIG_T1_TOTAL
            tblPr_t1 = t1.find(wt('tblPr'))
            if tblPr_t1 is not None:
                tblW = tblPr_t1.find(wt('tblW'))
                if tblW is None:
                    tblW = etree.SubElement(tblPr_t1, wt('tblW'))
                tblW.set(f'{{{W}}}w', str(TARGET_WIDTH))
                tblW.set(f'{{{W}}}type', 'dxa')
                tblInd = tblPr_t1.find(wt('tblInd'))
                if tblInd is None:
                    tblInd = etree.SubElement(tblPr_t1, wt('tblInd'))
                tblInd.set(f'{{{W}}}w', '0')
                tblInd.set(f'{{{W}}}type', 'dxa')
            tblGrid_t1 = t1.find(wt('tblGrid'))
            if tblGrid_t1 is not None:
                for gc in tblGrid_t1.findall(wt('gridCol')):
                    w = gc.get(f'{{{W}}}w')
                    if w:
                        gc.set(f'{{{W}}}w', str(int(int(w) * scale_t1)))
            for tr in t1.findall(wt('tr')):
                for tc in tr.findall(wt('tc')):
                    tcPr = tc.find(wt('tcPr'))
                    if tcPr is not None:
                        tcW = tcPr.find(wt('tcW'))
                        if tcW is not None:
                            w = tcW.get(f'{{{W}}}w')
                            if w:
                                tcW.set(f'{{{W}}}w', str(int(int(w) * scale_t1)))
            
            # ── Bo3/Bo5: redystrybucja szerokości kolumn w T1 ──
            # Dłuższe nazwy drużyn (np. "SIERMOLKKY-LESZNO TEAM" ~21 znaków) nie
            # mieszczą się w domyślnej szerokości Tor (5.94 cm po skalowaniu).
            # Przesuwamy 800 dxa = 1.41 cm do Tor: 300 z "Wygrane sety" (kol -2)
            # + 500 z "Podpis" (kol -1). Dzięki temu Podpis pozostaje szerszy
            # od Wygrane sety (po skalowaniu Wygrane sety ~1.93 cm, Podpis ~3.13 cm
            # → po transferze: Wygrane sety ~1.40 cm, Podpis ~2.25 cm).
            if template_type in ('TROJKA_Bo3', 'TROJKA_Bo5'):
                # Bo5: NIE kradniemy z Wygr — komórka „Wygr. sety" przy sz=20
                # potrzebuje pełnej szerokości (~1150 dxa), inaczej zawija na
                # 3 linie („Wyg/r./sety"). Cały transfer (800) bierzemy z Podpis
                # (szeroki ~1890 dxa → zostaje ~1090, „Podpis" mieści się).
                # Bo3: bez zmian (szablon ma inne proporcje, user: działa OK).
                if template_type == 'TROJKA_Bo5':
                    TRANSFER_FROM_PODPIS = 800
                    TRANSFER_FROM_WYGRANE = 0
                else:
                    TRANSFER_FROM_PODPIS = 500
                    TRANSFER_FROM_WYGRANE = 300
                TOTAL_TRANSFER = TRANSFER_FROM_PODPIS + TRANSFER_FROM_WYGRANE
                tblGrid_t1 = t1.find(wt('tblGrid'))
                if tblGrid_t1 is not None:
                    gcols = tblGrid_t1.findall(wt('gridCol'))
                    if len(gcols) >= 3:
                        first = gcols[0]
                        wygrane = gcols[-2]
                        podpis = gcols[-1]
                        first.set(f'{{{W}}}w', str(int(first.get(f'{{{W}}}w','0')) + TOTAL_TRANSFER))
                        wygrane.set(f'{{{W}}}w', str(int(wygrane.get(f'{{{W}}}w','0')) - TRANSFER_FROM_WYGRANE))
                        podpis.set(f'{{{W}}}w', str(int(podpis.get(f'{{{W}}}w','0')) - TRANSFER_FROM_PODPIS))
                # Dla każdego wiersza: dopasuj komórki tc[0], tc[-2], tc[-1].
                for tr in t1.findall(wt('tr')):
                    cells = tr.findall(wt('tc'))
                    if len(cells) >= 3:
                        deltas = [
                            (cells[0],  TOTAL_TRANSFER),
                            (cells[-2], -TRANSFER_FROM_WYGRANE),
                            (cells[-1], -TRANSFER_FROM_PODPIS),
                        ]
                        for tc, delta in deltas:
                            tcPr = tc.find(wt('tcPr'))
                            if tcPr is None: continue
                            tcW = tcPr.find(wt('tcW'))
                            if tcW is None: continue
                            cw = tcW.get(f'{{{W}}}w')
                            if cw:
                                tcW.set(f'{{{W}}}w', str(int(cw) + delta))
            
            # Tabela 2: ZALEŻNIE OD WARIANTU
            # - TROJKA (Bo2): lewa kolumna powiększona do 2700 dxa (na grafiki)
            # - TROJKA_Bo3/Bo5: jednorodne skalowanie (puchar trójki nie używa
            #   grafik — col0 zostaje wąska, nie poszerzamy)
            if template_type in ('TROJKA_Bo3', 'TROJKA_Bo5', 'DWOJKA_Bo3', 'DWOJKA_Bo5'):
                # Puchar nie używa grafik (is_no_graphics=True) — col0 zostaje
                # wąska, cała szerokość idzie na tabelę wyników (inaczej 2-cyfrowe
                # numery wierszy 10-18 zawijały się w za wąskiej kolumnie).
                NEW_LEFT_COL_DXA = ORIG_LEFT_COL_DXA
            else:
                NEW_LEFT_COL_DXA = 2700  # 4.76 cm — większy obszar na grafiki
            NEW_REST_TOTAL = TARGET_WIDTH - NEW_LEFT_COL_DXA
            ORIG_REST_TOTAL = ORIG_T2_TOTAL - ORIG_LEFT_COL_DXA
            scale_rest = NEW_REST_TOTAL / ORIG_REST_TOTAL
            
            tblPr_t2 = t2.find(wt('tblPr'))
            if tblPr_t2 is not None:
                tblW = tblPr_t2.find(wt('tblW'))
                if tblW is None:
                    tblW = etree.SubElement(tblPr_t2, wt('tblW'))
                tblW.set(f'{{{W}}}w', str(TARGET_WIDTH))
                tblW.set(f'{{{W}}}type', 'dxa')
                tblInd = tblPr_t2.find(wt('tblInd'))
                if tblInd is None:
                    tblInd = etree.SubElement(tblPr_t2, wt('tblInd'))
                tblInd.set(f'{{{W}}}w', '0')
                tblInd.set(f'{{{W}}}type', 'dxa')
            
            # tblGrid: col 0 → NEW_LEFT_COL_DXA, reszta × scale_rest
            tblGrid_t2 = t2.find(wt('tblGrid'))
            if tblGrid_t2 is not None:
                cols = tblGrid_t2.findall(wt('gridCol'))
                for i, gc in enumerate(cols):
                    w = gc.get(f'{{{W}}}w')
                    if w:
                        if i == 0:
                            gc.set(f'{{{W}}}w', str(NEW_LEFT_COL_DXA))
                        else:
                            gc.set(f'{{{W}}}w', str(int(int(w) * scale_rest)))
            
            # Każdy wiersz: tc[0] → NEW_LEFT_COL_DXA, reszta × scale_rest
            for tr in t2.findall(wt('tr')):
                cells = tr.findall(wt('tc'))
                for i, tc in enumerate(cells):
                    tcPr = tc.find(wt('tcPr'))
                    if tcPr is not None:
                        tcW = tcPr.find(wt('tcW'))
                        if tcW is not None:
                            w = tcW.get(f'{{{W}}}w')
                            if w:
                                w_int = int(w)
                                if i == 0 and abs(w_int - ORIG_LEFT_COL_DXA) <= 50:
                                    # To jest komórka lewa (R1.tc[0] etc) — daj nową szer.
                                    # TOLERANCJA 50 dxa: gridCol i tcW komórek bywają
                                    # o kilka dxa różne (DWÓJKA: grid 1191 vs tcW 1192).
                                    # Exact == gubił DWÓJKĘ → tcW col0 nie szło na 2700
                                    # → grafiki nachodziły na tabelę w Wordzie (PDF OK).
                                    tcW.set(f'{{{W}}}w', str(NEW_LEFT_COL_DXA))
                                else:
                                    tcW.set(f'{{{W}}}w', str(int(w_int * scale_rest)))
            
            # ── Normalizacja ramek kolumn SUMA w tabeli wyników ──────────
            # PROBLEM: w szablonach DWÓJKI (Grupa i Bo3) niektóre kolumny SUMA
            # miały niepasujące ramki vs inne: dodatkowe top/bottom (sz 4 lub 6)
            # w wierszach danych dawały widoczne poziome linie WEWNĄTRZ kolumny,
            # podczas gdy inne SUMA były „czyste" (tylko L/R). Skutek: ta jedna
            # kolumna wyglądała na pogrubioną/ze szczeblami, inne nie.
            # FIX: znajdź wszystkie SUMA-cells w R2 (label row); w każdym wierszu
            # DANYCH skopiuj borders z pierwszej SUMA-cell (c7 — pewnie czysty
            # styl: L=12, R=12, no T/B) do pozostałych SUMA-cells. Ostatnia SUMA
            # dostaje dodatkowo grube prawe border (krawędź tabeli).
            t2_rows = t2.findall(wt('tr'))
            from copy import deepcopy
            # Krok 1: zidentyfikuj indeksy SUMA-cells w r2 (label row)
            suma_indices = []
            r2_cells = t2_rows[2].findall(wt('tc')) if len(t2_rows) > 2 else []
            for ci, c in enumerate(r2_cells):
                txt = ''.join((t_el.text or '') for t_el in c.iter(wt('t'))).strip()
                if txt == 'SUMA':
                    suma_indices.append(ci)
            # Krok 2: w każdym wierszu danych skopiuj borders z REFERENCYJNEJ SUMA
            # (drugiej — bo pierwsza ma top=4/8 podpięte do nagłówka R2, ale i tak
            # gwarantuje to spójny wygląd; wybieramy drugą bo c4 czasem ma top=8
            # specjalnie pasujące do bottom=8 z r2 — kopiowanie tego do dalszych
            # SUMA jest bezpieczne wizualnie).
            if len(suma_indices) >= 2:
                ref_idx = suma_indices[1]  # druga SUMA (np. c7 w DWÓJCE)
                last_idx = suma_indices[-1]
                for r_idx, tr in enumerate(t2_rows):
                    if r_idx < 3 or r_idx == len(t2_rows) - 1:
                        continue
                    cells = tr.findall(wt('tc'))
                    if ref_idx >= len(cells): continue
                    ref_tc = cells[ref_idx]
                    ref_tcPr = ref_tc.find(wt('tcPr'))
                    if ref_tcPr is None: continue
                    ref_borders = ref_tcPr.find(wt('tcBorders'))
                    if ref_borders is None: continue
                    for tgt_idx in suma_indices:
                        if tgt_idx == ref_idx or tgt_idx >= len(cells):
                            continue
                        tgt_tc = cells[tgt_idx]
                        tgt_tcPr = tgt_tc.find(wt('tcPr'))
                        if tgt_tcPr is None: continue
                        # Usuń istniejące borders
                        old_b = tgt_tcPr.find(wt('tcBorders'))
                        if old_b is not None:
                            tgt_tcPr.remove(old_b)
                        new_borders = deepcopy(ref_borders)
                        # Ostatnia SUMA: prawa krawędź = krawędź tabeli, sz=12 —
                        # MUSI być gruba jak prawa krawędź WEWNĘTRZNYCH kolumn SUMA
                        # (right=12), inaczej ostatnia kolumna SUMA wygląda inaczej
                        # niż reszta (user: „ta krawędź miała być pogrubiona").
                        if tgt_idx == last_idx:
                            right_el = new_borders.find(wt('right'))
                            if right_el is None:
                                right_el = etree.SubElement(new_borders, wt('right'))
                            right_el.set(f'{{{W}}}val', 'single')
                            right_el.set(f'{{{W}}}sz', '12')
                            right_el.set(f'{{{W}}}space', '0')
                            right_el.set(f'{{{W}}}color', 'auto')
                        tgt_tcPr.insert(0, new_borders)

            # ── Jednolita PRAWA KRAWĘDŹ tabeli wynikowej ──────────────────────
            # PROBLEM (mierzone w px na renderze 300dpi): wiersze danych z JAWNYM
            # `tcBorders right=12` renderują się 7px (grubo), ale wiersz nagłówka
            # SET (r0, last cell gridSpan=6, BEZ tcBorders) polegał na `tblBorders
            # right=12` — a LibreOffice rysuje krawędź zewnętrzną z tblBorders
            # CIENKO (2px), ignorując sz. Stąd: prawa krawędź gruba w środku,
            # cienka u góry (SET header) — user: „przy sumie inna niż reszta".
            # FIX: jawny `tcBorders right sz=12` na OSTATNIEJ komórce KAŻDEGO
            # wiersza (r0..r21). Jawne per-cell borders LO renderuje poprawnie
            # (7px), więc cała krawędź jest jednolita od góry do dołu.
            for tr in t2_rows:
                cells = tr.findall(wt('tc'))
                if not cells: continue
                last_tc = cells[-1]
                tcPr = last_tc.find(wt('tcPr'))
                if tcPr is None:
                    tcPr = etree.Element(wt('tcPr'))
                    last_tc.insert(0, tcPr)
                tcBorders = tcPr.find(wt('tcBorders'))
                if tcBorders is None:
                    # tcBorders MUSI iść we właściwym miejscu schematu CT_TcPr —
                    # PO cnfStyle/tcW/gridSpan/hMerge/vMerge, PRZED shd/tcMar/
                    # vAlign. Wstawienie na pozycję 0 (przed tcW/gridSpan) sprawia
                    # że LibreOffice/Word IGNORUJĄ element → krawędź zostaje cienka
                    # (to był bug: r0 nagłówka SET renderował 2px mimo right=12).
                    tcBorders = etree.Element(wt('tcBorders'))
                    PRECEDING = {wt('cnfStyle'), wt('tcW'), wt('gridSpan'),
                                 wt('hMerge'), wt('vMerge')}
                    insert_at = 0
                    for k, child in enumerate(list(tcPr)):
                        if child.tag in PRECEDING:
                            insert_at = k + 1
                        else:
                            break
                    tcPr.insert(insert_at, tcBorders)
                # Grubość prawej krawędzi zależna od typu wiersza:
                # • wiersze SUMA-body (ostatnia komórka = kolumna SUMA, ma left=12)
                #   → right=12, żeby ostatnia kolumna SUMA miała grubą prawą krawędź
                #   JAK wewnętrzne kolumny SUMA (user: „ta krawędź miała być pogrubiona").
                # • nagłówek SET / sub-header / wiersz PKT (narożniki) → right=6,
                #   bo tam pogrubienie było zbędne (wcześniejsza uwaga usera o rogach).
                _lastleft = last_tc.find(f'{wt("tcPr")}/{wt("tcBorders")}/{wt("left")}')
                _is_suma_row = _lastleft is not None and _lastleft.get(f'{{{W}}}sz') == '12'
                right_el = tcBorders.find(wt('right'))
                if right_el is None:
                    right_el = etree.SubElement(tcBorders, wt('right'))
                right_el.set(f'{{{W}}}val', 'single')
                right_el.set(f'{{{W}}}sz', '12' if _is_suma_row else '6')
                right_el.set(f'{{{W}}}space', '0')
                right_el.set(f'{{{W}}}color', 'auto')

        # ── BO5: dodatkowo skalujemy tabele 3 i 4 (strona 3 — extension sheet) ──
        # T3 (header s.3) skalujemy proporcjonalnie do tej samej szerokości co T1.
        # T4 (results s.3 z (SET 4)/(SET 5)) celowo NIE skalujemy do max — niech
        # zachowa naturalną proporcję (jest węższa od T2 bo ma tylko 2 sety).
        if template_type == 'TROJKA_Bo5' and len(tbls) >= 4:
            t3, t4 = tbls[2], tbls[3]
            
            # ── EXPLICIT PAGE BREAK przed T3 ──
            # Bo5 musi rozłożyć się na 2 strony: strona 1 = SET 1-3,
            # strona 2 = SET 4-5 (extension). W oryginalnym docx page break
            # między T2 a T3 powstaje przez naturalny overflow (template ma
            # paragraphs z dużymi spacing). Przy klonowaniu i niezależnie od
            # spacing wstawiamy explicit `<w:br type="page"/>` w paragrafie
            # między T2 a T3, żeby T3 ZAWSZE szedł na nową stronę.
            t2_idx = list(body).index(t2)
            for j in range(t2_idx + 1, list(body).index(t3)):
                el_between = list(body)[j]
                if el_between.tag == wt('p'):
                    # Add page break run as the first child of this paragraph
                    run = etree.SubElement(el_between, wt('r'))
                    br = etree.SubElement(run, wt('br'))
                    br.set(f'{{{W}}}type', 'page')
                    break  # tylko raz
            
            scale_t3 = TARGET_WIDTH / ORIG_T3_TOTAL
            # tblW musi też dostać TARGET_WIDTH — bez tego LO bierze stare
            # tblW (9450) jako autorytet i kompresuje gridCol proporcjonalnie
            # → Wygr cell renderowany jako ~768 dxa, „Wygrane" zawija na 3 linie.
            tblPr_t3 = t3.find(wt('tblPr'))
            if tblPr_t3 is not None:
                tblW3 = tblPr_t3.find(wt('tblW'))
                if tblW3 is None:
                    tblW3 = etree.SubElement(tblPr_t3, wt('tblW'))
                tblW3.set(f'{{{W}}}w', str(TARGET_WIDTH))
                tblW3.set(f'{{{W}}}type', 'dxa')
            grid3 = t3.find(wt('tblGrid'))
            if grid3 is not None:
                for col in grid3.findall(wt('gridCol')):
                    w = col.get(f'{{{W}}}w')
                    if w:
                        col.set(f'{{{W}}}w', str(int(int(w) * scale_t3)))
            for tr in t3.findall(wt('tr')):
                for tc in tr.findall(wt('tc')):
                    tcPr = tc.find(wt('tcPr'))
                    if tcPr is not None:
                        tcW = tcPr.find(wt('tcW'))
                        if tcW is not None:
                            w = tcW.get(f'{{{W}}}w')
                            if w:
                                tcW.set(f'{{{W}}}w', str(int(int(w) * scale_t3)))
            # Bo5: redystrybucja szerokości w T3 (header s.2) — analogicznie do T1.
            # Całe 800 z Podpis (jak T1) — Wygr „Wygr. sety" sz=20 potrzebuje
            # pełnej szerokości, nie kradniemy z niego (patrz komentarz T1).
            T3_FROM_PODPIS = 800
            T3_FROM_WYGRANE = 0
            T3_TOTAL = T3_FROM_PODPIS + T3_FROM_WYGRANE
            grid3 = t3.find(wt('tblGrid'))
            if grid3 is not None:
                gcols = grid3.findall(wt('gridCol'))
                if len(gcols) >= 3:
                    first = gcols[0]
                    wygrane = gcols[-2]
                    podpis = gcols[-1]
                    first.set(f'{{{W}}}w',
                              str(int(first.get(f'{{{W}}}w','0')) + T3_TOTAL))
                    wygrane.set(f'{{{W}}}w',
                                str(int(wygrane.get(f'{{{W}}}w','0')) - T3_FROM_WYGRANE))
                    podpis.set(f'{{{W}}}w',
                               str(int(podpis.get(f'{{{W}}}w','0')) - T3_FROM_PODPIS))
            for tr in t3.findall(wt('tr')):
                cells = tr.findall(wt('tc'))
                if len(cells) >= 3:
                    deltas = [
                        (cells[0],  T3_TOTAL),
                        (cells[-2], -T3_FROM_WYGRANE),
                        (cells[-1], -T3_FROM_PODPIS),
                    ]
                    for tc, delta in deltas:
                        tcPr = tc.find(wt('tcPr'))
                        if tcPr is None: continue
                        tcW = tcPr.find(wt('tcW'))
                        if tcW is None: continue
                        cw = tcW.get(f'{{{W}}}w')
                        if cw:
                            tcW.set(f'{{{W}}}w', str(int(cw) + delta))
            # T4 zostaje w oryginalnej szerokości (6425 dxa) — narrower bo tylko 2 sety
            # ale aplikujemy fix ramek dla ostatniej kolumny (jak w T2)
            t4_rows = t4.findall(wt('tr'))
            for r_idx, tr in enumerate(t4_rows):
                if r_idx < 3 or r_idx == len(t4_rows) - 1:
                    continue
                cells = tr.findall(wt('tc'))
                if not cells: continue
                last_tc = cells[-1]
                tcPr = last_tc.find(wt('tcPr'))
                if tcPr is None: continue
                tcBorders = tcPr.find(wt('tcBorders'))
                if tcBorders is None: continue
                for side in ('top', 'bottom'):
                    b = tcBorders.find(wt(side))
                    if b is not None:
                        tcBorders.remove(b)

    # ── Operacje SPECYFICZNE DLA CZWÓRKOWEGO szablonu (CZWORKA_Grupa):
    # T1 i T2 mają szerokość 9026/9014 dxa → skalujemy do TARGET_WIDTH (pełna strona).
    # Calibri normalizacja: Aptos (font szablonu) nie jest dostępny w LibreOffice
    # → fallback na szerszy font → etykiety T1 zawijają się na 2 linie.
    if template_type == 'CZWORKA':
        CZWORKA_LABELS = {'Tor', 'Godzina', 'Godz.', 'Grupa', 'Mecz', '#', 'Mecz #',
                          'Punkty', 'SET 1', 'SET 2', 'Wygrane', 'sety', 'Podpis',
                          'Wyniki turnieju', 'Wyniki turnieju:'}
        for r in body.iter(wt('r')):
            ts = r.findall(wt('t'))
            if not ts: continue
            text_content = ''.join((t.text or '') for t in ts).strip()
            if text_content in CZWORKA_LABELS:
                rPr = r.find(wt('rPr'))
                if rPr is None:
                    rPr = etree.Element(wt('rPr'))
                    r.insert(0, rPr)
                fonts = rPr.find(wt('rFonts'))
                if fonts is None:
                    fonts = etree.SubElement(rPr, wt('rFonts'))
                for a in ('ascii', 'hAnsi', 'eastAsia', 'cs'):
                    fonts.set(f'{{{W}}}{a}', 'Calibri')

        TARGET_WIDTH = 10466
        CZWORKA_T1_TOTAL = 9026
        CZWORKA_T2_TOTAL = 9014
        czworka_tbls = body.findall(wt('tbl'))
        if len(czworka_tbls) >= 2:
            ct1, ct2 = czworka_tbls[0], czworka_tbls[1]
            for tbl, orig_total in ((ct1, CZWORKA_T1_TOTAL), (ct2, CZWORKA_T2_TOTAL)):
                scale = TARGET_WIDTH / orig_total
                tblPr = tbl.find(wt('tblPr'))
                if tblPr is not None:
                    tblW = tblPr.find(wt('tblW'))
                    if tblW is None:
                        tblW = etree.SubElement(tblPr, wt('tblW'))
                    tblW.set(f'{{{W}}}w', str(TARGET_WIDTH))
                    tblW.set(f'{{{W}}}type', 'dxa')
                    tblInd = tblPr.find(wt('tblInd'))
                    if tblInd is None:
                        tblInd = etree.SubElement(tblPr, wt('tblInd'))
                    tblInd.set(f'{{{W}}}w', '0')
                    tblInd.set(f'{{{W}}}type', 'dxa')
                grid = tbl.find(wt('tblGrid'))
                if grid is not None:
                    for gc in grid.findall(wt('gridCol')):
                        w = gc.get(f'{{{W}}}w')
                        if w:
                            gc.set(f'{{{W}}}w', str(int(int(w) * scale)))
                for tr in tbl.findall(wt('tr')):
                    for tc in tr.findall(wt('tc')):
                        tcPr = tc.find(wt('tcPr'))
                        if tcPr is not None:
                            tcW = tcPr.find(wt('tcW'))
                            if tcW is not None:
                                w = tcW.get(f'{{{W}}}w')
                                if w:
                                    tcW.set(f'{{{W}}}w', str(int(int(w) * scale)))

        # Usuń <w:br/> przed "Wyniki turnieju" — szablon ma puste run-y z break
        # przed tekstem które dodają linijkę odstępu i wypychają stronę 2.
        # Lokalizujemy paragraph zawierający 'Wyniki turnieju' i usuwamy z niego
        # wszystkie <w:r> które zawierają TYLKO <w:br/> bez tekstu.
        # Pozycja QR z image_positions — żeby "Wyniki" mieściło się w jego szerokości
        # i było centrowane nad nim. Bez image_positions używamy defaultów (x=0, w=1.6).
        _qr_x = 0.0
        _qr_w = 1.6
        if image_positions and 'qr' in image_positions:
            _qr_pos = image_positions['qr']
            _qr_x = float(_qr_pos.get('x', 0.0))
            _qr_w = float(_qr_pos.get('w', 1.6))
        # Indent paragrafu tak, że tekst zajmuje DOKŁADNIE kolumnę QR — leftInd
        # przesuwa do x=_qr_x, rightInd zostawia (table_w - qr_right) z prawej.
        # Tabela CZWORKA: TARGET_WIDTH 10466 dxa = 18.46 cm.
        TBL_W_CM = 18.46
        _left_ind = max(0, int(_qr_x * 567))
        _right_ind = max(0, int((TBL_W_CM - _qr_x - _qr_w) * 567))

        for p in body.iter(wt('p')):
            txts = p.findall(f'.//{wt("t")}')
            full_text = ''.join((t.text or '') for t in txts)
            if 'Wyniki turnieju' not in full_text:
                continue
            # Usuń <br/> bez tekstu (puste run-y z break z szablonu)
            for r in list(p.findall(wt('r'))):
                if r.find(wt('br')) is not None and not r.findall(wt('t')):
                    p.remove(r)
            # Skróć tekst do "Wyniki" + wymuś rozmiar fontu (20 = 10pt) by zmieściło się
            # w 1.6 cm kolumny QR. Bez tego template-default może być za duży.
            for t in p.findall(f'.//{wt("t")}'):
                if t.text and 'Wyniki turnieju' in t.text:
                    t.text = t.text.replace('Wyniki turnieju', 'Wyniki')
            for r in p.findall(wt('r')):
                if not any((t.text or '').strip() for t in r.findall(wt('t'))):
                    continue
                rPr = r.find(wt('rPr'))
                if rPr is None:
                    rPr = etree.Element(wt('rPr'))
                    r.insert(0, rPr)
                for tag in ('sz', 'szCs'):
                    el = rPr.find(wt(tag))
                    if el is None:
                        el = etree.SubElement(rPr, wt(tag))
                    el.set(f'{{{W}}}val', '20')  # 10pt

            pPr = p.find(wt('pPr'))
            if pPr is None:
                pPr = etree.Element(wt('pPr'))
                p.insert(0, pPr)

            # Spacing: 400 dxa before (~0.7 cm) — przesuwa "Wyniki" w połowę drogi
            # między dolną krawędzią tabeli a rzędem QR/grafik.
            spacing = pPr.find(wt('spacing'))
            if spacing is None:
                spacing = etree.SubElement(pPr, wt('spacing'))
            spacing.set(f'{{{W}}}before', '40')   # ~0.07 cm — minimalna przerwa od PKT, w granicy bezpieczeństwa
            spacing.set(f'{{{W}}}after', '0')
            spacing.set(f'{{{W}}}line', '240')
            spacing.set(f'{{{W}}}lineRule', 'auto')

            # Ind: ogranicza szerokość paragrafu do kolumny QR (1.6 cm), reszta tabeli
            # zarezerwowana przez right indent. Tekst się centruje wewnątrz kolumny.
            ind = pPr.find(wt('ind'))
            if ind is None:
                ind = etree.SubElement(pPr, wt('ind'))
            ind.set(f'{{{W}}}left', str(_left_ind))
            ind.set(f'{{{W}}}right', str(_right_ind))

            jc = pPr.find(wt('jc'))
            if jc is None:
                jc = etree.SubElement(pPr, wt('jc'))
            jc.set(f'{{{W}}}val', 'center')

    # ── CZWORKA pucharowa (Bo3/Bo5): normalizacja fontów (Calibri) bez TROJKA-scalingu.
    # Templates CZWORKA_Bo3/CZWORKA_Bo5 mieszczą się w 18.46 cm (margines 720), więc
    # nie skalujemy tabel — tylko wymuszamy Calibri na labelach (bez tego Aptos fallback
    # w LibreOffice rozjeżdża etykiety) i pomniejszamy wiersz Pkt/SET do sz=20.
    if template_type in ('CZWORKA_Bo3', 'CZWORKA_Bo5', 'DWOJKA_Bo7'):
        # Tor/Godzina/Mecz# celowo NIE są w tym secie — template używa Aptos Narrow,
        # co powoduje zawijanie "Tor"→"To/r" w wąskiej kolumnie, dokładnie jak w wzorcu.
        # Konwersja na Calibri uniemożliwiałaby zawijanie i zmieniała wygląd.
        # (SET 6/7, Pkt SET 6/7, (SET 5-7) dodane dla DWÓJKA Bo7 — 7 setów.)
        CZW_BO_LABELS = {'PunktySET 1', 'PunktySET 2', 'PunktySET 3',
                         'PunktySET 4', 'PunktySET 5', 'PunktySET 6', 'PunktySET 7',
                         'PktSET 1', 'PktSET 2', 'PktSET 3', 'PktSET 4', 'PktSET 5',
                         'PktSET 6', 'PktSET 7', 'Pkt.SET 6', 'Pkt.SET 7',
                         'Punkty', 'Pkt', 'Wygrane', 'sety', 'Wygranesety',
                         'Wygr.sety', 'Wygr. sety', 'Podpis',
                         'SET 1', 'SET 2', 'SET 3', 'SET 4', 'SET 5', 'SET 6', 'SET 7',
                         '(SET 3)', '(SET 4)', '(SET 5)', '(SET 6)', '(SET 7)',
                         'SUMA', 'S U M A', 'I M I O N A'}
        for r in body.iter(wt('r')):
            ts = r.findall(wt('t'))
            if not ts: continue
            if ''.join((t.text or '') for t in ts).strip() in CZW_BO_LABELS:
                rPr = r.find(wt('rPr'))
                if rPr is None:
                    rPr = etree.Element(wt('rPr')); r.insert(0, rPr)
                fonts = rPr.find(wt('rFonts'))
                if fonts is None:
                    fonts = etree.SubElement(rPr, wt('rFonts'))
                for a in ('ascii', 'hAnsi', 'eastAsia', 'cs'):
                    fonts.set(f'{{{W}}}{a}', 'Calibri')
        # NUCLEAR: każdy wiersz z Pkt/Punkt + SET → Calibri + vAlign center.
        # Font zależny od szerokości komórki: wąskie (<800 dxa, np. 5 setów Bo5) → sz=16,
        # szersze → sz=20. Bez tego "PktSET2" łamie się na 3 linie w wąskich kolumnach.
        for tbl in body.findall(wt('tbl')):
            for tr in tbl.findall(wt('tr')):
                row_text = ''.join((t.text or '') for t in tr.iter(wt('t')))
                if 'SET' not in row_text or ('Pkt' not in row_text and 'Punkt' not in row_text):
                    continue
                for tc in tr.findall(wt('tc')):
                    # Tylko komórki LABEL'owe: Pkt/Punkt+SET (rozmiar zależny od szer.),
                    # oraz Wygr./sety i Podpis (tylko font Calibri żeby nie wrapsowały
                    # przez Aptos-fallback w LO). Wcześniej processowało wszystkie komórki
                    # w wierszu, w tym Tor (cw<800 → sz=16 zamiast 20).
                    cell_text = ''.join((t.text or '') for t in tc.iter(wt('t')))
                    is_pkt_set = ('SET' in cell_text and ('Pkt' in cell_text or 'Punkt' in cell_text))
                    is_label = ('Wygr' in cell_text or 'sety' in cell_text or 'Podpis' in cell_text)
                    if not (is_pkt_set or is_label):
                        continue
                    tcPr = tc.find(wt('tcPr'))
                    if tcPr is None:
                        tcPr = etree.Element(wt('tcPr')); tc.insert(0, tcPr)
                    vA = tcPr.find(wt('vAlign'))
                    if vA is None:
                        vA = etree.SubElement(tcPr, wt('vAlign'))
                    vA.set(f'{{{W}}}val', 'center')
                    tcW = tcPr.find(wt('tcW'))
                    cw = int(tcW.get(f'{{{W}}}w', '990')) if tcW is not None else 990
                    cell_sz = '16' if cw < 800 else '20'
                    for run in tc.iter(wt('r')):
                        ts = run.findall(wt('t'))
                        if not ts or not ''.join((t.text or '') for t in ts).strip():
                            continue
                        rPr = run.find(wt('rPr'))
                        if rPr is None:
                            rPr = etree.Element(wt('rPr')); run.insert(0, rPr)
                        # rozmiar wg szerokości komórki dla Pkt/SET; dla DWÓJKA Bo7
                        # TAKŻE dla Wygr./Podpis — bez tego „Wygr. sety" zostawała na
                        # template'owym sz=20 (większa niż Pkt SET sz=16 → user:
                        # „wygr. sety większa czcionka niepotrzebnie"). Width-based
                        # daje Wygr (733<800)=16 = jak Pkt SET, Podpis (1165)=20.
                        if is_pkt_set or (is_label and template_type == 'DWOJKA_Bo7'):
                            for tag in ('sz', 'szCs'):
                                el = rPr.find(wt(tag))
                                if el is None:
                                    el = etree.SubElement(rPr, wt(tag))
                                el.set(f'{{{W}}}val', cell_sz)
                        # Calibri ZAWSZE (rozwiązuje Aptos-fallback wrap)
                        fonts = rPr.find(wt('rFonts'))
                        if fonts is None:
                            fonts = etree.SubElement(rPr, wt('rFonts'))
                        for a in ('ascii', 'hAnsi', 'eastAsia', 'cs'):
                            fonts.set(f'{{{W}}}{a}', 'Calibri')

        # Tabele headerowe (Tor/Godzina/Mecz# + instrukcje + nazwy drużyn):
        # wymuś Calibri na WSZYSTKICH runach. Bez Aptos/Aptos Narrow LO bierze szeroki
        # fallback, przez co (a) "Tor" w wąskiej kol. 536 dxa zawija na 2 linie, (b) text
        # instrukcji (jc=right, sz=20) przelewa się poza komórkę i jest obcinany z lewej.
        # Calibri (Carlito) mieści się — Tor jednolinijny, instrukcje nie obcięte.
        # Bo5 ma 2 takie tabele (header str.1 + header str.2 — tbl[0] i tbl[2]),
        # Bo3 tylko 1 (tbl[0]). Wykrywamy po obecności "Tor" w pierwszym wierszu.
        tbls = list(body.findall(wt('tbl')))
        header_tbls = []
        for tbl in tbls:
            first_row = tbl.find(wt('tr'))
            if first_row is None:
                continue
            row_text = ''.join((t.text or '') for t in first_row.iter(wt('t')))
            if 'Tor' in row_text and ('Mecz' in row_text or 'Godzina' in row_text):
                header_tbls.append(tbl)
        for htbl in header_tbls:
            for run in htbl.iter(wt('r')):
                rPr = run.find(wt('rPr'))
                if rPr is None:
                    rPr = etree.Element(wt('rPr')); run.insert(0, rPr)
                fonts = rPr.find(wt('rFonts'))
                if fonts is None:
                    fonts = etree.SubElement(rPr, wt('rFonts'))
                for a in ('ascii', 'hAnsi', 'eastAsia', 'cs'):
                    fonts.set(f'{{{W}}}{a}', 'Calibri')

            # Usuń poziomą linię między row 1 a row 2 w komórce instrukcji
            # (vMerge). Row 1 c0 (restart) ma bottom=nil, ale row 2 c0 (cont.)
            # dziedziczy tblBorders top=single → LO rysuje cienką linię
            # pod tekstem instrukcji ("Set przegrany... wynik 0:50.").
            # Fix: wymuś top=nil i bottom=nil na komórkach vMerge w tej kolumnie.
            for tr in htbl.findall(wt('tr')):
                for tc in tr.findall(wt('tc')):
                    vm = tc.find(f'{wt("tcPr")}/{wt("vMerge")}')
                    if vm is None:
                        continue
                    tcPr = tc.find(wt('tcPr'))
                    bdr = tcPr.find(wt('tcBorders'))
                    if bdr is None:
                        bdr = etree.SubElement(tcPr, wt('tcBorders'))
                    for side in ('top', 'bottom'):
                        b = bdr.find(wt(side))
                        if b is None:
                            b = etree.SubElement(bdr, wt(side))
                        b.set(f'{{{W}}}val', 'nil')

        # Bo5: komórki "Pkt SET N" w nagłówku są wąskie (720-745 dxa) i 1-liniowy
        # "Pkt. SET 1" nie mieści się — różnie wrapuje (c7 ma '. ', c8-c11 ma '.' +
        # ' SET '), więc renderują się niespójnie. Plus poprzednio nuclear block
        # wymuszał sz=16, a Wygr.sety/Podpis zostawały sz=20 → wizualna niespójność.
        # Fix: rebuild WSZYSTKICH komórek "Pkt SET N" w wąskich (cw<900) do 2 linii
        # "Pkt" + "SET N" przy sz=20 (zgodnie z Wygr/Podpis).
        if template_type in ('CZWORKA_Bo5', 'DWOJKA_Bo7'):
            import re as _re
            XMLSPACE = '{http://www.w3.org/XML/1998/namespace}space'
            # DWÓJKA Bo7: komórki Pkt SET wąskie (720) → sz=16 (jak NUCLEAR);
            # CZWÓRKA Bo5 → sz=20. Rebuild ujednolica run-splity (c7 miał „SET 1"
            # jednym runem, c8+ „ SET "+„N") → spójne centrowanie „Pkt SET 7".
            _rebuild_sz = '16' if template_type == 'DWOJKA_Bo7' else '20'
            for htbl in header_tbls:
                for tc in htbl.iter(wt('tc')):
                    full = ''.join((t.text or '') for t in tc.iter(wt('t')))
                    m = _re.search(r'(Punkty|Pkt)\.?\s*SET\s*(\d)', full.replace('\n', ''))
                    if not m:
                        continue
                    tcW = tc.find(f'{wt("tcPr")}/{wt("tcW")}')
                    cw = int(tcW.get(f'{{{W}}}w', '990')) if tcW is not None else 990
                    if cw >= 900:
                        continue
                    prefix, setnum = m.group(1), m.group(2)
                    ps = tc.findall(wt('p'))
                    if not ps:
                        continue
                    p = ps[0]
                    for extra in ps[1:]:
                        tc.remove(extra)
                    for r in p.findall(wt('r')):
                        p.remove(r)
                    pPr = p.find(wt('pPr'))
                    if pPr is None:
                        pPr = etree.Element(wt('pPr')); p.insert(0, pPr)
                    jc = pPr.find(wt('jc'))
                    if jc is None:
                        jc = etree.SubElement(pPr, wt('jc'))
                    jc.set(f'{{{W}}}val', 'center')

                    def _mkrun(txt, with_break):
                        r = etree.SubElement(p, wt('r'))
                        rPr = etree.SubElement(r, wt('rPr'))
                        fonts = etree.SubElement(rPr, wt('rFonts'))
                        for a in ('ascii', 'hAnsi', 'eastAsia', 'cs'):
                            fonts.set(f'{{{W}}}{a}', 'Calibri')
                        for tag in ('b', 'bCs'):
                            etree.SubElement(rPr, wt(tag)).set(f'{{{W}}}val', '1')
                        for tag in ('sz', 'szCs'):
                            etree.SubElement(rPr, wt(tag)).set(f'{{{W}}}val', _rebuild_sz)
                        if with_break:
                            etree.SubElement(r, wt('br'))
                        t = etree.SubElement(r, wt('t'))
                        t.text = txt
                        t.set(XMLSPACE, 'preserve')

                    _mkrun(f'{prefix}.', False)
                    _mkrun(f'SET {setnum}', True)
                    tcPr = tc.find(wt('tcPr'))
                    if tcPr is not None:
                        vA = tcPr.find(wt('vAlign'))
                        if vA is None:
                            vA = etree.SubElement(tcPr, wt('vAlign'))
                        vA.set(f'{{{W}}}val', 'center')

        # ── DWÓJKA Bo7: NUKLEARNE Calibri na CAŁYM dokumencie ──────────────
        # User: „szeryfów dużo, trudno utrzymać czcionkę z szablonu?". Etykiety
        # SET 1-7 / (SET 5-7) / numery wierszy 1-18 / PKT w TABELI WYNIKÓW (nie
        # header) były rozbite na runy „SET"+„ 1" itd. z font=None/Aptos Narrow →
        # LO bierze szeryfowy fallback. header_tbls force łapie tylko nagłówek.
        # Tu wymuszamy Calibri na KAŻDYM runie w body — POZA szarym nagłówkiem
        # turnieju (color=666666, Aptos italic, user: „w prawym górnym może zostać").
        # Rozmiarów NIE ruszamy (zostają z szablonu / wcześniejszych bloków).
        if template_type == 'DWOJKA_Bo7':
            for run in body.iter(wt('r')):
                if not run.findall(wt('t')):
                    continue
                rPr = run.find(wt('rPr'))
                if rPr is not None:
                    clr = rPr.find(wt('color'))
                    if clr is not None and (clr.get(f'{{{W}}}val') or '').lower() == '666666':
                        continue  # szary nagłówek turnieju — zostaw Aptos
                else:
                    rPr = etree.Element(wt('rPr')); run.insert(0, rPr)
                fonts = rPr.find(wt('rFonts'))
                if fonts is None:
                    fonts = etree.SubElement(rPr, wt('rFonts'))
                for a in ('ascii', 'hAnsi', 'eastAsia', 'cs'):
                    fonts.set(f'{{{W}}}{a}', 'Calibri')

    # ── Operacje SPECYFICZNE DLA INDYWIDUALNEGO szablonu:
    # 1) Pomniejszenie fontów etykiet (24→20)
    # 2) Przesunięcie tabeli 1 w prawo (tblInd=600 dxa)
    # Trójkowy szablon ma inną strukturę i nie wymaga tych poprawek.
    if template_type in ('IND', 'IND_Bo3', 'IND_Bo5', 'IND_Bo7'):
        # IND_Bo3: zwęzić Wygrane sety (idx 4: 1275→990 dxa) na rzecz Podpis (idx 5, +285).
        # Tylko Bo3 — bo Bo3 T0 ma 6 kolumn (Tor Godz SET1-3 → idx 4=Wygrane, idx 5=Podpis).
        # Bo5 ma 8 kolumn (SET1-5) gdzie idx 4/5 to Pkt SET 4/5 — NIE ruszamy.
        if template_type == 'IND_Bo3':
            tbls_w = body.findall(wt('tbl'))
            if tbls_w:
                t1_w = tbls_w[0]
                grid = t1_w.find(wt('tblGrid'))
                if grid is not None:
                    gcs = grid.findall(wt('gridCol'))
                    if len(gcs) >= 6:
                        gcs[4].set(f'{{{W}}}w', '990')
                        old_podpis_w = int(gcs[5].get(f'{{{W}}}w', '2160'))
                        gcs[5].set(f'{{{W}}}w', str(old_podpis_w + 285))
                for tr in t1_w.findall(wt('tr')):
                    tcs = tr.findall(wt('tc'))
                    if len(tcs) >= 6:
                        tcW4 = tcs[4].find(f'{wt("tcPr")}/{wt("tcW")}')
                        if tcW4 is not None:
                            tcW4.set(f'{{{W}}}w', '990')
                        tcW5 = tcs[5].find(f'{wt("tcPr")}/{wt("tcW")}')
                        if tcW5 is not None:
                            old_w5 = int(tcW5.get(f'{{{W}}}w', '2160'))
                            tcW5.set(f'{{{W}}}w', str(old_w5 + 285))

        # IND_Bo3/Bo5: pomniejszenie WSZYSTKICH runs w row 1 T1 (Punkty/SET labels)
        # — bez sprawdzania text content. Template ma "Punkty" split na "Punkt"+"y"
        # jako osobne runs, więc selective matching nie działa.
        if template_type in ('IND_Bo3', 'IND_Bo5', 'IND_Bo7'):
            IND_BO_LABELS = {'Tor', 'Godz.', 'Godz', 'Godzina', 'Mecz', 'Mecz #', '#',
                             'PunktySET 1', 'PunktySET 2', 'PunktySET 3',
                             'PunktySET 4', 'PunktySET 5', 'PunktySET 6', 'PunktySET 7',
                             'Punkty', 'SET 1', 'SET 2', 'SET 3', 'SET 4', 'SET 5',
                             'SET 6', 'SET 7',
                             '(SET 3)', '(SET 4)', '(SET 5)', '(SET 6)', '(SET 7)',
                             'Wygrane sety', 'Wygranesety', 'Wygrane', 'sety',
                             'Podpis', 'Runda', 'I M I O N A', 'S U M A'}
            for r in body.iter(wt('r')):
                ts = r.findall(wt('t'))
                if not ts: continue
                text_content = ''.join((t.text or '') for t in ts).strip()
                # Exact label LUB prefix „Mecz #" — po fill etykieta to „Mecz #  1"
                # (IND_Bo7), single-text nie pasuje do żadnego stringa z setu → bez
                # tego LO bierze Aptos fallback (szeryf na numerze meczu).
                if text_content in IND_BO_LABELS or text_content.startswith('Mecz #'):
                    rPr = r.find(wt('rPr'))
                    if rPr is None:
                        rPr = etree.Element(wt('rPr'))
                        r.insert(0, rPr)
                    fonts = rPr.find(wt('rFonts'))
                    if fonts is None:
                        fonts = etree.SubElement(rPr, wt('rFonts'))
                    for a in ('ascii', 'hAnsi', 'eastAsia', 'cs'):
                        fonts.set(f'{{{W}}}{a}', 'Calibri')

            # NUCLEAR: znajdź KAŻDY wiersz tabeli zawierający label "SET" (Pkt SET / Punkty SET)
            # i pomniejsz WSZYSTKIE runs w nim do sz=20 Calibri + vAlign=center.
            # Table-agnostic: działa dla Bo3 (1 tabela header) i Bo5 (2 tabele header,
            # str.1 i str.2). Bez tego "Punkt"/"y" split zostają sz=24 i łamią na 3 linie.
            for tbl in body.findall(wt('tbl')):
                for tr in tbl.findall(wt('tr')):
                    row_text = ''.join((t.text or '') for t in tr.iter(wt('t')))
                    # Wiersz Punkty/SET: zawiera "SET" i "Pkt"/"Punkt"
                    if 'SET' not in row_text or ('Pkt' not in row_text and 'Punkt' not in row_text):
                        continue
                    for tc in tr.findall(wt('tc')):
                        # Tylko komórki zawierające „Pkt"/„Punkt" + „SET" — bez tego
                        # font na Tor/Godz/Mecz# w tym samym wierszu też się skraca.
                        cell_text = ''.join((t.text or '') for t in tc.iter(wt('t')))
                        if 'SET' not in cell_text or ('Pkt' not in cell_text and 'Punkt' not in cell_text):
                            continue
                        tcPr = tc.find(wt('tcPr'))
                        if tcPr is None:
                            tcPr = etree.Element(wt('tcPr'))
                            tc.insert(0, tcPr)
                        vAlign = tcPr.find(wt('vAlign'))
                        if vAlign is None:
                            vAlign = etree.SubElement(tcPr, wt('vAlign'))
                        vAlign.set(f'{{{W}}}val', 'center')
                        # Font zależny od szerokości: wąskie (<800 dxa, 5 setów Bo5) → 16, reszta 20.
                        tcW = tcPr.find(wt('tcW'))
                        cw = int(tcW.get(f'{{{W}}}w', '990')) if tcW is not None else 990
                        cell_sz = '16' if cw < 800 else '20'
                        for p in tc.findall(wt('p')):
                            for run in p.findall(wt('r')):
                                ts = run.findall(wt('t'))
                                if not ts: continue
                                txt = ''.join((t.text or '') for t in ts).strip()
                                if not txt: continue
                                rPr = run.find(wt('rPr'))
                                if rPr is None:
                                    rPr = etree.Element(wt('rPr'))
                                    run.insert(0, rPr)
                                for tag in ('sz', 'szCs'):
                                    el = rPr.find(wt(tag))
                                    if el is None:
                                        el = etree.SubElement(rPr, wt(tag))
                                    el.set(f'{{{W}}}val', cell_sz)
                                fonts = rPr.find(wt('rFonts'))
                                if fonts is None:
                                    fonts = etree.SubElement(rPr, wt('rFonts'))
                                for a in ('ascii', 'hAnsi', 'eastAsia', 'cs'):
                                    fonts.set(f'{{{W}}}{a}', 'Calibri')

            # NUCLEAR 2: nagłówkowe komórki „Wygrane sety" i „Podpis" — normalizuj
            # WSZYSTKIE runy w nich do Calibri + bold + tym samym rozmiarem co
            # sąsiednie „Pkt SET N" w tym samym wierszu. W IND_Bo5 „Wygrane sety"
            # jest rozbite na runy „Wygr"+„."+„ sety"; selektywne matchowanie po
            # tekście łapało tylko „sety" (→Calibri sz=18), a „Wygr"+„." zostawały
            # w Aptos BEZ rozmiaru → dziedziczyły docDefault (~24pt) i serif
            # fallback w LO, przez co „Wygr." renderowało się wielką, dziwną
            # czcionką. ROZMIAR: dopasowujemy do tego co NUCLEAR-SET (wyżej)
            # wymusza na komórkach „Pkt SET N" w tym samym wierszu — `cw<800`
            # (wąskie Bo5) → sz=16, inaczej sz=20. Bez tego dopasowania użytkownik
            # widział wizualną różnicę: „Punkty SET" sz=20 vs „Wygrane sety" sz=18.
            for tbl in body.findall(wt('tbl')):
                rows_t = tbl.findall(wt('tr'))
                # tylko tabele nagłówkowe (mało wierszy) — nie tabele wynikowe
                if len(rows_t) > 6:
                    continue
                for tr in rows_t:
                    # 1) Wyznacz rozmiar wiersza patrząc na pierwszą Pkt-SET komórkę
                    row_size = None
                    for tc in tr.findall(wt('tc')):
                        ctext = ''.join((t.text or '') for t in tc.iter(wt('t')))
                        if ('Pkt' in ctext or 'Punkt' in ctext) and 'SET' in ctext:
                            tcPr_sib = tc.find(wt('tcPr'))
                            tcW_sib = tcPr_sib.find(wt('tcW')) if tcPr_sib is not None else None
                            cw = int(tcW_sib.get(f'{{{W}}}w', '990')) if tcW_sib is not None else 990
                            row_size = '16' if cw < 800 else '20'
                            break
                    if not row_size:
                        continue
                    # 2) Aplikuj do komórek Wygrane sety / Podpis w tym wierszu
                    for tc in tr.findall(wt('tc')):
                        cell_text = ''.join((t.text or '') for t in tc.iter(wt('t')))
                        cell_strip = cell_text.replace(' ', '')
                        is_wygr = ('Wygr' in cell_text and 'set' in cell_text.lower())
                        is_podpis = (cell_strip == 'Podpis')
                        if not (is_wygr or is_podpis):
                            continue
                        for run in tc.iter(wt('r')):
                            ts = run.findall(wt('t'))
                            if not ts: continue
                            if not ''.join((t.text or '') for t in ts).strip():
                                continue
                            rPr = run.find(wt('rPr'))
                            if rPr is None:
                                rPr = etree.Element(wt('rPr'))
                                run.insert(0, rPr)
                            for tag in ('sz', 'szCs'):
                                el = rPr.find(wt(tag))
                                if el is None:
                                    el = etree.SubElement(rPr, wt(tag))
                                el.set(f'{{{W}}}val', row_size)
                            fonts = rPr.find(wt('rFonts'))
                            if fonts is None:
                                fonts = etree.SubElement(rPr, wt('rFonts'))
                            for a in ('ascii', 'hAnsi', 'eastAsia', 'cs'):
                                fonts.set(f'{{{W}}}{a}', 'Calibri')
                            # Wymuś bold + boldCs — w IND_Bo5 runy „Wygr"+„." MAJĄ
                            # bold w szablonie, ale dla pewności i spójności
                            # z sąsiednimi etykietami nagłówka (które są bold).
                            for btag in ('b', 'bCs'):
                                if rPr.find(wt(btag)) is None:
                                    etree.SubElement(rPr, wt(btag))

            # Instrukcje (paragrafy "Sety 1 i 2..." i "Set przegrany..."): wymusz italic + szary
            # + Calibri. Bez tego LibreOffice renderuje czarny non-italic (fallback z Aptos Narrow).
            for p in body.findall(wt('p')):
                full_text = ''.join((t.text or '') for t in p.iter(wt('t')))
                if not ('rozpoczynamy naprzemiennie' in full_text or 'kolejne chybienia' in full_text):
                    continue
                for run in p.findall(wt('r')):
                    if not run.findall(wt('t')): continue
                    rPr = run.find(wt('rPr'))
                    if rPr is None:
                        rPr = etree.Element(wt('rPr'))
                        run.insert(0, rPr)
                    # italic
                    for tag in ('i', 'iCs'):
                        if rPr.find(wt(tag)) is None:
                            etree.SubElement(rPr, wt(tag))
                    # gray color
                    color = rPr.find(wt('color'))
                    if color is None:
                        color = etree.SubElement(rPr, wt('color'))
                    color.set(f'{{{W}}}val', '595959')
                    # Calibri 9pt (18)
                    fonts = rPr.find(wt('rFonts'))
                    if fonts is None:
                        fonts = etree.SubElement(rPr, wt('rFonts'))
                    for a in ('ascii', 'hAnsi', 'eastAsia', 'cs'):
                        fonts.set(f'{{{W}}}{a}', 'Calibri')
                    for tag in ('sz', 'szCs'):
                        el = rPr.find(wt(tag))
                        if el is None:
                            el = etree.SubElement(rPr, wt(tag))
                        el.set(f'{{{W}}}val', '18')

        # LABELS_BIGGER pomniejszanie (→18) TYLKO dla IND grupowej — Bo3/Bo5 zostawiają 12pt.
        # LABELS_HEADER (→18) stosowane dla wszystkich (Punkty SET / Wygrane sety / Podpis).
        # Redukujemy gdy efektywny rozmiar = 24, czyli JAWNE sz=24 LUB BRAK sz (dziedziczy
        # docDefaults=24). Nowe szablony nie mają jawnych rozmiarów na etykietach — bez tej
        # obsługi "if None" etykiety zostawały w 12pt i łamały się na 2 wiersze w PDF.
        # Etykiet SET=28 z tabeli wynikowej NIE ruszamy (jawne 28 ≠ 24).
        apply_bigger_reduction = (template_type == 'IND')
        for r in body.iter(wt('r')):
            ts = r.findall(wt('t'))
            if not ts: continue
            text_content = ''.join((t.text or '') for t in ts).strip()
            new_size = None
            if apply_bigger_reduction and text_content in LABELS_BIGGER:
                new_size = '20'   # 10pt — Tor/Godzina/Grupa/Mecz # (TYLKO IND grupowa).
                                  # Mieści się dzięki zmniejszonym marginesom komórek
                                  # nagłówka (blok first_tbl niżej) — bez tego "Mecz #"
                                  # łamałoby się w wąskiej komórce 960 dxa.
            elif text_content in LABELS_HEADER:
                new_size = '18'   # 9pt — Punkty SET 1, Wygrane sety, Podpis
            if not new_size:
                continue
            rPr = r.find(wt('rPr'))
            cur = rPr.find(wt('sz')) if rPr is not None else None
            cur_val = cur.get(f'{{{W}}}val') if cur is not None else None
            if cur_val not in (None, '24'):
                continue
            if rPr is None:
                rPr = etree.Element(wt('rPr')); r.insert(0, rPr)
            for tag in ('sz', 'szCs'):
                el = rPr.find(wt(tag))
                if el is None:
                    el = etree.SubElement(rPr, wt(tag))
                el.set(f'{{{W}}}val', new_size)
            # Calibri na etykietach nagłówka grupowej IND — Aptos w LibreOffice ma
            # szerszy fallback i przy 12pt łamał "Tor"/"Punkty"/"Wygrane" na 2 wiersze.
            # (Pucharowe mają własne wymuszanie Calibri w bloku IND_BO_LABELS.)
            if template_type == 'IND':
                fonts = rPr.find(wt('rFonts'))
                if fonts is None:
                    fonts = etree.SubElement(rPr, wt('rFonts'))
                for a in ('ascii', 'hAnsi', 'eastAsia', 'cs'):
                    fonts.set(f'{{{W}}}{a}', 'Calibri')

        # ── Wyrównanie tabeli 1 z tabelą 2 (jak we wzorcu z gridlinami):
        # Tabela 1 ma być WĘŻSZA (9090 DXA z oryginału - bez zmian) ale PRZESUNIĘTA W PRAWO
        # przez tblInd=600 DXA. Wtedy:
        #   - lewy brzeg tabeli 1 = 600 DXA (= ~1.06 cm)
        #   - prawy brzeg tabeli 1 = 600 + 9090 = 9690 DXA = prawy brzeg tabeli 2 ✓
        # Pierwsza pusta komórka tabeli 1 (z nazwiskami) zaczyna się od pozycji 600 DXA,
        # czyli mniej więcej tam gdzie kolumna IMIONA w tabeli 2.
        # UWAGA: tylko dla IND grupowej. Bo3/Bo5 mają inne wymiary tabel.
        first_tbl = body.find(wt('tbl')) if template_type == 'IND' else None
        if first_tbl is not None:
            # Dodaj tblInd = 600 DXA (przesunięcie w prawo)
            tblPr = first_tbl.find(wt('tblPr'))
            if tblPr is not None:
                tblInd = tblPr.find(wt('tblInd'))
                if tblInd is None:
                    tblInd = etree.SubElement(tblPr, wt('tblInd'))
                tblInd.set(f'{{{W}}}w', '600')
                tblInd.set(f'{{{W}}}type', 'dxa')
            # Zmniejsz L/R marginesy komórek TYLKO w 1. wierszu nagłówka (Tor/Godzina/
            # Grupa/Mecz #) → etykiety mieszczą się przy sz=20, w szczególności "Mecz #"
            # w wąskiej komórce 960 dxa. NIE ruszamy dalszych wierszy (nazwiska zawodników
            # są wyrównane do prawej i przy zmniejszonym marginesie były za blisko krawędzi).
            _r0 = first_tbl.find(wt('tr'))
            for tc in (_r0.findall(wt('tc')) if _r0 is not None else []):
                tcPr = tc.find(wt('tcPr'))
                if tcPr is None:
                    tcPr = etree.Element(wt('tcPr')); tc.insert(0, tcPr)
                tcMar = tcPr.find(wt('tcMar'))
                if tcMar is None:
                    tcMar = etree.SubElement(tcPr, wt('tcMar'))
                for side in ('left', 'right'):
                    m = tcMar.find(wt(side))
                    if m is None:
                        m = etree.SubElement(tcMar, wt(side)); m.set(f'{{{W}}}type', 'dxa')
                    m.set(f'{{{W}}}w', '30')
            # Szerokości gridCols i komórek - bez zmian (oryginalne 9090 DXA)
            # Wyrównaj "Tor" do lewej
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

    # ── Tabela 2: ZOSTAWIAMY ORYGINAŁ (9690 DXA z szablonu, IMIONA=840 DXA, WYNIK=840 DXA).
    # Komórka WYNIK (Indywidualny) lub PKT (Trójkowy) używa fontu Aptos Narrow size 20.
    # Wymuszamy Calibri size 18 (9pt) + zmniejszony padding (50 dxa) żeby się zawsze mieściło.
    tbls = body.findall(wt('tbl'))
    if len(tbls) >= 2:
        score_tbl = tbls[1]
        rows = score_tbl.findall(wt('tr'))
        if rows:
            last_row = rows[-1]
            for tc in last_row.findall(wt('tc')):
                txts = tc.findall(f'.//{wt("t")}')
                txt = ''.join((t.text or '') for t in txts).strip()
                if txt in ('WYNIK', 'PKT'):
                    # Rozmiar zależny od szerokości komórki: wąska (<820 dxa, np. nowy
                    # szablon IND ma 780) → 16, żeby "WYNIK" nie łamał się na 2 linie.
                    _tcW = tc.find(f'{wt("tcPr")}/{wt("tcW")}')
                    _cw = int(_tcW.get(f'{{{W}}}w', '840')) if _tcW is not None else 840
                    wynik_sz = '16' if _cw < 820 else '18'
                    # 1) Wymuś font Calibri + size dla KAŻDEGO 'r' (także bez rPr)
                    for r in tc.iter(wt('r')):
                        rPr = r.find(wt('rPr'))
                        if rPr is None:
                            rPr = etree.Element(wt('rPr'))
                            r.insert(0, rPr)
                        # font
                        fonts = rPr.find(wt('rFonts'))
                        if fonts is None:
                            fonts = etree.SubElement(rPr, wt('rFonts'))
                        for a in ('ascii', 'hAnsi', 'eastAsia', 'cs'):
                            fonts.set(f'{{{W}}}{a}', 'Calibri')
                        # size (16 wąska / 18 normalna)
                        sz = rPr.find(wt('sz'))
                        if sz is None:
                            sz = etree.SubElement(rPr, wt('sz'))
                        sz.set(f'{{{W}}}val', wynik_sz)
                        szCs = rPr.find(wt('szCs'))
                        if szCs is None:
                            szCs = etree.SubElement(rPr, wt('szCs'))
                        szCs.set(f'{{{W}}}val', wynik_sz)
                        # bold
                        b = rPr.find(wt('b'))
                        if b is None:
                            b = etree.SubElement(rPr, wt('b'))
                        b.set(f'{{{W}}}val', '1')
                    # 2) Zmniejsz padding komórki (105 → 50 dxa po obu stronach)
                    tcPr = tc.find(wt('tcPr'))
                    if tcPr is not None:
                        tcMar = tcPr.find(wt('tcMar'))
                        if tcMar is not None:
                            for side in ('left', 'right'):
                                m = tcMar.find(wt(side))
                                if m is not None:
                                    m.set(f'{{{W}}}w', '50')
                    break

    # ── Pucharowa (Bo3/Bo5): napraw komórki Pkt SET (3 linie → 2) + Calibri na
    # etykietach tabel wynikowych (IMIONA/SUMA/WYNIK krzywe przez Aptos Narrow).
    if template_type in ('IND_Bo3', 'IND_Bo5', 'IND_Bo7', 'TROJKA_Bo3', 'TROJKA_Bo5',
                         'CZWORKA_Bo3', 'CZWORKA_Bo5'):
        # CZWORKA_Bo5: pomijamy _fix_pkt_set_cells — komórki Pkt SET są już
        # zbudowane wcześniej w bloku CZWORKA z sz=20 (zgodnie z Wygr/Podpis).
        # Inaczej fix_pkt_set_cells zobaczyłby <w:br/> i przebudował ponownie
        # z narrow='16' (cw<800), przez co Pkt SET znów byłyby mniejsze.
        if template_type != 'CZWORKA_Bo5':
            _fix_pkt_set_cells(body)
        _force_calibri_score_labels(body)

    # ── IND_Bo5: wyrównaj tabelę wyników str.2 ((SET 4)/(SET 5)) do lewej, do tego
    # samego wcięcia (715 dxa) co tabela str.1. Szablon ma ją wyśrodkowaną (jc=center,
    # tblW=auto). KLUCZ: trzeba ustawić JAWNE jc=left — samo usunięcie jc nie wystarcza,
    # bo LibreOffice wtedy i tak centruje tabelę, ignorując tblInd.
    if template_type == 'IND_Bo5':
        _TBLPR_ORDER = ['tblStyle', 'tblpPr', 'tblOverlap', 'bidiVisual',
                        'tblStyleRowBandSize', 'tblStyleColBandSize', 'tblW', 'jc',
                        'tblCellSpacing', 'tblInd', 'tblBorders', 'shd', 'tblLayout',
                        'tblCellMar', 'tblLook']

        def _set_pr(pr, tag, attrs):
            for _o in pr.findall(wt(tag)):
                pr.remove(_o)
            _el = etree.Element(wt(tag))
            for _k, _v in attrs.items():
                _el.set(f'{{{W}}}{_k}', _v)
            _rank = _TBLPR_ORDER.index(tag) if tag in _TBLPR_ORDER else 99
            _idx = len(list(pr))
            for _i, _ch in enumerate(pr):
                _ct = _ch.tag.replace(f'{{{W}}}', '')
                if (_TBLPR_ORDER.index(_ct) if _ct in _TBLPR_ORDER else 99) > _rank:
                    _idx = _i
                    break
            pr.insert(_idx, _el)

        for _t in body.findall(wt('tbl')):
            _txt = ''.join(x.text or '' for x in _t.iter(wt('t')))
            if '(SET 4)' not in _txt and '(SET 5)' not in _txt:
                continue
            _pr = _t.find(wt('tblPr'))
            if _pr is None:
                continue
            # Skaluj kolumny do szerokości tabeli wyników str.1 (8550 dxa), żeby PRAWA
            # krawędź str.2 pokrywała się ze str.1 (str.2 ma mniej setów = naturalnie węższa).
            _SCORE_W = 8550
            _grid = _t.find(wt('tblGrid'))
            _cur = sum(int(c.get(f'{{{W}}}w', '0')) for c in _grid.findall(wt('gridCol'))) if _grid is not None else 0
            if _cur:
                _f = _SCORE_W / _cur
                for _c in _grid.findall(wt('gridCol')):
                    _c.set(f'{{{W}}}w', str(int(int(_c.get(f'{{{W}}}w', '0')) * _f)))
                for _tr in _t.findall(wt('tr')):
                    for _tc in _tr.findall(wt('tc')):
                        _tcpr = _tc.find(wt('tcPr'))
                        _tcw = _tcpr.find(wt('tcW')) if _tcpr is not None else None
                        if _tcw is not None and _tcw.get(f'{{{W}}}w'):
                            _tcw.set(f'{{{W}}}w', str(int(int(_tcw.get(f'{{{W}}}w')) * _f)))
                _newsum = sum(int(c.get(f'{{{W}}}w', '0')) for c in _grid.findall(wt('gridCol')))
                _set_pr(_pr, 'tblW', {'w': str(_newsum), 'type': 'dxa'})
            # KLUCZ: jawne jc=left — bez tego LibreOffice centruje tabelę ignorując tblInd.
            _set_pr(_pr, 'jc', {'val': 'left'})
            _set_pr(_pr, 'tblInd', {'w': '715', 'type': 'dxa'})
            _set_pr(_pr, 'tblLayout', {'type': 'fixed'})

    # ── IND_Bo7: wyrównaj tabelę wyników str.2 ((SET 5)/(SET 6)/(SET 7)) do tabeli
    # str.1 (SET 1-4). W szablonie usera str.2 ma osobne wcięcie (tblInd=355) i jest
    # naturalnie węższa (3 sety vs 4) → lewa krawędź przesunięta w prawo, prawa za
    # krótka względem str.1. Analogicznie do IND_Bo5, ale geometrię str.1 czytamy
    # DYNAMICZNIE z szablonu (Bo7 ma inne wymiary niż Bo5).
    if template_type == 'IND_Bo7':
        _TBLPR_ORDER = ['tblStyle', 'tblpPr', 'tblOverlap', 'bidiVisual',
                        'tblStyleRowBandSize', 'tblStyleColBandSize', 'tblW', 'jc',
                        'tblCellSpacing', 'tblInd', 'tblBorders', 'shd', 'tblLayout',
                        'tblCellMar', 'tblLook']

        def _set_pr7(pr, tag, attrs):
            for _o in pr.findall(wt(tag)):
                pr.remove(_o)
            _el = etree.Element(wt(tag))
            for _k, _v in attrs.items():
                _el.set(f'{{{W}}}{_k}', _v)
            _rank = _TBLPR_ORDER.index(tag) if tag in _TBLPR_ORDER else 99
            _idx = len(list(pr))
            for _i, _ch in enumerate(pr):
                _ct = _ch.tag.replace(f'{{{W}}}', '')
                if (_TBLPR_ORDER.index(_ct) if _ct in _TBLPR_ORDER else 99) > _rank:
                    _idx = _i
                    break
            pr.insert(_idx, _el)

        _all_t = body.findall(wt('tbl'))
        # Tabela WYNIKÓW str.1 = zawiera 'IMIONA' (nagłówek lewej kolumny) i NIE '(SET'
        # (nawias = str.2). UWAGA: tabele nagłówkowe zawierają 'PktSET 1' (→ podłańcuch
        # 'SET 1'), więc NIE filtrujemy po 'SET 1' — bralibyśmy nagłówek (szerszy, inny ind).
        _p1 = None
        for _t in _all_t:
            _txt = ''.join(x.text or '' for x in _t.iter(wt('t')))
            if 'IMIONA' in _txt.replace(' ', '') and '(SET' not in _txt:
                _p1 = _t
                break
        _target_w = 10710   # fallback = znana szerokość str.1 z szablonu
        if _p1 is not None:
            _g1 = _p1.find(wt('tblGrid'))
            if _g1 is not None:
                _s1 = sum(int(c.get(f'{{{W}}}w', '0')) for c in _g1.findall(wt('gridCol')))
                if _s1:
                    _target_w = _s1

        # KROK 1: skaluj WSZYSTKIE pozostałe tabele do szerokości tabeli wyników
        # str.1 (_target_w = 10710):
        #  • str.2 wyniki (SET 5/6/7) — naturalnie węższa (3 sety vs 4).
        #  • nagłówki str.1 i str.2 — naturalnie SZERSZE (11475 vs 10710), bo szablon
        #    ma negatywny tblInd kompensujący różnicę w Wordzie. Po jc=center
        #    (KROK 2) wystawały symetrycznie poza wyniki na obu stronach — user
        #    chciał, żeby krawędzie się pokrywały. Po skalowaniu obie strony =
        #    nagłówek i wyniki tej samej szerokości, wszystko wyrównane.
        for _t in _all_t:
            if _t is _p1:   # str.1 wyniki = źródło prawdy, nie ruszamy
                continue
            _pr = _t.find(wt('tblPr'))
            if _pr is None:
                continue
            _grid = _t.find(wt('tblGrid'))
            _cur = sum(int(c.get(f'{{{W}}}w', '0')) for c in _grid.findall(wt('gridCol'))) if _grid is not None else 0
            if _cur:
                _f = _target_w / _cur
                for _c in _grid.findall(wt('gridCol')):
                    _c.set(f'{{{W}}}w', str(int(int(_c.get(f'{{{W}}}w', '0')) * _f)))
                for _tr in _t.findall(wt('tr')):
                    for _tc in _tr.findall(wt('tc')):
                        _tcpr = _tc.find(wt('tcPr'))
                        _tcw = _tcpr.find(wt('tcW')) if _tcpr is not None else None
                        if _tcw is not None and _tcw.get(f'{{{W}}}w'):
                            _tcw.set(f'{{{W}}}w', str(int(int(_tcw.get(f'{{{W}}}w')) * _f)))
                _newsum = sum(int(c.get(f'{{{W}}}w', '0')) for c in _grid.findall(wt('gridCol')))
                _set_pr7(_pr, 'tblW', {'w': str(_newsum), 'type': 'dxa'})

        # KROK 1b: rozszerz nagłówek (t0/t2) do PRAWEJ krawędzi wyników.
        # PROBLEM: w nagłówku wiersz Tor/Godz/Mecz (r0) pokrywa wszystkie 19 kolumn,
        # ale wiersze Pkt SET / nazwiska (r1-r3) tylko 18 — ostatnia („sieroca")
        # kolumna ~765 dxa jest pusta. LibreOffice CENTRUJE tabelę po PEŁNEJ
        # szerokości (z sierocą kolumną), ale RYSUJE wiersze Pkt SET tylko do
        # pokrytych kolumn → prawa krawędź nagłówka kończy się przed wynikami
        # (user: „rozszerz górną tabelkę do żółtego prostokąta"). Próby z gridSpan
        # autofitowały. ROZWIĄZANIE: przenosimy szerokość sierocej kolumny do
        # POPRZEDNiej (ostatniej pokrywanej przez Pkt SET) i zerujemy sierocą do 1
        # dxa. Wtedy wiersze Pkt SET renderują pełną szerokość = jak r0 = jak wyniki,
        # a całkowita szer. tabeli bez zmian (centrowanie nienaruszone). Nazwiska
        # są w tc[0] (lewa strona) — nietknięte.
        for _t in _all_t:
            _ttxt = ''.join((x.text or '') for x in _t.iter(wt('t')))
            if 'Podpis' not in _ttxt:   # tylko tabele nagłówkowe
                continue
            _grid = _t.find(wt('tblGrid'))
            if _grid is None:
                continue
            _gcols = _grid.findall(wt('gridCol'))
            _N = len(_gcols)
            _min_cov = _N
            for _tr in _t.findall(wt('tr')):
                _cov = 0
                for _tc in _tr.findall(wt('tc')):
                    _gs = _tc.find(f'{wt("tcPr")}/{wt("gridSpan")}')
                    _cov += int(_gs.get(f'{{{W}}}val')) if _gs is not None else 1
                _min_cov = min(_min_cov, _cov)
            if _min_cov >= _N or _min_cov < 1:
                continue
            # Sieroce kolumny = ostatnie (N - min_cov). Ich szer. → kolumna min_cov-1.
            _orphan_w = sum(int(_gcols[i].get(f'{{{W}}}w', '0')) for i in range(_min_cov, _N))
            _keep = _gcols[_min_cov - 1]
            _keep.set(f'{{{W}}}w', str(int(_keep.get(f'{{{W}}}w', '0')) + _orphan_w))
            for i in range(_min_cov, _N):
                _gcols[i].set(f'{{{W}}}w', '1')
            # tcW ostatniej komórki w wierszach o min pokryciu += orphan (żeby tcW
            # zgadzało się z poszerzoną kolumną — LO/Word czytają tcW per komórka).
            for _tr in _t.findall(wt('tr')):
                _cells = _tr.findall(wt('tc'))
                _cov = 0
                for _tc in _cells:
                    _gs = _tc.find(f'{wt("tcPr")}/{wt("gridSpan")}')
                    _cov += int(_gs.get(f'{{{W}}}val')) if _gs is not None else 1
                if _cov != _min_cov or not _cells:
                    continue
                _ltcpr = _cells[-1].find(wt('tcPr'))
                _ltcw = _ltcpr.find(wt('tcW')) if _ltcpr is not None else None
                if _ltcw is not None and _ltcw.get(f'{{{W}}}w'):
                    _ltcw.set(f'{{{W}}}w', str(int(_ltcw.get(f'{{{W}}}w')) + _orphan_w))

        # KROK 2: WYŚRODKOWANIE wszystkich tabel (nagłówki + wyniki, obie strony).
        # Szablon używa UJEMNEGO tblInd (-720/-725) żeby tabela szersza niż obszar
        # tekstu „wystawała" symetrycznie w marginesy — Word liczy tblInd od marginesu
        # tekstu i renderuje OK. Ale LibreOffice (silnik docx→pdf) liczy tblInd od
        # krawędzi STRONY i przycina ujemny do 0 → tabela dosuwa się do lewej krawędzi
        # (user: „siada na lewej krawędzi"). `jc=center` jest niezależne od układu
        # odniesienia → centruje tak samo w Word i w LO. Usuwamy tblInd (zostawiony
        # ujemny i tak by przesuwał). Po fixie: równe marginesy L/R w PDF, obie strony
        # wyrównane (str.2 wyskalowana w KROKU 1 do szer. str.1).
        for _t in _all_t:
            _pr = _t.find(wt('tblPr'))
            if _pr is None:
                continue
            _set_pr7(_pr, 'jc', {'val': 'center'})
            for _o in _pr.findall(wt('tblInd')):
                _pr.remove(_o)
            _set_pr7(_pr, 'tblLayout', {'type': 'fixed'})

        # Kropka po „Pkt" w etykietach nagłówka (user: „po Pkt dodaj kropki").
        # Szablon ma runy split: „Pkt" + <w:br/> + „SET N". Bottom-left „PKT" (suma,
        # wielkie litery) NIE jest łapane (case-sensitive). Font już wymuszony wyżej
        # przez NUCLEAR-SET (match po podłańcuchu 'Pkt' — kropka go nie psuje).
        for _r in body.iter(wt('r')):
            _ts = _r.findall(wt('t'))
            if len(_ts) == 1 and (_ts[0].text or '').strip() == 'Pkt':
                _ts[0].text = (_ts[0].text or '').replace('Pkt', 'Pkt.')

    # ── DWÓJKA Bo3/Bo5: centruj tabele (ten sam bug Word/LO co IND Bo7).
    # Szablon ma ujemny tblInd (-540 nagłówek, -1530 wyniki) — w Wordzie tabele
    # wynikowe wystają symetrycznie poza nagłówek (user: „ładnie powyśrodkowywane"),
    # w LibreOffice (PDF) ujemny tblInd jest clampowany i tabele dosuwają się
    # asymetrycznie. Fix: jc=center + usunięcie tblInd. Po fixie wyniki (szersze)
    # wystają symetrycznie poza nagłówek na obu stronach (Bo5 = 4 tabele).
    if template_type in ('DWOJKA_Bo3', 'DWOJKA_Bo5'):
        _TBLPR_ORDER_D = ['tblStyle', 'tblpPr', 'tblOverlap', 'bidiVisual',
                          'tblStyleRowBandSize', 'tblStyleColBandSize', 'tblW', 'jc',
                          'tblCellSpacing', 'tblInd', 'tblBorders', 'shd', 'tblLayout',
                          'tblCellMar', 'tblLook']
        def _set_pr_d(pr, tag, attrs):
            for _o in pr.findall(wt(tag)): pr.remove(_o)
            _el = etree.Element(wt(tag))
            for _k, _v in attrs.items(): _el.set(f'{{{W}}}{_k}', _v)
            _rank = _TBLPR_ORDER_D.index(tag) if tag in _TBLPR_ORDER_D else 99
            _idx = len(list(pr))
            for _i, _ch in enumerate(pr):
                _ct = _ch.tag.replace(f'{{{W}}}', '')
                if (_TBLPR_ORDER_D.index(_ct) if _ct in _TBLPR_ORDER_D else 99) > _rank:
                    _idx = _i; break
            pr.insert(_idx, _el)
        for _t in body.findall(wt('tbl')):
            _pr = _t.find(wt('tblPr'))
            if _pr is None: continue
            _set_pr_d(_pr, 'jc', {'val': 'center'})
            for _o in _pr.findall(wt('tblInd')): _pr.remove(_o)

    # ── Sectpr i template
    sectPr = body.find(wt('sectPr'))
    template_elements = [copy.deepcopy(el) for el in body if el.tag != wt('sectPr')]
    while template_elements and template_elements[-1].tag == wt('p'):
        last = template_elements[-1]
        ts = last.findall(f'.//{wt("t")}')
        if not ts or not any((t.text or '').strip() for t in ts):
            template_elements.pop()
        else:
            break

    # ── Media (QR + grafiki)
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
            qr_rid_info = (rid, 2.4, 2.4)  # cm - większy QR

    logo_rids = {}

    # Logo PFM z lokalnego asseta (domyślnie włączone)
    if include_pfm_logo:
        pfm_path = os.path.join(os.path.dirname(__file__), 'assets_pfm_logo.png')
        if os.path.exists(pfm_path):
            with open(pfm_path, 'rb') as fp:
                pfm_bytes = fp.read()
            from PIL import Image as PILImage
            rid = f'rId{next_rid[0]}'; next_rid[0] += 1
            fname = 'media/pfm_logo.png'
            pil = PILImage.open(io.BytesIO(pfm_bytes)).convert('RGBA')
            buf = io.BytesIO(); pil.save(buf, format='PNG')
            media_files[fname] = buf.getvalue()
            ratio = pil.width / pil.height
            target_w_cm = 3.5
            target_h_cm = target_w_cm / ratio
            if target_h_cm > 2.5:
                target_h_cm = 2.5
                target_w_cm = target_h_cm * ratio
            rel = etree.SubElement(rels_root, f'{{{REL}}}Relationship')
            rel.set('Id', rid); rel.set('Type', REL_IMG); rel.set('Target', fname)
            logo_rids['pfm'] = (rid, target_w_cm, target_h_cm)

    if logos:
        from PIL import Image as PILImage
        for key, img_bytes in logos.items():
            if not img_bytes: continue
            rid = f'rId{next_rid[0]}'; next_rid[0] += 1
            fname = f'media/{key}.png'
            pil = PILImage.open(io.BytesIO(img_bytes)).convert('RGBA')
            buf = io.BytesIO(); pil.save(buf, format='PNG')
            media_files[fname] = buf.getvalue()
            ratio = pil.width / pil.height
            target_w_cm = 3.5
            target_h_cm = target_w_cm / ratio
            if target_h_cm > 2.5:
                target_h_cm = 2.5
                target_w_cm = target_h_cm * ratio
            rel = etree.SubElement(rels_root, f'{{{REL}}}Relationship')
            rel.set('Id', rid); rel.set('Type', REL_IMG); rel.set('Target', fname)
            logo_rids[key] = (rid, target_w_cm, target_h_cm)

    for el in list(body): body.remove(el)

    # Szerokość lewej kolumny tabeli wynikowej (kotwica grafik) — z REALNEGO szablonu.
    # Stare szablony: 2970 dxa (5.24 cm). Nowe IND: 2694 dxa (4.75 cm). Liczymy
    # dynamicznie, żeby grafiki centrowały się w faktycznej kolumnie i nie wystawały
    # na tabelę. Fallback 4.75 (nowy IND). Tylko dla IND grupowej (TROJKA/CZWORKA
    # mają własne stałe col_width_cm).
    anchor_col_cm = 4.75
    for _t in template_elements:
        if _t.tag == wt('tbl') and len(_t.findall(wt('tr'))) > 10:
            _g = _t.find(wt('tblGrid'))
            _gcs = _g.findall(wt('gridCol')) if _g is not None else []
            if _gcs:
                _w0 = _gcs[0].get(f'{{{W}}}w')
                if _w0:
                    anchor_col_cm = int(_w0) / 567.0
            break

    # ── Generuj protokoły
    total_matches = sum(len(g[1]) for g in sheets_data)
    done = 0
    first = True
    for group_entry in sheets_data:
        # sheets_data może być listą (group_name, matches) lub
        # (group_name, matches, phase_text_override) — to ostatnie używamy gdy
        # generujemy MULTI-PHASE w 1 dokumencie i każda grupa ma osobną fazę.
        if len(group_entry) == 3:
            group_name, matches, group_phase_override = group_entry
        else:
            group_name, matches = group_entry
            group_phase_override = None
        # Phase text efektywny dla tego protokołu: override per grupa lub globalny
        effective_phase_text = group_phase_override if group_phase_override is not None else tournament_phase_text
        # "Mecz #" ma sens tylko gdy w fazie jest WIELE meczów (np. 8× 1/16).
        # Dla pojedynczego meczu (Finał, Mecz o 3. miejsce) numer jest zbędny → ukrywamy.
        hide_mecz_num = len(matches) <= 1
        for match in matches:
            _is_first_match = first
            # DWÓJKA Bo7 (landscape, score s.2 wypełnia stronę): standalone
            # page-break-para tworzy PUSTĄ stronę między meczami (break ląduje na
            # już-pełnej stronie → przeskok o 2). Zamiast tego damy pageBreakBefore
            # na akapicie nagłówka meczu (hp) — patrz niżej. Inne typy: bez zmian.
            if not first and template_type != 'DWOJKA_Bo7':
                body.append(_make_page_break_para())
            first = False
            if hide_mecz_num and match.get('mecz'):
                match = dict(match)
                match['mecz'] = ''
            done += 1
            if progress_cb:
                try:
                    label_parts = [str(group_name)]
                    if match.get('mecz'): label_parts.append(f"mecz {match['mecz']}")
                    if match.get('tor'):  label_parts.append(f"tor {match['tor']}")
                    progress_cb(done, total_matches, ' · '.join(label_parts))
                except Exception:
                    pass

            cloned = [copy.deepcopy(el) for el in template_elements]

            # Wstaw nazwę turnieju + datę + fazę jako paragraf w prawym górnym rogu
            # (przed pierwszą tabelą, wyrównany do prawej, małą czcionką)
            if tournament_name or tournament_date or effective_phase_text:
                header_parts = []
                if tournament_name:
                    header_parts.append(tournament_name.strip())
                if tournament_date:
                    header_parts.append(tournament_date.strip())
                if effective_phase_text:
                    header_parts.append(effective_phase_text.strip())
                header_text = ' · '.join(header_parts)

                hp = etree.Element(wt('p'))
                hpPr = etree.SubElement(hp, wt('pPr'))
                # DWÓJKA Bo7: kolejne mecze startują na nowej stronie przez
                # pageBreakBefore na nagłówku (zamiast standalone break-para,
                # który dawał pustą stronę — patrz wyżej).
                if template_type == 'DWOJKA_Bo7' and not _is_first_match:
                    etree.SubElement(hpPr, wt('pageBreakBefore'))
                hjc = etree.SubElement(hpPr, wt('jc'))
                hjc.set(f'{{{W}}}val', 'right')
                # IND_Bo3: dodaj right indent (~1.4 cm) by header nie wystawał za prawą krawędź tabeli.
                # pgMar.right=1440 dxa, tabela szerokość 9360 dxa, więc tabela kończy się
                # na pgMar.left+tableW = 1440+9360 = 10800 dxa. Right margin = 11906-1440 = 10466.
                # Header bez ind: kończy na 10466 (= right margin). Tabela: 10800. Header wystaje 334 dxa za tabelę.
                # Czekaj — odwrotnie: tabela kończy na 10800, header na 10466. Header KRÓTSZY niż tabela.
                # No wait — patrząc na screenshot user, header wystaje na prawo OD tabeli. Coś się nie zgadza
                # z moimi obliczeniami. Daję ind right=800 dxa (~1.4 cm) jako empiryczny fix.
                if template_type in ('IND_Bo3', 'IND_Bo5', 'IND_Bo7'):
                    hind = etree.SubElement(hpPr, wt('ind'))
                    hind.set(f'{{{W}}}right', '800')
                hsp = etree.SubElement(hpPr, wt('spacing'))
                hsp.set(f'{{{W}}}before', '0')
                hsp.set(f'{{{W}}}after', '60')
                hr = etree.SubElement(hp, wt('r'))
                hrPr = etree.SubElement(hr, wt('rPr'))
                hfonts = etree.SubElement(hrPr, wt('rFonts'))
                for a in ('ascii', 'hAnsi', 'eastAsia', 'cs'):
                    hfonts.set(f'{{{W}}}{a}', 'Aptos')
                for tag in ('i', 'iCs'):
                    etree.SubElement(hrPr, wt(tag)).set(f'{{{W}}}val', '1')
                for tag in ('sz', 'szCs'):
                    etree.SubElement(hrPr, wt(tag)).set(f'{{{W}}}val', '16')
                hcolor = etree.SubElement(hrPr, wt('color'))
                hcolor.set(f'{{{W}}}val', '666666')
                ht = etree.SubElement(hr, wt('t'))
                ht.text = header_text
                cloned.insert(0, hp)
                
                # ── BO5: header też na drugiej stronie (przed T3) ──
                # Bo5 jest 2-stronicowe per mecz. Header w prawym górnym
                # rogu chcemy na obu stronach, więc klonujemy paragraf
                # i wstawiamy przed tabelą strony 2.
                # TROJKA_Bo5: str.2 zaczyna się od tabeli idx 2.
                # IND_Bo5 POMINIĘTE: nowy szablon ma nagłówek str.2 (Tor/Godz/Runda)
                # wbudowany w T1, a podział strony jest WEWNĄTRZ T1 — wstawianie hp_page2
                # przed T2 trafiłoby w środek str.2 (między nagłówek a SET 4/5).
                # Bo5: header też na drugiej stronie. Włączone dla IND_Bo5
                # również — szablon ma 4 tabele (str.1 header, str.1 score,
                # str.2 header, str.2 score), więc wstawiamy hp_page2 przed
                # tbls[2] (header str.2) — zgodnie z TROJKA_Bo5 / CZWORKA_Bo5.
                # DWOJKA_Bo5 dodane na zapas — gdy szablon powstanie i będzie miał
                # układ 4-tabelowy (header/score str.1 + header/score str.2),
                # nagłówek str.2 zadziała automatycznie. Guard `len >= 3` chroni
                # przed wstawieniem gdy szablon ma inną strukturę.
                if template_type in ('TROJKA_Bo5', 'CZWORKA_Bo5', 'IND_Bo5', 'IND_Bo7', 'DWOJKA_Bo5', 'DWOJKA_Bo7'):
                    cloned_tbls = [el for el in cloned if el.tag == wt('tbl')]
                    if len(cloned_tbls) >= 3:
                        t3_in_cloned = cloned_tbls[2]
                        from copy import deepcopy
                        hp_page2 = deepcopy(hp)
                        # pageBreakBefore — gwarantuje że str.2 zaczyna się na nowej stronie.
                        # IND_Bo5 ma już wbudowany page break w paragrafie; dodanie
                        # pageBreakBefore jest idempotentne (LO i Word obsłużą poprawnie).
                        hpPr2 = hp_page2.find(wt('pPr'))
                        if hpPr2 is None:
                            hpPr2 = etree.Element(wt('pPr'))
                            hp_page2.insert(0, hpPr2)
                        if hpPr2.find(wt('pageBreakBefore')) is None:
                            etree.SubElement(hpPr2, wt('pageBreakBefore'))
                        # DWÓJKA Bo7: szablon ma między score s.1 (t1) a header s.2
                        # (t2) DWA puste akapity + osobny akapit z jawnym page-break.
                        # Score s.1 (22 wiersze landscape) wypełnia stronę 1 prawie
                        # do końca → te akapity SPADAJĄ na nową stronę PRZED jawnym
                        # breakiem → pusta strona (1 mecz = 3 strony zamiast 2).
                        # Fix: usuwamy puste/break-only akapity między t1 a t2 i
                        # polegamy na pageBreakBefore (powyżej) na hp_page2.
                        if template_type == 'DWOJKA_Bo7':
                            _t1c = cloned_tbls[1]
                            _i1 = list(cloned).index(_t1c)
                            _i2 = list(cloned).index(t3_in_cloned)
                            for _el in list(cloned)[_i1 + 1:_i2]:
                                if _el.tag != wt('p'):
                                    continue
                                _has_txt = any((x.text or '').strip() for x in _el.iter(wt('t')))
                                if not _has_txt:
                                    cloned.remove(_el)
                        t3_idx = list(cloned).index(t3_in_cloned)
                        cloned.insert(t3_idx, hp_page2)
            _fill_protocol(cloned, match,
                           hide_grupa_mecz=hide_grupa_mecz,
                           phase_label=phase_label,
                           template_type=template_type)

            # Lista elementów do wstawienia w lewym obszarze
            # Bo3/Bo5 nie używają grafik (lewa kolumna jest wąska, ~0.7 cm).
            if template_type in ('TROJKA_Bo3', 'TROJKA_Bo5', 'CZWORKA_Bo3', 'CZWORKA_Bo5', 'DWOJKA_Bo7'):
                order = []
            else:
                order = image_order if image_order else (
                    (['qr'] if qr_rid_info else []) +
                    sorted(logo_rids.keys())
                )

            anchored = []
            czworka_specs = []    # CZWORKA: [(x_cm, rid, w_cm, h_cm, key)] do strip table
            cell_w_cm = anchor_col_cm   # realna szerokość kolumny "Wyniki turnieju"
            cur_y_cm = 0.1        # od góry komórki
            spacing_cm = 0.15

            label_after_qr = False

            for key in order:
                rid_info = None
                if key == 'qr' and qr_rid_info:
                    rid_info = qr_rid_info
                elif key in logo_rids:
                    rid_info = logo_rids[key]
                if rid_info is None:
                    continue
                rid, w_cm, h_cm = rid_info

                # Pozycja własna z image_positions, jeśli jest.
                # Akceptujemy oba zestawy kluczy: 'w'/'h' (z app.py) lub 'width'/'height' (legacy).
                if image_positions and key in image_positions:
                    pos = image_positions[key]
                    x_cm = pos.get('x', (cell_w_cm - w_cm) / 2)
                    y_cm = pos.get('y', cur_y_cm)
                    # User podał szerokość — przyjmujemy ją 1:1 (bez wymuszania aspect ratio
                    # obrazu, bo user może chcieć skalować nieproporcjonalnie).
                    new_w = pos.get('w', pos.get('width'))
                    new_h = pos.get('h', pos.get('height'))
                    if new_w is not None:
                        w_cm = new_w
                    if new_h is not None:
                        h_cm = new_h
                    elif new_w is not None:
                        # Tylko width — zachowaj proporcje (dla QR zachowaj 1:1)
                        if key == 'qr':
                            h_cm = w_cm
                        else:
                            orig_w, orig_h = rid_info[1], rid_info[2]
                            h_cm = w_cm * (orig_h / orig_w)
                else:
                    # Domyślnie jedna pod drugą, wycentrowane
                    x_cm = (cell_w_cm - w_cm) / 2
                    y_cm = cur_y_cm

                cx_emu = int(w_cm * 360000)
                cy_emu = int(h_cm * 360000)
                px_emu = int(x_cm * 360000)
                py_emu = int(y_cm * 360000)

                # ── DEFENSE IN DEPTH: clamp tak, żeby obraz nie wyszedł poza
                # kolumnę kotwicy (R1.tc[0]). Bez tego stare zapisane pozycje
                # albo wartości spoza zakresu w UI mogłyby przelać obraz na
                # tabelę wyników po prawej. Skalujemy proporcjonalnie żeby
                # nie zniekształcić obrazu.
                # Trójka: kolumna 4.76 cm. IND: kolumna 5.24 cm. Kompensacja
                # cellMargin (-0.185 cm) ze strony app.py oznacza że dopuszczalny
                # zakres X to ok. -0.2…(col_width - w).
                col_width_cm = (4.76 if template_type in ('TROJKA', 'TROJKA_Bo3', 'TROJKA_Bo5', 'DWOJKA')
                                else 18.46 if template_type == 'CZWORKA'
                                else anchor_col_cm)
                # Effective right edge (z kompensacją cellMargin po lewej)
                # Image left edge może być nawet -0.2 (kompensacja), wtedy max w to col_width.
                # Ale clampujemy x do >= -0.25 żeby nie wszedł za daleko w lewo.
                if x_cm < -0.25:
                    x_cm = -0.25
                    px_emu = int(x_cm * 360000)
                # Maks. szerokość przy aktualnym X tak, żeby right edge <= col_width
                max_w_at_x = col_width_cm - x_cm
                if w_cm > max_w_at_x and max_w_at_x > 0.3:
                    # Skalowanie zachowujące aspect ratio
                    scale = max_w_at_x / w_cm
                    w_cm = max_w_at_x
                    h_cm = h_cm * scale
                    cx_emu = int(w_cm * 360000)
                    cy_emu = int(h_cm * 360000)

                # CZWORKA: zbieramy specs do strip table (inline images, nie anchored).
                # Anchory w cell context były nieprzewidywalne; strip table to deterministyczne
                # pozycjonowanie przez table cell widths.
                # Inne template'y zachowują behavioral kompatybilność (column anchor).
                if template_type == 'CZWORKA':
                    # Zbierz spec do późniejszego strip table (inline images w tabeli)
                    czworka_specs.append((x_cm, rid, w_cm, h_cm, key))
                else:
                    _px_emu_final = px_emu
                    anchored.append(_make_anchored_image_drawing(
                        rid, cx_emu, cy_emu, py_emu, _px_emu_final,
                        posX_relative_to='column'))

                cur_y_cm = y_cm + h_cm + spacing_cm

                # Po QR — flaga że napis idzie pod QR
                if key == 'qr':
                    label_after_qr = True
                    cur_y_cm += 0.4  # miejsce na napis "Wyniki turnieju"

            # Napis "Wyniki turnieju" - tylko gdy jest QR.
            # Bo3/Bo5 nie używają lewej kolumny na grafiki (zostaje oryginalna),
            # więc całkowicie pomijamy populate_left_area.
            # Czwórka — nowy layout: grafiki w poziomie pod tabelą. Template ma już
            # własny napis "Wyniki turnieju" więc tylko placujemy anchored w R1.tc[0]
            # (kotwica) — pozycje z compute_default_positions dla CZWORKA mają Y~22.5cm
            # co spycha je pod tabelę dzięki layoutInCell=0.
            if template_type in ('TROJKA_Bo3', 'TROJKA_Bo5', 'IND_Bo3', 'IND_Bo5', 'IND_Bo7', 'CZWORKA_Bo3', 'CZWORKA_Bo5'):
                pass  # pucharowa — nie ruszamy lewej kolumny T2 (zostaje IMIONA / oryginalna treść)
            elif template_type == 'CZWORKA':
                # Strip table: inline images w tabeli na pełną szerokość strony.
                # Sortuj specs po X (lewa→prawa), zbuduj tabelę.
                czworka_specs.sort(key=lambda s: s[0])
                strip_specs = [(rid, w, h) for (_, rid, w, h, _) in czworka_specs]
                strip_tbl = _make_czworka_strip_table(strip_specs)
                if strip_tbl is not None:
                    # Wstaw strip table NA KOŃCU cloned (po "Wyniki" body para).
                    # cloned struktura: [tbl1, p, p, p, tbl2_score, p_wyniki].
                    # Wstawiamy po p_wyniki, czyli na końcu.
                    cloned.append(strip_tbl)
            elif qr_rid_info and include_qr:
                # Wyciągnij faktyczną pozycję i wysokość QR z image_positions
                if image_positions and 'qr' in image_positions:
                    qr_pos = image_positions['qr']
                    qr_y = qr_pos.get('y', 0.2)
                    # Akceptuj oba klucze (app.py używa 'h')
                    qr_h_cm = qr_pos.get('h', qr_pos.get('height',
                                qr_pos.get('w', qr_pos.get('width', 2.4))))
                    label_y_cm = qr_y + qr_h_cm + 0.1
                else:
                    label_y_cm = 0.2 + 2.4 + 0.1
                _populate_left_area(cloned, anchored, 'Wyniki turnieju', label_y_cm)
            else:
                # Bez QR - tylko obrazy bez napisu
                _populate_left_area(cloned, anchored, '', 0)

            # DWÓJKA Bo7: usuń KOŃCOWE puste akapity z bloku meczu. Szablon ma po
            # score s.2 puste akapity, które (gdy score s.2 wypełnia stronę) spadają
            # na nową stronę i razem z page-breakiem MIĘDZY meczami (_make_page_break_para)
            # tworzą pustą stronę (2 mecze = 5 stron zamiast 4). Bez nich score s.2
            # kończy stronę, a break czysto przechodzi do kolejnego meczu.
            if template_type == 'DWOJKA_Bo7':
                while cloned and cloned[-1].tag == wt('p') and \
                        not any((x.text or '').strip() for x in cloned[-1].iter(wt('t'))):
                    cloned.pop()

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
