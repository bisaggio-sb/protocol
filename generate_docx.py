"""
generate_docx.py – Generator protokołów meczowych Mölkky.
"""

import io, os, csv, re, copy, zipfile, string
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
        tor=g(col_tor); godz=g(col_godz); z1=g(col_z1); z2=g(col_z2)
        if not _is_valid_match_row(tor,godz,z1,z2): continue
        matches.append({'tor':tor,'godz':godz,'grupa':grupa,
                        'mecz':g(col_mecz),'z1':z1,'z2':z2})
    return matches

def fetch_all_group_sheets(sheet_id):
    """Skanuje wszystkie zakładki Gr. A do Gr. Z. Zwraca tylko te z meczami."""
    gid_map = get_sheet_gids(sheet_id)
    results = []
    for letter in string.ascii_uppercase:   # A..Z, nie tylko A..P
        name = f"Gr. {letter}"
        try:
            rows = fetch_sheet(sheet_id, name, gid_map)
            if rows is None: continue
            matches = parse_group_rows(rows)
            if matches: results.append((name, matches))
        except Exception:
            continue
    return results


# ─────────────────────────────────────────────────────────────────────
# Parser zakładki "Drabinka" - dla meczów fazy pucharowej
# ─────────────────────────────────────────────────────────────────────

# Rozpoznawane nazwy faz (fragment do wyszukania w arkuszu)
PHASE_NAMES = {
    '1/64': '1/64 FINAŁU',
    '1/32': '1/32 FINAŁU',
    '1/16': '1/16 FINAŁU',
    '1/8':  '1/8 FINAŁU',
    '1/4':  '1/4 FINAŁU',
    '1/2':  '1/2 FINAŁU',
    'półfinał': '1/2 FINAŁU',
    'mecz o 3': 'MECZ O 3 MIEJSCE',
    'finał': 'FINAŁ',
}


def parse_drabinka_rows(rows, target_phase=None):
    """
    Parsuje wiersze zakładki Drabinka i zwraca mecze z konkretnej fazy.
    
    Format zakładki:
    - Wiersz nagłówka zawiera: A=("ok" lub puste), B="Tor", C="Grupa",
      D=nazwa fazy z godziną np. "1/32 FINAŁU (13:45)", H="SETY"
    - Co 2 wiersze: para zawodników jednego meczu
      (kolumna Tor jest scalona dla obu wierszy, lub jest tylko w pierwszym)
    
    target_phase: np. "1/32 FINAŁU" - jeśli None, zwraca pierwszą znalezioną fazę.
    
    Zwraca: (phase_name, time, [matches])
        matches: [{tor, godz, grupa, mecz, z1, z2}]
    """
    if not rows:
        return None, None, []
    
    # Szukamy wiersza nagłówka - musi mieć "Tor" w którejś kolumnie
    header_idx = None
    header = []
    for i, row in enumerate(rows):
        norm = [c.strip().lower() for c in row]
        if 'tor' in norm:
            header_idx = i
            header = norm
            break
    
    if header_idx is None:
        return None, None, []
    
    # Indeksy kolumn
    raw_header = rows[header_idx]
    def ci(name):
        try: return header.index(name)
        except ValueError: return None
    
    col_tor = ci('tor')
    col_grupa = ci('grupa')
    
    # Kolumna z nazwą fazy + zawodnikami: szukamy w nagłówku komórki która
    # zawiera nazwę fazy (np. "1/32 finału" lub "FINAŁ" lub "1/2 finału")
    col_player = None
    phase_full_name = None
    phase_time = None
    for i, h in enumerate(raw_header):
        h_lower = h.strip().lower()
        # Sprawdź czy zawiera nazwę fazy
        for phase_key, phase_full in PHASE_NAMES.items():
            if phase_key in h_lower:
                col_player = i
                # Wyciągnij nazwę fazy + godzinę
                # np. "1/32 FINAŁU (13:45)" -> phase="1/32 FINAŁU", time="13:45"
                phase_full_name = phase_full
                m = re.search(r'\((\d{1,2}:\d{2})\)', h)
                if m:
                    phase_time = m.group(1)
                break
        if col_player is not None:
            break
    
    if col_player is None:
        # Fallback: kolumna zawodnika to ta po Grupie lub po Tor
        if col_grupa is not None:
            col_player = col_grupa + 1
        elif col_tor is not None:
            col_player = col_tor + 2
        else:
            col_player = 3  # default kol D
    
    # Filtr: czy ta faza nas interesuje?
    if target_phase and phase_full_name:
        # target_phase np. "Pucharowa 1/32" lub "1/32 FINAŁU"
        # Szukamy kluczowego fragmentu jak "1/32"
        match = False
        for key in PHASE_NAMES:
            if key in target_phase.lower() and key in phase_full_name.lower():
                match = True
                break
        if not match and target_phase.lower() not in (phase_full_name or '').lower():
            return phase_full_name, phase_time, []
    
    # Czytamy pary wierszy. Zatrzymujemy się gdy:
    # - Pojawia się nowy nagłówek (kolejna faza, "MECZE O MIEJSCA" itp.)
    # - Pojawiają się 2+ kolejne puste wiersze
    # - Tekst w kolumnie zawodnika to nagłówek (zawiera "FINAŁ", "MIEJSC", "FINAŁOWY")
    matches = []
    data_rows = rows[header_idx + 1:]
    
    # Słowa-stop: pojawienie się któregokolwiek w komórce = koniec sekcji
    stop_keywords = ['miejsc', 'finał', 'finałow', 'mecz o', 'play-off',
                     'ranking', 'klasyfikacja', '1/64', '1/32', '1/16',
                     '1/8', '1/4', '1/2']
    
    def is_section_header(row):
        """Czy ten wiersz wygląda jak nagłówek nowej sekcji (a nie dane)?"""
        if not row: return False
        joined = ' '.join(c.strip().lower() for c in row if c)
        # Jeśli zawiera słowo-stop i nie zaczyna się od liczby (jak np. "16 Jérémy")
        for kw in stop_keywords:
            if kw in joined:
                # Sprawdź czy to faktycznie nagłówek (np. "MECZE O MIEJSCA 33-64"
                # ma "miejsc" - ale wiersz danych z "Robert Bąścik" nie ma).
                # Wyklucz wiersze gdzie kolumna zawodnika to normalne imię.
                player_text = row[col_player].strip() if col_player < len(row) else ''
                # Jeśli player_text zawiera słowo-stop = to nagłówek
                if kw in player_text.lower():
                    return True
                # Jeśli słowo-stop jest w jakiejś innej komórce z dużymi literami
                # i player_text jest pusty = nagłówek
                if not player_text and kw in joined:
                    return True
        return False
    
    i = 0
    match_num = 1
    empty_streak = 0
    last_known_tor = ''  # cache dla scalonych komórek tora
    while i < len(data_rows):
        row1 = data_rows[i]
        # Pomijamy puste wiersze
        if not any(c.strip() for c in row1):
            empty_streak += 1
            if empty_streak >= 2:
                # 2+ puste z rzędu = koniec sekcji
                break
            i += 1
            continue
        empty_streak = 0
        
        # Sprawdź czy to nie jest nagłówek nowej sekcji
        if is_section_header(row1):
            break
        
        # Następny wiersz to drugi zawodnik
        row2 = data_rows[i+1] if i+1 < len(data_rows) else []
        
        def g(row, c):
            if c is None or c >= len(row): return ''
            return row[c].strip()
        
        # Tor: hierarchia źródeł
        # 1) row1.col_tor (najczęściej tu jest gdy pierwsza komórka scalona)
        # 2) row2.col_tor (gdy Google export zwraca tor w drugim wierszu)
        # 3) ostatni znany tor + 1 (gdy scalona obejmuje wiele wierszy a CSV zwraca tylko w 1.)
        # 4) numer meczu (fallback)
        tor_raw = g(row1, col_tor) or g(row2, col_tor)
        if tor_raw and tor_raw.strip().isdigit():
            tor = tor_raw.strip()
            last_known_tor = tor
        elif last_known_tor and last_known_tor.isdigit():
            # Inkrementuj ostatni znany — typowy wzorzec dla drabinki gdzie tory rosną
            tor = str(int(last_known_tor) + 1)
            last_known_tor = tor
        else:
            # Fallback na numer meczu (drabinka zwykle ma tory 1,2,3,... wg kolejności)
            tor = str(match_num)
            last_known_tor = tor
        
        z1 = g(row1, col_player)
        z2 = g(row2, col_player)
        
        # Walidacja: oba imiona muszą być nie-puste i mieć sens
        if z1 and z2 and len(z1) >= 3 and len(z2) >= 3:
            # Imię/Nazwisko musi mieć choć jedną wielką literę
            # I nie zawierać słów-stop (zabezpieczenie)
            z1_lower = z1.lower()
            z2_lower = z2.lower()
            has_stop = any(kw in z1_lower or kw in z2_lower for kw in stop_keywords)
            if not has_stop and any(c.isupper() for c in z1) and any(c.isupper() for c in z2):
                matches.append({
                    'tor': tor,
                    'godz': phase_time or '',
                    'grupa': '',  # w drabince nie ma grupy
                    'mecz': str(match_num),
                    'z1': z1,
                    'z2': z2,
                })
                match_num += 1
        
        i += 2  # przeskakujemy 2 wiersze (mecz)
    
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


# ─────────────────────────────────────────────────────────────────────


def get_sheet_names_debug(sheet_id):
    """Czytelny debug: liczba grup, liczba meczów per grupa, total + info o Drabince."""
    info = []
    gid_map = get_sheet_gids(sheet_id)
    found_groups = []
    total_matches = 0
    for letter in string.ascii_uppercase:
        name = f"Gr. {letter}"
        rows = fetch_via_gviz(sheet_id, name)
        if not rows and name in gid_map:
            rows = fetch_via_gid(sheet_id, gid_map[name])
        if not rows:
            rows = fetch_via_export_name(sheet_id, name)
        if not rows:
            continue
        matches = parse_group_rows(rows)
        if not matches:
            continue
        found_groups.append((name, len(matches)))
        total_matches += len(matches)

    # Sprawdź też zakładkę Drabinka
    drabinka_phase = None
    drabinka_count = 0
    for tab_name in ('Drabinka', 'drabinka', 'DRABINKA'):
        try:
            rows = fetch_sheet(sheet_id, tab_name, gid_map)
            if rows:
                phase, _, m = parse_drabinka_rows(rows)
                if m:
                    drabinka_phase = phase
                    drabinka_count = len(m)
                    break
        except Exception:
            continue

    if not found_groups and not drabinka_phase:
        info.append("❌ Nie znaleziono żadnych grup z meczami ani zakładki Drabinka.")
        info.append("   Sprawdź czy arkusz jest publiczny.")
        return info

    if found_groups:
        info.append(f"✅ Faza grupowa: {len(found_groups)} grup, {total_matches} meczów")
        for name, count in found_groups:
            info.append(f"  • {name}: {count} meczów")
    
    if drabinka_phase:
        info.append("")
        info.append(f"✅ Drabinka: faza {drabinka_phase}, {drabinka_count} meczów")
    
    return info


# ─── XML helpers: tekst w komórce ─────────────────────────────────────────────

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
    for a in ('ascii','hAnsi','eastAsia','cs'):
        fonts.set(f'{{{W}}}{a}', 'Aptos')
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


def _make_page_break_para():
    p = etree.Element(wt('p'))
    pPr = etree.SubElement(p, wt('pPr'))
    sp = etree.SubElement(pPr, wt('spacing'))
    sp.set(f'{{{W}}}before','0'); sp.set(f'{{{W}}}after','0')
    r = etree.SubElement(p, wt('r'))
    br = etree.SubElement(r, wt('br'))
    br.set(f'{{{W}}}type','page')
    return p


def _fill_protocol(elements, match, hide_grupa_mecz=False, phase_label=None):
    tbls = [el for el in elements if el.tag == wt('tbl')]
    if not tbls: return
    rows = tbls[0].findall(wt('tr'))
    if len(rows) > 0:
        tcs = rows[0].findall(wt('tc'))
        if len(tcs) > 1: _set_cell_value(tcs[1], match.get('tor',''),  size=28)
        if len(tcs) > 3: _set_cell_value(tcs[3], match.get('godz',''), size=28)
        if hide_grupa_mecz:
            # Faza pucharowa: zamień etykietę "Grupa" → "Faza" zachowując format,
            # wartość fazy (np. "1/32") wstaw jak inne wartości (Aptos size 28 bold).
            # Numer meczu # nie jest dostępny w arkuszu Drabinka — ukrywamy
            # zarówno etykietę "Mecz #" jak i wartość.
            if len(tcs) > 4: _set_cell_label(tcs[4], 'Faza')
            if len(tcs) > 5: _set_cell_value(tcs[5], phase_label or '', size=28)
            if len(tcs) > 6: _set_cell_label(tcs[6], '')  # ukryj etykietę 'Mecz #'
            if len(tcs) > 7: _set_cell_value(tcs[7], '', size=28)  # ukryj numer
        else:
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


def _make_anchored_image_drawing(rel_id, cx_emu, cy_emu, posY_emu, posX_emu=0):
    """Pływający obraz z layoutInCell=0 — KLUCZOWE: pozwala obrazom wystawać
    poza komórkę kotwicy (w dół, do końca tabeli wynikowej). Bez tego obrazy
    byłyby ucinane na granicy komórki R1 'Wyniki turnieju' (wysokość ~2.14 cm)."""
    uid = _next_uid()
    return etree.fromstring(f'''<w:drawing xmlns:w="{W}"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
  <wp:anchor distT="0" distB="0" distL="114300" distR="114300"
             simplePos="0" relativeHeight="{251659264 + uid}"
             behindDoc="0" locked="0" layoutInCell="0" allowOverlap="1"
             wp14:anchorId="0000{uid:04X}" wp14:editId="0000{uid:04X}">
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
                         tournament_date=None, include_qr=False,
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
        include_qr=include_qr and bool(sheets_url),  # QR tylko jeśli jest URL
        include_pfm_logo=include_pfm_logo,
        image_order=image_order, image_positions=image_positions,
        template_type=template_type
    )


def build_document(sheet_id, sheets_url, sheets_data, logos=None,
                   tournament_name=None, tournament_date=None,
                   include_qr=True, include_pfm_logo=True,
                   image_order=None, image_positions=None,
                   hide_grupa_mecz=False, phase_label=None,
                   template_type='IND'):
    """
    `tournament_date`: string (np. "10.05.2026") wyświetlany w nagłówku obok nazwy.
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

    template_files = {
        'IND': 'Grupa_IND.docx',
        'TROJKA': 'Grupa_TROJKA.docx',
        # CZWORKA, DRUZYNA - nie zaimplementowane jeszcze (czeka na szablony)
    }
    tpl_filename = template_files.get(template_type, 'Grupa_IND.docx')
    tpl_path = os.path.join(os.path.dirname(__file__), tpl_filename)
    with open(tpl_path, 'rb') as f:
        tpl_bytes = f.read()

    zin = zipfile.ZipFile(io.BytesIO(tpl_bytes))
    doc_xml = zin.read('word/document.xml')
    doc_root = etree.fromstring(doc_xml)
    body = doc_root.find(wt('body'))

    # ── Marginesy
    # Indywidualny: 720 dxa = 1.27 cm wszędzie (oryginał ma więcej, ścieśnimy żeby zostało
    # więcej miejsca dla obrazów po lewej).
    # Trójka: 567 dxa = 1.0 cm (jeszcze ciaśniej) bo szablon trójki jest większy
    # i trzeba zmieścić na 1 stronie.
    margin_dxa = '567' if template_type == 'TROJKA' else '720'
    sectPr_check = body.find(wt('sectPr'))
    if sectPr_check is not None:
        pgMar = sectPr_check.find(wt('pgMar'))
        if pgMar is not None:
            for side in ('top','bottom','left','right'):
                pgMar.set(f'{{{W}}}{side}', margin_dxa)

    # ── Fonty etykiet: pomniejsz zbyt duże (24→20 dla głównych etykiet,
    # zachowaj 24 dla nagłówków SET 1/SET 2/Wyniki turnieju w tabeli wyników).
    # Tor/Godzina/Grupa/Mecz# (sz=24) → sz=22 (czytelne, mieszczą się w linii)
    # Punkty SET 1/SET 2/Wygrane sety/Podpis (sz=24) → sz=20
    LABELS_BIGGER = {'Tor','Godzina','Grupa','Mecz','#'}
    LABELS_HEADER = {'Punkty','SET 1','SET 2','Wygrane','sety','Podpis'}
    # ── Operacje SPECYFICZNE DLA TRÓJKOWEGO szablonu:
    # Wymuszenie Calibri jako fontu etykiet (Tor/Godzina/Grupa/Mecz#/Punkty SET 1/2/
    # Wygrane sety/Podpis). Bez tego LibreOffice (i Word bez Aptos) używa fallback
    # który jest znacznie szerszy i wszystko rozjeżdża się na 2 wiersze.
    # Zachowujemy oryginalne size (24) - Calibri w tym rozmiarze mieści się normalnie.
    if template_type == 'TROJKA':
        TROJKA_LABELS = {'Tor','Godzina','Grupa','Mecz','#',
                         'Punkty','SET 1','SET 2','Wygrane','sety','Podpis',
                         'PunktySET 1','PunktySET 2','Wygranesety',
                         'Wyniki turnieju','Wyniki turnieju:'}
        for r in body.iter(wt('r')):
            ts = r.findall(wt('t'))
            if not ts: continue
            text_content = ''.join((t.text or '') for t in ts).strip()
            if text_content in TROJKA_LABELS:
                rPr = r.find(wt('rPr'))
                if rPr is None:
                    rPr = etree.Element(wt('rPr'))
                    r.insert(0, rPr)
                fonts = rPr.find(wt('rFonts'))
                if fonts is None:
                    fonts = etree.SubElement(rPr, wt('rFonts'))
                for a in ('ascii','hAnsi','eastAsia','cs'):
                    fonts.set(f'{{{W}}}{a}', 'Calibri')
        
        # Zmniejszenie odstępów po tabeli wynikowej żeby mieściło się na 1 stronie.
        # Ostatni paragraf przed sectPr ma duży before - usuwamy.
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

    # ── Operacje SPECYFICZNE DLA INDYWIDUALNEGO szablonu:
    # 1) Pomniejszenie fontów etykiet (24→20)
    # 2) Przesunięcie tabeli 1 w prawo (tblInd=600 dxa)
    # Trójkowy szablon ma inną strukturę i nie wymaga tych poprawek.
    if template_type == 'IND':
        for r in body.iter(wt('r')):
            ts = r.findall(wt('t'))
            if not ts: continue
            text_content = ''.join((t.text or '') for t in ts).strip()
            sz_el  = r.find(f'{wt("rPr")}/{wt("sz")}')
            szCs_el = r.find(f'{wt("rPr")}/{wt("szCs")}')
            if sz_el is not None and sz_el.get(f'{{{W}}}val') == '24':
                new_size = None
                if text_content in LABELS_BIGGER:
                    new_size = '20'   # 10pt — Tor/Godzina/Mecz #
                elif text_content in LABELS_HEADER:
                    new_size = '18'   # 9pt — Punkty SET 1, Wygrane sety, Podpis
                if new_size:
                    sz_el.set(f'{{{W}}}val', new_size)
                    if szCs_el is not None:
                        szCs_el.set(f'{{{W}}}val', new_size)

        # ── Wyrównanie tabeli 1 z tabelą 2 (jak we wzorcu z gridlinami):
        # Tabela 1 ma być WĘŻSZA (9090 DXA z oryginału - bez zmian) ale PRZESUNIĘTA W PRAWO
        # przez tblInd=600 DXA. Wtedy:
        #   - lewy brzeg tabeli 1 = 600 DXA (= ~1.06 cm)
        #   - prawy brzeg tabeli 1 = 600 + 9090 = 9690 DXA = prawy brzeg tabeli 2 ✓
        # Pierwsza pusta komórka tabeli 1 (z nazwiskami) zaczyna się od pozycji 600 DXA,
        # czyli mniej więcej tam gdzie kolumna IMIONA w tabeli 2.
        first_tbl = body.find(wt('tbl'))
        if first_tbl is not None:
            # Dodaj tblInd = 600 DXA (przesunięcie w prawo)
            tblPr = first_tbl.find(wt('tblPr'))
            if tblPr is not None:
                tblInd = tblPr.find(wt('tblInd'))
                if tblInd is None:
                    tblInd = etree.SubElement(tblPr, wt('tblInd'))
                tblInd.set(f'{{{W}}}w', '600')
                tblInd.set(f'{{{W}}}type', 'dxa')
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
                    # 1) Wymuś font Calibri + size 18 dla KAŻDEGO 'r' (także bez rPr)
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
                        # size 18 (9pt)
                        sz = rPr.find(wt('sz'))
                        if sz is None:
                            sz = etree.SubElement(rPr, wt('sz'))
                        sz.set(f'{{{W}}}val', '18')
                        szCs = rPr.find(wt('szCs'))
                        if szCs is None:
                            szCs = etree.SubElement(rPr, wt('szCs'))
                        szCs.set(f'{{{W}}}val', '18')
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

    # ── Generuj protokoły
    first = True
    for group_name, matches in sheets_data:
        for match in matches:
            if not first:
                body.append(_make_page_break_para())
            first = False

            cloned = [copy.deepcopy(el) for el in template_elements]

            # Wstaw nazwę turnieju + datę jako paragraf w prawym górnym rogu
            # (przed pierwszą tabelą, wyrównany do prawej, małą czcionką)
            if tournament_name or tournament_date:
                header_parts = []
                if tournament_name:
                    header_parts.append(tournament_name.strip())
                if tournament_date:
                    header_parts.append(tournament_date.strip())
                header_text = ' · '.join(header_parts)

                hp = etree.Element(wt('p'))
                hpPr = etree.SubElement(hp, wt('pPr'))
                hjc = etree.SubElement(hpPr, wt('jc'))
                hjc.set(f'{{{W}}}val', 'right')
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
            _fill_protocol(cloned, match,
                           hide_grupa_mecz=hide_grupa_mecz,
                           phase_label=phase_label)

            # Lista elementów do wstawienia w lewym obszarze
            order = image_order if image_order else (
                (['qr'] if qr_rid_info else []) +
                sorted(logo_rids.keys())
            )

            anchored = []
            cell_w_cm = 5.24      # szerokość kolumny "Wyniki turnieju"
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

                # Pozycja własna z image_positions, jeśli jest
                if image_positions and key in image_positions:
                    pos = image_positions[key]
                    x_cm = pos.get('x', (cell_w_cm - w_cm) / 2)
                    y_cm = pos.get('y', cur_y_cm)
                    # Jeśli user podał konkretną szerokość i wysokość, użyj ich
                    # bezpośrednio (bez wymuszania proporcji obrazu)
                    if 'width' in pos:
                        w_cm = pos['width']
                    if 'height' in pos:
                        h_cm = pos['height']
                    elif 'width' in pos:
                        # Tylko width podana - dla QR zachowaj 1:1, inaczej zachowaj proporcje
                        if key == 'qr':
                            h_cm = w_cm
                        else:
                            # Zachowaj proporcje obrazu z rid_info
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

                anchored.append(_make_anchored_image_drawing(
                    rid, cx_emu, cy_emu, py_emu, px_emu))

                cur_y_cm = y_cm + h_cm + spacing_cm

                # Po QR — flaga że napis idzie pod QR
                if key == 'qr':
                    label_after_qr = True
                    cur_y_cm += 0.4  # miejsce na napis "Wyniki turnieju"

            # Napis "Wyniki turnieju" - tylko gdy jest QR.
            # Bez QR napis nie ma sensu (jest tylko dla podpisania QR).
            if qr_rid_info and include_qr:
                # Wyciągnij faktyczną pozycję i wysokość QR z image_positions
                if image_positions and 'qr' in image_positions:
                    qr_pos = image_positions['qr']
                    qr_y = qr_pos.get('y', 0.2)
                    qr_w_cm = qr_pos.get('width', 2.4)
                    label_y_cm = qr_y + qr_w_cm + 0.1
                else:
                    label_y_cm = 0.2 + 2.4 + 0.1
                _populate_left_area(cloned, anchored, 'Wyniki turnieju', label_y_cm)
            else:
                # Bez QR - tylko obrazy bez napisu
                _populate_left_area(cloned, anchored, '', 0)

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
