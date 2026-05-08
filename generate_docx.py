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

# Rozpoznawane nazwy faz — funkcja detect_phase rozpoznaje dynamicznie:
#   • "1/N finału" / "1/N FINAŁU"
#   • "PÓŁFINAŁY" / "półfinał"
#   • "FINAŁ"
#   • "MECZ O N. MIEJSCE" (gdzie N = 3, 5, 7, ..., 31)
#   • "MIEJSCA X-Y" (dowolny zakres)
# Zwraca (phase_key, phase_full_name) lub (None, None).
# phase_key jest unikatowym kluczem fazy (do dopasowania target_phase).

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
            # Znaleźliśmy fazę. Szukamy "Tor" w lewo (max 4 kolumny).
            col_tor = None
            for offset in range(1, 5):
                if c_idx - offset >= 0:
                    t = row[c_idx - offset].strip().lower()
                    if t == 'tor':
                        col_tor = c_idx - offset
                        break
            # "Grupa" między col_tor a c_idx (jeśli jest)
            col_grupa = None
            if col_tor is not None:
                for j in range(col_tor + 1, c_idx):
                    if j < len(row) and row[j].strip().lower() == 'grupa':
                        col_grupa = j
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
    
    # ── KROK 2: wybór bloku pasującego do target_phase ──
    # Używamy detect_phase() do parsowania target_phase (np. "Pucharowa 1/4 finału",
    # "Mecz o 9. miejsce", "Miejsca 17-20"). Potem szukamy bloku którego phase_key
    # == target_key, ALBO którego phase_full pasuje (np. 'półfinał' z arkusza i '1/2'
    # z UI oba mapują na '1/2 FINAŁU').
    chosen = None
    target_key = None
    target_full = None
    if target_phase:
        target_key, target_full = detect_phase(target_phase)
        if target_key:
            # Bezpośrednie: blok ma identyczny phase_key
            for block in phase_blocks:
                if block['phase_key'] == target_key:
                    chosen = block
                    break
            # Pośrednie: blok ma ten sam phase_full (np. PÓŁFINAŁY i 1/2 FINAŁU)
            if not chosen:
                for block in phase_blocks:
                    if block['phase_full'] == target_full:
                        chosen = block
                        break
            # Jeśli target_phase ma sens ale brak takiego bloku w arkuszu —
            # zwracamy pustą listę z poprawnym phase_full_name. User dostanie
            # czytelny komunikat "Nie znaleziono fazy".
            if not chosen:
                return target_full, None, []
    
    # target_phase=None lub niezrozumiały → fallback do pierwszego bloku
    if not chosen:
        chosen = phase_blocks[0]
    
    col_player = chosen['col']
    col_tor = chosen['col_tor'] if chosen['col_tor'] is not None else max(0, col_player - 1)
    col_grupa = chosen['col_grupa']
    phase_full_name = chosen['phase_full']
    phase_time = chosen['time']
    header_idx = chosen['header_row']
    
    # ── KROK 3: czytanie meczów z wybranego bloku ──
    matches = []
    data_rows = rows[header_idx + 1:]
    
    def g(row, c):
        if c is None or c >= len(row): return ''
        return row[c].strip()
    
    i = 0
    match_num = 1
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
                break  # 5+ pustych = koniec bloku (więcej tolerancji
                # na luki między sekcjami w arkuszu z wieloma fazami obok siebie)
            i += 1
            continue
        empty_streak = 0
        
        # Czy ta komórka to nagłówek nowej sekcji (np. "MIEJSCA 17-24" pod 1/8)?
        if any(kw in z1_lower for kw in stop_keywords):
            break
        
        # Drugi zawodnik w następnym wierszu
        row2 = data_rows[i+1] if i+1 < len(data_rows) else []
        z2_raw = g(row2, col_player)
        z2_lower = z2_raw.lower()
        
        # Tor: hierarchia źródeł — row1.col_tor → row2.col_tor → last+1 → match_num
        tor_raw = g(row1, col_tor) or g(row2, col_tor)
        if tor_raw and tor_raw.strip().isdigit():
            tor = tor_raw.strip()
            last_known_tor = tor
        elif last_known_tor and last_known_tor.isdigit():
            tor = str(int(last_known_tor) + 1)
            last_known_tor = tor
        else:
            tor = str(match_num)
            last_known_tor = tor
        
        # Walidacja: oba nazwiska niepuste, sensowna długość, brak stop-keywords
        if z1_raw and z2_raw and len(z1_raw) >= 3 and len(z2_raw) >= 3:
            has_stop = any(kw in z1_lower or kw in z2_lower for kw in stop_keywords)
            if not has_stop and any(c.isupper() for c in z1_raw) and any(c.isupper() for c in z2_raw):
                matches.append({
                    'tor': tor,
                    'godz': phase_time or '',
                    'grupa': '',
                    'mecz': str(match_num),
                    'z1': z1_raw,
                    'z2': z2_raw,
                })
                match_num += 1
        
        i += 2  # przeskakujemy 2 wiersze (mecz)
    
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

    # Sprawdź też zakładkę Drabinka — pokazujemy WSZYSTKIE wykryte fazy
    drabinka_phases = []  # [(phase_full_name, n_matches, time)]
    for tab_name in ('Drabinka', 'drabinka', 'DRABINKA'):
        try:
            rows = fetch_sheet(sheet_id, tab_name, gid_map)
            if not rows:
                continue
            # Iteruj wszystkie zarejestrowane fazy i sprawdź każdą
            seen_phase_full = set()
            for phase_key in PHASE_NAMES:
                phase_full = PHASE_NAMES[phase_key]
                if phase_full in seen_phase_full:
                    continue  # różne klucze mogą mapować na ten sam phase_full
                phase_name, phase_time, matches = parse_drabinka_rows(rows, phase_key)
                if matches:
                    drabinka_phases.append((phase_name, len(matches), phase_time))
                    seen_phase_full.add(phase_full)
            break
        except Exception:
            continue

    if not found_groups and not drabinka_phases:
        info.append("❌ Nie znaleziono żadnych grup z meczami ani zakładki Drabinka.")
        info.append("   Sprawdź czy arkusz jest publiczny.")
        return info

    if found_groups:
        n = len(found_groups)
        # Polski plural dla "grupa": 1 grupa, 2-4 grupy, 5+ grup, 22-24 grupy itd.
        last = n % 10
        last2 = n % 100
        if n == 1:
            grupa_word = "grupa"
        elif last in (2, 3, 4) and last2 not in (12, 13, 14):
            grupa_word = "grupy"
        else:
            grupa_word = "grup"
        # Polski plural dla "mecz": 1 mecz, 2-4 mecze, 5+ meczów
        m = total_matches
        m_last = m % 10
        m_last2 = m % 100
        if m == 1:
            mecz_word = "mecz"
        elif m_last in (2, 3, 4) and m_last2 not in (12, 13, 14):
            mecz_word = "mecze"
        else:
            mecz_word = "meczów"
        info.append(f"✅ Faza grupowa: {n} {grupa_word}, {m} {mecz_word}")
        for name, count in found_groups:
            count_word = "mecz" if count == 1 else (
                "mecze" if (count % 10) in (2, 3, 4) and (count % 100) not in (12, 13, 14)
                else "meczów")
            info.append(f"  • {name}: {count} {count_word}")
    
    if drabinka_phases:
        info.append("")
        n_phases = len(drabinka_phases)
        faza_word = "faza" if n_phases == 1 else (
            "fazy" if (n_phases % 10) in (2, 3, 4) and (n_phases % 100) not in (12, 13, 14)
            else "faz")
        info.append(f"✅ Drabinka: {n_phases} {faza_word}")
        for phase_name, count, time_str in drabinka_phases:
            count_word = "mecz" if count == 1 else (
                "mecze" if (count % 10) in (2, 3, 4) and (count % 100) not in (12, 13, 14)
                else "meczów")
            time_part = f" ({time_str})" if time_str else ""
            info.append(f"  • {phase_name}{time_part}: {count} {count_word}")
    
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


def _fill_protocol(elements, match, hide_grupa_mecz=False, phase_label=None,
                   template_type='IND'):
    """`hide_grupa_mecz` to historyczna nazwa — obecnie kontroluje tylko czy etykieta
    'Grupa' i jej wartość są ukrywane (dla pucharowej). 'Mecz #' i jego numer
    są ZAWSZE pokazywane (też w pucharowej — numerujemy mecze sekwencyjnie)."""
    tbls = [el for el in elements if el.tag == wt('tbl')]
    if not tbls: return
    rows = tbls[0].findall(wt('tr'))
    
    # ── Bo3 trójka: inna struktura tabeli 1 ──────────────────────────────
    # Bo3 ma 4 wiersze w tabeli 1:
    #   R1: [Tor] [Godz.] [empty] [empty] [empty] [Runda]   ← etykiety
    #   R2: [empty] [PunktySET 1] [PunktySET 2] [PunktySET 3] [Wygrane sety] [Podpis]
    #   R3: [team A name] [score] [score] [score] [sets won] [signature]
    #   R4: [team B name] [score] [score] [score] [sets won] [signature]
    # Wartości Tor/Godz/Runda doklejamy do etykiet (np. "Tor  1", "Godz. 09:00").
    # 'mecz' (numer) ZAWSZE w 'Runda' obok nazwy fazy.
    if template_type == 'TROJKA_Bo3':
        # Bo3 template R1 cells (nowa wersja, prawdziwa trójka):
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
                _set_cell_value(tcs[2], godz_val, size=24, bold=True, align='left')
            mecz_val = match.get('mecz', '').strip()
            if len(tcs) > 5:
                if mecz_val:
                    # Podmieniamy "Runda" na "Mecz X"
                    _set_cell_label(tcs[5], f'Mecz  {mecz_val}')
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
        
        # ── Strona 1 (Table 1, R1.tc[0..5]) — taka sama logika jak w Bo3 ──
        if len(rows) >= 1:
            tcs = rows[0].findall(wt('tc'))
            if len(tcs) > 0 and tor_val:
                _set_cell_label(tcs[0], f'Tor  {tor_val}')
            if len(tcs) > 2 and godz_val:
                _set_cell_value(tcs[2], godz_val, size=24, bold=True, align='left')
            if len(tcs) > 5:
                if mecz_val:
                    _set_cell_label(tcs[5], f'Mecz  {mecz_val}')
                else:
                    _set_cell_label(tcs[5], '')
        # Drużyny w T1.R3.tc[0] i T1.R4.tc[0]
        if len(rows) > 2:
            tcs = rows[2].findall(wt('tc'))
            if tcs: _set_cell_value(tcs[0], z1, size=24, align='right')
        if len(rows) > 3:
            tcs = rows[3].findall(wt('tc'))
            if tcs: _set_cell_value(tcs[0], z2, size=24, align='right')
        
        # ── Strona 3 (Table 3) — inna struktura: 7 widocznych kolumn w R1 ──
        # tc[0]=Tor, tc[1]=empty(wide), tc[2-4]=empty, tc[5]=empty, tc[6]=Runda(last)
        # Bez "Godz." labela (na drugiej stronie nie ma czasu, jest już na pierwszej)
        all_tbls = [el for el in elements if el.tag == wt('tbl')]
        if len(all_tbls) >= 3:
            t3 = all_tbls[2]
            t3_rows = t3.findall(wt('tr'))
            if t3_rows:
                tcs = t3_rows[0].findall(wt('tc'))
                if len(tcs) > 0 and tor_val:
                    _set_cell_label(tcs[0], f'Tor  {tor_val}')
                # "Runda" jest w ostatniej komórce R1 — szukamy ostatniej z 7
                if len(tcs) >= 7:
                    if mecz_val:
                        _set_cell_label(tcs[-1], f'Mecz  {mecz_val}')
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
                   template_type='IND'):
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

    template_files = {
        'IND': 'Grupa_IND.docx',
        'TROJKA': 'Grupa_TROJKA.docx',
        'TROJKA_Bo3': 'Bo3_TROJKA.docx',
        'TROJKA_Bo5': 'Bo5_TROJKA.docx',
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
    # Indywidualny: 720 dxa = 1.27 cm wszędzie.
    # Trójka: 720 dxa (1.27 cm) + zsynchronizowanie tabeli 1 z tabelą 2 (oba 9251 dxa
    # od lewego brzegu marginesu) — bez tego tabela 1 'auto' rozciąga się na całą
    # użyteczną szerokość a tabela 2 zostaje fixed 9251 → rozjazd.
    sectPr_check = body.find(wt('sectPr'))
    if sectPr_check is not None:
        pgMar = sectPr_check.find(wt('pgMar'))
        if pgMar is not None:
            for side in ('top','bottom','left','right'):
                pgMar.set(f'{{{W}}}{side}', '720')

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
    if template_type in ('TROJKA', 'TROJKA_Bo3', 'TROJKA_Bo5'):
        # Bo2 i Bo3/Bo5 mają wspólny zestaw etykiet, ale Bo3/Bo5 NIE potrzebują
        # normalizacji fontu dla 'SET 1/2/3/4/5' / '(SET N)' bo w nowym szablonie
        # te etykiety są inaczej zbudowane (różne run-e) i normalizacja powoduje
        # niespójność (np. SET 2 robi się grubsze niż SET 1/(SET 3)).
        if template_type in ('TROJKA_Bo3', 'TROJKA_Bo5'):
            TROJKA_LABELS = {'Tor','Godz.','Godzina','Mecz','Runda',
                             'PunktySET 1','PunktySET 2','PunktySET 3',
                             'PunktySET 4','PunktySET 5',
                             'PktSET 1','PktSET 2','PktSET 3','PktSET 4','PktSET 5',
                             'Punkty','Pkt','Wygrane','sety','Podpis','Wygranesety'}
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
            # Bo3_TROJKA.docx (nowa wersja): tabela 1 = 9239 dxa (6 kol),
            # tabela 2 = 9439 dxa (25 kol, col0=391). Struktura trójkowa
            # (wiersze 1-18 + WYNIK), 3 sety, SUMA per zawodnik per set.
            ORIG_T1_TOTAL = 9239
            ORIG_T2_TOTAL = 9439
            ORIG_LEFT_COL_DXA = 391
        elif template_type == 'TROJKA_Bo5':
            # Bo5_TROJKA.docx: 4 tabele. Strona 1 = SET 1-3 (jak Bo3),
            # strona 3 = SET 4-5 (warunkowe sety + przepisanie punktacji 1-5).
            # T1 = 9026 dxa (header s.1), T2 = 9439 dxa (results s.1, 25 kol),
            # T3 = 9026 dxa (header s.3, 8 kol z PktSET 1-5),
            # T4 = 6425 dxa (results s.3, 17 kol z (SET 4)/(SET 5)).
            ORIG_T1_TOTAL = 9026
            ORIG_T2_TOTAL = 9439
            ORIG_T3_TOTAL = 9026
            ORIG_T4_TOTAL = 6425
            ORIG_LEFT_COL_DXA = 393  # T2.col0
        else:
            # Grupa_TROJKA.docx: tabela 1 = 9026 dxa, tabela 2 = 9251 dxa, col0=1186
            ORIG_T1_TOTAL = 9026
            ORIG_T2_TOTAL = 9251
            ORIG_LEFT_COL_DXA = 1186
        
        tbls = body.findall(wt('tbl'))
        if len(tbls) >= 2:
            t1, t2 = tbls[0], tbls[1]
            
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
            # Dłuższe nazwy drużyn (np. "SIERMOLKKY-LESZNO TEAM" ~21 znaków)
            # nie mieszczą się w domyślnej szerokości komórki Tor (5.94 cm po
            # skalowaniu). Przesuwamy 800 dxa = 1.41 cm z ostatniej kolumny
            # (Podpis) do pierwszej (Tor / nazwa drużyny).
            if template_type in ('TROJKA_Bo3', 'TROJKA_Bo5'):
                TEAM_COL_TRANSFER = 800  # dxa
                tblGrid_t1 = t1.find(wt('tblGrid'))
                if tblGrid_t1 is not None:
                    gcols = tblGrid_t1.findall(wt('gridCol'))
                    if len(gcols) >= 2:
                        first = gcols[0]
                        last = gcols[-1]
                        first_w = int(first.get(f'{{{W}}}w', '0'))
                        last_w = int(last.get(f'{{{W}}}w', '0'))
                        first.set(f'{{{W}}}w', str(first_w + TEAM_COL_TRANSFER))
                        last.set(f'{{{W}}}w', str(last_w - TEAM_COL_TRANSFER))
                # Dla każdego wiersza: pierwsza komórka +TRANSFER, ostatnia -TRANSFER
                for tr in t1.findall(wt('tr')):
                    cells = tr.findall(wt('tc'))
                    if len(cells) >= 2:
                        for tc, delta in ((cells[0], TEAM_COL_TRANSFER),
                                          (cells[-1], -TEAM_COL_TRANSFER)):
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
            if template_type in ('TROJKA_Bo3', 'TROJKA_Bo5'):
                NEW_LEFT_COL_DXA = ORIG_LEFT_COL_DXA  # zostaje wąsko (~393 dxa)
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
                                if i == 0 and w_int == ORIG_LEFT_COL_DXA:
                                    # To jest komórka lewa (R1.tc[0] etc) — daj nową szer.
                                    tcW.set(f'{{{W}}}w', str(NEW_LEFT_COL_DXA))
                                else:
                                    tcW.set(f'{{{W}}}w', str(int(w_int * scale_rest)))
            
            # ── Normalizacja ramek w ostatniej kolumnie tabeli wyników ──────────
            # PROBLEM: w obu szablonach trójki ostatnia kolumna SUMA miała pełne
            # ramki (top+bottom+left+right), podczas gdy inne kolumny SUMA mają
            # tylko left+right. Efekt: w ostatniej kolumnie widać poziome linie
            # między wierszami (wygląda na "pogrubione"), w innych nie.
            # FIX: w wierszach DANYCH (od R4 do przedostatniego) usuń top/bottom
            # ramki w OSTATNIEJ komórce, żeby kolumna była "czysta" jak inne SUMA.
            t2_rows = t2.findall(wt('tr'))
            for r_idx, tr in enumerate(t2_rows):
                # Pomijamy: R1 (SET headers), R2 (SUMA labels), R3 (continuation
                # vMerge), oraz ostatni wiersz (PKT/WYNIK).
                if r_idx < 3 or r_idx == len(t2_rows) - 1:
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
            # Bo5: redystrybucja szerokości w T3 (header s.2) — analogicznie do T1
            T3_TEAM_TRANSFER = 800  # dxa
            grid3 = t3.find(wt('tblGrid'))
            if grid3 is not None:
                gcols = grid3.findall(wt('gridCol'))
                if len(gcols) >= 2:
                    first = gcols[0]
                    last = gcols[-1]
                    first.set(f'{{{W}}}w',
                              str(int(first.get(f'{{{W}}}w', '0')) + T3_TEAM_TRANSFER))
                    last.set(f'{{{W}}}w',
                             str(int(last.get(f'{{{W}}}w', '0')) - T3_TEAM_TRANSFER))
            for tr in t3.findall(wt('tr')):
                cells = tr.findall(wt('tc'))
                if len(cells) >= 2:
                    for tc, delta in ((cells[0], T3_TEAM_TRANSFER),
                                      (cells[-1], -T3_TEAM_TRANSFER)):
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

            # Wstaw nazwę turnieju + datę + fazę jako paragraf w prawym górnym rogu
            # (przed pierwszą tabelą, wyrównany do prawej, małą czcionką)
            if tournament_name or tournament_date or tournament_phase_text:
                header_parts = []
                if tournament_name:
                    header_parts.append(tournament_name.strip())
                if tournament_date:
                    header_parts.append(tournament_date.strip())
                if tournament_phase_text:
                    header_parts.append(tournament_phase_text.strip())
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
                
                # ── BO5: header też na drugiej stronie (przed T3) ──
                # Bo5 jest 2-stronicowe per mecz. Header w prawym górnym
                # rogu chcemy na obu stronach, więc klonujemy paragraf
                # i wstawiamy przed tabelą 3 (header s.2).
                if template_type == 'TROJKA_Bo5':
                    cloned_tbls = [el for el in cloned if el.tag == wt('tbl')]
                    if len(cloned_tbls) >= 3:
                        t3_in_cloned = cloned_tbls[2]
                        from copy import deepcopy
                        hp_page2 = deepcopy(hp)
                        t3_idx = list(cloned).index(t3_in_cloned)
                        cloned.insert(t3_idx, hp_page2)
            _fill_protocol(cloned, match,
                           hide_grupa_mecz=hide_grupa_mecz,
                           phase_label=phase_label,
                           template_type=template_type)

            # Lista elementów do wstawienia w lewym obszarze
            # Bo3/Bo5 nie używają grafik (lewa kolumna jest wąska, ~0.7 cm).
            if template_type in ('TROJKA_Bo3', 'TROJKA_Bo5'):
                order = []
            else:
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
                col_width_cm = 4.76 if template_type in ('TROJKA', 'TROJKA_Bo3', 'TROJKA_Bo5') else 5.24
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

                anchored.append(_make_anchored_image_drawing(
                    rid, cx_emu, cy_emu, py_emu, px_emu))

                cur_y_cm = y_cm + h_cm + spacing_cm

                # Po QR — flaga że napis idzie pod QR
                if key == 'qr':
                    label_after_qr = True
                    cur_y_cm += 0.4  # miejsce na napis "Wyniki turnieju"

            # Napis "Wyniki turnieju" - tylko gdy jest QR.
            # Bo3/Bo5 nie używają lewej kolumny na grafiki (zostaje oryginalna),
            # więc całkowicie pomijamy populate_left_area.
            if template_type in ('TROJKA_Bo3', 'TROJKA_Bo5'):
                pass  # nie ruszamy lewej kolumny
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
