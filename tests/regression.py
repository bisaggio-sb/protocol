"""Smoke regression for generate_docx.build_document.

Buduje wszystkie 9 szablonów na danych syntetycznych. Test filtra placeholderów
dla IND. Exit 0 jeśli wszystko OK, 1 jeśli jakikolwiek crash/anomalia.
"""
import sys
import traceback

import generate_docx as g

DATA = [("Grupa A", [
    {'tor': '1', 'godz': '09:30', 'mecz': '1', 'grupa': 'A',
     'z1': 'Anna Kordecka', 'z2': 'Bartosz Koziński'},
    {'tor': '2', 'godz': '09:30', 'mecz': '2', 'grupa': 'A',
     'z1': 'Marek Rogalski', 'z2': 'Konrad Rudnik'},
])]
TEMPLATES = ['IND', 'IND_Bo3', 'IND_Bo5',
             'TROJKA', 'TROJKA_Bo3', 'TROJKA_Bo5',
             'CZWORKA', 'CZWORKA_Bo3', 'CZWORKA_Bo5']

failures = []
for t in TEMPLATES:
    is_puch = ('Bo3' in t or 'Bo5' in t)
    try:
        d = g.build_document(
            sheet_id='', sheets_url='https://e.com/s', sheets_data=DATA,
            include_qr=False, include_pfm_logo=False,
            tournament_name='T', tournament_date='27.05.2026',
            tournament_phase_text=('1/4' if is_puch else 'Grupowa'),
            hide_grupa_mecz=is_puch,
            phase_label=('1/4' if is_puch else None),
            template_type=t,
        )
        if len(d) < 10000:
            failures.append(f'{t}: docx podejrzanie mały ({len(d)} bajtów)')
    except Exception as e:
        failures.append(f'{t}: {type(e).__name__}: {e}')
        traceback.print_exc(file=sys.stderr)

# Filtr placeholderów (IND only)
try:
    m_placeholder = [{'tor': '1', 'godz': '09:00', 'mecz': '1', 'grupa': 'A',
                      'z1': 'Anna', 'z2': 'Gracz 5'}]
    d_skip = g.build_document(
        sheet_id='', sheets_url='https://e.com/s',
        sheets_data=[('Gr. A', m_placeholder)],
        include_qr=False, include_pfm_logo=False,
        template_type='IND', skip_placeholders=True,
    )
    d_keep = g.build_document(
        sheet_id='', sheets_url='https://e.com/s',
        sheets_data=[('Gr. A', m_placeholder)],
        include_qr=False, include_pfm_logo=False,
        template_type='IND', skip_placeholders=False,
    )
    if not (len(d_skip) < len(d_keep)):
        failures.append(
            f'filtr placeholderów: skip ({len(d_skip)}B) NIE jest mniejszy '
            f'od keep ({len(d_keep)}B) — filtr nie zadziałał')
except Exception as e:
    failures.append(f'filtr placeholderów: {type(e).__name__}: {e}')

# Helpers do detekcji
for name, expected in [('Gracz 5', True), ('bye', True), ('Mariusz Kordecki', False), ('', False)]:
    if g._is_placeholder_name(name) is not expected:
        failures.append(f'_is_placeholder_name({name!r}) zwróciło źle')

# ── Symulacja gviz w parse_drabinka_rows ──
# gviz (Google Sheets fetch produkcyjny) ma 2 znane patologie:
# 1) Gubi nagłówek 'Tor' gdy kolumna jest czysto numeryczna z scaleniami.
# 2) Floaty (1.0, 2.0) zostają jako string '1.0', '2.0'.
# Lokalnie testujemy z openpyxl (data_only=True) który TYCH problemów nie ma —
# regresja musi je SYMULOWAĆ żeby reprodukować bugi produkcyjne (np. M4U 2026:
# "Tor SeedID Player" — bez header 'Tor', col_player-1 = SeedID col).
def _simulate_gviz_drop_tor_header(rows):
    """Drop header 'Tor' z kolumn gdzie data poniżej jest czysto numeryczna."""
    import re
    rows = [list(r) for r in rows]
    if not rows: return rows
    for ci, h in enumerate(rows[0]):
        if str(h).strip().lower() != 'tor':
            continue
        n_num = 0
        n_other = 0
        for ri in range(1, min(20, len(rows))):
            if ci >= len(rows[ri]): continue
            v = str(rows[ri][ci]).strip()
            if not v: continue
            if re.match(r'^\d{1,2}(\.0+)?$', v): n_num += 1
            else: n_other += 1
        if n_num > 0 and n_other == 0:
            rows[0][ci] = ''  # gviz drop
    return rows

# Scenariusz M4U 2026: kolumny [Tor, SeedID, Player, Set1..]. SeedID ma 'A1'/'B4' itp.
_m4u_rows = [
    ['Tor', '',     '1/4 FINAŁU (14:00)', 'Set 1', 'Set 2', 'Set 3', 'Set 4', 'Set 5', 'SETY'],
    ['1',   'A1',   'Lionheart',           '', '', '', '', '', ''],
    ['',    'B4',   'Puszczyki Chomiki',   '', '', '', '', '', ''],
    ['2',   'B2',   'INNER BULL FORMAT',   '', '', '', '', '', ''],
    ['',    'A3',   'REAL SZTUM',          '', '', '', '', '', ''],
    ['3',   'A2',   'Dwunastka MM',        '', '', '', '', '', ''],
    ['',    'B3',   'ZaGRYFka PAraBoLe',   '', '', '', '', '', ''],
    ['4',   'B1',   'ZaGRYFka Origins',    '', '', '', '', '', ''],
    ['',    'A4',   'Stowarzyszenie A.O.', '', '', '', '', '', ''],
]
try:
    # (a) Z headerem 'Tor' (openpyxl): musi dać 1,2,3,4
    _, _, m_with = g.parse_drabinka_rows(_m4u_rows, target_phase='Pucharowa 1/4 finału')
    tors_with = [m['tor'] for m in m_with]
    if tors_with != ['1', '2', '3', '4']:
        failures.append(f'parse_drabinka_rows (z header Tor): Tor={tors_with}, oczek. [1,2,3,4]')
    # (b) Symulacja gviz (header 'Tor' usunięty): MUSI dalej dać 1,2,3,4 — nie seed IDs
    _gviz_rows = _simulate_gviz_drop_tor_header(_m4u_rows)
    assert _gviz_rows[0][0] == '', 'symulator nie zdjął headera'
    _, _, m_gviz = g.parse_drabinka_rows(_gviz_rows, target_phase='Pucharowa 1/4 finału')
    tors_gviz = [m['tor'] for m in m_gviz]
    if tors_gviz != ['1', '2', '3', '4']:
        failures.append(
            f'parse_drabinka_rows (gviz symulacja, header Tor zgubiony): '
            f'Tor={tors_gviz}, oczek. [1,2,3,4] (bug produkcyjny M4U 2026)')
except Exception as e:
    failures.append(f'gviz drabinka symulacja: {type(e).__name__}: {e}')
    traceback.print_exc(file=sys.stderr)

# ── Faza rozbita na KILKA bloków nagłówkowych (split phase) ──
# Bug produkcyjny FMC 2026: "MIEJSCA 5-8" (drabinka o miejsca, 4 zawodników,
# 2 półfinały) była zapisana jako DWA osobne bloki "MIEJSCA 5-8 (16:45)" w tej
# samej kolumnie, każdy z 1 meczem, oddzielone pustym wierszem. Stary parser brał
# tylko pierwszy blok → 1 mecz zamiast 2. Test reprodukuje strukturę.
_split_rows = [
    ['Tor', 'MIEJSCA 5-8 (16:45)', 'Set 1', 'Set 2', 'Set 3', 'SETY'],
    ['13',  'Marcin Czech',        '0',  '50', '50', '2'],
    ['',    'Rafał Wesołowski',    '50', '20', '40', '1'],
    ['',    '',                    '',   '',   '',   ''],
    ['Tor', 'MIEJSCA 5-8 (16:45)', 'Set 1', 'Set 2', 'Set 3', 'SETY'],
    ['14',  'Sebastian Bisaga',    '30', '34', '',   '0'],
    ['',    'Mateusz Walasik',     '50', '50', '',   '2'],
    ['',    '',                    '',   '',   '',   ''],
]
try:
    _, _, m_split = g.parse_drabinka_rows(_split_rows, target_phase='Miejsca 5-8')
    if len(m_split) != 2:
        failures.append(
            f'parse_drabinka_rows (split phase MIEJSCA 5-8): {len(m_split)} meczów, '
            f'oczek. 2 (bug produkcyjny FMC 2026 — 2 bloki nagłówkowe tej samej fazy)')
    else:
        tors_split = [m['tor'] for m in m_split]
        z1_split = [m['z1'] for m in m_split]
        if tors_split != ['13', '14']:
            failures.append(f'split phase: Tor={tors_split}, oczek. [13, 14]')
        if z1_split != ['Marcin Czech', 'Sebastian Bisaga']:
            failures.append(f'split phase: z1={z1_split}')
except Exception as e:
    failures.append(f'split phase drabinka: {type(e).__name__}: {e}')
    traceback.print_exc(file=sys.stderr)

# ── Dedup: przypadkowo zduplikowany nagłówek+treść (błąd arkusza) ──
# Gdy autor arkusza przez pomyłkę skopiuje ten SAM mecz pod drugim nagłówkiem
# tej samej fazy, nie chcemy 2 identycznych protokołów. Para zawodników w fazie
# knockout jest unikalna → dedup po nieuporządkowanej parze.
_dup_rows = [
    ['Tor', 'MIEJSCA 5-8 (16:45)', 'Set 1', 'Set 2', 'SETY'],
    ['13',  'Marcin Czech',        '0',  '50', '2'],
    ['',    'Rafał Wesołowski',    '50', '20', '1'],
    ['',    '',                    '',   '',   ''],
    ['Tor', 'MIEJSCA 5-8 (16:45)', 'Set 1', 'Set 2', 'SETY'],
    ['13',  'Marcin Czech',        '0',  '50', '2'],
    ['',    'Rafał Wesołowski',    '50', '20', '1'],
]
try:
    _, _, m_dup = g.parse_drabinka_rows(_dup_rows, target_phase='Miejsca 5-8')
    if len(m_dup) != 1:
        failures.append(
            f'parse_drabinka_rows (dedup zduplikowanego meczu): {len(m_dup)} meczów, '
            f'oczek. 1 (ta sama para Czech/Wesołowski pod 2 nagłówkami)')
except Exception as e:
    failures.append(f'dedup drabinka: {type(e).__name__}: {e}')
    traceback.print_exc(file=sys.stderr)

# ── Rozpiski meczowe per zawodnik (IND) ──
# Test: parse_group_rows → matches_to_player_schedules → build_player_schedules_doc.
# Grupa 4-osobowa → każdy ma N-1=3 mecze.
try:
    _g4 = [
        {'tor': '1', 'godz': '11:00', 'grupa': 'A', 'mecz': '1', 'z1': 'Anna', 'z2': 'Bartosz'},
        {'tor': '2', 'godz': '11:00', 'grupa': 'A', 'mecz': '2', 'z1': 'Czesław', 'z2': 'Damian'},
        {'tor': '3', 'godz': '11:30', 'grupa': 'A', 'mecz': '3', 'z1': 'Anna', 'z2': 'Czesław'},
        {'tor': '4', 'godz': '11:30', 'grupa': 'A', 'mecz': '4', 'z1': 'Bartosz', 'z2': 'Damian'},
        {'tor': '5', 'godz': '12:00', 'grupa': 'A', 'mecz': '5', 'z1': 'Anna', 'z2': 'Damian'},
        {'tor': '6', 'godz': '12:00', 'grupa': 'A', 'mecz': '6', 'z1': 'Bartosz', 'z2': 'Czesław'},
    ]
    sch = g.matches_to_player_schedules(_g4, 'A')
    if len(sch) != 4:
        failures.append(f'matches_to_player_schedules: zwróciło {len(sch)} zawodników, oczek. 4')
    for p in sch:
        if len(p['matches']) != 3:
            failures.append(f"rozpiski: {p['name']} ma {len(p['matches'])} meczów, oczek. 3")
    names = [p['name'] for p in sch]
    if names != sorted(names, key=str.lower):
        failures.append(f'rozpiski: zawodnicy nie posortowani alfabetycznie: {names}')
    docx_bytes = g.build_player_schedules_doc(
        sch, tournament_name='Test', tournament_date='30.05.2026')
    if len(docx_bytes) < 5000:
        failures.append(f'build_player_schedules_doc: docx podejrzanie mały ({len(docx_bytes)} B)')
except Exception as e:
    failures.append(f'rozpiski: {type(e).__name__}: {e}')
    traceback.print_exc(file=sys.stderr)

if failures:
    print('REGRESJA FAIL:', file=sys.stderr)
    for f in failures:
        print(f'  - {f}', file=sys.stderr)
    sys.exit(1)

print(f'REGRESJA OK: {len(TEMPLATES)} szablonów + filtr placeholderów + helpers + gviz drabinka + split-phase + rozpiski')
sys.exit(0)
