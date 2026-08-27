"""Smoke regression for generate_docx.build_document.

Buduje wszystkie 12 szablonów na danych syntetycznych. Test filtra placeholderów
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
TEMPLATES = ['IND', 'IND_Bo3', 'IND_Bo5', 'IND_Bo7',
             'DWOJKA', 'DWOJKA_Bo3', 'DWOJKA_Bo5', 'DWOJKA_Bo7',
             'TROJKA', 'TROJKA_Bo3', 'TROJKA_Bo5',
             'CZWORKA', 'CZWORKA_Bo3', 'CZWORKA_Bo5']

failures = []
for t in TEMPLATES:
    is_puch = ('Bo3' in t or 'Bo5' in t or 'Bo7' in t)
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

# ── Nierozstrzygnięte pary NIE przesuwają parowania (bug IMP 2026 1/16) ──
# Struktura z produkcji: mecz = 2 kolejne wiersze, tor SCALONY na pierwszym
# wierszu pary. Tory 15 i 8 mają pustą komórkę pierwszego zawodnika (czekamy na
# wynik poprzedniej fazy). Stary parser robił `i += 1` przy pustym z1 → sklejał
# zawodników z RÓŻNYCH meczów (Skoracki+Sajkowski) i wymyślał tory (+1).
_imp_rows = [
    ['Tor', '1/16 FINAŁU (14:15)', 'Set 1', 'Set 2', 'Set 3', 'Set 4', 'Set 5', 'SETY'],
    ['16',  'Łukasz Szulc',        '50', '50', '50', '', '', '3'],
    ['',    'Karina Wittmann',     '',   '',   '',   '', '', '0'],
    ['15',  '',                    '50', '50', '50', '', '', '3'],
    ['',    'Tomasz Skoracki',     '',   '',   '',   '', '', '0'],
    ['14',  'Kamil Sajkowski',     '50', '50', '50', '', '', '3'],
    ['',    'Rafał Ściepuro',      '',   '',   '',   '', '', '0'],
    ['13',  'Bartosz Kordecki',    '50', '50', '50', '', '', '3'],
    ['',    'Mariusz Goeck',       '',   '',   '',   '', '', '0'],
    ['8',   '',                    '50', '50', '50', '', '', '3'],
    ['',    'Robert Baścik',       '',   '',   '',   '', '', '0'],
    ['1',   'Jakub Karwowski',     '50', '50', '50', '', '', '3'],
    ['',    'Damian Szostak',      '',   '',   '',   '', '', '0'],
]
try:
    for _label, _rows_v in (('openpyxl', _imp_rows),
                            ('gviz-drop', _simulate_gviz_drop_tor_header(_imp_rows))):
        _, _, m_imp = g.parse_drabinka_rows(_rows_v, target_phase='Pucharowa 1/16 finału')
        _complete = [(m['z1'], m['z2'], m['tor']) for m in m_imp if m['z1'] and m['z2']]
        _incomplete = [(m['z1'], m['z2'], m['tor']) for m in m_imp
                       if not (m['z1'] and m['z2'])]
        _want = [('Łukasz Szulc', 'Karina Wittmann', '16'),
                 ('Kamil Sajkowski', 'Rafał Ściepuro', '14'),
                 ('Bartosz Kordecki', 'Mariusz Goeck', '13'),
                 ('Jakub Karwowski', 'Damian Szostak', '1')]
        if _complete != _want:
            failures.append(f'IMP unresolved ({_label}): kompletne={_complete}, oczek. {_want} '
                            '(shift-bug: sklejanie zawodników z różnych meczów)')
        _want_inc = [('', 'Tomasz Skoracki', '15'), ('', 'Robert Baścik', '8')]
        if _incomplete != _want_inc:
            failures.append(f'IMP unresolved ({_label}): niekompletne={_incomplete}, '
                            f'oczek. {_want_inc}')
except Exception as e:
    failures.append(f'IMP unresolved: {type(e).__name__}: {e}')
    traceback.print_exc(file=sys.stderr)

# ── Puste komórki torów: tor='' (NIE wymyślamy last_known+1) ──
_notor_rows = [
    ['Tor', '1/4 FINAŁU (16:15)', 'Set 1', 'Set 2', 'Set 3', 'SETY'],
    ['7',   'Anna Pierwsza',      '', '', '', ''],
    ['',    'Beata Druga',        '', '', '', ''],
    ['',    'Cezary Trzeci',      '', '', '', ''],
    ['',    'Dorota Czwarta',     '', '', '', ''],
]
try:
    _, _, m_nt = g.parse_drabinka_rows(_notor_rows, target_phase='Pucharowa 1/4 finału')
    _tors_nt = [m['tor'] for m in m_nt]
    _pairs_nt = [(m['z1'], m['z2']) for m in m_nt]
    if _pairs_nt != [('Anna Pierwsza', 'Beata Druga'), ('Cezary Trzeci', 'Dorota Czwarta')]:
        failures.append(f'no-tor pary: {_pairs_nt}')
    if _tors_nt != ['7', '']:
        failures.append(f"no-tor: Tor={_tors_nt}, oczek. ['7', ''] — pusty tor ma "
                        "ZOSTAĆ pusty (nie wymyślamy '8')")
except Exception as e:
    failures.append(f'no-tor drabinka: {type(e).__name__}: {e}')
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

# Test: parse_group_rows odporne na gviz-drop nagłówka 'Tor' (bug IMP 2026 →
# „0 grup"). Layout grup PFM: '#' | Godzina | Tor | dup | Grupa X | set | set | z2…
# Po zgubieniu nagłówka 'Tor' przez gviz parser MUSI dalej znajdować mecze
# (przedtem 'tor' in norm było jedynym wyzwalaczem detekcji nagłówka → 0).
try:
    _grp_rows = [
        ['#', 'Godzina', 'Tor', '', 'Grupa A', '1. set', '2. set', '', '1.set', '2. set'],
        ['1', '09:00:00', '5',  'Krzysztof D', 'Krzysztof D', '39', '50', 'Małgorzata H', '50', '43'],
        ['2', '09:00:00', '6',  'Wojciech K',  'Wojciech K',  '50', '50', 'Przemysław L', '48', '11'],
        ['3', '09:30:00', '7',  'Bartosz K',   'Bartosz K',   '32', '50', 'Leszek Ś',     '50', '27'],
    ]
    _gm_norm = g.parse_group_rows(_grp_rows)
    if len(_gm_norm) != 3:
        failures.append(f'parse_group_rows (z header Tor): {len(_gm_norm)} meczów, oczek. 3')
    _gm_drop = g.parse_group_rows(_simulate_gviz_drop_tor_header(_grp_rows))
    if len(_gm_drop) != 3:
        failures.append(f'parse_group_rows (gviz drop Tor): {len(_gm_drop)} meczów, oczek. 3 '
                         '(regresja bugu IMP 2026 „0 grup")')
    # Tor musi być poprawny (numer), nie pusty/nazwisko
    if _gm_drop and _gm_drop[0].get('tor') != '5':
        failures.append(f"parse_group_rows (gviz drop Tor): Tor={_gm_drop[0].get('tor')}, oczek. '5'")
    # Najgorszy przypadek (IMP 2026 live): gviz KASUJE cały nagłówek poza 'Grupa A'
    # (znika '#'/'Godzina'/'Tor'/etykiety setów). Fallback pozycyjny MUSI ratować.
    _grp_killed = [
        ['', '', '', 'Grupa A', '', '', '', '', '15'],
        ['1', '09:00', '5',  'Krzysztof D', 'Krzysztof D', '39', '50', 'Małgorzata H', '50', '43', '-1'],
        ['2', '09:00', '6',  'Wojciech K',  'Wojciech K',  '50', '50', 'Przemysław L', '48', '11', '-1'],
        ['3', '09:30', '7',  'Bartosz K',   'Bartosz K',   '32', '50', 'Leszek Ś',     '50', '27', '-1'],
    ]
    _gm_killed = g.parse_group_rows(_grp_killed)
    if len(_gm_killed) != 3:
        failures.append(f'parse_group_rows (gviz skasowany nagłówek, IMP 2026 live): '
                        f'{len(_gm_killed)} meczów, oczek. 3 (fallback pozycyjny)')
    if _gm_killed and (_gm_killed[0].get('z1') != 'Krzysztof D' or
                       _gm_killed[0].get('z2') != 'Małgorzata H' or
                       _gm_killed[0].get('grupa') != 'A'):
        failures.append(f'parse_group_rows (skasowany nagłówek): zła detekcja kolumn: {_gm_killed[0]}')
except Exception as e:
    failures.append(f'parse_group_rows gviz-drop: {type(e).__name__}: {e}')
    traceback.print_exc(file=sys.stderr)

# ── Helpery filtra torów z app.py ────────────────────────────────────────
# app.py w całości zaimportować się nie da (uruchomiłby aplikację Streamlit),
# ale czyste funkcje pomocnicze da się wyciąć po AST i przetestować osobno.
# Dzięki temu logika dzielenia wydruku po torach ma realną osłonę regresyjną.
try:
    import ast as _ast
    import os as _os
    _app_path = _os.path.join(
        _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__))), 'app.py')
    _app_src = open(_app_path, encoding='utf-8').read()
    _want = {'_tor_sort_key', '_tor_label', '_collect_tors', '_apply_tor_filter',
             '_tor_name_suffix', '_pages_per_match'}
    _tree = _ast.parse(_app_src)
    _picked = [n for n in _tree.body
               if (isinstance(n, _ast.FunctionDef) and n.name in _want)
               or (isinstance(n, _ast.Assign)
                   and getattr(n.targets[0], 'id', '') == 'TOR_NONE')]
    _missing = _want - {n.name for n in _picked if isinstance(n, _ast.FunctionDef)}
    if _missing:
        failures.append(f'app.py: brak helperów filtra torów: {sorted(_missing)}')
    else:
        _ns = {}
        exec(compile(_ast.Module(body=_picked, type_ignores=[]), 'app.py', 'exec'), _ns)
        _NONE = _ns['TOR_NONE']

        def _m(t):
            return {'tor': t, 'z1': 'A Kowalski', 'z2': 'B Nowak'}

        # Tory sortujemy liczbowo — alfabetycznie dałoby 1,10,11,2.
        if _ns['_collect_tors']([_m('10'), _m('2'), _m('1'), _m('11')]) != ['1', '2', '10', '11']:
            failures.append('_collect_tors: złe sortowanie torów (ma być liczbowe)')
        # Mecz bez toru MUSI mieć własny kubełek, inaczej zniknąłby po cichu.
        if _ns['_collect_tors']([_m('3'), _m(''), _m(None)]) != ['3', _NONE]:
            failures.append('_collect_tors: brak kubełka na mecze bez toru')

        _sd = [('Gr. A', [_m('1'), _m('2'), _m('3')]), ('Gr. B', [_m('2'), _m('')])]
        _cases = [
            (['2'], 2, 'filtr po torze 2'),
            ([], 5, 'pusty filtr = bez filtrowania'),
            ([_NONE], 1, 'filtr na mecze bez toru'),
            (['99'], 0, 'nieistniejący tor'),
        ]
        for _sel, _exp, _desc in _cases:
            _got = _ns['_apply_tor_filter'](_sd, _sel)[1]
            if _got != _exp:
                failures.append(f'_apply_tor_filter ({_desc}): {_got} meczów, oczek. {_exp}')
        # Grupa bez trafień nie może zostać jako pusty wpis.
        if [l for l, _ in _ns['_apply_tor_filter'](_sd, ['3'])[0]] != ['Gr. A']:
            failures.append('_apply_tor_filter: pusta grupa nie została odfiltrowana')

        for _sel, _exp in [(['1', '2', '3'], '_tory_1-2-3'),
                           (['1', '2', '3', '4', '5'], '_tory_1-5'),
                           ([], ''),
                           ([_NONE], '_bez_toru')]:
            _got = _ns['_tor_name_suffix'](_sel)
            if _got != _exp:
                failures.append(f'_tor_name_suffix({_sel}): {_got!r}, oczek. {_exp!r}')

        # Bo5/Bo7 = 2 strony na protokół (estymata czasu + timeout konwersji PDF).
        for _tt, _exp in [('IND', 1), ('IND_Bo3', 1), ('IND_Bo5', 2),
                          ('DWOJKA_Bo7', 2), ('CZWORKA_Bo5', 2), (None, 1)]:
            if _ns['_pages_per_match'](_tt) != _exp:
                failures.append(f'_pages_per_match({_tt}): oczek. {_exp}')
except Exception as e:
    failures.append(f'helpery filtra torów: {type(e).__name__}: {e}')
    traceback.print_exc(file=sys.stderr)

# ── PIN-y Mölkkify na rozpiskach ────────────────────────────────────────────
try:
    # Nagłówek rozpoznawany po braku cyfry w kolumnie PIN-u; zera wiodące
    # i artefakt '1234.0' z Excela nie mogą zepsuć PIN-u.
    _rows = [['Imię i nazwisko', 'PIN'], ['Jan Kowalski', '0042'],
             ['Anna Nowak', '1234.0'], ['', ''], ['Piotr Zych', '7777']]
    _pairs, _info = g.pin_rows_to_pairs(_rows)
    if [p[0] for p in _pairs] != ['Jan Kowalski', 'Anna Nowak', 'Piotr Zych']:
        failures.append(f'pin_rows_to_pairs: zle nazwy {[p[0] for p in _pairs]}')
    if dict(_pairs).get('Anna Nowak') != '1234':
        failures.append('pin_rows_to_pairs: artefakt .0 nie usuniety przy parsowaniu')
    if g.clean_pin('1234.0') != '1234':
        failures.append('clean_pin: artefakt .0 z Excela nie usuniety')
    if g.clean_pin('0042') != '0042':
        failures.append('clean_pin: zero wiodace zgubione')
    # Bez nagłówka pierwszy wiersz musi zostać danymi (tryb awaryjny).
    if len(g.pin_rows_to_pairs([['Jan Kowalski', '1111']])[0]) != 1:
        failures.append('pin_rows_to_pairs: wiersz danych potraktowany jak naglowek')

    # read_pin_file: xlsx (openpyxl) i csv. Regresja na buga produkcyjnego
    # "name 'pd' is not defined" - pandas NIE jest w requirements.txt, wiec
    # odczyt nie moze sie na nim opierac.
    try:
        import io as _io2
        from openpyxl import Workbook as _WB
        _wb = _WB(); _ws = _wb.active
        _ws.append(['Team Name', 'Country', 'PIN', 'Email', 'Members'])
        _ws.append(['Sebastian Bisaga', '', '0013', '', ''])   # tekst -> zero zostaje
        _ws.append(['Szymon Szulc', '', 2137, '', ''])          # liczba -> bez zera
        _buf = _io2.BytesIO(); _wb.save(_buf)
        _xr = g.read_pin_file(_buf.getvalue(), 'plik.xlsx')
        _xp, _xi = g.pin_rows_to_pairs(_xr)
        if _xp != [('Sebastian Bisaga', '0013'), ('Szymon Szulc', '2137')]:
            failures.append(f'read_pin_file (xlsx): {_xp}')
        if (_xi['name_col'], _xi['pin_col']) != ('Team Name', 'PIN'):
            failures.append(f'read_pin_file (xlsx): zle naglowki {_xi}')
        _cp, _ = g.pin_rows_to_pairs(
            g.read_pin_file(b'Team Name;PIN\nAla Makocka;0013\n', 'x.csv'))
        if _cp != [('Ala Makocka', '0013')]:
            failures.append(f'read_pin_file (csv, separator sredniki): {_cp}')
    except Exception as e:
        failures.append(f'read_pin_file: {type(e).__name__}: {e}')

    # Realny układ eksportu z Mölkkify: kolumny nazwy i PIN-u NIE sasiaduja
    # ('Country' miedzy nimi) i nie sa pierwsze. Wiersze z PIN-em bez nazwy
    # nie maja do kogo trafic - musza byc policzone, nie polkniete po cichu.
    _mk = [['Team Name', 'Country', 'PIN', 'Email', 'Members'],
           ['Ala Makocka', 'PL', '0013', 'a@b.pl', ''],
           ['', '', '2137', '', ''],
           ['', '', '1303', '', ''],
           ['Druga Druzyna', 'PL', '0500', '', '']]
    _mp, _mi = g.pin_rows_to_pairs(_mk)
    if _mp != [('Ala Makocka', '0013'), ('Druga Druzyna', '0500')]:
        failures.append(f'pin_rows_to_pairs (uklad Molkkify): {_mp}')
    if _mi['skipped_no_name'] != 2:
        failures.append(f"pin_rows_to_pairs: skipped_no_name={_mi['skipped_no_name']}, oczek. 2")
    if (_mi['name_col'], _mi['pin_col']) != ('Team Name', 'PIN'):
        failures.append(f'pin_rows_to_pairs: zle rozpoznane naglowki {_mi}')
    # Zero wiodace z kolumny tekstowej musi przetrwac caly przeplyw.
    if dict(_mp).get('Ala Makocka') != '0013':
        failures.append('pin_rows_to_pairs: zero wiodace zgubione w ukladzie Molkkify')

    # Dopasowanie odporne na ogonki, wielkosc liter, spacje i kolejnosc czlonow.
    if g.name_match_key('Kowalski Jan') != g.name_match_key('jan  KOWALSKI'):
        failures.append('name_match_key: kolejnosc/wielkosc liter nie znormalizowana')
    if g.normalize_person_name('Łukasz Żak') != g.normalize_person_name('Lukasz Zak'):
        failures.append('normalize_person_name: ogonki nie usuniete')

    _sched = [{'name': 'Jan Kowalski', 'group': 'A', 'matches': []},
              {'name': 'Łukasz Żak', 'group': 'A', 'matches': []},
              {'name': 'Ewa Bez Pinu', 'group': 'B', 'matches': []}]
    _out, _rep = g.match_pins_to_schedules(_sched, [
        ('kowalski  jan', '0042'),      # inna kolejnosc + podwojna spacja
        ('Lukasz Zak', '5555'),         # bez ogonkow
        ('Ktos Spoza Arkusza', '9999'),  # nadmiarowy wpis
    ])
    if _out[0].get('pin') != '0042':
        failures.append('match_pins: nie dopasowano po odwroconej kolejnosci')
    if _out[1].get('pin') != '5555':
        failures.append('match_pins: nie dopasowano mimo braku ogonkow')
    if _rep['without_pin'] != ['Ewa Bez Pinu']:
        failures.append(f"match_pins: zly without_pin {_rep['without_pin']}")
    if _rep['unused'] != ['Ktos Spoza Arkusza']:
        failures.append(f"match_pins: zly unused {_rep['unused']}")
    if _rep['matched'] != 2:
        failures.append(f"match_pins: matched={_rep['matched']}, oczek. 2")
    # Wejscie nie moze byc modyfikowane w miejscu.
    if 'pin' in _sched[0]:
        failures.append('match_pins: zmodyfikowal wejsciowa liste')
    # Ta sama osoba z dwoma roznymi PIN-ami = konflikt, wpis pomijany.
    _o2, _r2 = g.match_pins_to_schedules(
        [{'name': 'Jan Kowalski', 'group': 'A', 'matches': []}],
        [('Jan Kowalski', '1111'), ('Jan Kowalski', '2222')])
    if _r2['conflicts'] != ['Jan Kowalski'] or _o2[0].get('pin') != '1111':
        failures.append(f'match_pins: konflikt obsluzony zle ({_r2})')

    # PIN musi realnie trafic do docx (i tylko wtedy, gdy jest dopiety).
    _doc_pin = g.build_player_schedules_doc(
        [{'name': 'Jan Kowalski', 'group': 'A', 'pin': '0042',
          'matches': [{'godzina': '10:00', 'tor': '1', 'z1': 'Jan Kowalski', 'z2': 'X'}]}])
    _doc_no = g.build_player_schedules_doc(
        [{'name': 'Jan Kowalski', 'group': 'A',
          'matches': [{'godzina': '10:00', 'tor': '1', 'z1': 'Jan Kowalski', 'z2': 'X'}]}])
    import zipfile as _zf, io as _io
    _xml = _zf.ZipFile(_io.BytesIO(_doc_pin)).read('word/document.xml').decode('utf-8')
    _xml_no = _zf.ZipFile(_io.BytesIO(_doc_no)).read('word/document.xml').decode('utf-8')
    if '0042' not in _xml:
        failures.append('build_player_schedules_doc: PIN nie trafil do dokumentu')
    if 'PIN' in _xml_no:
        failures.append('build_player_schedules_doc: PIN drukuje sie mimo braku danych')
except Exception as e:
    failures.append(f'PIN-y Molkkify: {type(e).__name__}: {e}')
    traceback.print_exc(file=sys.stderr)

if failures:
    print('REGRESJA FAIL:', file=sys.stderr)
    for f in failures:
        print(f'  - {f}', file=sys.stderr)
    sys.exit(1)

print(f'REGRESJA OK: {len(TEMPLATES)} szablonów + filtr placeholderów + helpers + gviz drabinka + grupy gviz-drop + split-phase + rozpiski + filtr torów + PIN-y Mölkkify')
sys.exit(0)
