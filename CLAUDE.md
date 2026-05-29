# CLAUDE.md

Wskazówki dla Claude Code przy pracy w tym repo. Dotyczy każdej sesji.
Opis funkcjonalny i instrukcja użytkownika: patrz `README.md` (nie duplikuj tego tutaj).

## Czym jest projekt

Aplikacja Streamlit generująca protokoły meczowe `.docx` dla turniejów Mölkky
(Polska Federacja Mölkky). Pobiera dane z publicznego arkusza Google Sheets,
wykrywa fazy grupowe i drabinkę pucharową, składa dokument z szablonów `.docx`
i opcjonalnie konwertuje do PDF.

Produkcja: https://protocol.streamlit.app/ (deploy z `main` przez Streamlit Cloud).

## Workflow gita (WAŻNE)

- **Właściciel autoryzował push BEZPOŚREDNIO na `main`.** Każdy push na `main`
  = natychmiastowy deploy na produkcję (Streamlit Cloud odświeża ~1–2 min).
  Można commitować i pushować na `main` bez dopytywania.
- Mimo to: pushuj świadomie. Większe/ryzykowne zmiany warto najpierw odpalić
  lokalnie (`streamlit run app.py`), bo na `main` nie ma siatki bezpieczeństwa.
- Operacje destrukcyjne (force-push, `reset --hard`, kasowanie branchy/historii)
  tylko na wyraźne polecenie.
- Commituj i pushuj często — środowisko sesji jest efemeryczne, niezacommitowane
  zmiany przepadają przy starcie kolejnej sesji.

**Definicja ukończenia zadania (standing rule — rób to BEZ proszenia):**
zadanie nie jest skończone, dopóki nie: (1) zaktualizowano sekcję
**Roadmap i status** poniżej (Co teraz / Backlog / wpis w Log zmian),
(2) zrobiono commit, (3) wypchnięto na `main`. To rutyna każdej sesji —
użytkownik nie musi o to prosić.

## Architektura

Dwa duże pliki, bez pakietów/modułów:

### `app.py` (~2200 linii) — UI Streamlit
- Wejście użytkownika, edytor pozycji grafik (X/Y/szer./wys.), podgląd HTML 1:1.
- Ładowanie i cache'owanie grafik (logo PFM, dodatkowe grafiki), QR.
- `compute_default_positions(...)` — domyślne rozmieszczenie elementów.
- `docx_to_pdf(...)` — konwersja przez **LibreOffice** (CLI `soffice`), nie czysty Python.
- `build_image_args()` — spina grafiki z UI w argumenty dla generatora.

### `generate_docx.py` (~2870 linii) — dane + składanie docx
- **Pobieranie z Google Sheets:** `get_sheet_gids`, `fetch_sheet`,
  `fetch_via_gviz` / `fetch_via_gid` / `fetch_via_export_name`,
  `fetch_all_group_sheets`. Gviz bywa kapryśny przy nagłówkach — patrz Pułapki.
- **Faza grupowa:** `parse_group_rows`, `_is_valid_match_row`.
- **Drabinka pucharowa:** `detect_phase` (rozpoznaje `1/8`, `1/4`, `Półfinał`,
  `Finał`, `Mecz o 3 miejsce`, godziny w nawiasach), `parse_drabinka_rows`,
  `detect_drabinka_phases`, `fetch_drabinka_phase`.
- **Składanie docx (XML python-docx + ręczny lxml):** `_fill_protocol`,
  `build_document`, `build_blank_document`, helpery od obrazków
  (`_make_inline_image_drawing`, `_make_anchored_image_drawing`,
  `_make_czworka_strip_table`), QR (`make_qr_bytes`), formatowanie komórek
  (`_set_cell_value`, `_fix_pkt_set_cells`, `_force_calibri_score_labels`).

### Szablony `.docx`
Pliki binarne w katalogu repo (nazwy: `IND_*`, `TROJKA_*`, `CZWORKA_*` ×
`Grupa`/`Bo3`/`Bo5`). Generator wypełnia placeholdery — przy zmianie układu
często trzeba edytować szablon, nie tylko kod. Diff gita ich nie pokaże sensownie.

## Uruchomienie i weryfikacja

```bash
pip install -r requirements.txt   # streamlit, requests, lxml, pillow, qrcode
streamlit run app.py
```
- Systemowo (z `packages.txt`): **libreoffice** (docx→pdf) + **fonts-crosextra-carlito**
  (substytut Calibri — bez tego font w docx/pdf się rozjeżdża).
- Python 3.9+.
- **Brak testów automatycznych.** Weryfikacja = odpalenie aplikacji i wygenerowanie
  docx na realnym/przykładowym arkuszu. Po zmianach w generatorze zawsze sprawdź
  wygenerowany plik (paginacja, czcionki, pozycje grafik), bo regresje są wizualne.

### Smoke regresja (automatyczna via hooki)

`tests/regression.py` buduje wszystkie 9 szablonów + testuje filtr placeholderów.
Uruchamiana automatycznie przez hooki:
- **PostToolUse** (po Edit/Write `generate_docx.py` lub `app.py`): py_compile +
  regresja. Exit 2 jeśli FAIL — Claude widzi błąd i naprawia od razu.
- **PreToolUse Bash** (przed `git push ...`): regresja. Blokuje push jeśli FAIL.

Ręcznie: `PYTHONPATH=. python3 tests/regression.py`

## Pułapki (gotchas)

- **`app.py` ma WIELE gałęzi sekcji „4. Domyślne elementy"** (puchar vs grupa).
  Dwa różne `st.header("4. Domyślne elementy")` o linie ~851 i ~860. Każda
  zmiana wprowadzająca nową zmienną używaną w `build_document(...)` MUSI
  zdefiniować ją w **OBU** gałęziach (puchar `is_no_graphics=True` i grupa
  `else`), inaczej puchar wybucha `NameError`. Sprawdzaj:
  `grep -n 'st.header(.4\. Dom' app.py` → muszą być 2 linie, edytuj obie.
  Statyczna analiza (ruff, pylint, pyflakes) tego NIE łapie — to dyscyplina,
  nie tooling. Nie ma realistycznego testu na app.py (Streamlit + sieć).

- **Gviz gubi/przesuwa nagłówki** (np. kolumnę `Tor`) — parser musi być odporny;
  patrz dotychczasowe fixy w historii commitów. **TESTOWANIE:** openpyxl
  (lokalnie z xlsx) NIE reprodukuje tego bugu — daje pełny header. Produkcja
  używa gviz przez Google Sheets (sandbox blokuje docs.google.com). Dlatego
  regresja `tests/regression.py` MA helper `_simulate_gviz_drop_tor_header()`
  który blanksuje header 'Tor' gdy kolumna pod nim jest czysto numeryczna
  (główne źródło bugów produkcyjnych typu M4U 2026: layout "Tor SeedID Player",
  bez headera col_player-1=SeedID 'A1'/'B4'). **Przed pushem zawsze sprawdź czy
  `parse_drabinka_rows` przechodzi też przez ten symulator** — odpalana
  automatycznie w regresji, ale nie ignoruj jeśli przy testach realnych danych
  xlsx wynik jest „za dobry" w porównaniu do skarg usera.
- **Czcionka:** docelowo Calibri → renderowane przez Carlito; bez `fonts-crosextra-carlito`
  szeryfy/metryki się psują. Nazwiska celowo bywają serif — sprawdzaj świadomie.
- **Paginacja Bo5/IND** bywała problematyczna (treść str.2 wchodziła na str.1) —
  uważaj przy zmianach łamania stron i wysokości elementów.
- Arkusz Google Sheets musi być publiczny ("każdy z linkiem").

## Roadmap i status (źródło prawdy między sesjami)

> **Zasada:** to jest pamięć projektu. Każda sesja zaczyna od przeczytania tej
> sekcji. Po skończeniu zadania **zaktualizuj** „Co teraz", „Backlog" i dopisz
> wpis do „Log zmian". Tabelę statusu szablonów trzymaj zsynchronizowaną
> z `README.md`.

### Status szablonów
✅ zweryfikowane · 🟡 w trakcie weryfikacji · 🔴 niedostępne

| Typ | Grupowa | Pucharowa (Bo3/Bo5) |
|---|---|---|
| Indywidualny | ✅ | ✅ (Bo3 ✅, Bo5 ✅) |
| Drużynowy 2-os. | 🟡 (klon TROJKA, tylko blank) | 🔴 |
| Drużynowy 3-os. | ✅ | ✅ (Bo3 ✅, Bo5 ✅) |
| Drużynowy 4-os. | ✅ | ✅ (Bo3 ✅, Bo5 ✅) |

**Rozwiązany problem „linie" (str.2 — wyrównanie tabeli wyników):** tabela `(SET 4)/(SET 5)`
na str.2 renderowała się (a) wyśrodkowana (lewa x≈197 zamiast 110) ORAZ (b) za wąska
(prawa krawędź 624 zamiast 763 jak str.1). Fix w `build_document` (blok
`if template_type=='IND_Bo5'` po `_force_calibri_score_labels`):
- **lewa krawędź:** KLUCZ to **JAWNE `jc=left`** — samo usunięcie `jc=center` NIE działa
  (LibreOffice i tak centruje, ignorując `tblInd`). Trzeba: `jc=left` + `tblInd=715`
  + `tblLayout=fixed` (wszystko w kolejności schematu CT_TblPrBase).
- **prawa krawędź:** skalujemy kolumny (gridCol + tcW) do `tblW=8550` (= szer. tabeli str.1),
  bo str.2 ma mniej setów i była naturalnie węższa. Po fixie obie tabele: L=110, R=763.

**Weryfikacja real-data — jak (Google zablokowany w sandboxie):** network policy zwraca
„Host not in allowlist" 403 dla docs.google.com → live-fetch testować TYLKO na produkcji.
W sandboxie: user dał xlsx (`GP2_2026_wyniki_2.xlsx`, zakładka `Drabinka`, turniej
INDYWIDUALNY). Czytanie: `pip install openpyxl` (NIE w hooku) → `load_workbook(data_only=True)`
→ wartości OK. Dane meczu = dict `{tor,godz,mecz,z1,z2}`; `build_document(sheets_data=[(label,[...])],
template_type='IND_Bo5', hide_grupa_mecz=True, ...)` zwraca bajty docx (bez sieci, bo dane
podajemy ręcznie). Render→PDF: `soffice --headless -env:UserInstallation=file:///tmp/lo_p2
--convert-to pdf`; PDF→PNG: `pdftoppm -r 110`. Pomiar wyrównania: pierwszy ciemny piksel
w wierszu tabeli (PIL).

### Co teraz (current focus)
- Wszystkie 3 typy (IND, TROJKA, CZWORKA) × wszystkie 3 fazy (Grupa, Bo3, Bo5)
  zweryfikowane wizualnie u usera. Blank-formularze dostępne dla wszystkich 9.
- DWOJKA_Grupa: klon TROJKA_Grupa.docx jako bazowy szablon. Wystarcza dla
  blank-formularzy; pełna obsługa real-data (parsing 2-os. składów) — backlog.

### Backlog (kolejność = priorytet)
- [x] IND_Bo5 — szablon, fill, wyrównanie str.2, weryfikacja real-data (xlsx GP2 2026) ✅
- [x] CZWORKA Bo3/Bo5 — wzorzec templates, poprawny Tor, formatowanie nagłówka ✅
- [x] IND_Bo3 + CZWORKA Grupa zweryfikowane wizualnie ✅
- [x] DWOJKA_Grupa blank — klon TROJKA, dostępny w UI „Pusty formularz" ✅
- [ ] DWOJKA — real-data fill (parsing 2-osobowych składów drużyn)
- [ ] DWOJKA Bo3/Bo5 — szablony pucharowe

### Log zmian (najnowsze u góry)
- 2026-05-29 — Iteracja statusu + nowości:
  (a) Wszystkie IND/TROJKA/CZWORKA fazy oznaczone ✅ w tabeli (i README).
  (b) Pusty formularz: dodane CZWORKA Bo3, CZWORKA Bo5, DWOJKA Grupa do
  selectbox „Format setów" w sekcji „Pobierz pusty formularz". Mapping
  blank_type → blank_template rozszerzony.
  (c) DWOJKA_Grupa.docx: klon TROJKA_Grupa.docx (binary copy) zarejestrowany
  w `template_files`. Działa dla blank-formularzy; dla real-data fill TROJKA
  fill code rysuje 3 sloty graczy — adaptacja na 2 graczy = backlog.
  (d) Rozpiski: auto-shrink mniej agresywny (75 dxa/char zamiast 110).
  „Warsaw Adventure Team" (21 znaków, bold) mieści się przy default sz=8
  bez skalowania; tylko naprawdę długie nazwy (>26 znaków) skalują się.
  (e) Komunikaty „rozpiski dla N zawodników" → „N drużyn" gdy turniej
  drużynowy (poprawione 2 zapomniane miejsca w app.py — sukces po
  generacji DOCX i PDF).
- 2026-05-29 — Trzy fixy z jednej iteracji usera:
  (1) Rozpiski: auto-shrink czcionki w kolumnach drużyna/gracz 1/2 gdy
  tekst się nie mieści w komórce. Heurystyka 110 dxa/char przy sz=16 bold;
  skalujemy proporcjonalnie do min 6pt. Przykład: „Stowarzyszenie Aktywny
  Orlik" (28 znaków) przy bold rozjeżdżało się na 2 linie — teraz mieści się
  w 1 linii pomniejszone. `_new_para` przyjmuje float size_pt (round do
  half-pt) bo wcześniej psuł sz=14.0 jako string.
  (2) CZWORKA Bo5: usunięcie cienkiej poziomej linii pod instrukcjami
  („Set przegrany... 0:50."). Row 1 c0 instrukcji ma vMerge=restart
  bottom=nil, ale row 2 c0 (vMerge cont.) dziedziczył tblBorders
  top=single — LO rysował linię. Fix: dla wszystkich komórek z `<w:vMerge/>`
  w tabelach headerowych wymuś top=nil + bottom=nil.
  (3) CZWORKA Bo5: pominięcie `_fix_pkt_set_cells` dla CZWORKA_Bo5 —
  poprzednio nasz custom rebuild Pkt SET (sz=20) był nadpisywany przez
  `_fix_pkt_set_cells` (sz=16 dla cw<800), przez co Pkt SET 1-5 robiły się
  mniejsze niż Wygr.sety / Podpis (sz=20). Po skipie wszystkie 7 komórek
  nagłówka (Pkt SET 1-5 + Wygr + Podpis) renderują się tym samym sz=20.
- 2026-05-29 — Rozpiski meczowe rozszerzone na turnieje drużynowe
  (3-os i 4-os). UI expander pokazuje się też dla drużynowych; label
  „zawodników"/„drużyn" dobierany dynamicznie. `build_player_schedules_doc`
  ma nowy flag `is_team` — header tabeli „gracz 1/2" → „drużyna 1/2".
  Pogrubienie nazwy własnej drużyny w wierszach działa tak samo jak dla
  zawodników (matches_to_player_schedules nie zmienia logiki — bierze
  z1/z2 jak są).
- 2026-05-29 — CZWORKA Bo5: dwa fixy nagłówka. (1) Calibri-force rozszerzony
  z tylko `tbls[0]` na WSZYSTKIE tabele headerowe (heurystyka: 'Tor' + 'Mecz'
  /'Godzina' w pierwszym wierszu). Bo5 ma 2 takie tabele (str.1 + str.2) —
  bez tego na str.2 Tor/Godzina/Mecz# znów wrapowały w Aptos-fallback.
  (2) Komórki Pkt SET 1-5 w Bo5 są wąskie (720-745 dxa), 1-liniowy
  „Pkt. SET N" nie mieści się — c7 ('Pkt'+'. '+'SET 1') i c8-c11
  ('Pkt'+'.'+' SET '+'N') wrapowały NIESPÓJNIE (różny układ runów ⇒ różne
  miejsca łamania). Plus nuclear block forsował sz=16, a Wygr.sety/Podpis
  zostawały sz=20 ⇒ wizualnie różnej wielkości. Fix: dla CZWORKA_Bo5
  rebuilduję komórki w nagłówku (cw<900) do CZYSTYCH 2 linii „Pkt." +
  „SET N" przy sz=20 — zgodnie z Wygr.sety/Podpis.
- 2026-05-29 — Rozpiski zawodników: zagęszczone do 10 kart/stronę (2×5
  zamiast 2×4). Header tabeli „godzina"→„godz." (mieści się w 1 linii
  w wąskiej kol.), nazwisko 14pt→12pt, subtitle 8pt→7pt, body 10/9pt→8/7pt,
  marginesy kart 140/160→80/140, tcMar 40/60→10/50. Weryfikacja na 12
  zawodnikach: 10 na str.1, 2 na str.2, czytelne.
- 2026-05-29 — CZWORKA Bo3/Bo5: pierwsza tabela (header z Tor/Godzina/Mecz# +
  instrukcje + nazwy drużyn) — wymuszony Calibri na wszystkich runach. Bez tego
  produkcyjny LO (bez Aptos/Aptos Narrow) brał szeroki fallback: (a) „Tor"
  w komórce 536 dxa zawijało na 2 linie („To"/„r"), (b) instrukcje (jc=right,
  sz=20) przelewały się poza komórkę i były obcinane z lewej strony. Carlito
  (Calibri) mieści się — weryfikacja przez render docx→pdf→png pokazuje
  jednolinijny Tor i pełne instrukcje. Cofa fragment polityki z 2026-05-29
  „Tor zostaje w Aptos Narrow żeby wrapował jak we wzorcu" — wzorzec usera
  miał Tor jednolinijny, wrap był artefaktem braku Aptos w LO.
- 2026-05-29 — `requirements.txt`: dodano `python-docx>=1.0.0`. Rozpiski
  zawodników (`build_player_schedules_doc`) importują `from docx import Document`,
  ale pakiet nie był w requirements → na Streamlit Cloud rzucało
  „Błąd budowania rozpisek: No module named 'docx'" zarówno dla DOCX jak i PDF.
- 2026-05-30 — Rozpiski meczowe per zawodnik (IND, button 🃏 w UI). Każda
  karta = imię + Grupa/turniej/data + tabela 4-kol (godz | tor | gracz 1 |
  gracz 2). Grid 2 kolumny na A4 portrait. Derywowane z `parse_group_rows`
  (NIE z ukrytych kolumn AC-AI — to były duplikaty). Builder używa
  python-docx jako skeletonu (raw OOXML nie rendrował multi-col w LO).
- 2026-05-30 — `parse_drabinka_rows`: fallback Tor (gdy gviz zgubi nagłówek)
  scan leftward 1..3 offsety z walidacją `_looks_like_tor_col` (odrzuca seed IDs
  typu 'A1'/'B4'). Wcześniejszy fallback `col_player-1` brał kolumnę seed ID
  (M4U 2026: układ "Tor SeedID Player" → bez 'Tor' header, col_player-1=SeedID).
  Reprodukcja: openpyxl daje 'Tor' header → OK; po `rows[0][1]=''` (symulacja
  gviz) — przed fixem Tor='A1'/'B2'/'A2'/'B1', po fixie Tor='1'/'2'/'3'/'4'.
- 2026-05-29 — CZWORKA Bo3/Bo5: podmiana szablonów na wzorzec usera (M4U 2026).
  Usunięcie bloku 'Instrukcje italic gray' — kod nadpisywał Aptos Narrow 10pt czarny
  na Calibri 8pt szary kursywa, co dawało inny wygląd niż wzorzec. Usunięcie
  Tor/Godzina/Mecz# z CZW_BO_LABELS — etykiety zostają w Aptos Narrow (narrow
  column → zawijanie "Tor"→"To/r" jak w wzorcu). Tor poprawny: weryfikacja na xlsx
  M4U 2026 daje Tor=1/2/3/4 (nie A1/B4) — detekcja _looks_like_tor_col działa.
- 2026-05-29 — `parse_drabinka_rows`: robust detekcja kolumny `Tor`/`Grupa`.
  Wcześniej sprawdzało tylko 2 sąsiednie kolumny na lewo od fazy — gdy arkusz
  miał dodatkową kolumnę (np. seed ID „A1/B4" w A, Tor w B, Grupa w C, faza w D)
  parser brał A jako Tor i wszystko się rozjeżdżało. Teraz skanuje wszystkie
  kolumny na lewo aż do innej fazy. Plus: grupa z `col_grupa` faktycznie
  zwracana w match (była hardcoded ''). Test scenariuszy: IND klasyczny,
  CZWORKA z Grupą, CZWORKA z seed ID, wiele faz obok siebie — wszystkie ✓.
- 2026-05-29 — `_set_cell_value`: width-aware auto-shrink też dla wąskich
  komórek (próg z 1500→400 dxa). „13:30" w wąskim Godz cell (743 dxa)
  w nowym Bo5 szablonie nie wrapsuje już na „13:3"/„0".
- 2026-05-29 — CZWORKA_Bo5: podmiana szablonu (user — więcej komórek w
  górnej tabelce dla 5 SETów). IND: oznaczony ✅ (Bo3/Bo5 zweryfikowane).
- 2026-05-28 — CZWORKA Bo3/Bo5: nowe szablony usera z równymi R1=R2=346 +
  węższym score table (R1: 450→360, dane: 374→360) — Lionheart/Stowarzyszenie
  nie ściśnięte, balans wizualny zachowany.
- 2026-05-28 — CZWORKA Bo3/Bo5: napraw `_fix_pkt_set_cells` (skip gdy brak
  `<w:br/>` — nowy szablon ma już czyste „Pkt. SET 1"), NUCLEAR block ogranicz
  do komórek z „Pkt"/„Punkt"+„SET" (Tor sz=20 nie był resetowany na 16),
  Calibri force na Wygr./sety/Podpis (rozwiązuje Aptos-fallback wrap),
  pageBreakBefore na header str.2 Bo5 (R0 TBL2 nie wycieka na koniec str.1).
- 2026-05-28 — IND: filtr placeholderów. `_is_placeholder_name` rozpoznaje
  „Gracz N" (grupowa) i „bye" (drabinka, case-insensitive). `build_document` ma
  nowy param `skip_placeholders=True` — filtruje mecze z placeholderem po
  dowolnej stronie. Aktywny tylko dla `IND`/`IND_Bo3`/`IND_Bo5`; dla zespołowych
  no-op. UI: checkbox w app.py „Ignoruj protokoły dla placeholderów…" widoczny
  tylko gdy `is_individual`, default ON.
- 2026-05-28 — CZWORKA Bo3/Bo5: lock wysokości R2 (`hRule=exact` na trows[2])
  — w produkcyjnym LO Carlito renderował szerzej, „Stowarzyszenie Aktywny Orlik I"
  zawijało → R2 puchnął 2x. Lock tylko R2; R1 zostaje auto-grow (instrukcje na 4 linie).
- 2026-05-28 — CZWORKA Bo3/Bo5: podmiana szablonów na poprawione przez usera (split
  komórki przy `Mecz #`), `_set_cell_value` przerobiony na width-aware auto-shrink
  (czytamy `tcW` komórki, dobieramy największy rozmiar przy którym tekst mieści się
  w 1 linii; factor 5.0 dxa/sz_unit empirycznie skalibrowany na „Stowarzyszenie
  Aktywny Orlik I" w komórce 3060 — bez wrapowania). Tor/Godz/Mecz# wartości
  wyrównane do lewej w CZWORKA Bo3/Bo5 (komórka tc[5] jest 2520 dxa, przy center
  „1" lądował daleko od „Mecz #"). Regresja 9/9: poprawne strony (grupowe 2, Bo3 2,
  Bo5 4 dla 2 meczów).
- 2026-05-27 — CZWORKA Bo3/Bo5 (landscape): fix overflow — paragraf nagłówka turnieju
  dodawany przez build_document (~280 dxa) pchał ostatnie wiersze tabeli wyników na nową
  stronę, bo bottom margin szablonu = 26 dxa (~0). Kompensacja: gdy header występuje
  (tournament_name/date/phase), zmniejszamy top margin z 540 do 260 dxa. Po fixie:
  Bo3 1 mecz = 1 strona (było 1, stabilniej), Bo5 1 mecz = 2 strony (było 3).
- 2026-05-27 — IND_Bo5: tabela wyników str.2 ((SET 4)/(SET 5)) — pełne wyrównanie do str.1:
  lewa krawędź (jawne `jc=left` + `tblInd=715` + `tblLayout=fixed`) ORAZ prawa krawędź
  (skalowanie kolumn do `tblW=8550`). Po fixie obie tabele L=110/R=763. Regresja 9/9 OK.
- 2026-05-27 — IND_Bo5: nowy szablon z wbudowanym page breakiem + przeróbka kodu
  (fill str.2 przez `all_tbls[2]`, usunięcie splitu, godz=24). Hook: doinstalowanie
  `libreoffice-writer` (obraz ma tylko `libreoffice-core` — bez Writera docx→pdf pada
  na „source file could not be loaded") + `poppler-utils` (PDF→PNG do weryfikacji).
- 2026-05-27 — Dodano CLAUDE.md (pamięć projektu), session-start hook
  (auto-instalacja pip-deps + Carlito) oraz tę sekcję roadmap.
