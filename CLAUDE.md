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
Konwencja ikon (zmiana 2026-05-30): zielonych ✅ NIE używamy — działające
typy są BEZ ikony. 🟡 = budujemy/testujemy. 🔴 = niedostępne (wkrótce).

| Typ | Grupowa | Pucharowa (Bo3/Bo5) |
|---|---|---|
| Indywidualny | działa | działa (Bo3, Bo5, Bo7; w tym 1/32 i 1/64) |
| Drużynowy 2-os. | działa | działa (Bo3, Bo5, Bo7) |
| Drużynowy 3-os. | działa | działa (Bo3, Bo5) |
| Drużynowy 4-os. | działa | działa (Bo3, Bo5) |

**Nazewnictwo plików** (zgodne z PFM SharePoint, UNDERSCORE od 2026-06-01):
`IND_Grupa.docx`, `IND_Bo3.docx`, `IND_Bo5.docx`, `DWÓJKA_Grupa.docx`,
`IND_Bo7.docx`, `DWÓJKA_Bo3.docx`, `TRÓJKA_Grupa.docx`, `TRÓJKA_Bo3.docx`,
`TRÓJKA_Bo5.docx`, `CZWÓRKA_Grupa.docx`, `CZWÓRKA_Bo3.docx`, `CZWÓRKA_Bo5.docx`.
Kod używa stałych
`IND` / `IND_Bo3` / `IND_Bo7` / `DWOJKA` / `DWOJKA_Bo3` / `TROJKA*` / `CZWORKA*` w
identyfikatorach Python, ale mapuje je na nazwy plików przez dict `template_files`.

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
- DWÓJKA Grupa: DZIAŁA (zatwierdzone — status bez ikony). Własny szablon
  `DWÓJKA_Grupa.docx` (DRUŻYNY pion, 4 SUMA/SET). Fix grafik w DOCX zrobiony
  (patrz log 2026-06-01 col0).
- DWÓJKA Bo3 (puchar): DZIAŁA. Szablon `DWÓJKA_Bo3.docx` (SET 1/2/(SET 3),
  6 SUMA). Routing w app.py (puchar + Bo3), format Bo3 odblokowany, fazy
  drabinki bez ikon. Col0 wąska (jak TRÓJKA_Bo3 — puchar bez grafik).
- DWÓJKA Bo5 (puchar): DZIAŁA (patrz log 2026-06-15 fix3).
- IND Bo7 (puchar/finały): DZIAŁA. Szablon `IND_Bo7.docx` od usera (4 tabele:
  header s.1 + score SET 1-4 + header s.2 + score (SET 5)(SET 6)(SET 7)).
  User CELOWO usunął etykietę „Mecz #" (Bo7 grany praktycznie tylko w finale —
  jeden mecz, numer zbędny). Format „Best of 7" w selectboxie TYLKO dla
  Indywidualnego. Weryfikacja renderem: obie strony WYŚRODKOWANE z równymi
  marginesami (L≈45px, R≈863px, gap L/R≈47px), fonty Calibri spójne, etykiety
  „Pkt." z kropką, brak Mecz #. (Centrowanie + kropki: log 2026-06-15 fix5.)

### Backlog (kolejność = priorytet)
- [x] IND Bo5 — szablon, fill, wyrównanie str.2, weryfikacja real-data
- [x] CZWÓRKA Bo3/Bo5 — szablony, Tor, formatowanie nagłówka
- [x] IND Bo3 + CZWÓRKA Grupa zweryfikowane wizualnie
- [x] DWÓJKA Grupa real-data fill (TMP 2026 xlsx) — zatwierdzone
- [x] Renaming plików docx na konwencję PFM SharePoint z UNDERSCORE
- [x] DWÓJKA Bo3 — szablon pucharowy, routing, format/fazy odblokowane
- [x] DWÓJKA Bo5 — szablon pucharowy, routing, guard zdjęty
- [x] IND Bo7 — szablon pucharowy/finały (bez „Mecz #"), routing, wyrównanie str.2
- [x] DWÓJKA Bo7 — szablon pucharowy (landscape, fill+paginacja+routing) DZIAŁA
- [ ] IND Bo7 — nagłówek str.1 do prawej krawędzi w PDF: wymaga EDYCJI SZABLONU
  (kod nie daje rady, patrz log fix5). Nice-to-have.

### Log zmian (najnowsze u góry)
- 2026-06-16 (fix7) — **Numery wierszy zawijały pionowo + TRÓJKA Bo5 Godz str.2.**
  (A) **Numery wierszy (10-18) zawijały pionowo** („1"/„0") w wąskiej kolumnie
  numerów (TRÓJKA Bo3/Bo5 numcol ~391 dxa). W Wordzie „10" mieści się poziomo,
  ale LO (Carlito) renderuje szerzej + tcMar 105/105 → zawijanie. Fix uniwersalny:
  w tabelach wyników (z SUMA) komórki z liczbą 1-18 → tcMar lewo/prawo = 10 dxa.
  Sprawdzone na wszystkich szablonach (TRÓJKA/DWÓJKA/CZWÓRKA Bo3/Bo5, DWÓJKA Bo7).
  (B) **TRÓJKA Bo5: brak Godz. na str.2** — fill T3 miał stary komentarz „NO
  Godz." (poprzednia wersja szablonu bez tej etykiety). Nowy szablon ma „Godz."
  na obu str. → dodano fill wartości godz do tcs[2] na str.2 (jak str.1).
  (C) **Auto-shrink `_set_cell_value`** rozszerzony do sz=14 (było min 18) —
  user: „wolę ciut mniejszą czcionkę niż zawijanie". ≤35 znaków bez zawijania
  w komórce nazwy drużyny DWÓJKA Bo7 (3009 dxa).
- 2026-06-16 (fix6) — **DWÓJKA Bo7 szeryfy/rozmiary + ramki SUMA + jednolity Tor/Godz/Mecz.**
  (1) **DWÓJKA Bo7 szeryfy** — etykiety SET 1-7/(SET 5-7)/numery wierszy/PKT w
  TABELI WYNIKÓW (nie header) były rozbite na runy z font=None/Aptos Narrow →
  LO szeryf. `header_tbls` force łapał tylko nagłówek. Fix: NUKLEARNE Calibri na
  KAŻDYM runie body DWÓJKA Bo7 POZA szarym nagłówkiem (color=666666). Skan
  potwierdza: tylko szary nagłówek non-Calibri. (2) **Wygr. sety za duże** —
  NUCLEAR is_label trzymał template sz=20; dodano width-based size dla is_label
  w DWÓJKA Bo7 → Wygr (733<800)=16 jak Pkt SET. (3) **Pkt SET 7 centrowanie** —
  dodano DWÓJKA Bo7 do rebuildu Pkt SET (czyste 2-linie „Pkt."+„SET N", sz=16)
  → spójne run-splity. (4) **Nowy szablon** (więcej miejsca na nazwę drużyny);
  auto-shrink w `_set_cell_value` (min 18) obsługuje długie nazwy, bardzo długie
  zawijają na 2 linie. (5) **DWÓJKA Bo3/Bo5 prawa krawędź** — POPRZEDNIO za
  bardzo ściennłem (cała na 6). User: krawędź ostatniej SUMA MIAŁA być gruba
  (jak wewnętrzne SUMA right=12). Fix: wiersze SUMA-body (ostatnia komórka
  left=12) → right=12; nagłówek SET/sub-header/PKT (narożniki) → right=6.
  Ostatnia SUMA w normalizacji też 12. (6) **TRÓJKA Bo5 „Mecz #" rozjechane na
  str.2** — komórka Mecz# str.2 węższa (1093 vs 1314) → „Mecz #  1" przy
  docDefault ~24 zawijało. Fix: rozszerzono normalizację nagłówka (force Calibri
  + JEDNOLITY sz=22 na WSZYSTKICH runach r0) na TROJKA_Bo3/Bo5 — mieści się +
  jednolite. (7) **Jednolity rozmiar Tor/Godz/Mecz** (user: „różne rozmiary,
  powinno być jednolicie"). Etykiety r0 → sz=22 (force, normalizacja na
  szablonie). Wartości przez `_set_cell_value` (DWÓJKA Bo3 tor/godz/mecz=28,
  godz wszędzie=24, IND_Bo3 godz=24) → ujednolicone do sz=22 w fillach. Teraz
  DWÓJKA/TRÓJKA Bo3/Bo5 mają Tor/Godz/Mecz jednolicie 22 (wartości bold). UWAGA:
  normalizacja nagłówka działa na SZABLONIE (przed fillem) → wartości
  `_set_cell_value` trzeba zmieniać W FILLU (nie łapie ich normalizacja);
  wartości przez `_set_cell_label` dziedziczą rozmiar znormalizowanego runu.
  Landscape (CZWÓRKA, DWÓJKA Bo7) zostają na sz=20 (własny spójny zestaw).
- 2026-06-16 (fix5) — **DWÓJKA Bo7 (NOWY TYP) + dokończenie iteracji 5-pkt.**
  (1) **DWÓJKA Bo7** — landscape jak CZWÓRKA Bo5, 2 drużyny, 7 setów (score
  s.1=SET 1-4, s.2=SET 5-7). Fill przez gałąź CZWORKA_Bo3/Bo5 (r0[1]=tor,
  r0[3]=godz, r0[5]=mecz, r1[1]/r2[1]=nazwy drużyn). **Fix paginacji:** puste
  akapity + standalone page-break w szablonie dawały PUSTE strony (1 mecz=3,
  2 mecze=5). Fix: pageBreakBefore na hp_page2 + usunięcie pustych akapitów
  między t1-t2; między meczami pageBreakBefore na nagłówku meczu zamiast
  standalone break + usuwanie końcowych pustych akapitów. Wynik: 1 mecz=2 str,
  2 mecze=4 str. **app.py:** „Best of 7" dla 2-os. (bo_options + blank form +
  oba bloki routingu). Dodane też dwójka Bo3/Bo5 do multi-phase bloku (luka).
  (2) **DWÓJKA Bo3 + TRÓJKA Bo5: cienka prawa krawędź** — pętla wymuszała
  `right sz=12` na ostatniej komórce każdego wiersza → pogrubione narożniki.
  Fix: prawa krawędź zewn. sz=12→6 (jednolita z top/left/bottom) + tblBorders
  right→6 + ostatnia SUMA right→6. Separatory SUMA (left=12) bez zmian.
  (3) **DWÓJKA Bo5: „Godz." zawijało + Tor/Mecz nie-bold.** „Godz." = runy
  „Godz"(sz22)+„."(BEZ sz). Cell-level norm nagłówka (Calibri+sz22 na r0) dla
  DWOJKA_Bo3/Bo5 + `_bold_cell` na Tor/Mecz.
  (4) **TRÓJKA Bo5: „Wygr. sety" sz=20.** Nowy szablon + patch t2 „Wygrane"→
  „Wygr." (spójność str.); redystrybucja 800 dxa cała z Podpis (Wygr ~1150,
  mieści sz=20). Cofnięty hack sz=16.
  (5) **CZWÓRKA Bo3/Bo5: podmiana szablonów** (kosmetyka wording, bez kodu).
  **IND Bo7 header do prawej krawędzi — NIE UDAŁO SIĘ (LO-owe ograniczenie).**
  r0 ma inny układ kolumn niż r1-r3 na wspólnym gridzie → ostatnia kolumna jest
  tylko spanowana (nie startuje komórki) → LO ją zwija (header R≈789 vs wyniki
  863). 5 podejść bez skutku (transfer+zero, gridSpan++, usunięcie gridCol+skrót
  span, pusta komórka startująca, autofit). WYMAGA EDYCJI SZABLONU (r1-r3 muszą
  dzielić grid r0). Cofnięte do oryginalnego KROK1b. NIE walczyć kodem.
- 2026-06-16 (fix4) — **Iteracja 5-punktowa: CZWÓRKA swap, TRÓJKA Bo5 Wygr.+ramki,
  DWÓJKA Bo3 ramki, DWÓJKA Bo5 Godz./bold, IND Bo7 header == wyniki.**
  (1) **DWÓJKA Bo3 + TRÓJKA Bo5: zbędne pogrubienia prawego górnego/dolnego
  narożnika.** Pętla „jednolita prawa krawędź" wymuszała `right sz=12` na
  ostatniej komórce każdego wiersza tabeli wyników → gruba prawa krawędź
  (asymetryczna vs cienka top/left/bottom sz=6). Fix: prawa krawędź ZEWNĘTRZNA
  → sz=6 (jednolita ramka), + tblBorders right 12→6, + ostatnia SUMA right 12→6.
  Wewnętrzne separatory SUMA (left=12) bez zmian.
  (2) **DWÓJKA Bo5: „Godz." zawijało + Tor/Mecz nie-bold.** „Godz." rozbite na
  runy „Godz"(sz22)+„."(BEZ sz → dziedziczy duży docDefault, Aptos) — match
  per-run w TROJKA_LABELS nie łapał. Dodano cell-level normalizację nagłówka
  (Calibri + jawny sz=22 na KAŻDYM runie r0) dla DWOJKA_Bo3/Bo5. Tor/Mecz # teraz
  bold (`_bold_cell` po `_set_cell_label`) — jak w CZWÓRCE/DWÓJCE Bo3.
  (3) **TRÓJKA Bo5: „Wygrane sety" → „Wygr. sety" + sz=20 (jak inne kolumny).**
  User wgrał nowy szablon z „Wygr. sety" (t0) — patch binarny też na t2 (str.2
  miała wciąż „Wygrane"). Redystrybucja kolumn w T1/T3 dla Bo5: całe 800 dxa
  z Podpis (nie 300 z Wygr) → Wygr zostaje ~1150 dxa, „Wygr. sety" mieści się
  sz=20 (2 linie). Cofnięty hack sz=16. (Bo3 bez zmian — user: działa.)
  (4) **CZWÓRKA Bo3/Bo5: podmiana szablonów (kosmetyczny wording usera).** Bez
  zmian w kodzie. (5) **IND Bo7: nagłówek == wyniki (prawa krawędź).** Szablon
  ma nagłówki z „sierocą" ostatnią kolumną (r0 pokrywa 19 kol, Pkt SET/nazwiska
  18) → LO centrował tabelę po pełnej szer. ale rysował Pkt SET krócej (R≈809 vs
  wyniki 863). Fix (KROK 1b): przeniesienie szer. sierocej kolumny do poprzedniej
  + zerowanie sierocej do 1 dxa → wiersze Pkt SET renderują pełną szer. = wyniki
  (R=863, obie strony). Zmierzone: header L=46/R=863 = wyniki. **(DWÓJKA Bo7
  szablon dodany do repo, wiring w osobnym commicie.)**
- 2026-06-16 (fix3) — **DWÓJKA Bo5 str.2 = SET 4-5 + TRÓJKA Bo5 dwa bugi.**
  (A) **DWÓJKA Bo5: str.2 pokazywała SET 1-3 zamiast SET 4-5.** Stary szablon
  w repo miał T3 jako duplikat T1 (oba SET 1/2/(SET 3)) — komentarz w logu
  6-15 fix3 błędnie nazwał to „backup sheet, celowo". User wgrał nowy
  `DWÓJKA_Bo5.docx` gdzie T3 ma poprawnie SET 4/5/(SET 4)/(SET 5). Podmiana
  szablonu, fill bez zmian (i tak dotyka tylko nagłówków T0/T2).
  (B) **TRÓJKA Bo5 #1: zduplikowany „Mecz #" w nagłówku str.1.** Szablon
  ma etykietę „Mecz #" w `tcs[7]` (ostatnia komórka), ale fill `_fill_protocol`
  pisał wartość do `tcs[5]` (pustej) → renderowały się DWA „Mecz #": jeden
  z numerem, drugi pusty. Fix: zmiana target na `tcs[-1]` (jak w str.2, która
  była OK). Bo3 nieruszany — działa.
  (C) **TRÓJKA Bo5 #2: „Wygrane sety" zawijało na 3 linie („Wygra/ne/sety").**
  Dwie złożone przyczyny: (1) blok redystrybucji w T1 (linia ~2491) przesuwa
  300 dxa z Wygrane do Tor → Wygr cell ~850 dxa po skalowaniu, za mało dla
  sz=20 single-word „Wygrane" (7 znaków). Fix: w NUCLEAR-2 dla TROJKA_Bo5
  hard-set sz=16 dla is_wygr (Bo3 zostaje na sz=20 bo szablon ma szerszą Wygr
  cell, działa). (2) Na str.2 dodatkowo T3 (tbls[2]) miał `tblW=9450`
  oryginalne ale gridCol już zaktualizowane do 10462 → LO bierze tblW jako
  autorytet i kompresuje proporcjonalnie 0.903 → faktyczna Wygr renderowana
  ~768 dxa, „Wygran/e/sety" 3 linie nawet przy sz=16. Fix: dodano set tblW
  T3 do TARGET_WIDTH (10466) razem ze skalowaniem gridCol. Po obu fixach:
  Wygr cell na obu str. = 850 dxa rzeczywiste, sz=16, „Wygrane/sety" 2 linie.
- 2026-06-16 (fix2) — **IND Bo7 Mecz # (nowy szablon usera) + nowy CZWÓRKA_Bo5.**
  (a) User podesłał nowy `IND_Bo7.docx` z dodaną komórką „Mecz #" w nagłówku
  (r0: tcs[4]='Mecz #' s3, tcs[5]=wartość s2). Fill: numer dopisywany do
  ETYKIETY `_set_cell_label(tcs[4], f'Mecz #  {n}')` (jak TRÓJKA/CZWÓRKA) —
  komórka c4 ma vAlign=center, więc „Mecz #  1" jest pionowo wyśrodkowane;
  oddzielna komórka wartości c5 NIE ma vAlign → wartość lądowała u góry (źle).
  Ukrywanie: dla faz z 1 meczem (Finał, „Mecz o N. miejsce") build_document
  zeruje match['mecz'] (`hide_mecz_num = len(matches)<=1`) → czyścimy całą
  etykietę `_set_cell_label(tcs[4], '')`. Działa na obu nagłówkach (str.1+str.2).
  Font: dodany prefix-match `startswith('Mecz #')` w bloku IND_BO_LABELS
  (etykieta „Mecz #" jest Aptos → bez tego LO serif fallback na numerze).
  Weryfikacja renderem: 2 mecze → „Mecz #  1"/„Mecz #  2" inline Calibri na obu
  str.; 1 mecz → brak etykiety. (b) Nowy `CZWÓRKA_Bo5.docx` od usera (fix
  wcześniejszego błędu) — podmiana bez zmian w kodzie, render OK, regresja 13/13.
  **UWAGA — wyrównanie nagłówka str.1 w PDF NADAL ~53px krótsze od wyników**
  (HEADER R≈809 vs BOTTOM R≈863). Nowy szablon NIE naprawił tego: wiersze
  nazwisk + Pkt SET (r1-r3) wciąż pokrywają 18 z 19 gridCol (Podpis kończy się
  1 kolumnę przed wierszem Tor/Godz). W Wordzie OK (ujemny tblInd), w LO render
  krótki. User: „z obecną wersją można żyć". Fix wymaga edycji SZABLONU
  (rozszerzyć ostatnią komórkę r1-r3 by spinała pełny grid) — moje próby przez
  kod (gridSpan/tcW/orphan-col) za każdym razem powodowały autofit LO i zwijanie
  nazwisk na 3 linie. NIE ruszać kodem.
- 2026-06-16 — **3 fixy: IND Bo7 nagłówek == wyniki, DWÓJKA Bo3/Bo5 centrowanie,
  TRÓJKA Bo3 „Mecz #" Calibri.**
  (A) **IND Bo7 nagłówek wystawał poza wyniki** (po fixie z 06-15 nagłówek
  11475 dxa vs wyniki 10710 dxa, centrowane osobno → wystawanie symetryczne).
  Fix: blok KROK 1 skaluje WSZYSTKIE tabele (oprócz `_p1` = źródło prawdy)
  do `_target_w` = 10710. Wcześniej skalował tylko str.2 (SET 5/6/7). Po fixie
  nagłówki s.1+s.2 = wyniki = 10710 dxa, wszystko wyrównane.
  (B) **DWÓJKA Bo3/Bo5: wyniki dosuwały się w prawo w PDF** (ten sam bug Word/LO
  co IND Bo7: ujemny tblInd -540 nagłówek / -1530 wyniki). Word: tabele
  wynikowe wystają symetrycznie poza nagłówek (user: „ładnie powyśrodkowywane"
  w Word view). LO: ujemny tblInd clampowany → wyniki dosuwają się
  asymetrycznie w prawo. Fix: jc=center + usunięcie tblInd na WSZYSTKICH
  tabelach (Bo5 = 4 tabele). Po fixie wyniki wystają symetrycznie poza
  nagłówek na obu stronach — identycznie jak w Word.
  (C) **TRÓJKA Bo3 „Mecz # 1" w szeryfowej czcionce** (prawy górny róg).
  Przyczyna: template c5 ma „Mecz #" w Aptos; `_set_cell_label` zmieniał tekst
  na „Mecz #  1" (TWO spaces) ale zachowywał font (Aptos). Font-normalizacja
  TROJKA_LABELS matchowała po EXACT string — „Mecz #  1" nie pasowało (set
  zawiera „Mecz" i „#" osobno, oraz pełne „Mecz #" ale BEZ wartości). Bez
  matchu Aptos zostawał → LO fallback na szeryf. Fix: dodano prefix match
  `text_content.startswith('Mecz #')` w pętli normalizacji. Działa też dla
  TRÓJKA Bo5 / DWÓJKA Bo3 / DWÓJKA Bo5 (wszystkie używają tego samego bloku).
- 2026-06-15 (fix5) — **IND Bo7: centrowanie tabel w PDF + kropki po „Pkt".**
  (A) **Bug „strona siada na lewej krawędzi" (PDF).** Szablon używa UJEMNEGO
  `tblInd` (-720 nagłówki / -725 wyniki) — tabela jest szersza niż obszar tekstu
  (wyniki 10710 dxa vs usable 9026) i ma wystawać symetrycznie w marginesy.
  WORD liczy tblInd od marginesu tekstu → renderuje z równymi marginesami (OK,
  „mój docx"). LIBREOFFICE (silnik docx→pdf) liczy tblInd od krawędzi STRONY
  i przycina ujemny do 0 → tabela dosuwa się do lewej krawędzi (zmierzone L=0,
  prawy gap 92px — asymetria). Poprzedni „wyrównanie str.2" kopiował ten ujemny
  tblInd ze str.1, więc OBIE strony siadały na lewej. FIX: `jc=center` na
  WSZYSTKICH tabelach Bo7 (nagłówki + wyniki, obie strony) + usunięcie tblInd.
  Centrowanie jest niezależne od układu odniesienia → identyczne w Word i LO.
  Skalowanie str.2 (SET 5/6/7) do szer. str.1 (10710) zostaje (KROK 1), więc
  obie strony równo szerokie i wyrównane. Zmierzone po fixie: L=45/46px,
  R=863px, gap≈47px na obu str. (B) **Kropka po „Pkt"**: runy nagłówka to
  split „Pkt"+<br/>+„SET N" (gołe „Pkt", 14×=7/str). Dodany loop zamieniający
  run o tekście dokładnie „Pkt" → „Pkt." (bottom-left suma „PKT" nietknięta —
  case-sensitive; font już wymuszony przez NUCLEAR-SET match po podłańcuchu
  'Pkt' — kropka go nie psuje). Weryfikacja renderem: „Pkt. SET 1..7", layout
  wyśrodkowany na obu str. Regresja 13/13 OK.
- 2026-06-15 (fix4) — **IND Bo7 dodany (finały).** Nowy szablon `IND_Bo7.docx`
  od usera: 4 tabele (header s.1, score SET 1-4, header s.2, score (SET 5/6/7))
  z twardym page-breakiem w T1→T2. Struktura analogiczna do IND_Bo5, ale user
  USUNĄŁ etykietę „Mecz #" (Bo7 grany w praktyce tylko w finale — jeden mecz).
  Zmiany:
  (a) `generate_docx`: `template_files['IND_Bo7']`; nowy `_fill_protocol` branch
  IND_Bo7 (godz→tcs[3], imiona→hrows[2/3].tc[0], BEZ Mecz #); IND_Bo7 dopisane
  do list font-normalizacji (IND_BO_LABELS rozszerzone o PktSET 6/7, SET 6/7,
  (SET 6)/(SET 7), 'Godz'), `_fix_pkt_set_cells`/`_force_calibri_score_labels`,
  skip_placeholders, header right-indent, no-graphics left-col, insert nagłówka
  str.2 (cloned_tbls[2]).
  (b) **Wyrównanie str.2** (SET 5/6/7): osobny blok `if template_type=='IND_Bo7'`
  (Bo5 ma swój, inne wymiary). Geometrię str.1 czytamy DYNAMICZNIE z szablonu —
  tabelę wyników str.1 wykrywamy po 'IMIONA' (NIE po 'SET 1', bo nagłówek ma
  'PktSET 1' i jest szerszy/inny ind → str.2 wyrównywałaby się do nagłówka,
  nie do tabeli wyników). Skalujemy gridCol+tcW str.2 do tblW str.1 (10710),
  jawne jc=left + tblInd = ind str.1 (-725). Po fixie: L=0, R≈817px na obu str.
  (c) `app.py`: „Best of 7" w `bo_options` TYLKO dla Indywidualnego; routing
  w obu blokach template_type (grupowy ~1787 i pucharowy ~2040); fmt_suffix
  '_Bo7'; pusty formularz (osobny selectbox dla IND z Bo7 + mapping + label).
  (d) Regresja: IND_Bo7 w TEMPLATES + `is_puch` łapie 'Bo7' → 13/13 OK.
- 2026-06-15 (fix3) — **DWÓJKA Bo5 dodana + TRÓJKA_Bo5 podmieniona + bug 'SET 4/5'.**
  (a) Nowy szablon `DWÓJKA_Bo5.docx` od usera (4 tabele jak TROJKA_Bo5:
  header s.1 + score s.1 + header s.2 + score s.2). Header r0 ma 8 komórek:
  c0=Tor (label+val 5.27cm), c1=Godz./c2=val, c3-c5/c6=filler, c7=Mecz #.
  Score 20 cols/22 rows, SET 1/SET 2/(SET 3) + 6 SUMA na OBU stronach
  (page 2 = duplikat — backup sheet jak user zaprojektował). Dodany
  `_fill_protocol` branch dla DWOJKA_Bo5 wzorowany na TROJKA_Bo5, plus
  wypełnia oba nagłówki (T1 i T3). Usunięte guards w app.py (bo_options,
  warning Bo5, blank-form). Dodany routing template_type='DWOJKA_Bo5'
  dla `is_dwojka and is_pucharowa and Best of 5`. Dodany do regression
  TEMPLATES → 12/12 OK.
  (b) **Bug 'SET 4'/'SET 5' w Aptos**: TROJKA_LABELS (font normalization
  dla TROJKA/DWOJKA Bo3/Bo5) miało 'SET 1/2/3' ale BRAKOWAŁO 'SET 4'/
  'SET 5'. Stąd w DWÓJKA_Bo5 Pkt SET 4/5 zostawały w Aptos (serif
  fallback) zamiast Calibri jak SET 1-3. Dodane do listy. Wcześniej
  nie wyłapało bo żaden poprzedni szablon nie miał 'SET 4'/'SET 5' jako
  osobnych run text (TROJKA_Bo5 też ma — czyli mogło być nawet wcześniej
  złe, ale szablon TROJKA może miał inny split runów).
  (c) Nowy szablon `TRÓJKA_Bo5.docx` od usera podmieniony bez zmian
  w kodzie (regresja 11/11 → 12/12 OK).
- 2026-06-15 (fix2) — **3 uwagi usera po fixach z 06-15: faktycznie nie były naprawione.**
  (A) **IND_Bo3/Bo5 „Wygrane sety" mniejsze niż „Punkty SET N"**: poprzedni
  NUCLEAR 2 hardcodował sz=18 dla Wygr/Podpis, podczas gdy NUCLEAR-SET dla
  Pkt-SET komórek używa sz=20 (cw≥800) lub sz=16 (cw<800 w Bo5). Stąd
  „Punkty SET" sz=20 + „Wygrane sety" sz=18 → wizualna różnica wielkości.
  FIX: NUCLEAR 2 w każdym wierszu znajduje sąsiednią Pkt-SET komórkę,
  czyta jej tcW i stosuje TĘ SAMĄ regułę rozmiaru do Wygr/Podpis. Plus
  wymuszamy `<w:b/>` + `<w:bCs/>` dla pewności (runy „Wygr"+„." w IND_Bo5
  miały bold w szablonie ale dla spójności forsujemy). Zweryfikowane:
  IND_Bo3 wszystkie etykiety = sz=20 bold; IND_Bo5 wszystkie = sz=16 bold.
  (B) **DWÓJKA Grupa/Bo3 prawa krawędź: r1 (pusty wiersz nad SUMA) miał
  jawne `tcBorders right=6`** które nadpisuje `tblBorders=12` z 06-15.
  Stąd: r0 + r2..r20 + r21 = grube, ale r1 cienkie — wizualna przerwa.
  User: „przeciągnąłeś je na dół, a wyżej jest słabo" — dokładnie ten
  efekt. FIX: patch binarny szablonów DWÓJKA_Grupa.docx + DWÓJKA_Bo3.docx
  — r1 ostatnia komórka `right sz=6 → sz=12`. Po fixie krawędź jednolita
  od góry do dołu.
  (C) **Punkty SET 3 — user potwierdził że ten fix z 06-15 zadziałał.** Bez zmian.
- 2026-06-15 (fix) — **Follow-up po 3 uwagach usera (screeny) do iteracji 06-14:**
  (A) **IND_Bo5 „Wygr." wielką dziwną czcionką**: komórka „Wygrane sety" w
  szablonie jest rozbita na runy „Wygr"+„."+„ sety". Selektywne matchowanie
  fontu/rozmiaru (IND_BO_LABELS, LABELS_HEADER, NUCLEAR-SET) łapało tylko
  „sety" → Calibri sz=18; „Wygr"+„." zostawały Aptos BEZ sz → dziedziczyły
  docDefault (~24pt) + serif fallback w LO. FIX: nowy blok „NUCLEAR 2" —
  cell-level rebuild komórek nagłówka zawierających „Wygr…set" oraz „Podpis"
  (tabele ≤6 wierszy) na Calibri sz=18 dla WSZYSTKICH runów. Zweryfikowane
  renderem: „Wygr. sety" = ten sam rozmiar co „Pkt SET N"/„Podpis".
  (B) **DWÓJKA Grupa/Bo3 prawa krawędź za cienka**: cell tcBorders right=12
  na ostatniej kolumnie SUMA NIE renderuje się na ZEWNĘTRZNEJ krawędzi tabeli
  w LibreOffice — LO bierze krawędź zewnętrzną z `tblBorders` (było sz=6).
  Stąd: gruby separator SUMA (left, sz=12) vs cienka krawędź tabeli (right).
  FIX: patch binarny `tblBorders right` 6→12 w tabeli wynikowej obu szablonów.
  Teraz prawa krawędź = gruba, spójna z separatorami SUMA. (Lewa/górna/dolna
  krawędź zostają sz=6 — zgodnie z designem.)
  (C) **DWÓJKA Bo3 „Punkty SET 3" wyrównane do GÓRY**: prawdziwa przyczyna
  (≠ tcMar z 06-14): komórka c3 w r2 NIE miała `<w:vAlign w:val="center"/>`
  podczas gdy c1/c2/c4/c5 miały → tekst lądował przy górnej krawędzi. To był
  błąd w szablonie usera. FIX: patch binarny — wstrzyknięty `vAlign=center`
  do c3 (po tcMar, przed </tcPr>). Zweryfikowane renderem: SET 3 wyśrodkowane
  pionowo jak reszta. (tcMar z 06-14 zostaje — i tak pasuje do sąsiadów.)
  (D) **DWÓJKA Bo5 nagłówek str.2 na zapas**: dodane 'DWOJKA_Bo5' do gałęzi
  insertu hp_page2 — gdy szablon powstanie (4-tabelowy), nagłówek str.2
  zadziała automatycznie (guard `len(cloned_tbls) >= 3`).
- 2026-06-14 (fix) — **6 problemów zgłoszonych przez usera w jednej iteracji:**
  (#1) **IND grupowa 240 vs 230 protokołów**: NIE bug — `skip_placeholders=True`
  (default dla IND) filtruje mecze z „Gracz N" po którejkolwiek stronie. 10
  meczów z niepełnej grupy zostało pominiętych. Wyjaśnione, bez zmiany kodu.
  (#2) **Brak nagłówka „turniej · data · faza" na str.2 IND_Bo5**: gałąź
  inserting hp_page2 obejmowała tylko TROJKA_Bo5/CZWORKA_Bo5. Dodane IND_Bo5
  — szablon ma 4 tabele (header str.1, score str.1, header str.2, score str.2),
  wstawiamy hp_page2 przed `cloned_tbls[2]`. Zweryfikowane renderem.
  (#3) **PFM logo mniejsze niż inne** w DWÓJCE Grupa (gdy user wgrywa PFM jako
  logo1..4): `PFM_TARGET_W = QR_W*1.09 = 2.18 cm` vs `OTHER_MAX_W = 2.7 cm`.
  Wyrównane: `PFM_TARGET_W = OTHER_MAX_W = 2.7 cm` w trybie TROJKA. Teraz
  „system" PFM i „user" PFM mają tę samą szerokość docelową.
  (#4) **Ostatnia kolumna SUMA w DWÓJCE Grupa**: po build_document c13 miała
  top/bottom usuwane SZERZEJ niż c4/c7/c10 — w R3 (pierwszy data row) c4/c7/c10
  miały top=4-8 (pasujące do bottom SUMA-label), c13 miało top usunięte. W
  R20 (ostatni data) c4/c7/c10 miały bottom=12 (gruba krawędź ramy SUMA), c13
  bottom usunięte. Naprawione: zamiast usuwać tcBorders z c13, kopiujemy z
  REFERENCYJNEJ SUMA-cell (drugiej — suma_indices[1]) i dodajemy right=12 do
  ostatniej. Generalizacja: znajdujemy WSZYSTKIE SUMA-cells w R2 i normalizujemy
  każdą do tej samej ramki — działa dla DWÓJKI Bo3 (6 SUMA) i Grupa (4 SUMA).
  (#5) **Ostatnie 3 SUMA w DWÓJCE Bo3**: c13/c16/c19 miały top=4-6, bottom=4-6
  w wierszach danych (poziome linie wewnątrz kolumny SUMA), inne SUMA czyste.
  Naprawione tą samą generalną normalizacją (#4). User podesłał też nowy
  szablon `DWÓJKA_Bo3.docx` z poprawionymi szerokościami kolumn.
  (#6) **Punkty SET 3 inne wyrównanie**: c3 w R2 nie miało `<w:tcMar>` podczas
  gdy c1/c2/c4/c5 miały `<w:tcMar w:left="105" w:right="105"/>`. Padded text
  centrowanie różniło się. Patch binarny — wstrzyknięte tcMar.
  (a) Punkty SET 3 wyglądała szerzej + lekko inaczej niż SET 1/2. Przyczyna:
  w szablonie tblGrid + tcW dawały c1=991, c2=1035, c3=1173 dxa (różnica
  ~18% między c1 a c3). Plus tekst SET 3 był rozbity na 2 runy „SET "+„3"
  zamiast jednego „SET 3" jak SET 1/2 — co przy braku Aptos w LO mogło
  dawać subtelne różnice odstępów. FIX (patch binarny `word/document.xml`):
  grid[7] 675→631, grid[9] 810→628 (każdy traci tyle by c2/c3 miały 991);
  zaoszczędzone 226 dxa wpada do grid[10] (Wygrane sety col1) 355→581.
  tcW c1/c2/c3 w r2-r4 wyrównane do 991, c4 (Wygrane sety) do 1591.
  Runy SET „SET "+„3" zmergowane w jedno „SET 3" (regex single match).
  (b) **TROJKA_AREA_HEIGHT_CM 15.0 → 17.5.** Z 4-5 grafikami logi kończyły
  się ok. wiersza 14 zamiast PKT. Wcześniejszy komentarz „18.5 ucinało
  4-grafikę przez overflow:hidden" dotyczył PODGLĄDU HTML w streamlitcie,
  nie realnego DOCX. 17.5 obejmuje pełną wysokość tabeli wynikowej; w
  podglądzie ostatnia grafika może wystawać poza overflow:hidden — to
  kosmetyka, realna pozycja w DOCX/PDF priorytetem. Dotyczy DWÓJKI Grupa
  i TRÓJKI Grupa (CZWORKA ma strip pod tabelą, IND własną geometrię).
- 2026-06-01 (fix) — **DWÓJKA_Bo3 v2: Mecz # label + custom fill + font.**
  (a) Nowy plik usera `DWÓJKA_Bo3.docx` z prawdziwą etykietą „Mecz #"
  w nagłówku (poprzedni nie miał labelki — pokazywała się sama liczba).
  Layout r0 = 9 komórek z gridSpan: [Tor | Tor-val | Godzina | Godz-val(2) |
  filler(2)×3 | Mecz # | Mecz-val(2)]. Standardowy fill kasował etykietę
  „Mecz #" bo pisał wartość do tcs[7] (label cell). FIX: dedykowana gałąź
  `if template_type == 'DWOJKA_Bo3'` w `_fill_protocol` — wartości lecą do
  właściwych komórek (tor→tcs[1], godz→tcs[3], mecz→tcs[8], drużyny→r3/r4 c0).
  (b) Font normalizacji: DWOJKA_Bo3 przeniesione do gałęzi Bo3/Bo5
  (`PunktySET 1/2/3`, `(SET 3)`, `Mecz`, `#`, `DRUŻYNY` itp. wymuszone na
  Calibri). Wcześniej padało do else-branchu z mniejszym zestawem etykiet,
  przez co „Punkty SET 3" + „(SET 3)" wisiały w Aptos → szeryfowy fallback
  w LO. Po fixie cały nagłówek jednolicie Calibri.
- 2026-06-01 — **DWÓJKA Bo3 + rename UNDERSCORE + KRYTYCZNY fix grafik DOCX.**
  (a) **Rename plików docx** ze spacji na underscore (konwencja PFM SharePoint
  z ostatniego skrina): `IND Grupa.docx` → `IND_Grupa.docx` itd. (git mv ×10).
  Zmieniony tylko `template_files` (jedyne źródło nazw) + help/README.
  (b) **DWÓJKA Bo3** dodana: nowy szablon `DWÓJKA_Bo3.docx` (od usera, SET
  1/2/(SET 3), DRUŻYNY pion, 6 SUMA). Routing app.py: `elif is_dwojka and
  is_pucharowa and Best of 3 → 'DWOJKA_Bo3'`. Format Bo3 odblokowany (Bo5
  🔴 wkrótce + guard st.stop). Fazy drabinki dla dwójki bez ikon. is_supported_type
  += is_dwojka. Blank-form: dodany Bo3. generate_docx: DWOJKA_Bo3 przez branch
  TROJKA (font, dynamiczne wymiary), ale lewa kolumna WĄSKA jak TRÓJKA_Bo3
  (puchar bez grafik — inaczej 2-cyfrowe numery 10-18 zawijały się).
  (c) **TRÓJKA_Bo3.docx** odświeżony (nowy plik od usera). Regresja 11/11 OK.
  (d) **KRYTYCZNY BUG grafiki w DOCX (DWÓJKA Grupa):** w PDF OK, w Wordzie logo
  nachodziło na tabelę. Przyczyna: blok skalowania ustawiał `tblGrid` col0=2700,
  ale per-komórkowe `tcW` col0 tylko gdy `w_int == ORIG_LEFT_COL_DXA` (hardcode
  1186 z TRÓJKI). DWÓJKA ma col0=1192 → warunek nie trafiał → tcW zostawało ~1148.
  LibreOffice (PDF) czyta gridCol (szeroko, OK), Word czyta tcW per komórka
  (wąsko) → tabela przesunięta, grafiki nachodzą. FIX: (1) wymiary
  ORIG_T1/T2_TOTAL + ORIG_LEFT_COL_DXA czytane DYNAMICZNIE z szablonu dla
  TROJKA/DWOJKA/DWOJKA_Bo3; (2) dopasowanie komórki lewej z TOLERANCJĄ 50 dxa
  (grid vs tcW bywają o kilka dxa różne). Audyt wszystkich 11 szablonów:
  brak mismatchu grid-vs-tcW na komórkach bez gridSpan.
  (e) **Logo PFM nachodziło na „Wyniki turnieju"** (DWÓJKA Grupa): `compute_default_positions`
  dostawała `_tpl_type='IND'` dla dwójki (geometria IND), choć tabela jest
  TRÓJKA-style (kolumna 4.76 cm). Fix: `is_dwojka → 'TROJKA'` w obliczaniu
  `_tpl_type` (3 miejsca). Teraz QR → „Wyniki turnieju" → logo bez kolizji.
- 2026-06-01 (fix) — **DWÓJKA Grupa.docx: 2 fixy wizualne w szablonie.**
  (1) „Wyniki turnieju" miało font `Aptos Narrow` → LO bez Aptos brał serif
  fallback. Podmieniony run na `Calibri` (Carlito w prod). (2) Ostatnia
  kolumna `SUMA` (tbl[1] r2 c13) miała tylko `right` border sz=12; trzy
  pozostałe SUMA (c4/c7/c10) mają top+bottom borders. Dodane top sz=12 +
  bottom sz=4 do c13 — pogrubione obramowanie spójne z resztą. Patch
  binarny `word/document.xml` (regex + str.replace, dokładnie 1 trafienie
  per fix).
- 2026-06-01 (fix) — **KRYTYCZNY: DWÓJKA ładowała szablon IND.** Oba bloki
  wyboru `template_type` w app.py (grupowy ~1811, pucharowy/uniwersalny ~2058)
  NIE miały gałęzi DWÓJKI — `else` spychał 2-os. na `'IND'` (przy fazie
  grupowej) → produkcja renderowała szablon indywidualny (IMIONA, 1 kolumna)
  zamiast nowego DWÓJKA (DRUŻYNY pion, 2 kolumny/drużynę, 4 SUMA/SET).
  Mój wcześniejszy test fałszywie „przechodził" bo wołałem `build_document`
  z `template_type='DWOJKA'` ręcznie, omijając mapowanie z app.py. Fix:
  dodane `elif is_dwojka: template_type = 'DWOJKA'` w OBU blokach. Generator
  już był OK (`template_files['DWOJKA']` + branch TROJKA). Zweryfikowane
  renderem przez ścieżkę 'DWOJKA' (DMP 2026): poprawny layout DRUŻYNY.
- 2026-06-01 (fix) — **Rozpiski pomijają placeholder „X".** `_is_placeholder_name`
  rozpoznaje teraz lone `X` (sentinel pustej drużyny, obok `bye`/`Gracz N`);
  `matches_to_player_schedules` filtruje go, więc drużyna „X" nie dostaje
  karty rozpiski. Działa też dla protokołów (skip_placeholders).
- 2026-06-01 — (a) Podmieniony **DWÓJKA Grupa.docx** na właściwy szablon
  użytkownika (poprzedni był klonem TRÓJKA Grupa; nowy ma osobny layout
  tabeli wyników: 14 kolumn, 4 SUMA per SET, etykieta „DRUŻYNY" pionowo
  zamiast „IMIONA"). Build path bez zmian — TROJKA branch obsługuje
  skalowanie po istniejących gridCol/tcW, więc działa od razu.
  (b) **Rozpiski meczowe** rozszerzone na DWÓJKĘ — dodany flag `is_dwojka`
  do bramki `if is_individual or is_trojka or is_czworka` w app.py.
  Builder już wcześniej przyjmował `is_team=True` dla drużynowych —
  nazwy drużyn pogrubione po swojej stronie.
- 2026-05-30 — Iteracja użytkownika TMP 2026: (a) **DWÓJKA Grupa real-data
  fill** działa na xlsx TMP 2026. Istniejący `parse_group_rows` już radzi
  sobie z layoutem 2-os. (każdy mecz = 1 wiersz: tor/godz/z1/score/z2/score)
  — nazwa drużyny w kolumnie z1/z2, brak listy graczy w arkuszu. Build
  ścieżka: DWOJKA dodany do branchy `if template_type in ('TROJKA', …)`
  bo template = identyczny klon TRÓJKA Grupa.docx (font normalizacja,
  skalowanie tabel — wszystko TROJKA). (b) **Bug fix `Mecz # = "1.0"`**:
  arkusz formatuje kolumnę # jako liczbę → gviz/openpyxl daje "1.0". Dodany
  cleanup `r'^\d+\.0+$'` przed wstawieniem do match dict (analogicznie do
  tor). (c) **Rename plików docx** zgodnie z konwencją PFM SharePoint:
  spacje zamiast podkreślników, polskie znaki w nazwach typów drużynowych
  („CZWÓRKA Bo3.docx", „TRÓJKA Grupa.docx", „DWÓJKA Grupa.docx"). Stałe
  Python (IND, TROJKA, DWOJKA…) zostały bez zmian — tylko mapping w
  `template_files` zaktualizowany. (d) **Wywalenie zielonych ✅ ikon**
  z UI: rodzaj turnieju, drabinka, faza, format — wszystkie selectboxy
  pokazują działające opcje BEZ ikony, 🟡 dla testowanych (DWÓJKA Grupa),
  🔴 dla niedostępnych (DWÓJKA puchar). `st.success(…)` z ✅ zostają bo to
  potwierdzenia akcji, nie etykiety statusu. (e) Weryfikacja IND 1/32 i 1/64
  — działają (regex `1/(\d+)` matchuje, parser zwraca poprawne mecze; xlsx
  FMC 2026 ma 24 mecze w 1/32 = 32 par minus 8 bye, czyli OK).
- 2026-05-30 — KRYTYCZNY fix: faza rozbita na KILKA bloków nagłówkowych.
  Bug FMC 2026: „MIEJSCA 5-8" (drabinka o miejsca, 4 zawodników = 2 półfinały)
  była zapisana jako DWA osobne bloki „MIEJSCA 5-8 (16:45)" w tej samej kolumnie
  (puste wiersze między nimi). Stary `parse_drabinka_rows` brał TYLKO pierwszy
  blok (KROK 2 wybierał 1 `chosen`), czytał do nagłówka drugiego (stop-keyword)
  → 1 mecz zamiast 2. Podobnie `detect_drabinka_phases` PASS 2 dedupował po
  `seen_keys` → UI pokazywało „1 mecz". FIX: (1) KROK 2/3 zbiera WSZYSTKIE bloki
  o danym phase_key (`chosen_blocks`) i skleja mecze; (2) PASS 2 sumuje n_matches
  po kluczu zamiast dedupować markery. ROBUSTNOŚĆ (na życzenie usera — błąd
  arkusza, ale uodparniamy): dedup par zawodników w obrębie fazy (frozenset
  z1/z2) — przypadkowo zduplikowany nagłówek+treść daje 1 mecz, nie 2 kopie;
  legit split (2 różne półfinały) nietknięty bo pary się różnią. Hard-cap
  (expected = (Y-X+1)//2 dla MIEJSCA) bez zmian. Weryfikacja: xlsx FMC 2026
  — MIEJSCA 5-8 = 2 mecze (Czech/Wesołowski tor13, Bisaga/Walasik tor14),
  pozostałe fazy bez regresji (1/32=24 z bye, 9-16=4, 17-32=8). Test regresji:
  `_split_rows` + `_dup_rows` w tests/regression.py.
- 2026-05-29 — Drobne fixy UI: (a) `is_ok_pre` w app.py uwzględnia teraz
  CZWORKA (drabinka i Drabinka B labelki dostają ✅ zamiast 🟡 dla 4-os.).
  (b) `fetch_all_player_schedules`: gdy gid_map padnie, fallback A..Z
  zatrzymuje się po 2 z rzędu pustych zakładkach (jeśli już coś znaleziono),
  zamiast brnąć przez całe A-Z. (c) Komunikat progresu rozpisek: gdy
  total jest nieznany (no-gid fallback), nie pokazujemy „/N" — tylko
  „Pobieram zakładki Gr. * — Gr. A…". Progress bar w tym trybie skaluje się
  do heurystycznych 8 grup. (d) DWOJKA — user przygotuje szablon
  samodzielnie; backlog usunięty.
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
