# Generator protokołów meczowych Mölkky

Aplikacja webowa generująca protokoły meczowe (`.docx`) dla turniejów Mölkky organizowanych przez Polską Federację Mölkky.

🌐 **Aplikacja online:** [protocol.streamlit.app](https://protocol.streamlit.app/)

## Funkcjonalności

- 📊 **Pobieranie danych z arkusza Google Sheets** — automatyczne wczytanie meczów z faz grupowych (`Gr. A`–`Gr. Z`) i drabinki pucharowej
- 📄 **Trzy szablony** — indywidualny, drużynowy 3-osobowy, drużynowy 4-osobowy
- 🥇 **Faza pucharowa** — detekcja faz z zakładki `Drabinka` (1/8, 1/4, półfinały, finał, mecze o miejsca) + filtr godzinowy
- 🔗 **Kod QR** — link do arkusza wyników do skanowania telefonem z wydrukowanej karty
- 🏆 **Logo PFM** + do **4 dodatkowych grafik** (sponsor, klub, banner)
- 📐 **Edytor pozycji** — X/Y/szerokość/wysokość każdego elementu, z podglądem na żywo
- 👁️ **Podgląd HTML** — schemat strony pokazuje rozłożenie elementów 1:1 z generowanym docx
- 📑 **Filtrowanie** — pojedynczy mecz, cała grupa, dany tor, dane okno czasowe, dana faza

## Jak używać

1. Wejdź na [protocol.streamlit.app](https://protocol.streamlit.app/)
2. Wpisz nazwę i datę turnieju, wybierz rodzaj (indywidualny / 3-os. / 4-os.)
3. Wklej link do publicznego arkusza Google Sheets i kliknij **Wczytaj zakładki**
4. (Opcjonalnie) Wyłącz QR/logo PFM lub dodaj własne grafiki
5. Dostosuj pozycje grafik suwakami pod podglądem (lub zostaw domyślne)
6. Wybierz fazę i zakres (np. konkretna grupa, godzina, faza pucharu)
7. Kliknij **Generuj protokoły .docx** i pobierz plik

## Struktura arkusza Google Sheets

Arkusz musi być **publiczny** ("każdy z linkiem może wyświetlać").

### Faza grupowa
Zakładki nazwane `Gr. A`, `Gr. B`, … `Gr. Z` (akceptowane też `Gr.A`, `Grupa A`).

Każda zakładka zawiera nagłówek z kolumnami: `Tor`, `Godzina`, `Mecz #` (lub `#`/`Lp`/`Nr`), `Grupa X` (gdzie X to litera grupy — komórka zawiera nazwy zawodników 1) i 3 kolumny dalej — zawodnik 2.

### Faza pucharowa (drabinka)
Zakładka nazwana `Drabinka` (akceptowane też `drabinka`, `DRABINKA`). W arkuszu wystarczy oznaczyć nagłówki faz tekstem typu `1/8 (10:00)`, `1/4`, `Półfinał`, `Finał`, `Mecz o 3 miejsce` — kod automatycznie:
- Wykryje wszystkie obecne fazy
- Pogrupuje je po godzinie (jeśli podana w nawiasie)
- Zliczy mecze schodząc w dół kolumny od nagłówka

## Szablony i status

Legenda: bez ikony = działa · 🟡 w testach · 🔴 niedostępne

| Rodzaj turnieju | Faza grupowa | Faza pucharowa (Bo3 / Bo5) | Szablony docx |
|---|---|---|---|
| Indywidualny | działa | działa (Bo3 / Bo5) | `IND Grupa`, `IND Bo3`, `IND Bo5` |
| Drużynowy 2-os. | 🟡 w testach | 🔴 wkrótce | `DWÓJKA Grupa` |
| Drużynowy 3-os. | działa | działa (Bo3 / Bo5) | `TRÓJKA Grupa`, `TRÓJKA Bo3`, `TRÓJKA Bo5` |
| Drużynowy 4-os. | działa | działa (Bo3 / Bo5) | `CZWÓRKA Grupa`, `CZWÓRKA Bo3`, `CZWÓRKA Bo5` |

- **Wszystkie typy (poza 2-os.)** obsługują wszystkie fazy: grupowa, drabinka główna (1/64…finał), drabinka B (mecze o miejsca).
- **4-os. grupowa** ma poziomy strip z QR + logami pod tabelą.

## Uruchomienie lokalne

```bash
git clone https://github.com/polska-federacja-molkky/protocol.git
cd protocol
pip install -r requirements.txt
streamlit run app.py
```

Wymagania:
- Python 3.9+
- Szablony w katalogu: `IND Grupa.docx`, `IND Bo3.docx`, `IND Bo5.docx`, `DWÓJKA Grupa.docx`, `TRÓJKA Grupa.docx`, `TRÓJKA Bo3.docx`, `TRÓJKA Bo5.docx`, `CZWÓRKA Grupa.docx`, `CZWÓRKA Bo3.docx`, `CZWÓRKA Bo5.docx`
- Logo PFM: `assets_pfm_logo.png`

## Branche i workflow

- `main` — produkcja (deployowane na `protocol.streamlit.app`)
- `develop` — integracja, tu lądują nowe zmiany do testów
- `claude/*` — efemeryczne branche z sesji Claude Code (po review mergowane do `develop`)

**Przepływ zmiany dev → prod:**

```bash
git checkout main
git pull origin main
git merge develop
git push origin main
```

Lub przez PR na GitHubie: `develop → main`, review, merge. Po pushu na `main`: Streamlit Cloud sam odświeży aplikację (zwykle 1-2 min). W razie potrzeby: **Manage app → Reboot app**.

## Deploy na Streamlit Cloud

1. Połącz repozytorium z [share.streamlit.io](https://share.streamlit.io)
2. Wskaż `app.py` jako entry point
3. Branch: `main` (produkcja) lub `develop` (preview)
4. **Deploy**

## Struktura projektu

```
protocol/
├── app.py                  # Streamlit UI + podgląd HTML
├── generate_docx.py        # Generowanie docx, pobieranie z Google Sheets, detekcja drabinki
├── IND Grupa.docx          # Szablon indywidualny (grupowa)
├── IND Bo3.docx            # Indywidualny pucharowa Best of 3
├── IND Bo5.docx            # Indywidualny pucharowa Best of 5
├── DWÓJKA Grupa.docx       # Szablon 2-osobowy (grupowa) — w testach
├── TRÓJKA Grupa.docx       # Szablon 3-osobowy (grupowa)
├── TRÓJKA Bo3.docx         # 3-os. pucharowa Best of 3
├── TRÓJKA Bo5.docx         # 3-os. pucharowa Best of 5
├── CZWÓRKA Grupa.docx      # Szablon 4-osobowy (grupowa)
├── CZWÓRKA Bo3.docx        # 4-os. pucharowa Best of 3
├── CZWÓRKA Bo5.docx        # 4-os. pucharowa Best of 5
├── assets_pfm_logo.png     # Logo Polskiej Federacji Mölkky
├── requirements.txt        # Zależności Python
├── .streamlit/config.toml  # Motyw Streamlit
└── README.md
```

---

*Polska Federacja Mölkky · 2026*
