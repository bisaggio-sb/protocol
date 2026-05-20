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

## Szablony

| Rodzaj turnieju | Szablon docx | Faza grupowa | Faza pucharowa |
|---|---|---|---|
| Indywidualny | `Grupa_IND.docx` | ✅ | ✅ (Bo3 / Bo5) |
| Drużynowy 3-os. | `Grupa_TROJKA.docx` | ✅ | ✅ (Bo3 / Bo5) |
| Drużynowy 4-os. | `Grupa_CZWORKA.docx` | ✅ | 🔜 |

## Uruchomienie lokalne

```bash
git clone https://github.com/polska-federacja-molkky/protocol.git
cd protocol
pip install -r requirements.txt
streamlit run app.py
```

Wymagania:
- Python 3.9+
- Szablony w katalogu: `Grupa_IND.docx`, `Grupa_TROJKA.docx`, `Grupa_CZWORKA.docx`, `TROJKA_Bo3.docx`, `TROJKA_Bo5.docx`
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
├── Grupa_IND.docx          # Szablon indywidualny
├── Grupa_TROJKA.docx       # Szablon 3-osobowy
├── Grupa_CZWORKA.docx      # Szablon 4-osobowy
├── TROJKA_Bo3.docx         # Pucharowa Bo3 (3-os.)
├── TROJKA_Bo5.docx         # Pucharowa Bo5 (3-os.)
├── assets_pfm_logo.png     # Logo Polskiej Federacji Mölkky
├── requirements.txt        # Zależności Python
├── .streamlit/config.toml  # Motyw Streamlit
└── README.md
```

---

*Polska Federacja Mölkky · 2026*
