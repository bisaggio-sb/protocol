# Generator protokołów meczowych Mölkky

Aplikacja webowa generująca protokoły meczowe (`.docx`) dla turniejów Mölkky organizowanych przez Polską Federację Mölkky.

🌐 **Aplikacja online:** [protocol.streamlit.app](https://protocol.streamlit.app/)

## Funkcjonalności

- 📊 **Pobieranie danych z arkusza Google Sheets** — automatyczne wczytanie meczów ze wszystkich grup (Gr. A – Gr. P)
- 📄 **Generowanie protokołów .docx** — jeden protokół na stronę, dokładnie według wzorca PFM
- 🔗 **Kod QR** — automatycznie generowany z linka do arkusza wyników (do skanowania telefonem z wydrukowanej karty)
- 🏆 **Logo PFM** — domyślnie dodawane do każdego protokołu (z opcją wyłączenia)
- 🖼️ **Dodatkowe grafiki** — do 4 logo (np. logo klubu organizującego, sponsora, banner turnieju)
- 📐 **Edytor pozycji** — dostosuj pozycję X/Y i rozmiar każdej grafiki w protokole
- 👁️ **Podgląd na żywo** — schemat strony pokazuje rozłożenie elementów przed generowaniem
- 📅 **Nazwa i data turnieju** — wyświetlane małą czcionką w prawym górnym rogu każdego protokołu

## Jak używać

1. Wejdź na [protocol.streamlit.app](https://protocol.streamlit.app/)
2. Wpisz nazwę i datę turnieju
3. Wklej link do arkusza Google Sheets (musi być publiczny)
4. (Opcjonalnie) Wyłącz QR/logo PFM lub dodaj własne grafiki
5. Dostosuj pozycje grafik (lub zostaw domyślne)
6. Kliknij **Generuj protokoły .docx** i pobierz plik

## Struktura arkusza Google Sheets

Arkusz musi być **publiczny** ("każdy z linkiem może wyświetlać").

Zakładki z meczami muszą mieć nazwy `Gr. A`, `Gr. B`, ..., do `Gr. P`.

Każda zakładka powinna zawierać kolumny:

| # | Godzina | Tor | Grupa X | 1.set | 2.set | _Z2_ | 1.set | 2.set |
|---|---------|-----|---------|-------|-------|------|-------|-------|
| 1 | 09:30   | 7   | Jan Kowalski | | | Anna Nowak | | |

Kolumna z nazwą "Grupa X" zawiera nazwiska zawodnika 1, a kolumna 3 pozycje dalej (po `1.set` i `2.set`) zawiera zawodnika 2.

## Stan rozwoju

**Obecnie zaimplementowane:**
- ✅ Turniej indywidualny, faza grupowa (2 sety)

**Planowane (UI już dostępny):**
- 🔜 Turnieje drużynowe (2-, 3-, 4-osobowe)
- 🔜 Faza pucharowa: best of 3, best of 5
- 🔜 Drag & drop pozycjonowania grafik z uchwytami zmiany rozmiaru

## Uruchomienie lokalne

```bash
git clone https://github.com/polska-federacja-molkky/protocol.git
cd protocol
pip install -r requirements.txt
streamlit run app.py
```

Wymagania:
- Python 3.9+
- Pliki w katalogu: `Grupa_IND.docx` (szablon protokołu), `assets_pfm_logo.png` (logo federacji)

## Deploy na Streamlit Cloud

1. Fork lub push na GitHub (np. `polska-federacja-molkky/protocol`)
2. Wejdź na [share.streamlit.io](https://share.streamlit.io)
3. Połącz repozytorium → wskaż `app.py` jako główny plik → **Deploy**

Po zmianach w repozytorium: **Manage app → Reboot app** żeby wyczyścić cache.

## Struktura projektu

```
protocol/
├── app.py                  # Streamlit UI
├── generate_docx.py        # Logika generowania docx
├── Grupa_IND.docx          # Szablon protokołu (1 mecz)
├── assets_pfm_logo.png     # Logo Polskiej Federacji Mölkky
├── requirements.txt        # Zależności Python
├── .streamlit/config.toml  # Konfiguracja motywu Streamlit
└── README.md
```

---

*Polska Federacja Mölkky · 2026*
