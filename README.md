# GP2 – Generator protokołów meczowych

Aplikacja webowa generująca protokoły meczowe dla turniejów Mölkky GP2.

## Jak używać

1. Wejdź na stronę aplikacji (link poniżej)
2. Wklej link do arkusza Google Sheets z danymi turnieju
3. Opcjonalnie wgraj grafiki (logo, banner)
4. Kliknij **Generuj** i pobierz gotowy `.docx`

## Struktura arkusza Google Sheets

Arkusz musi być **publiczny** (udostępniony jako „każdy z linkiem może wyświetlać").  
Zakładki z meczami grupowymi muszą mieć nazwy `Gr. A`, `Gr. B`, ..., `Gr. P`.

Każda zakładka powinna zawierać kolumny:

| Tor | Godzina | Grupa | Mecz | Zawodnik 1 | Zawodnik 2 |
|-----|---------|-------|------|------------|------------|
| 1   | 09:30   | A     | 1    | Jan Kowalski | Anna Nowak |

## Uruchomienie lokalne

```bash
pip install -r requirements.txt
# Wgraj plik szablonu Grupa_IND.docx do katalogu aplikacji
streamlit run app.py
```

## Deploy na Streamlit Cloud

1. Fork lub push repo na GitHub (np. `polska-federacja-molkky/gp2-protokoly`)
2. Wejdź na [share.streamlit.io](https://share.streamlit.io)
3. Połącz repozytorium i kliknij **Deploy**

Plik `Grupa_IND.docx` (szablon) musi być w repozytorium.

---

*Polska Federacja Mölkky*
