#!/bin/bash
# Instaluje zależności potrzebne do uruchomienia aplikacji (Streamlit + docx→pdf).
# Idempotentny i nieinteraktywny. Uruchamiany tylko w środowisku web (Claude Code on the web).
set -euo pipefail

# Tylko w zdalnym środowisku (web). Lokalnie nic nie robimy.
if [ "${CLAUDE_CODE_REMOTE:-}" != "true" ]; then
  exit 0
fi

PROJECT_DIR="${CLAUDE_PROJECT_DIR:-$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)}"

# 1) Zależności Python (+ openpyxl do lokalnych testów xlsx — patrz przypomnienie niżej).
pip install --quiet --break-system-packages -r "$PROJECT_DIR/requirements.txt" openpyxl

# 2) Pakiety systemowe: LibreOffice do docx→pdf + Carlito (substytut Calibri).
need_pkgs=()
command -v soffice >/dev/null 2>&1 || need_pkgs+=(libreoffice)
# Sam libreoffice-core (bywa w obrazie) NIE wystarcza do konwersji docx→pdf —
# potrzebny moduł Writer, inaczej soffice rzuca "source file could not be loaded".
dpkg -s libreoffice-writer >/dev/null 2>&1 || need_pkgs+=(libreoffice-writer)
# poppler-utils (pdftoppm/pdfinfo) — rasteryzacja PDF→PNG do wizualnej weryfikacji protokołów.
command -v pdftoppm >/dev/null 2>&1 || need_pkgs+=(poppler-utils)
fc-list 2>/dev/null | grep -qi carlito || need_pkgs+=(fonts-crosextra-carlito)
if [ "${#need_pkgs[@]}" -gt 0 ]; then
  export DEBIAN_FRONTEND=noninteractive
  # Nie przerywaj, jeśli któreś zewnętrzne repo (PPA) jest niedostępne —
  # główne archiwum Ubuntu (skąd pochodzą te pakiety) zwykle aktualizuje się OK.
  apt-get update -qq || true
  apt-get install -y --no-install-recommends "${need_pkgs[@]}" \
    || echo "session-start: ostrzeżenie — nie udało się zainstalować: ${need_pkgs[*]}"
fi

echo "session-start: zależności gotowe."
# ── PRZYPOMNIENIE dla Claude (widoczne w transkrypcie sesji) ────────────────
# Produkcja pobiera arkusze przez gviz (Google Sheets) — sandbox NIE ma
# dostępu do docs.google.com. Lokalne testy z xlsx (openpyxl) NIE reprodukują
# kluczowych patologii gviz, m.in.:
#   • gubienie nagłówka 'Tor' gdy kolumna jest czysto numeryczna,
#   • float "1.0"/"2.0" zamiast int "1"/"2".
# Bug M4U 2026 ("Tor=A1/B2") wynikał DOKŁADNIE z tej różnicy — happy path
# z openpyxl pokazał zielone, produkcja była czerwona.
# → ZAWSZE testuj `parse_drabinka_rows` przez `tests/regression.py` (ma
#   `_simulate_gviz_drop_tor_header`). Przy nowych bugach drabinki dorób tam
#   scenariusz nim ogłosisz „naprawione".
echo "session-start: REMINDER → testy drabinki idą przez symulator gviz (regression.py)."
