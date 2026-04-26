"""
GP2 – Generator protokołów meczowych
Streamlit app: wczytuje dane z Google Sheets, generuje .docx
"""

import streamlit as st
import io
import requests
import csv
import re
from PIL import Image
import generate_docx

st.set_page_config(
    page_title="GP2 – Protokoły meczowe",
    page_icon="🎯",
    layout="centered",
)

st.title("🎯 GP2 – Generator protokołów meczowych")
st.markdown("Podaj link do arkusza Google Sheets z danymi turnieju, opcjonalnie dodaj grafiki, i pobierz gotowy .docx ze wszystkimi protokołami.")

# ─── Input ────────────────────────────────────────────────────────────────────

st.header("1. Link do arkusza Google Sheets")
sheets_url = st.text_input(
    "URL arkusza",
    placeholder="https://docs.google.com/spreadsheets/d/XXXX/edit...",
    help="Arkusz musi być publiczny (udostępniony jako 'każdy z linkiem może wyświetlać')"
)

# ─── Grafiki ──────────────────────────────────────────────────────────────────

st.header("2. Grafiki (opcjonalnie)")
st.markdown("Grafiki będą umieszczone na każdym protokole w wybranych miejscach.")

col1, col2 = st.columns(2)
with col1:
    logo_top_left = st.file_uploader("Logo górny lewy róg", type=["png","jpg","jpeg"], key="logo_tl")
    logo_top_right = st.file_uploader("Logo górny prawy róg", type=["png","jpg","jpeg"], key="logo_tr")
with col2:
    logo_bottom_left = st.file_uploader("Logo dolny lewy róg", type=["png","jpg","jpeg"], key="logo_bl")
    banner = st.file_uploader("Banner (szerokość całej strony, góra)", type=["png","jpg","jpeg"], key="banner")

logos = {
    "top_left":     logo_top_left,
    "top_right":    logo_top_right,
    "bottom_left":  logo_bottom_left,
    "banner":       banner,
}

# Podgląd wgranych grafik
uploaded = {k: v for k, v in logos.items() if v is not None}
if uploaded:
    st.markdown("**Podgląd wgranych grafik:**")
    cols = st.columns(len(uploaded))
    labels = {"top_left":"Góra lewo","top_right":"Góra prawo","bottom_left":"Dół lewo","banner":"Banner"}
    for col, (key, file) in zip(cols, uploaded.items()):
        with col:
            img = Image.open(file)
            st.image(img, caption=labels[key], use_container_width=True)
            file.seek(0)

# ─── Generowanie ──────────────────────────────────────────────────────────────

st.header("3. Generuj")

def extract_sheet_id(url):
    m = re.search(r'/spreadsheets/d/([a-zA-Z0-9_-]+)', url)
    return m.group(1) if m else None

if st.button("🚀 Generuj protokoły .docx", type="primary", use_container_width=True):
    if not sheets_url.strip():
        st.error("Podaj link do arkusza Google Sheets.")
        st.stop()

    sheet_id = extract_sheet_id(sheets_url.strip())
    if not sheet_id:
        st.error("Nie można wyciągnąć ID arkusza z podanego linku.")
        st.stop()

    with st.spinner("Pobieram dane z arkusza..."):
        try:
            sheets_data = generate_docx.fetch_all_group_sheets(sheet_id)
        except Exception as e:
            st.error(f"Błąd pobierania danych: {e}")
            st.stop()

    total_matches = sum(len(m) for _, m in sheets_data)
    st.info(f"Pobrano {len(sheets_data)} grup, {total_matches} meczów.")

    if total_matches == 0:
        st.error("Nie znaleziono żadnych meczów. Sprawdź czy arkusz jest publiczny i zawiera zakładki 'Gr. A', 'Gr. B' itp.")
        st.stop()

    with st.spinner(f"Generuję {total_matches} protokołów..."):
        try:
            logos_bytes = {}
            for key, f in logos.items():
                if f is not None:
                    f.seek(0)
                    logos_bytes[key] = f.read()

            docx_bytes = generate_docx.build_document(
                sheet_id=sheet_id,
                sheets_url=sheets_url.strip(),
                sheets_data=sheets_data,
                logos=logos_bytes,
            )
        except Exception as e:
            st.error(f"Błąd generowania dokumentu: {e}")
            raise

    st.success(f"✅ Gotowe! {total_matches} protokołów, {len(sheets_data)} grup.")
    st.download_button(
        label="⬇️ Pobierz GP2_protokoly.docx",
        data=docx_bytes,
        file_name="GP2_protokoly.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True,
    )

st.divider()
st.caption("Polska Federacja Mölkky · github.com/polska-federacja-molkky")
