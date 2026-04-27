"""
Generator protokołów meczowych Mölkky
"""
import streamlit as st, io, re
from PIL import Image
import generate_docx

st.set_page_config(page_title="Protokoły Mölkky", page_icon="🎯", layout="centered")
st.title("🎯 Generator protokołów meczowych Mölkky")
st.markdown("Podaj nazwę turnieju, link do arkusza Google Sheets, opcjonalnie dodaj grafiki — pobierz gotowy `.docx`.")

def extract_id(url):
    m = re.search(r'/spreadsheets/d/([a-zA-Z0-9_-]+)', url)
    return m.group(1) if m else None

st.header("1. Nazwa turnieju")
tournament_name = st.text_input("Nazwa turnieju (pojawi się na stronie tytułowej)",
    value="GP2 2026", placeholder="np. GP2 2026, Mistrzostwa Polski 2026")

st.header("2. Link do arkusza Google Sheets")
sheets_url = st.text_input("URL arkusza",
    placeholder="https://docs.google.com/spreadsheets/d/XXXX/edit...",
    help="Arkusz musi być publiczny. Zakładki grup: 'Gr. A', 'Gr. B', ..., 'Gr. P'")

st.header("3. Grafiki (opcjonalnie)")
col1, col2 = st.columns(2)
with col1:
    logo_tl = st.file_uploader("Logo górny lewy róg",  type=["png","jpg","jpeg"], key="tl")
    logo_bl = st.file_uploader("Logo dolny lewy róg",  type=["png","jpg","jpeg"], key="bl")
with col2:
    logo_tr = st.file_uploader("Logo górny prawy róg", type=["png","jpg","jpeg"], key="tr")
    banner  = st.file_uploader("Banner (góra strony)", type=["png","jpg","jpeg"], key="bn")

logos_raw = {"top_left":logo_tl,"top_right":logo_tr,"bottom_left":logo_bl,"banner":banner}
uploaded = {k:v for k,v in logos_raw.items() if v}
if uploaded:
    cols = st.columns(len(uploaded))
    labels = {"top_left":"Góra lewo","top_right":"Góra prawo","bottom_left":"Dół lewo","banner":"Banner"}
    for col,(key,f) in zip(cols, uploaded.items()):
        with col:
            st.image(Image.open(f), caption=labels[key], use_container_width=True)
            f.seek(0)

with st.expander("🔍 Debug – sprawdź zakładki arkusza"):
    if st.button("Sprawdź zakładki"):
        sid = extract_id(sheets_url.strip()) if sheets_url.strip() else None
        if not sid:
            st.error("Wklej najpierw poprawny link do arkusza.")
        else:
            with st.spinner("Sprawdzam zakładki Gr. A – Gr. P..."):
                info = generate_docx.get_sheet_names_debug(sid)
            st.code("\n".join(info))

st.header("4. Generuj")
if st.button("🚀 Generuj protokoły .docx", type="primary", use_container_width=True):
    if not sheets_url.strip():
        st.error("Podaj link do arkusza."); st.stop()
    sid = extract_id(sheets_url.strip())
    if not sid:
        st.error("Nieprawidłowy link."); st.stop()

    with st.spinner("Pobieram dane z grup..."):
        try:
            sheets_data = generate_docx.fetch_all_group_sheets(sid)
        except Exception as e:
            st.error(f"Błąd pobierania: {e}"); st.stop()

    total = sum(len(m) for _,m in sheets_data)
    st.info(f"Pobrano {len(sheets_data)} grup, {total} meczów.")

    if total == 0:
        st.error("0 meczów. Użyj przycisku Debug żeby sprawdzić zakładki.")
        st.stop()

    with st.spinner(f"Generuję {total} protokołów..."):
        logos_bytes = {k: (f.read(), f.seek(0))[0] for k,f in logos_raw.items() if f}
        docx_bytes = generate_docx.build_document(
            sid, sheets_url.strip(), sheets_data,
            logos_bytes or None,
            tournament_name=tournament_name.strip() or "Turniej Mölkky")

    st.success(f"✅ Gotowe! {total} protokołów w {len(sheets_data)} grupach.")
    safe_name = re.sub(r'[^\w\s-]','', tournament_name).strip().replace(' ','_') or "protokoly"
    st.download_button(f"⬇️ Pobierz {safe_name}.docx", data=docx_bytes,
        file_name=f"{safe_name}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True)

st.divider()
st.caption("Polska Federacja Mölkky · github.com/polska-federacja-molkky/protocol")
