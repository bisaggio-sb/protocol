import streamlit as st, io, re
from PIL import Image
import generate_docx

st.set_page_config(page_title="GP2 – Protokoły meczowe", page_icon="🎯", layout="centered")
st.title("🎯 GP2 – Generator protokołów meczowych")
st.markdown("Podaj link do arkusza Google Sheets, opcjonalnie dodaj grafiki i pobierz gotowy `.docx`.")

def extract_id(url):
    m = re.search(r'/spreadsheets/d/([a-zA-Z0-9_-]+)', url)
    return m.group(1) if m else None

st.header("1. Link do arkusza Google Sheets")
sheets_url = st.text_input("URL arkusza",
    placeholder="https://docs.google.com/spreadsheets/d/XXXX/edit...")

st.header("2. Grafiki (opcjonalnie)")
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

st.header("3. Debug – sprawdź zakładki")
if st.button("🔍 Sprawdź zakładki arkusza"):
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

    with st.spinner("Pobieram dane z grup Gr. A – Gr. P..."):
        try:
            sheets_data = generate_docx.fetch_all_group_sheets(sid)
        except Exception as e:
            st.error(f"Błąd: {e}"); st.stop()

    total = sum(len(m) for _,m in sheets_data)
    st.info(f"Pobrano {len(sheets_data)} grup, {total} meczów.")

    if total == 0:
        st.error("0 meczów. Użyj przycisku 'Sprawdź zakładki' żeby zobaczyć co apka widzi.")
        st.stop()

    with st.spinner(f"Generuję {total} protokołów..."):
        logos_bytes = {k: (f.read(), f.seek(0))[0] for k,f in logos_raw.items() if f}
        docx_bytes = generate_docx.build_document(
            sid, sheets_url.strip(), sheets_data, logos_bytes or None)

    st.success(f"✅ Gotowe! {total} protokołów w {len(sheets_data)} grupach.")
    st.download_button("⬇️ Pobierz GP2_protokoly.docx", data=docx_bytes,
        file_name="GP2_protokoly.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True)

st.divider()
st.caption("Polska Federacja Mölkky · github.com/polska-federacja-molkky/protocol")
