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

# ─── 1. Nazwa turnieju ────────────────────────────────────────────────────────
st.header("1. Nazwa turnieju")
tournament_name = st.text_input("Nazwa turnieju",
    value="GP2 2026", placeholder="np. GP2 2026, Mistrzostwa Polski 2026")

# ─── 2. Link do arkusza ───────────────────────────────────────────────────────
st.header("2. Link do arkusza Google Sheets")
sheets_url = st.text_input("URL arkusza",
    placeholder="https://docs.google.com/spreadsheets/d/XXXX/edit...",
    help="Arkusz musi być publiczny. Zakładki grup: 'Gr. A', 'Gr. B', ..., 'Gr. P'")

# ─── 3. Kod QR ────────────────────────────────────────────────────────────────
st.header("3. Kod QR z linkiem do arkusza")
include_qr = st.checkbox("✅ Generuj kod QR z linkiem do arkusza wyników",
                          value=True,
                          help="QR pojawi się w lewym obszarze obok tabeli wyników (jak na wzorcu)")

# ─── 4. Grafiki ───────────────────────────────────────────────────────────────
st.header("4. Grafiki (opcjonalnie, max 4)")
st.markdown("Grafiki pojawią się w lewym obszarze obok tabeli wyników, "
            "jedna pod drugą — najpierw QR (jeśli włączony), potem grafiki w wybranej kolejności.")

NUM_LOGOS = 4
logo_files = []
cols = st.columns(2)
for i in range(NUM_LOGOS):
    with cols[i % 2]:
        f = st.file_uploader(f"Grafika {i+1}", type=["png","jpg","jpeg"], key=f"logo_{i}")
        logo_files.append(f)

# Podgląd uploadowanych
uploaded = [(i, f) for i, f in enumerate(logo_files) if f is not None]
if uploaded:
    st.markdown("**Podgląd grafik:**")
    pcols = st.columns(len(uploaded))
    for col, (i, f) in zip(pcols, uploaded):
        with col:
            st.image(Image.open(f), caption=f"Grafika {i+1}", use_container_width=True)
            f.seek(0)

# ─── 5. Kolejność (drag-and-drop substitute) ─────────────────────────────────
st.header("5. Kolejność elementów w lewym obszarze")
st.markdown("Każdy element pojawi się od góry do dołu. Zmień numer porządkowy aby przesunąć element.")

# Zbudujmy listę elementów do uporządkowania
elements_available = []
if include_qr:
    elements_available.append(('qr', 'Kod QR'))
for i, f in enumerate(logo_files):
    if f is not None:
        elements_available.append((f'logo{i+1}', f'Grafika {i+1}'))

if elements_available:
    st.markdown("Pozycje (1 = na górze, większe = niżej):")
    positions = {}
    pcols = st.columns(min(len(elements_available), 4))
    for idx, (key, label) in enumerate(elements_available):
        with pcols[idx % len(pcols)]:
            positions[key] = st.number_input(label, min_value=1, max_value=10,
                                              value=idx+1, key=f"pos_{key}")
    # Sortuj
    image_order = [k for k, _ in sorted(elements_available,
                                          key=lambda x: positions[x[0]])]
else:
    image_order = []

# ─── 6. Debug ─────────────────────────────────────────────────────────────────
with st.expander("🔍 Debug – sprawdź zakładki arkusza"):
    if st.button("Sprawdź zakładki"):
        sid = extract_id(sheets_url.strip()) if sheets_url.strip() else None
        if not sid:
            st.error("Wklej najpierw poprawny link do arkusza.")
        else:
            with st.spinner("Sprawdzam zakładki Gr. A – Gr. P..."):
                info = generate_docx.get_sheet_names_debug(sid)
            st.code("\n".join(info))

# ─── 7. Generuj ───────────────────────────────────────────────────────────────
st.header("6. Generuj")
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
        logos_bytes = {}
        for i, f in enumerate(logo_files):
            if f is not None:
                f.seek(0)
                logos_bytes[f'logo{i+1}'] = f.read()

        docx_bytes = generate_docx.build_document(
            sid, sheets_url.strip(), sheets_data,
            logos=logos_bytes or None,
            tournament_name=tournament_name.strip() or "Turniej Mölkky",
            include_qr=include_qr,
            image_order=image_order or None)

    st.success(f"✅ Gotowe! {total} protokołów w {len(sheets_data)} grupach.")
    safe_name = re.sub(r'[^\w\s-]','', tournament_name).strip().replace(' ','_') or "protokoly"
    st.download_button(f"⬇️ Pobierz {safe_name}.docx", data=docx_bytes,
        file_name=f"{safe_name}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True)

st.divider()
st.caption("Polska Federacja Mölkky · github.com/polska-federacja-molkky/protocol")
