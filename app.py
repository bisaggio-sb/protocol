"""
Generator protokołów meczowych Mölkky
"""
import streamlit as st, io, re
from PIL import Image, ImageDraw, ImageFont
import generate_docx

st.set_page_config(page_title="Protokoły Mölkky", page_icon="🎯", layout="wide")
st.title("🎯 Generator protokołów meczowych Mölkky")
st.markdown("Podaj nazwę turnieju, link do arkusza Google Sheets, opcjonalnie dodaj grafiki — pobierz gotowy `.docx`.")

def extract_id(url):
    m = re.search(r'/spreadsheets/d/([a-zA-Z0-9_-]+)', url)
    return m.group(1) if m else None

# Layout: lewa kolumna z formularzem, prawa z podglądem
col_form, col_preview = st.columns([3, 2])

with col_form:
    st.header("1. Nazwa turnieju")
    tournament_name = st.text_input("Nazwa turnieju",
        value="GP2 2026", placeholder="np. GP2 2026, Mistrzostwa Polski 2026")

    st.header("2. Link do arkusza Google Sheets")
    sheets_url = st.text_input("URL arkusza",
        placeholder="https://docs.google.com/spreadsheets/d/XXXX/edit...",
        help="Arkusz musi być publiczny. Zakładki grup: 'Gr. A', 'Gr. B', ..., 'Gr. P'")

    st.header("3. Kod QR")
    include_qr = st.checkbox("✅ Generuj kod QR z linkiem do arkusza",
                              value=True)

    st.header("4. Grafiki (max 4)")
    NUM_LOGOS = 4
    logo_files = []
    cols_log = st.columns(2)
    for i in range(NUM_LOGOS):
        with cols_log[i % 2]:
            f = st.file_uploader(f"Grafika {i+1}", type=["png","jpg","jpeg"], key=f"logo_{i}")
            logo_files.append(f)

    uploaded = [(i, f) for i, f in enumerate(logo_files) if f is not None]
    if uploaded:
        st.markdown("**Podgląd:**")
        pcols = st.columns(len(uploaded))
        for col, (i, f) in zip(pcols, uploaded):
            with col:
                st.image(Image.open(f), caption=f"Grafika {i+1}", use_container_width=True)
                f.seek(0)

    # ─── Kolejność (rank_priorities lub po prostu lista) ─────────────
    st.header("5. Kolejność elementów")
    elements_available = []
    if include_qr:
        elements_available.append(('qr', 'Kod QR'))
    for i, f in enumerate(logo_files):
        if f is not None:
            elements_available.append((f'logo{i+1}', f'Grafika {i+1}'))

    image_order = []
    if elements_available:
        st.markdown("Każdy element pojawi się od góry do dołu w wybranej kolejności:")
        # Zamiast indywidualnych selectboxów (które pozwalały duplikaty)
        # robimy listę reorder przez przeciąganie
        # Streamlit nie ma natywnego drag&drop, więc używamy multiselect
        # gdzie kolejność ma znaczenie
        labels = [label for _, label in elements_available]
        reordered = st.multiselect(
            "Wybierz kolejność (klikaj kolejno od góry do dołu):",
            options=labels,
            default=labels,
            help="Kolejność klikania = kolejność wyświetlania od góry"
        )
        # Mapuj z powrotem na klucze
        label_to_key = {label: key for key, label in elements_available}
        image_order = [label_to_key[lbl] for lbl in reordered]
        if len(image_order) < len(elements_available):
            st.info(f"Pominięto {len(elements_available) - len(image_order)} elementów - nie pojawią się na protokole.")

    with st.expander("🔍 Debug – sprawdź zakładki"):
        if st.button("Sprawdź zakładki"):
            sid = extract_id(sheets_url.strip()) if sheets_url.strip() else None
            if not sid:
                st.error("Wklej najpierw poprawny link do arkusza.")
            else:
                with st.spinner("Sprawdzam zakładki Gr. A – Gr. P..."):
                    info = generate_docx.get_sheet_names_debug(sid)
                st.code("\n".join(info))

# ─── PRAWA KOLUMNA: Podgląd ─────────────────────────────────────────────────
with col_preview:
    st.header("📄 Podgląd strony")
    st.caption("Tak będzie wyglądał lewy obszar protokołu (uproszczony schemat)")

    # Symulacja podglądu - rysujemy pierwszą stronę protokołu jako PIL Image
    PAGE_W, PAGE_H = 600, 800
    preview = Image.new('RGB', (PAGE_W, PAGE_H), 'white')
    d = ImageDraw.Draw(preview)

    # Marginesy (1.27 cm na 600 px ≈ 36 px)
    LEFT, TOP = 36, 36

    # Tabela 1 (nagłówek mecz)
    d.rectangle([LEFT+50, TOP+30, PAGE_W-LEFT, TOP+50], outline='gray')
    d.text((LEFT+60, TOP+33), "Tor 1   Godzina 09:30   Grupa A   Mecz # 1", fill='black')

    # Tabela 1 dolna część (zawodnicy)
    d.rectangle([LEFT+50, TOP+60, PAGE_W-LEFT, TOP+120], outline='gray')
    d.text((LEFT+250, TOP+68), "Łukasz Szulc", fill='black')
    d.text((LEFT+250, TOP+95), "Anna Ściepuro", fill='black')

    # Linia "Każdy zawodnik..."
    d.text((LEFT+100, TOP+135), "Każdy zawodnik zaczyna po jednym secie...", fill='gray')

    # Tabela wyników
    TBL_TOP = TOP+170
    TBL_BOTTOM = PAGE_H - 60
    d.rectangle([LEFT, TBL_TOP, PAGE_W-LEFT, TBL_BOTTOM], outline='gray')

    # Linia rozdzielająca "Wyniki turnieju" od "IMIONA"
    LEFT_AREA_W = 160  # ~5.24 cm w skali
    d.line([(LEFT+LEFT_AREA_W, TBL_TOP), (LEFT+LEFT_AREA_W, TBL_BOTTOM)], fill='gray')

    # Lewy obszar - rysujemy obrazy
    cur_y = TBL_TOP + 10
    if include_qr:
        # QR placeholder
        qr_size = 60
        qr_x = LEFT + (LEFT_AREA_W - qr_size) // 2
        d.rectangle([qr_x, cur_y, qr_x+qr_size, cur_y+qr_size], outline='black', width=2)
        d.text((qr_x+18, cur_y+22), "QR", fill='black')
        cur_y += qr_size + 5
        # Napis "Wyniki turnieju" pod QR
        d.text((LEFT+30, cur_y), "Wyniki turnieju", fill='black')
        cur_y += 25
    else:
        d.text((LEFT+30, cur_y), "Wyniki turnieju", fill='black')
        cur_y += 25

    # Rysuj logo placeholdery
    logo_colors = ['red', 'orange', 'green', 'blue']
    for i, f in enumerate(logo_files):
        if f is not None and cur_y < TBL_BOTTOM - 50:
            try:
                f.seek(0)
                img = Image.open(f).convert('RGB')
                f.seek(0)
                # Wpisz miniatury w obszar
                logo_w = 60
                ratio = img.width / img.height
                logo_h = min(int(logo_w / ratio), 40)
                logo_w = int(logo_h * ratio)
                logo_x = LEFT + (LEFT_AREA_W - logo_w) // 2
                img_thumb = img.resize((logo_w, logo_h))
                preview.paste(img_thumb, (logo_x, cur_y))
                cur_y += logo_h + 8
            except Exception:
                # fallback
                d.rectangle([LEFT+50, cur_y, LEFT+110, cur_y+30],
                            fill=logo_colors[i % len(logo_colors)])
                cur_y += 35

    # Prawy obszar tabeli wyników (siatka)
    d.text((LEFT+LEFT_AREA_W+30, TBL_TOP+10), "SET 1     SET 2", fill='black')
    # Pusta siatka (uproszczona)
    for i in range(1, 18):
        y = TBL_TOP + 50 + i*30
        if y < TBL_BOTTOM - 30:
            d.line([(LEFT+LEFT_AREA_W, y), (PAGE_W-LEFT, y)], fill='lightgray')
    d.text((LEFT+LEFT_AREA_W+50, TBL_BOTTOM-25), "WYNIK", fill='black')

    st.image(preview, use_container_width=True)

    st.info("📝 To uproszczony podgląd. Dokładny wygląd zobaczysz w pobranym pliku .docx")

# ─── Generuj ────────────────────────────────────────────────────────────────
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
        st.error("0 meczów. Sprawdź czy arkusz jest publiczny."); st.stop()

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
