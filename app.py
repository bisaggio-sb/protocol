"""
Generator protokołów meczowych Mölkky
Polska Federacja Mölkky · github.com/polska-federacja-molkky/protocol
"""
import streamlit as st
import io, re, base64, os, subprocess, tempfile
from datetime import date, timedelta
from PIL import Image
import generate_docx

st.set_page_config(page_title="Protokoły Mölkky", page_icon="🎯", layout="wide")

APP_DIR = os.path.dirname(os.path.abspath(__file__))
PFM_PATH = os.path.join(APP_DIR, 'assets_pfm_logo.png')


def img_to_data_url(file_or_bytes):
    if hasattr(file_or_bytes, 'seek'):
        file_or_bytes.seek(0)
        img_bytes = file_or_bytes.read()
        file_or_bytes.seek(0)
    else:
        img_bytes = file_or_bytes
    img = Image.open(io.BytesIO(img_bytes)).convert("RGBA")
    buf = io.BytesIO()
    img.thumbnail((400, 400))
    img.save(buf, format="PNG")
    b64 = base64.b64encode(buf.getvalue()).decode()
    return f"data:image/png;base64,{b64}"


def get_image_aspect(file_or_bytes):
    """Zwraca aspect ratio (width/height) obrazu."""
    if hasattr(file_or_bytes, 'seek'):
        file_or_bytes.seek(0)
        img_bytes = file_or_bytes.read()
        file_or_bytes.seek(0)
    else:
        img_bytes = file_or_bytes
    img = Image.open(io.BytesIO(img_bytes))
    return img.width / img.height


# Header z logo PFM
if os.path.exists(PFM_PATH):
    with open(PFM_PATH, 'rb') as fp:
        pfm_data_url = img_to_data_url(fp.read())
        # Aspect dla PFM logo (562x509 ≈ 1.10)
    pfm_aspect = get_image_aspect(open(PFM_PATH, 'rb').read())
else:
    pfm_data_url = ""
    pfm_aspect = 1.10

st.markdown(f"""
<div style="display:flex; align-items:center; gap:16px; margin-bottom:8px;">
  <img src="{pfm_data_url}" style="height:64px; width:auto; flex-shrink:0;"/>
  <div>
    <h1 style="margin:0; padding:0; font-size:2.0rem; line-height:1.2;">
      Generator protokołów meczowych Mölkky
    </h1>
    <p style="margin:4px 0 0 0; color:#666; font-size:0.95rem;">
      Wygeneruj protokoły meczowe z arkusza Google Sheets.
    </p>
  </div>
</div>
""", unsafe_allow_html=True)

st.divider()


def extract_id(url):
    m = re.search(r'/spreadsheets/d/([a-zA-Z0-9_-]+)', url)
    return m.group(1) if m else None


# Obszar "Wyniki turnieju" ma szerokość 5.24 cm
LEFT_AREA_CM = 5.24    # Indywidualny: szerokość lewej kolumny tabeli wynikowej (R1.tc[0] = 2970 dxa)
TROJKA_LEFT_AREA_CM = 2.0  # Trójka: lewa kolumna jest dużo węższa (1186 dxa = 2.09 cm)
TROJKA_AREA_HEIGHT_CM = 17.5  # Trójka: lewa kolumna tabeli wynikowej, z R1 do R21


def compute_default_positions(active_keys, logos_aspect=None,
                              area_height_cm=17.0, area_width_cm=LEFT_AREA_CM,
                              template_type='IND'):
    """Rozkłada elementy w lewym obszarze równomiernie z odstępami.
    QR (jeśli jest) na górze. PFM ~109% szerokości QR. Pozostałe grafiki
    równomiernie rozłożone z zachowaniem proporcji obrazu.
    
    Indywidualny:
      area_height_cm=17.0, area_width_cm=5.24 - duży obszar po lewej.
    Trójka:
      area_width_cm=2.0 (znacznie węższy lewy pasek), area_height_cm=17.5.
      Wszystkie obrazy są dużo mniejsze (QR ~1.6 cm zamiast 2.4 cm).
    
    Z layoutInCell=0 obrazy mogą wystawać poza komórkę kotwicy R1
    i ciągnąć się aż do dołu tabeli."""
    is_trojka = (template_type == 'TROJKA')
    
    # Trójka ma węższy lewy pas - mniejsze rozmiary domyślne
    if is_trojka:
        area_width_cm = TROJKA_LEFT_AREA_CM
        area_height_cm = TROJKA_AREA_HEIGHT_CM
        QR_W = 1.6  # zamiast 2.4
        PFM_TARGET_W = QR_W * 1.09  # ~1.74 cm
        OTHER_MAX_W = 1.7   # bardzo wąsko
        SPACE_AFTER_QR = 0.4
    else:
        QR_W = 2.4
        PFM_TARGET_W = QR_W * 1.09  # ~2.6
        OTHER_MAX_W = 4.0
        SPACE_AFTER_QR = 0.5

    positions = {}
    
    other_keys = [k for k in active_keys if k != 'qr']
    
    # ── QR na górze (jeśli aktywny)
    if 'qr' in active_keys:
        qr_x = (area_width_cm - QR_W) / 2
        positions['qr'] = {'x': qr_x, 'y': 0.2, 'w': QR_W, 'h': QR_W}
        # Po QR: napis "Wyniki turnieju" + odstęp
        # W trójce napis jest mniejszy, więc krótszy odstęp
        text_h = 0.4 if is_trojka else 0.5
        first_logo_y = 0.2 + QR_W + text_h + SPACE_AFTER_QR
        pfm_w_target = PFM_TARGET_W
    else:
        first_logo_y = 0.2
        pfm_w_target = PFM_TARGET_W
    
    if not other_keys:
        return positions
    
    n = len(other_keys)
    available = area_height_cm - first_logo_y
    
    # Adaptacyjny SPACING — w trójce mniejszy bo wszystkie elementy są mniejsze
    if is_trojka:
        # Trójka: wąski pas, więc mniejsze odstępy
        if n <= 2:
            SPACING = 1.0
        elif n <= 4:
            SPACING = 0.6
        else:
            SPACING = 0.4
        max_slot_h_default_max = 2.2
        max_slot_h_default_min = 1.0
    else:
        # Indywidualny: duży obszar
        if n <= 2:
            SPACING = 1.5
        elif n <= 4:
            SPACING = 1.0
        else:
            SPACING = 0.6
        max_slot_h_default_max = 3.5
        max_slot_h_default_min = 1.5
    
    if n == 1:
        max_slot_h = min(max_slot_h_default_max + 0.5, available)
    else:
        max_slot_h = (available - (n - 1) * SPACING) / n
        max_slot_h = max(max_slot_h_default_min, min(max_slot_h_default_max, max_slot_h))
    
    cur_y = first_logo_y
    for key in other_keys:
        # Określ proporcje obrazu
        if key == 'pfm':
            aspect = pfm_aspect  # ~1.10
            target_w = pfm_w_target
        elif logos_aspect and key in logos_aspect:
            aspect = logos_aspect[key]
            target_w = min(OTHER_MAX_W, area_width_cm - 0.2)
        else:
            aspect = 1.5
            target_w = min(OTHER_MAX_W - 0.5, area_width_cm - 0.2)
        
        target_h = target_w / aspect
        
        if target_h > max_slot_h:
            target_h = max_slot_h
            target_w = target_h * aspect
            if target_w > area_width_cm - 0.2:
                target_w = area_width_cm - 0.2
                target_h = target_w / aspect
        
        slot_top = cur_y
        img_y = slot_top + (max_slot_h - target_h) / 2
        img_x = (area_width_cm - target_w) / 2
        
        positions[key] = {'x': img_x, 'y': img_y, 'w': target_w, 'h': target_h}
        cur_y += max_slot_h + SPACING
    
    return positions


# ════════════════════════════════════════════════════════════════════════
col_form, col_preview = st.columns([3, 2])
# ════════════════════════════════════════════════════════════════════════

with col_form:

    # ─── 1. Turniej ──────────────────────────────────────────────────────
    st.header("1. Turniej")
    cols_t1 = st.columns([2, 1])
    with cols_t1[0]:
        tournament_name = st.text_input("Nazwa", value="GP2 2026",
                                        placeholder="np. GP2 2026")
    with cols_t1[1]:
        next_sat = date.today() + timedelta(days=(5 - date.today().weekday()) % 7)
        tournament_date_d = st.date_input("Data", value=next_sat, format="DD.MM.YYYY")
    tournament_date = tournament_date_d.strftime("%d.%m.%Y") if tournament_date_d else ""

    # 3 dropdowny w 3 kolumnach (kompaktowo)
    # ✅ = w pełni zweryfikowane na danych z prawdziwych turniejów
    # 🟡 = zaimplementowane, ale nie zweryfikowane na realnych danych
    # 🔴 = jeszcze nieobsługiwane (zostanie pominięte z błędem)
    cols_t2 = st.columns(3)
    with cols_t2[0]:
        rodzaj_options = {
            "Indywidualny": "✅ Indywidualny",
            "Drużynowy 2-os.": "🔴 Drużynowy 2-os. (wkrótce)",
            "Drużynowy 3-os.": "🟡 Drużynowy 3-os.",
            "Drużynowy 4-os.": "🔴 Drużynowy 4-os. (wkrótce)",
        }
        tournament_type = st.selectbox(
            "Rodzaj",
            list(rodzaj_options.keys()),
            format_func=lambda k: rodzaj_options[k],
            index=0,
            help="✅ = w pełni zweryfikowane, 🟡 = zaimplementowane lecz niezweryfikowane, 🔴 = w trakcie przygotowywania"
        )
    with cols_t2[1]:
        faza_options = {
            "Grupowa": "✅ Grupowa",
            "Pucharowa 1/64": "🟡 Pucharowa 1/64",
            "Pucharowa 1/32": "✅ Pucharowa 1/32",
            "Pucharowa 1/16": "🟡 Pucharowa 1/16",
            "Pucharowa 1/8": "🟡 Pucharowa 1/8",
            "Pucharowa 1/4": "🟡 Pucharowa 1/4",
            "Pucharowa 1/2": "🟡 Pucharowa 1/2 (Półfinał)",
            "Mecz o 3. miejsce": "🟡 Mecz o 3. miejsce",
            "Finał": "🟡 Finał",
        }
        tournament_phase = st.selectbox(
            "Faza",
            list(faza_options.keys()),
            format_func=lambda k: faza_options[k],
            index=0,
            help="Zweryfikowane: Grupowa, Pucharowa 1/32"
        )
    with cols_t2[2]:
        # Format setów zależny od fazy:
        # Grupowa → tylko "2 sety"
        # Pucharowa → "best of 3" lub "best of 5"
        if tournament_phase == "Grupowa":
            sets_format = st.selectbox("Format", ["2 sety"],
                format_func=lambda k: f"✅ {k}",
                index=0,
                help="Faza grupowa zawsze 2 sety")
        else:
            bo_options = {"Best of 3": "🔴 Best of 3 (wkrótce)",
                         "Best of 5": "🔴 Best of 5 (wkrótce)",
                         "2 sety": "✅ 2 sety"}
            sets_format = st.selectbox("Format",
                list(bo_options.keys()),
                format_func=lambda k: bo_options[k],
                index=2,  # 2 sety jako sprawdzone
                help="Obecnie zweryfikowane: 2 sety. Best of 3/5 wkrótce."
            )

    # Walidacja - obecnie sprawdzone i zweryfikowane:
    # 1. Indywidualny + Grupowa + 2 sety
    # 2. Indywidualny + Pucharowa 1/32 + 2 sety (z zakładki Drabinka)
    # Pozostałe fazy pucharowe (1/64, 1/16, 1/8, 1/4, 1/2, 3. miejsce, Finał) są
    # zaimplementowane w kodzie ale nie zostały zweryfikowane na rzeczywistych
    # danych — pokazujemy ostrzeżenie.
    is_individual = tournament_type == "Indywidualny"
    is_2_sety = sets_format == "2 sety"
    is_grupowa = tournament_phase == "Grupowa"
    is_132 = tournament_phase == "Pucharowa 1/32"
    is_pucharowa = "Pucharowa" in tournament_phase or tournament_phase in (
        "Mecz o 3. miejsce", "Finał")
    
    is_fully_tested = is_individual and is_2_sety and (is_grupowa or is_132)
    
    # Krótka etykieta fazy do wyświetlenia (komórka tc[5] = 1275 dxa = 2.25 cm).
    # Używana zarówno w preview HTML jak i w docx.
    phase_label_short = ""
    if is_pucharowa:
        m = re.search(r'(\d+/\d+)', tournament_phase)
        if m:
            phase_label_short = "Półfinał" if m.group(1) == "1/2" else m.group(1)
        elif "3. miejsce" in tournament_phase:
            phase_label_short = "3. miejsce"
        elif tournament_phase == "Finał":
            phase_label_short = "Finał"
    
    is_trojka = tournament_type == "Drużynowy 3-os."
    is_supported_type = is_individual or is_trojka
    
    if not is_supported_type:
        st.warning(f"⚠️ Format **{tournament_type}** nie jest jeszcze obsługiwany. "
                   "Obecnie dostępne: Indywidualny, Drużynowy 3-os.")
    elif not is_2_sety:
        st.warning("⚠️ Obecnie zaimplementowane są tylko mecze na **2 sety**. "
                   "Best of 3/5 będą dostępne w kolejnej wersji.")
    elif is_trojka and (is_pucharowa or not is_grupowa):
        # Trójkowy zweryfikowany tylko dla grupowej
        st.warning(f"⚠️ Drużynowy 3-os. obecnie zweryfikowany **tylko dla fazy grupowej**. "
                   f"Faza **{tournament_phase}** może działać niepoprawnie.")
    elif is_individual and is_pucharowa and not is_132:
        # Indywidualny + 2 sety + faza pucharowa inna niż 1/32 → eksperymentalne
        st.warning(f"⚠️ Faza **{tournament_phase}** nie została jeszcze zweryfikowana "
                   "na rzeczywistych danych. Generator może działać poprawnie, ale "
                   "zalecamy sprawdzenie wyniku. **Zweryfikowane są tylko: Grupowa, "
                   "Pucharowa 1/32**.")
    
    if is_pucharowa and is_2_sety and is_individual:
        st.info(f"ℹ️ Wczytam mecze z fazy **{tournament_phase}** z zakładki **Drabinka**.")

    # ─── 2. Link do arkusza ─────────────────────────────────────────────
    st.header("2. Link do arkusza Google Sheets")
    cols_link = st.columns([4, 1])
    with cols_link[0]:
        sheets_url = st.text_input("URL arkusza",
            placeholder="https://docs.google.com/spreadsheets/d/XXXX/edit...",
            help="Arkusz musi być publiczny. Zakładki grup: 'Gr. A', 'Gr. B', ...",
            label_visibility="collapsed")
    with cols_link[1]:
        check_clicked = st.button("🔍 Sprawdź zakładki", use_container_width=True)
    
    if check_clicked:
        sid = extract_id(sheets_url.strip()) if sheets_url.strip() else None
        if not sid:
            st.error("Wklej najpierw poprawny link do arkusza.")
        else:
            with st.spinner("Sprawdzam zakładki Gr. A – Gr. Z..."):
                info = generate_docx.get_sheet_names_debug(sid)
            st.code("\n".join(info))

    # ─── 3. Domyślne elementy ───────────────────────────────────────────
    st.header("3. Domyślne elementy")
    cols_dom = st.columns(2)
    with cols_dom[0]:
        include_qr = st.checkbox("Kod QR (link do arkusza)", value=True)
        show_header_on_protocol = st.checkbox(
            "Nazwa i data turnieju w prawym górnym rogu", value=True)
    with cols_dom[1]:
        include_pfm_logo = st.checkbox("Logo Polskiej Federacji Mölkky", value=True)

    # ─── 4. Dodatkowe grafiki (zwijane) ─────────────────────────────────
    NUM_LOGOS = 4
    logo_files = []
    
    # Sprawdź ile grafik już dodanych żeby pokazać liczbę
    n_uploaded = sum(1 for k in [f"logo_{i}" for i in range(NUM_LOGOS)] 
                     if k in st.session_state and st.session_state[k] is not None)
    
    expander_label = "4. Dodatkowe grafiki"
    if n_uploaded > 0:
        expander_label += f" ({n_uploaded} dodano)"
    else:
        expander_label += " (max 4)"
    
    with st.expander(expander_label, expanded=(n_uploaded > 0)):
        cols_log = st.columns(2)
        for i in range(NUM_LOGOS):
            with cols_log[i % 2]:
                f = st.file_uploader(f"Grafika {i+1}", type=["png","jpg","jpeg"],
                                     key=f"logo_{i}", label_visibility="visible")
                logo_files.append(f)

    # ─── Aktywne elementy ────────────────────────────────────────────────
    elements_active = []
    if include_qr:
        elements_active.append(('qr', 'Kod QR'))
    if include_pfm_logo:
        elements_active.append(('pfm', 'Logo PFM'))
    for i, f in enumerate(logo_files):
        if f is not None:
            elements_active.append((f'logo{i+1}', f'Grafika {i+1}'))

    # Aspect ratios uploadowanych grafik
    logos_aspect = {}
    for i, f in enumerate(logo_files):
        if f is not None:
            try:
                f.seek(0)
                logos_aspect[f'logo{i+1}'] = get_image_aspect(f.read())
                f.seek(0)
            except Exception:
                logos_aspect[f'logo{i+1}'] = 1.5

    # Pozycje (uwzględniając typ szablonu - trójka ma węższy lewy pasek)
    default_pos = compute_default_positions([k for k, _ in elements_active],
                                            logos_aspect=logos_aspect,
                                            template_type=('TROJKA' if is_trojka else 'IND'))
    active_keys_signature = "|".join(sorted(k for k, _ in elements_active))
    
    image_positions = {}
    for key in default_pos:
        image_positions[key] = dict(default_pos[key])

    if elements_active:
        st.header("5. Pozycja i rozmiar elementów")
        cols_h5 = st.columns([4, 1])
        with cols_h5[0]:
            st.caption("Elementy są ułożone automatycznie. "
                       "Rozwiń element żeby dostosować pozycję ręcznie.")
        with cols_h5[1]:
            if st.button("↻ Rozmieść równomiernie",
                         help="Przywróć automatyczne rozmieszczenie"):
                st.session_state['reset_nonce'] = st.session_state.get('reset_nonce', 0) + 1
                st.rerun()

        nonce = st.session_state.get('reset_nonce', 0)
        for key, label in elements_active:
            with st.expander(f"📍 {label}", expanded=False):
                cols = st.columns(4)
                pos = image_positions[key]
                slider_key_base = f"{key}_{active_keys_signature}_{nonce}"
                with cols[0]:
                    new_x = st.number_input("X (cm)", value=float(pos['x']),
                                           min_value=0.0, max_value=5.0, step=0.1,
                                           key=f"x_{slider_key_base}")
                with cols[1]:
                    new_y = st.number_input("Y (cm)", value=float(pos['y']),
                                           min_value=0.0, max_value=17.0, step=0.1,
                                           key=f"y_{slider_key_base}")
                with cols[2]:
                    new_w = st.number_input("Szer. (cm)", value=float(pos['w']),
                                           min_value=0.5, max_value=5.0, step=0.1,
                                           key=f"w_{slider_key_base}")
                with cols[3]:
                    new_h = st.number_input("Wys. (cm)", value=float(pos['h']),
                                           min_value=0.5, max_value=5.0, step=0.1,
                                           key=f"h_{slider_key_base}")
                image_positions[key] = {
                    'x': new_x, 'y': new_y, 'w': new_w, 'h': new_h
                }


# ─── PODGLĄD HTML/CSS po prawej (przesunięty 5cm w dół) ───────────────
with col_preview:
    st.markdown("<div style='height:80px;'></div>", unsafe_allow_html=True)
    st.header("📄 Podgląd strony")
    st.caption("Schemat dokumentu w skali")

    SCALE = 22
    PAGE_W_CM = 18.46
    PAGE_H_CM = 27.16
    PAGE_W_PX = int(PAGE_W_CM * SCALE)
    PAGE_H_PX = int(PAGE_H_CM * SCALE)
    LEFT_AREA_PX = int(LEFT_AREA_CM * SCALE)

    image_urls = {}
    if include_qr:
        image_urls['qr'] = None
    if include_pfm_logo and pfm_data_url:
        image_urls['pfm'] = pfm_data_url
    for i, f in enumerate(logo_files):
        if f is not None:
            image_urls[f'logo{i+1}'] = img_to_data_url(f)

    elements_html = ""
    for key in image_urls:
        if key not in image_positions:
            continue
        pos = image_positions[key]
        x_px = int(pos['x'] * SCALE)
        y_px = int(pos['y'] * SCALE)
        w_px = int(pos['w'] * SCALE)
        h_px = int(pos['h'] * SCALE)
        if key == 'qr':
            elements_html += f"""
            <div style="position:absolute; left:{x_px}px; top:{y_px}px; 
                        width:{w_px}px; height:{h_px}px; 
                        background:white; border:2px solid #333;
                        display:flex; align-items:center; justify-content:center;
                        font-size:11px; color:#333; font-family:monospace;">
              <div style="text-align:center;"><div style="font-weight:bold;">QR</div></div>
            </div>"""
        else:
            elements_html += f"""
            <img src="{image_urls[key]}" 
                 style="position:absolute; left:{x_px}px; top:{y_px}px;
                        width:{w_px}px; height:{h_px}px; object-fit:contain;
                        border:1px dashed #ccc;"/>"""

    label_y_px = 5
    if include_qr and 'qr' in image_positions:
        qr_pos = image_positions['qr']
        label_y_px = int((qr_pos['y'] + qr_pos['h'] + 0.1) * SCALE)

    header_text = ""
    if show_header_on_protocol and (tournament_name or tournament_date):
        parts = [tournament_name] if tournament_name else []
        if tournament_date:
            parts.append(tournament_date)
        header_text = " · ".join(parts)

    TBL1_OFFSET_FRAC = 600 / 9690
    TBL1_OFFSET_PX = int(PAGE_W_PX * TBL1_OFFSET_FRAC)
    NAMES_FRAC = 3915 / 9090
    NAMES_PX = int((PAGE_W_PX - TBL1_OFFSET_PX) * NAMES_FRAC)

    html = f"""
    <div style="background:white; border:1px solid #ccc; 
                width:{PAGE_W_PX}px; height:{PAGE_H_PX}px;
                position:relative; font-family:Arial, sans-serif;
                box-shadow:0 2px 8px rgba(0,0,0,0.1); margin:0 auto;">
      
      <div style="position:absolute; right:8px; top:6px; 
                  font-size:9px; color:#666; font-style:italic;">
        {header_text}
      </div>
      
      <div style="position:absolute; left:{TBL1_OFFSET_PX}px; top:30px; right:0; height:30px;
                  display:flex; align-items:center; padding-left:6px;
                  font-size:11px; gap:18px;">
        <span>Tor <b>1</b></span>
        <span>Godzina <b>09:30</b></span>
        {f'<span>Faza <b>{phase_label_short}</b></span>' if is_pucharowa else '<span>Grupa <b>A</b></span>'}
        <span>Mecz # <b>1</b></span>
      </div>
      
      <div style="position:absolute; left:{TBL1_OFFSET_PX + NAMES_PX}px; top:65px; right:0; height:32px;
                  background:#f0f0f0; border:1px solid #999;
                  display:flex; font-size:9px; font-weight:bold; text-align:center;
                  align-items:center; justify-content:space-around;">
        <div style="flex:1; border-right:1px solid #999;">Punkty<br>SET 1</div>
        <div style="flex:1; border-right:1px solid #999;">Punkty<br>SET 2</div>
        <div style="flex:1.15; border-right:1px solid #999;">Wygrane<br>sety</div>
        <div style="flex:1.95;">Podpis</div>
      </div>
      
      <div style="position:absolute; left:{TBL1_OFFSET_PX}px; top:97px; right:0; height:24px;
                  border:1px solid #999; display:flex; font-size:11px;">
        <div style="flex:0 0 {NAMES_PX}px; padding-right:8px; 
                    text-align:right; line-height:24px; font-weight:bold;
                    border-right:1px solid #999;">
          Łukasz Szulc
        </div>
        <div style="flex:1; border-right:1px solid #999;"></div>
        <div style="flex:1; border-right:1px solid #999;"></div>
        <div style="flex:1.15; border-right:1px solid #999;"></div>
        <div style="flex:1.95;"></div>
      </div>
      <div style="position:absolute; left:{TBL1_OFFSET_PX}px; top:121px; right:0; height:24px;
                  border:1px solid #999; border-top:none; display:flex; font-size:11px;">
        <div style="flex:0 0 {NAMES_PX}px; padding-right:8px;
                    text-align:right; line-height:24px; font-weight:bold;
                    border-right:1px solid #999;">
          Anna Ściepuro
        </div>
        <div style="flex:1; border-right:1px solid #999;"></div>
        <div style="flex:1; border-right:1px solid #999;"></div>
        <div style="flex:1.15; border-right:1px solid #999;"></div>
        <div style="flex:1.95;"></div>
      </div>
      
      <div style="position:absolute; left:0; right:0; top:148px;
                  text-align:center; font-size:8px; color:#555; font-style:italic;">
        Każdy zawodnik zaczyna po jednym secie (w dowolnej kolejności)<br>
        Set przegrany przez 3 kolejne chybienia oznacza wynik 0:50
      </div>
      
      <div style="position:absolute; left:0; top:180px; right:0; height:375px;
                  border:1px solid #999;">
        <div style="position:absolute; left:0; top:0; bottom:0;
                    width:{LEFT_AREA_PX}px; border-right:1px solid #999;
                    overflow:hidden;">
          {elements_html}
          {f'''<div style="position:absolute; left:0; right:0; top:{label_y_px}px;
                      text-align:center; font-size:10px; font-weight:bold;">
            Wyniki turnieju
          </div>''' if include_qr else ''}
        </div>
        <div style="position:absolute; left:{LEFT_AREA_PX+1}px; top:0; right:0; bottom:0;">
          <div style="position:absolute; left:0; top:0; right:0; height:22px;
                      background:#f0f0f0; display:flex; font-size:10px; font-weight:bold;
                      align-items:center; text-align:center; border-bottom:1px solid #999;">
            <div style="flex:0.84; border-right:1px solid #999;"></div>
            <div style="flex:2.94; border-right:1px solid #999;">SET 1</div>
            <div style="flex:2.94;">SET 2</div>
          </div>
          <div style="position:absolute; left:0; top:22px; right:0; height:30px;
                      background:#f8f8f8; display:flex; font-size:8px; font-weight:bold;
                      align-items:center; text-align:center; 
                      border-bottom:1px solid #999;">
            <div style="flex:0.84; border-right:1px solid #999;
                        writing-mode:vertical-rl; transform:rotate(180deg);
                        line-height:1.2;">IMIONA</div>
            <div style="flex:0.735; border-right:1px solid #999;"></div>
            <div style="flex:0.735; border-right:1px solid #999;
                        writing-mode:vertical-rl; transform:rotate(180deg);
                        line-height:1.2;">SUMA</div>
            <div style="flex:0.735; border-right:1px solid #999;"></div>
            <div style="flex:0.735; border-right:1px solid #999;
                        writing-mode:vertical-rl; transform:rotate(180deg);
                        line-height:1.2;">SUMA</div>
            <div style="flex:0.735; border-right:1px solid #999;"></div>
            <div style="flex:0.735; border-right:1px solid #999;
                        writing-mode:vertical-rl; transform:rotate(180deg);
                        line-height:1.2;">SUMA</div>
            <div style="flex:0.735; border-right:1px solid #999;"></div>
            <div style="flex:0.735;
                        writing-mode:vertical-rl; transform:rotate(180deg);
                        line-height:1.2;">SUMA</div>
          </div>
          <div style="position:absolute; left:0; top:52px; right:0; bottom:24px;
                      background:repeating-linear-gradient(
                        to bottom, transparent 0, transparent 17px,
                        #ccc 17px, #ccc 18px);"></div>
          <div style="position:absolute; left:0; right:0; bottom:0; height:24px;
                      border-top:1px solid #999;
                      background:#f0f0f0; display:flex;
                      font-size:10px; font-weight:bold; text-align:center;
                      align-items:center;">
            <div style="flex:0.84; border-right:1px solid #999; line-height:24px;">WYNIK</div>
            <div style="flex:7.46;"></div>
          </div>
        </div>
      </div>
    </div>
    """
    st.components.v1.html(html, height=PAGE_H_PX + 30, scrolling=True)
    st.caption("📝 Podgląd schematyczny. Dokładny wygląd w pobranym pliku.")


# ─── Helper: konwersja docx → pdf ───────────────────────────────────────
def docx_to_pdf(docx_bytes, name):
    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            docx_path = os.path.join(tmpdir, f"{name}.docx")
            with open(docx_path, 'wb') as f:
                f.write(docx_bytes)
            # Filtr PDF z opcjami zmniejszającymi rozmiar
            pdf_filter = ('pdf:writer_pdf_Export:'
                          '{"SelectPdfVersion":{"type":"long","value":"0"},'
                          '"EmbedStandardFonts":{"type":"boolean","value":"false"},'
                          '"ReduceImageResolution":{"type":"boolean","value":"true"},'
                          '"MaxImageResolution":{"type":"long","value":"150"},'
                          '"UseLosslessCompression":{"type":"boolean","value":"false"},'
                          '"Quality":{"type":"long","value":"90"}}')
            result = subprocess.run(
                ['libreoffice', '--headless', '--convert-to', pdf_filter,
                 '--outdir', tmpdir, docx_path],
                capture_output=True, text=True, timeout=300
            )
            pdf_path = os.path.join(tmpdir, f"{name}.pdf")
            if os.path.exists(pdf_path):
                with open(pdf_path, 'rb') as f:
                    return f.read(), None
            return None, result.stderr[:500]
    except FileNotFoundError:
        return None, "LibreOffice nie jest zainstalowany na serwerze."
    except subprocess.TimeoutExpired:
        return None, "Konwersja PDF trwała zbyt długo."
    except Exception as e:
        return None, str(e)


def build_image_args():
    logos_bytes = {}
    for i, f in enumerate(logo_files):
        if f is not None:
            f.seek(0)
            logos_bytes[f'logo{i+1}'] = f.read()
    image_order = [k for k, _ in elements_active]
    img_pos_for_docx = {}
    for key in image_order:
        if key in image_positions:
            p = image_positions[key]
            img_pos_for_docx[key] = {
                'x': p['x'], 'y': p['y'],
                'width': p['w'], 'height': p['h']
            }
    return logos_bytes, image_order, img_pos_for_docx


# ─── Generuj ────────────────────────────────────────────────────────────
st.divider()
st.header("6. Generuj")

cols_fmt = st.columns([1, 1, 4])
with cols_fmt[0]:
    fmt_docx = st.checkbox("📄 Word (.docx)", value=True)
with cols_fmt[1]:
    fmt_pdf = st.checkbox("📕 PDF (.pdf)", value=True)

cols_main = st.columns([1, 2, 1])
with cols_main[1]:
    gen_clicked = st.button("🚀 Generuj protokoły z arkusza",
                            type="primary", use_container_width=True)

with st.expander("➕ Dodatkowe opcje"):
    st.caption("Pusty formularz przyda się gdy chcesz wydrukować protokół do ręcznego wypełnienia "
               "(bez pobierania danych z arkusza).")
    blank_clicked = st.button("📝 Pobierz pusty formularz",
                              use_container_width=False)

# ─── Akcje ──────────────────────────────────────────────────────────────
if gen_clicked:
    if not fmt_docx and not fmt_pdf:
        st.error("Wybierz co najmniej jeden format pliku."); st.stop()
    if not sheets_url.strip():
        st.error("Podaj link do arkusza."); st.stop()
    sid = extract_id(sheets_url.strip())
    if not sid:
        st.error("Nieprawidłowy link."); st.stop()

    # Wybierz źródło danych zależne od fazy:
    # - Faza grupowa → zakładki Gr. A/B/C...
    # - Faza pucharowa → zakładka Drabinka
    if is_pucharowa:
        with st.spinner(f"Pobieram mecze z zakładki Drabinka (faza {tournament_phase})..."):
            try:
                phase_name, matches = generate_docx.fetch_drabinka_phase(sid, tournament_phase)
            except Exception as e:
                st.error(f"Błąd pobierania zakładki Drabinka: {e}"); st.stop()
        if not matches:
            st.error(f"Nie znaleziono meczów fazy '{tournament_phase}' w zakładce Drabinka. "
                     "Upewnij się że arkusz ma zakładkę o nazwie 'Drabinka' i że nagłówek "
                     f"kolumny zawiera nazwę fazy (np. '1/32 FINAŁU')."); st.stop()
        # Format dla build_document: lista (group_name, matches)
        sheets_data = [(phase_name or tournament_phase, matches)]
    else:
        with st.spinner("Pobieram dane z grup..."):
            try:
                sheets_data = generate_docx.fetch_all_group_sheets(sid)
            except Exception as e:
                st.error(f"Błąd pobierania: {e}"); st.stop()

    total = sum(len(m) for _,m in sheets_data)
    if total == 0:
        st.error("0 meczów. Użyj 'Sprawdź zakładki' żeby sprawdzić."); st.stop()

    with st.spinner(f"Generuję {total} protokołów..."):
        logos_bytes, image_order, img_pos = build_image_args()
        show_name = tournament_name.strip() if show_header_on_protocol else ""
        show_date = tournament_date if show_header_on_protocol else ""
        # Mapowanie typu turnieju → szablon docx
        template_type = 'TROJKA' if is_trojka else 'IND'
        docx_bytes = generate_docx.build_document(
            sid, sheets_url.strip(), sheets_data,
            logos=logos_bytes or None,
            tournament_name=show_name, tournament_date=show_date,
            include_qr=include_qr, include_pfm_logo=include_pfm_logo,
            image_order=image_order or None, image_positions=img_pos or None,
            hide_grupa_mecz=is_pucharowa, phase_label=phase_label_short,
            template_type=template_type)

    safe_name = re.sub(r'[^\w\s-]','', tournament_name).strip().replace(' ','_') or "protokoly"
    if is_pucharowa:
        # Dodaj fazę do nazwy pliku, np. "GP2_2026_1_32"
        phase_suffix = re.sub(r'[^\w]', '_', tournament_phase).strip('_').lower()
        safe_name = f"{safe_name}_{phase_suffix}"
    
    pdf_bytes, pdf_err = (None, None)
    if fmt_pdf:
        with st.spinner("Konwertuję do PDF..."):
            pdf_bytes, pdf_err = docx_to_pdf(docx_bytes, safe_name)

    st.session_state['last_gen'] = {
        'docx': docx_bytes if fmt_docx else None,
        'pdf': pdf_bytes if fmt_pdf else None,
        'pdf_err': pdf_err,
        'name': safe_name,
        'total': total,
        'groups': len(sheets_data),
        'kind': 'full',
        'is_pucharowa': is_pucharowa,
        'phase_name': tournament_phase,
    }

if blank_clicked:
    if not fmt_docx and not fmt_pdf:
        st.error("Wybierz co najmniej jeden format pliku."); st.stop()
    with st.spinner("Generuję pusty formularz..."):
        logos_bytes, image_order, img_pos = build_image_args()
        show_name = tournament_name.strip() if show_header_on_protocol else ""
        show_date = tournament_date if show_header_on_protocol else ""
        template_type = 'TROJKA' if is_trojka else 'IND'
        docx_bytes = generate_docx.build_blank_document(
            num_pages=1, logos=logos_bytes or None,
            tournament_name=show_name, tournament_date=show_date,
            sheets_url=sheets_url.strip(),
            include_qr=include_qr, include_pfm_logo=include_pfm_logo,
            image_order=image_order or None, image_positions=img_pos or None,
            template_type=template_type)

    safe_name = "pusty_formularz"
    pdf_bytes, pdf_err = (None, None)
    if fmt_pdf:
        with st.spinner("Konwertuję do PDF..."):
            pdf_bytes, pdf_err = docx_to_pdf(docx_bytes, safe_name)

    st.session_state['last_gen'] = {
        'docx': docx_bytes if fmt_docx else None,
        'pdf': pdf_bytes if fmt_pdf else None,
        'pdf_err': pdf_err,
        'name': safe_name,
        'kind': 'blank',
    }

if 'last_gen' in st.session_state:
    gen = st.session_state['last_gen']
    if gen['kind'] == 'full':
        if gen.get('is_pucharowa'):
            st.success(f"✅ Gotowe! {gen['total']} protokołów dla fazy {gen.get('phase_name','')}.")
        else:
            st.success(f"✅ Gotowe! {gen['total']} protokołów w {gen['groups']} grupach.")
    else:
        st.success("✅ Pusty formularz gotowy!")

    # Jeśli oba formaty - 2 kolumny side by side. Jeśli tylko jeden - wycentrowany.
    has_docx = bool(gen['docx'])
    has_pdf = bool(gen['pdf'])
    if has_docx and has_pdf:
        cols_dl = st.columns(2)
        with cols_dl[0]:
            st.download_button(f"⬇️ Pobierz {gen['name']}.docx",
                data=gen['docx'], file_name=f"{gen['name']}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True, key=f"dl_docx_{gen['name']}")
        with cols_dl[1]:
            st.download_button(f"⬇️ Pobierz {gen['name']}.pdf",
                data=gen['pdf'], file_name=f"{gen['name']}.pdf",
                mime="application/pdf",
                use_container_width=True, key=f"dl_pdf_{gen['name']}")
    elif has_docx:
        cols_dl = st.columns([1, 2, 1])
        with cols_dl[1]:
            st.download_button(f"⬇️ Pobierz {gen['name']}.docx",
                data=gen['docx'], file_name=f"{gen['name']}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True, key=f"dl_docx_{gen['name']}")
    elif has_pdf:
        cols_dl = st.columns([1, 2, 1])
        with cols_dl[1]:
            st.download_button(f"⬇️ Pobierz {gen['name']}.pdf",
                data=gen['pdf'], file_name=f"{gen['name']}.pdf",
                mime="application/pdf",
                use_container_width=True, key=f"dl_pdf_{gen['name']}")
    elif gen.get('pdf_err'):
        st.error(f"Konwersja PDF nie powiodła się: {gen['pdf_err']}")

st.divider()
st.caption("Polska Federacja Mölkky · github.com/polska-federacja-molkky/protocol")
