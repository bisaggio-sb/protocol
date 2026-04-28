"""
Generator protokołów meczowych Mölkky
"""
import streamlit as st
import io, re, base64, os
from datetime import date, timedelta
from PIL import Image
import generate_docx

st.set_page_config(page_title="Protokoły Mölkky", page_icon="🎯", layout="wide")
st.title("🎯 Generator protokołów meczowych Mölkky")
st.markdown("Podaj nazwę turnieju, link do arkusza Google Sheets, opcjonalnie dodaj grafiki — pobierz gotowy `.docx`.")


def extract_id(url):
    m = re.search(r'/spreadsheets/d/([a-zA-Z0-9_-]+)', url)
    return m.group(1) if m else None


def img_to_data_url(file_or_bytes):
    """uploaded file lub bytes → data URL."""
    if hasattr(file_or_bytes, 'seek'):
        file_or_bytes.seek(0)
        img_bytes = file_or_bytes.read()
        file_or_bytes.seek(0)
    else:
        img_bytes = file_or_bytes
    img = Image.open(io.BytesIO(img_bytes)).convert("RGBA")
    buf = io.BytesIO()
    img.thumbnail((300, 300))
    img.save(buf, format="PNG")
    b64 = base64.b64encode(buf.getvalue()).decode()
    return f"data:image/png;base64,{b64}"


def next_saturday():
    """Zwraca string DD.MM.YYYY najbliższej soboty (lub dziś jeśli sobota)."""
    today = date.today()
    days_until_sat = (5 - today.weekday()) % 7  # 5 = sobota
    sat = today + timedelta(days=days_until_sat)
    return sat.strftime("%d.%m.%Y")


# Stan pozycji
if 'image_positions' not in st.session_state:
    st.session_state.image_positions = {
        'qr':    {'x': 1.6, 'y': 0.2, 'w': 1.8, 'h': 1.8},
        'pfm':   {'x': 1.5, 'y': 2.4, 'w': 2.0, 'h': 1.2},
        'logo1': {'x': 1.5, 'y': 3.7, 'w': 2.0, 'h': 1.2},
        'logo2': {'x': 1.5, 'y': 5.0, 'w': 2.0, 'h': 1.2},
        'logo3': {'x': 1.5, 'y': 6.3, 'w': 2.0, 'h': 1.2},
        'logo4': {'x': 1.5, 'y': 7.6, 'w': 2.0, 'h': 1.2},
    }


col_form, col_preview = st.columns([3, 2])

with col_form:
    # ─── 1. Turniej (nazwa + data w jednym rzędzie, kompaktowo) ──────────
    st.header("1. Turniej")
    cols_t = st.columns([2, 1])
    with cols_t[0]:
        tournament_name = st.text_input("Nazwa turnieju",
            value="GP2 2026", placeholder="np. GP2 2026")
    with cols_t[1]:
        tournament_date_d = st.date_input("Data turnieju",
            value=date.today() + timedelta(days=(5 - date.today().weekday()) % 7),
            format="DD.MM.YYYY")
    tournament_date = tournament_date_d.strftime("%d.%m.%Y") if tournament_date_d else ""

    # ─── 2. Link ────────────────────────────────────────────────────────
    st.header("2. Link do arkusza Google Sheets")
    sheets_url = st.text_input("URL arkusza",
        placeholder="https://docs.google.com/spreadsheets/d/XXXX/edit...",
        help="Arkusz musi być publiczny. Zakładki grup: 'Gr. A', 'Gr. B', ...")

    # ─── 3. QR + PFM (kompaktowo, w jednym rzędzie) ─────────────────────
    st.header("3. Domyślne elementy")
    cols_dom = st.columns(2)
    with cols_dom[0]:
        include_qr = st.checkbox("✅ Kod QR (link do arkusza)", value=True)
    with cols_dom[1]:
        include_pfm_logo = st.checkbox("✅ Logo Polskiej Federacji Mölkky", value=True)

    # ─── 4. Dodatkowe grafiki ───────────────────────────────────────────
    st.header("4. Dodatkowe grafiki (max 4)")
    NUM_LOGOS = 4
    logo_files = []
    cols_log = st.columns(2)
    for i in range(NUM_LOGOS):
        with cols_log[i % 2]:
            f = st.file_uploader(f"Grafika {i+1}", type=["png","jpg","jpeg"], key=f"logo_{i}")
            logo_files.append(f)

    # ─── 5. Lista aktywnych elementów + edycja pozycji ──────────────────
    elements_active = []
    if include_qr:
        elements_active.append(('qr', 'Kod QR'))
    if include_pfm_logo:
        elements_active.append(('pfm', 'Logo PFM'))
    for i, f in enumerate(logo_files):
        if f is not None:
            elements_active.append((f'logo{i+1}', f'Grafika {i+1}'))

    if elements_active:
        st.header("5. Pozycja i rozmiar elementów")
        st.caption("Domyślnie elementy są ułożone jeden pod drugim w lewym obszarze. "
                   "Możesz dostosować pozycję (cm od lewego górnego rogu obszaru) i rozmiar.")
        for key, label in elements_active:
            with st.expander(f"📍 {label}", expanded=False):
                cols = st.columns(4)
                pos = st.session_state.image_positions.get(key, {'x':1.5,'y':0.2,'w':2.0,'h':1.2})
                with cols[0]:
                    new_x = st.number_input("X (cm)", value=float(pos['x']),
                                           min_value=0.0, max_value=5.0, step=0.1,
                                           key=f"x_{key}")
                with cols[1]:
                    new_y = st.number_input("Y (cm)", value=float(pos['y']),
                                           min_value=0.0, max_value=12.0, step=0.1,
                                           key=f"y_{key}")
                with cols[2]:
                    new_w = st.number_input("Szerokość (cm)", value=float(pos['w']),
                                           min_value=0.5, max_value=5.0, step=0.1,
                                           key=f"w_{key}")
                with cols[3]:
                    new_h = st.number_input("Wysokość (cm)", value=float(pos['h']),
                                           min_value=0.5, max_value=5.0, step=0.1,
                                           key=f"h_{key}")
                st.session_state.image_positions[key] = {
                    'x': new_x, 'y': new_y, 'w': new_w, 'h': new_h
                }

    with st.expander("🔍 Debug – sprawdź zakładki"):
        if st.button("Sprawdź zakładki"):
            sid = extract_id(sheets_url.strip()) if sheets_url.strip() else None
            if not sid:
                st.error("Wklej najpierw poprawny link do arkusza.")
            else:
                with st.spinner("Sprawdzam zakładki Gr. A – Gr. P..."):
                    info = generate_docx.get_sheet_names_debug(sid)
                st.code("\n".join(info))


# ─── PODGLĄD HTML/CSS po prawej ─────────────────────────────────────────
with col_preview:
    st.header("📄 Podgląd strony")
    st.caption("Schemat dokumentu w skali")

    SCALE = 22
    PAGE_W_CM = 18.46
    PAGE_H_CM = 27.16
    PAGE_W_PX = int(PAGE_W_CM * SCALE)
    PAGE_H_PX = int(PAGE_H_CM * SCALE)
    LEFT_AREA_CM = 5.24
    LEFT_AREA_PX = int(LEFT_AREA_CM * SCALE)

    # Buduj data URLs
    image_urls = {}
    if include_qr:
        image_urls['qr'] = None  # placeholder
    if include_pfm_logo:
        pfm_path = os.path.join(os.path.dirname(__file__), 'assets_pfm_logo.png')
        if os.path.exists(pfm_path):
            with open(pfm_path, 'rb') as fp:
                image_urls['pfm'] = img_to_data_url(fp.read())
    for i, f in enumerate(logo_files):
        if f is not None:
            image_urls[f'logo{i+1}'] = img_to_data_url(f)

    # HTML elementów w lewym obszarze
    elements_html = ""
    for key in image_urls:
        if key not in st.session_state.image_positions:
            continue
        pos = st.session_state.image_positions[key]
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
                        font-size:10px; color:#333; font-family:monospace;">
              <div style="text-align:center;">
                <div style="font-weight:bold;">QR</div>
              </div>
            </div>"""
        else:
            elements_html += f"""
            <img src="{image_urls[key]}" 
                 style="position:absolute; left:{x_px}px; top:{y_px}px;
                        width:{w_px}px; height:{h_px}px; object-fit:contain;
                        border:1px dashed #aaa;"/>"""

    # Etykieta "Wyniki turnieju" pod QR (lub na górze jeśli QR off)
    label_y_px = 5
    if include_qr and 'qr' in st.session_state.image_positions:
        qr_pos = st.session_state.image_positions['qr']
        label_y_px = int((qr_pos['y'] + qr_pos['h'] + 0.1) * SCALE)

    # Header w prawym górnym rogu
    header_text = ""
    if tournament_name or tournament_date:
        parts = [tournament_name] if tournament_name else []
        if tournament_date:
            parts.append(tournament_date)
        header_text = " · ".join(parts)

    html = f"""
    <div style="background:white; border:1px solid #ccc; 
                width:{PAGE_W_PX}px; height:{PAGE_H_PX}px;
                position:relative; font-family:Arial, sans-serif;
                box-shadow:0 2px 8px rgba(0,0,0,0.1); margin:0 auto;">
      
      <!-- Header w prawym górnym rogu -->
      <div style="position:absolute; right:8px; top:6px; 
                  font-size:9px; color:#666; font-style:italic;">
        {header_text}
      </div>
      
      <!-- Tabela 1: nagłówek meczu -->
      <div style="position:absolute; left:6px; top:30px; right:0; height:30px;
                  display:flex; align-items:center; padding-left:6px;
                  font-size:11px; gap:18px;">
        <span>Tor <b>1</b></span>
        <span>Godzina <b>09:30</b></span>
        <span>Grupa <b>A</b></span>
        <span>Mecz # <b>1</b></span>
      </div>
      
      <!-- Tabela 1: nagłówki kolumn -->
      <div style="position:absolute; left:{LEFT_AREA_PX}px; top:65px; right:0; height:32px;
                  background:#f0f0f0; border:1px solid #999;
                  display:flex; font-size:9px; font-weight:bold; text-align:center;
                  align-items:center; justify-content:space-around;">
        <div style="flex:1; border-right:1px solid #999;">Punkty<br>SET 1</div>
        <div style="flex:1; border-right:1px solid #999;">Punkty<br>SET 2</div>
        <div style="flex:1; border-right:1px solid #999;">Wygrane<br>sety</div>
        <div style="flex:2;">Podpis</div>
      </div>
      
      <!-- Zawodnik 1 -->
      <div style="position:absolute; left:6px; top:97px; right:0; height:24px;
                  border:1px solid #999; display:flex; font-size:11px;">
        <div style="flex:0 0 {LEFT_AREA_PX-6}px; padding-right:8px; 
                    text-align:right; line-height:24px; font-weight:bold;">
          Łukasz Szulc
        </div>
        <div style="flex:1; border-left:1px solid #999;"></div>
        <div style="flex:1; border-left:1px solid #999;"></div>
        <div style="flex:1; border-left:1px solid #999;"></div>
        <div style="flex:2; border-left:1px solid #999;"></div>
      </div>
      
      <!-- Zawodnik 2 -->
      <div style="position:absolute; left:6px; top:121px; right:0; height:24px;
                  border:1px solid #999; border-top:none; display:flex; font-size:11px;">
        <div style="flex:0 0 {LEFT_AREA_PX-6}px; padding-right:8px;
                    text-align:right; line-height:24px; font-weight:bold;">
          Anna Ściepuro
        </div>
        <div style="flex:1; border-left:1px solid #999;"></div>
        <div style="flex:1; border-left:1px solid #999;"></div>
        <div style="flex:1; border-left:1px solid #999;"></div>
        <div style="flex:2; border-left:1px solid #999;"></div>
      </div>
      
      <!-- Zasady -->
      <div style="position:absolute; left:0; right:0; top:148px;
                  text-align:center; font-size:8px; color:#555; font-style:italic;">
        Każdy zawodnik zaczyna po jednym secie (w dowolnej kolejności)<br>
        Set przegrany przez 3 kolejne chybienia oznacza wynik 0:50
      </div>
      
      <!-- Tabela 2 -->
      <div style="position:absolute; left:0; top:180px; right:0; bottom:30px;
                  border:1px solid #999;">
        <!-- Lewy obszar -->
        <div style="position:absolute; left:0; top:0; bottom:0;
                    width:{LEFT_AREA_PX}px; border-right:1px solid #999;
                    overflow:hidden;">
          {elements_html}
          <div style="position:absolute; left:0; right:0; top:{label_y_px}px;
                      text-align:center; font-size:10px; font-weight:bold;">
            Wyniki turnieju
          </div>
        </div>
        
        <!-- Prawa część tabeli wyników -->
        <div style="position:absolute; left:{LEFT_AREA_PX+1}px; top:0; right:0; bottom:0;">
          <!-- Header SET 1, SET 2 -->
          <div style="position:absolute; left:0; top:0; right:0; height:22px;
                      background:#f0f0f0; display:flex; font-size:10px; font-weight:bold;
                      align-items:center; text-align:center; border-bottom:1px solid #999;">
            <div style="flex:1.4; border-right:1px solid #999;"></div>
            <div style="flex:3.0; border-right:1px solid #999;">SET 1</div>
            <div style="flex:3.0;">SET 2</div>
          </div>
          
          <!-- Header IMIONA / SUMA -->
          <div style="position:absolute; left:0; top:22px; right:0; height:30px;
                      background:#f8f8f8; display:flex; font-size:8px; font-weight:bold;
                      align-items:center; text-align:center; 
                      border-bottom:1px solid #999;">
            <div style="flex:1.4; border-right:1px solid #999;
                        writing-mode:vertical-rl; transform:rotate(180deg);
                        line-height:1.2;">IMIONA</div>
            <div style="flex:1.5; border-right:1px solid #999;"></div>
            <div style="flex:1.5; border-right:1px solid #999;
                        writing-mode:vertical-rl; transform:rotate(180deg);
                        line-height:1.2;">SUMA</div>
            <div style="flex:1.5; border-right:1px solid #999;"></div>
            <div style="flex:1.5;
                        writing-mode:vertical-rl; transform:rotate(180deg);
                        line-height:1.2;">SUMA</div>
          </div>
          
          <!-- Pusta siatka -->
          <div style="position:absolute; left:0; top:52px; right:0; bottom:24px;
                      background:repeating-linear-gradient(
                        to bottom, transparent 0, transparent 17px,
                        #ccc 17px, #ccc 18px);"></div>
          
          <!-- WYNIK -->
          <div style="position:absolute; left:0; right:0; bottom:0; height:24px;
                      border-top:1px solid #999;
                      background:#f0f0f0; display:flex;
                      font-size:10px; font-weight:bold; text-align:center;
                      align-items:center;">
            <div style="flex:1.4;"></div>
            <div style="flex:1.5; border-left:1px solid #999; line-height:24px;">WYNIK</div>
            <div style="flex:4.5; border-left:1px solid #999;"></div>
          </div>
        </div>
      </div>
    </div>
    """

    st.components.v1.html(html, height=PAGE_H_PX + 30, scrolling=True)
    st.caption("📝 Podgląd schematyczny. Dokładny wygląd zobaczysz w pobranym .docx")


# ─── Generuj ────────────────────────────────────────────────────────────
st.divider()
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
        st.error("0 meczów."); st.stop()

    with st.spinner(f"Generuję {total} protokołów..."):
        logos_bytes = {}
        for i, f in enumerate(logo_files):
            if f is not None:
                f.seek(0)
                logos_bytes[f'logo{i+1}'] = f.read()

        image_order = [k for k, _ in elements_active]

        image_positions = {}
        for key in image_order:
            if key in st.session_state.image_positions:
                p = st.session_state.image_positions[key]
                image_positions[key] = {
                    'x': p['x'], 'y': p['y'], 'width': p['w']
                }

        docx_bytes = generate_docx.build_document(
            sid, sheets_url.strip(), sheets_data,
            logos=logos_bytes or None,
            tournament_name=tournament_name.strip() or "Turniej Mölkky",
            tournament_date=tournament_date,
            include_qr=include_qr,
            include_pfm_logo=include_pfm_logo,
            image_order=image_order or None,
            image_positions=image_positions or None)

    st.success(f"✅ Gotowe! {total} protokołów w {len(sheets_data)} grupach.")
    safe_name = re.sub(r'[^\w\s-]','', tournament_name).strip().replace(' ','_') or "protokoly"
    st.download_button(f"⬇️ Pobierz {safe_name}.docx", data=docx_bytes,
        file_name=f"{safe_name}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True)

st.divider()
st.caption("Polska Federacja Mölkky · github.com/polska-federacja-molkky/protocol")
