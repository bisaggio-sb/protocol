"""
Generator protokołów meczowych Mölkky
"""
import streamlit as st
import io, re, base64
from PIL import Image
import generate_docx

st.set_page_config(page_title="Protokoły Mölkky", page_icon="🎯", layout="wide")
st.title("🎯 Generator protokołów meczowych Mölkky")
st.markdown("Podaj nazwę turnieju, link do arkusza Google Sheets, opcjonalnie dodaj grafiki — pobierz gotowy `.docx`.")


def extract_id(url):
    m = re.search(r'/spreadsheets/d/([a-zA-Z0-9_-]+)', url)
    return m.group(1) if m else None


def img_to_data_url(file_obj):
    """PIL image lub uploaded file → data URL."""
    file_obj.seek(0)
    img = Image.open(file_obj).convert("RGB")
    file_obj.seek(0)
    buf = io.BytesIO()
    img.thumbnail((300, 300))
    img.save(buf, format="PNG")
    b64 = base64.b64encode(buf.getvalue()).decode()
    return f"data:image/png;base64,{b64}"


# Inicjalizacja stanu pozycji w session_state
if 'image_positions' not in st.session_state:
    # domyślne pozycje: jedna pod drugą w lewym obszarze
    # X w cm od lewej krawędzi obszaru "Wyniki turnieju" (5.24 cm)
    # Y w cm od góry komórki
    st.session_state.image_positions = {
        'qr':    {'x': 1.6, 'y': 0.2, 'w': 1.8, 'h': 1.8},
        'logo1': {'x': 1.5, 'y': 2.4, 'w': 2.0, 'h': 1.2},
        'logo2': {'x': 1.5, 'y': 3.7, 'w': 2.0, 'h': 1.2},
        'logo3': {'x': 1.5, 'y': 5.0, 'w': 2.0, 'h': 1.2},
        'logo4': {'x': 1.5, 'y': 6.3, 'w': 2.0, 'h': 1.2},
    }


# Layout: 2 kolumny - formularz po lewej, podgląd po prawej
col_form, col_preview = st.columns([3, 2])

with col_form:
    st.header("1. Nazwa turnieju")
    tournament_name = st.text_input("Nazwa turnieju",
        value="GP2 2026", placeholder="np. GP2 2026")

    st.header("2. Link do arkusza Google Sheets")
    sheets_url = st.text_input("URL arkusza",
        placeholder="https://docs.google.com/spreadsheets/d/XXXX/edit...",
        help="Arkusz musi być publiczny. Zakładki grup: 'Gr. A', 'Gr. B', ..., 'Gr. P'")

    st.header("3. Kod QR")
    include_qr = st.checkbox("✅ Generuj kod QR z linkiem do arkusza", value=True)

    st.header("4. Grafiki (max 4)")
    NUM_LOGOS = 4
    logo_files = []
    cols_log = st.columns(2)
    for i in range(NUM_LOGOS):
        with cols_log[i % 2]:
            f = st.file_uploader(f"Grafika {i+1}", type=["png","jpg","jpeg"], key=f"logo_{i}")
            logo_files.append(f)

    # Lista aktywnych elementów
    elements_active = []
    if include_qr:
        elements_active.append(('qr', 'Kod QR'))
    for i, f in enumerate(logo_files):
        if f is not None:
            elements_active.append((f'logo{i+1}', f'Grafika {i+1}'))

    # Edycja pozycji
    if elements_active:
        st.header("5. Pozycja i rozmiar elementów")
        st.caption("Domyślnie elementy są ułożone jeden pod drugim w lewym obszarze. "
                   "Możesz dostosować pozycję (cm od lewego górnego rogu obszaru) i rozmiar.")
        
        for key, label in elements_active:
            with st.expander(f"📍 {label}", expanded=False):
                cols = st.columns(4)
                pos = st.session_state.image_positions.get(key, {'x':1.5,'y':0.2,'w':2.0,'h':1.2})
                with cols[0]:
                    new_x = st.number_input(f"X (cm)", value=float(pos['x']),
                                           min_value=0.0, max_value=5.0, step=0.1,
                                           key=f"x_{key}")
                with cols[1]:
                    new_y = st.number_input(f"Y (cm)", value=float(pos['y']),
                                           min_value=0.0, max_value=12.0, step=0.1,
                                           key=f"y_{key}")
                with cols[2]:
                    new_w = st.number_input(f"Szerokość (cm)", value=float(pos['w']),
                                           min_value=0.5, max_value=5.0, step=0.1,
                                           key=f"w_{key}")
                with cols[3]:
                    new_h = st.number_input(f"Wysokość (cm)", value=float(pos['h']),
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


# ─── PODGLĄD HTML/CSS po prawej ─────────────────────────────────────────────
with col_preview:
    st.header("📄 Podgląd strony")
    st.caption("Schemat dokumentu w skali")

    # Wymiary A4 z marginesami 1.27cm: obszar tekstu 18.46×27.16 cm
    # Skala podglądu: 1 cm = 22 px → 406×598 px
    SCALE = 22  # px na cm
    PAGE_W_CM = 18.46
    PAGE_H_CM = 27.16
    PAGE_W_PX = int(PAGE_W_CM * SCALE)
    PAGE_H_PX = int(PAGE_H_CM * SCALE)

    # Lewy obszar "Wyniki turnieju" = 5.24 cm
    LEFT_AREA_CM = 5.24
    LEFT_AREA_PX = int(LEFT_AREA_CM * SCALE)

    # Buduj data URLs dla wgranych grafik
    image_urls = {}
    if include_qr:
        # Symuluj QR jako wzór (prawdziwy QR będzie tylko w pobranym docx)
        image_urls['qr'] = None  # narysuje placeholder
    for i, f in enumerate(logo_files):
        if f is not None:
            image_urls[f'logo{i+1}'] = img_to_data_url(f)

    # Buduj HTML podglądu
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
            # Placeholder QR
            elements_html += f"""
            <div style="position:absolute; left:{x_px}px; top:{y_px}px; 
                        width:{w_px}px; height:{h_px}px; 
                        background:white; border:2px solid #333;
                        display:flex; align-items:center; justify-content:center;
                        font-size:10px; color:#333; font-family:monospace;">
              <div style="text-align:center;">
                <div style="font-weight:bold;">QR</div>
                <div style="font-size:7px;">code</div>
              </div>
            </div>"""
        else:
            elements_html += f"""
            <img src="{image_urls[key]}" 
                 style="position:absolute; left:{x_px}px; top:{y_px}px;
                        width:{w_px}px; height:{h_px}px; object-fit:contain;
                        border:1px dashed #aaa;"/>"""

    # Pozycja "Wyniki turnieju" napisu (pod ostatnim QR jeśli jest)
    label_y_px = 22 if include_qr else 5
    if include_qr and 'qr' in st.session_state.image_positions:
        qr_pos = st.session_state.image_positions['qr']
        label_y_px = int((qr_pos['y'] + qr_pos['h'] + 0.1) * SCALE)

    html = f"""
    <div style="background:white; border:1px solid #ccc; 
                width:{PAGE_W_PX}px; height:{PAGE_H_PX}px;
                position:relative; font-family:Arial, sans-serif;
                box-shadow:0 2px 8px rgba(0,0,0,0.1); margin:0 auto;">
      
      <!-- Tabela 1: nagłówek meczu (Tor/Godzina/Grupa/Mecz) -->
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
      
      <!-- Tabela 2: wyniki turnieju -->
      <div style="position:absolute; left:0; top:180px; right:0; bottom:30px;
                  border:1px solid #999;">
        <!-- Pierwsza kolumna "Wyniki turnieju" -->
        <div style="position:absolute; left:0; top:0; bottom:0;
                    width:{LEFT_AREA_PX}px; border-right:1px solid #999;
                    overflow:hidden;">
          <!-- Tu będą floating images -->
          {elements_html}
          
          <!-- Napis "Wyniki turnieju" -->
          <div style="position:absolute; left:0; right:0; top:{label_y_px}px;
                      text-align:center; font-size:10px; font-weight:bold;">
            Wyniki turnieju
          </div>
        </div>
        
        <!-- Prawa część: SET 1, SET 2, kolumny -->
        <div style="position:absolute; left:{LEFT_AREA_PX+1}px; top:0; right:0; bottom:0;">
          <!-- Header SET 1, SET 2 -->
          <div style="position:absolute; left:0; top:0; right:0; height:22px;
                      background:#f0f0f0; display:flex; font-size:10px; font-weight:bold;
                      align-items:center; text-align:center; border-bottom:1px solid #999;">
            <div style="flex:1.4;"></div>
            <div style="flex:3.0; border-left:1px solid #999;
                        border-right:1px solid #999;">SET 1</div>
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
            <div style="flex:1.5; border-right:1px solid #999;
                        writing-mode:vertical-rl; transform:rotate(180deg);
                        line-height:1.2;">SUMA</div>
            <div style="flex:1.5; border-right:1px solid #999;"></div>
            <div style="flex:1.5; border-right:1px solid #999;
                        writing-mode:vertical-rl; transform:rotate(180deg);
                        line-height:1.2;">SUMA</div>
            <div style="flex:1.5; border-right:1px solid #999;"></div>
            <div style="flex:1.5;
                        writing-mode:vertical-rl; transform:rotate(180deg);
                        line-height:1.2;">SUMA</div>
          </div>
          
          <!-- Pusta siatka 18 wierszy -->
          <div style="position:absolute; left:0; top:52px; right:0; bottom:24px;
                      display:flex; flex-direction:column;">
            <!-- Generujemy 18 pustych wierszy jako stripes -->
            <div style="flex:1; background:repeating-linear-gradient(
                        to bottom, transparent 0, transparent 17px,
                        #ccc 17px, #ccc 18px);
                        background-position: 0 -1px;"></div>
          </div>
          
          <!-- WYNIK na dole -->
          <div style="position:absolute; left:0; right:0; bottom:0; height:24px;
                      border-top:1px solid #999;
                      background:#f0f0f0; display:flex;
                      font-size:10px; font-weight:bold; text-align:center;
                      align-items:center;">
            <div style="flex:1.4;"></div>
            <div style="flex:1.5; border-left:1px solid #999; line-height:24px;">WYNIK</div>
            <div style="flex:9; border-left:1px solid #999;"></div>
          </div>
        </div>
      </div>
    </div>
    """

    st.components.v1.html(html, height=PAGE_H_PX + 30, scrolling=True)
    st.caption("📝 Podgląd schematyczny. Dokładny wygląd zobaczysz w pobranym .docx")


# ─── Generuj ────────────────────────────────────────────────────────────────
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

        # Kolejność: wszystkie aktywne elementy w kolejności dodawania
        image_order = [k for k, _ in elements_active]

        # Pozycje z session_state przekształcone na format generate_docx
        image_positions = {}
        for key in image_order:
            if key in st.session_state.image_positions:
                p = st.session_state.image_positions[key]
                image_positions[key] = {
                    'x': p['x'], 'y': p['y'],
                    'width': p['w']
                }

        docx_bytes = generate_docx.build_document(
            sid, sheets_url.strip(), sheets_data,
            logos=logos_bytes or None,
            tournament_name=tournament_name.strip() or "Turniej Mölkky",
            include_qr=include_qr,
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
