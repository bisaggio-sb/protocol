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


# Header z logo PFM
if os.path.exists(PFM_PATH):
    with open(PFM_PATH, 'rb') as fp:
        pfm_data_url = img_to_data_url(fp.read())
else:
    pfm_data_url = ""

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


# Obszar "Wyniki turnieju" ma szerokość 5.24 cm (gridCol[0] = 2970 DXA - oryginał)
LEFT_AREA_CM = 5.24


def compute_default_positions(active_keys, area_height_cm=16.0, area_width_cm=LEFT_AREA_CM):
    """Rozkłada elementy w lewym obszarze równomiernie z 1cm odstępami.
    QR (jeśli jest) na górze, napis pod QR (renderowany w docx),
    PFM + grafiki rozłożone równomiernie poniżej."""
    positions = {}
    SPACING = 1.0
    
    other_keys = [k for k in active_keys if k != 'qr']
    
    if 'qr' in active_keys:
        qr_w = 2.4
        qr_x = (area_width_cm - qr_w) / 2
        positions['qr'] = {'x': qr_x, 'y': 0.2, 'w': qr_w, 'h': qr_w}
        first_logo_y = 0.2 + qr_w + 0.5 + SPACING
    else:
        first_logo_y = 0.2
    
    if not other_keys:
        return positions
    
    n = len(other_keys)
    available = area_height_cm - first_logo_y
    
    if n == 1:
        img_h = min(3.5, available)
    else:
        img_h = (available - (n - 1) * SPACING) / n
        img_h = max(1.5, min(3.5, img_h))
    
    img_w = min(4.0, area_width_cm - 0.6)
    img_x = (area_width_cm - img_w) / 2
    
    cur_y = first_logo_y
    for key in other_keys:
        positions[key] = {'x': img_x, 'y': cur_y, 'w': img_w, 'h': img_h}
        cur_y += img_h + SPACING
    
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

    show_header_on_protocol = st.checkbox(
        "Pokaż nazwę i datę na protokole (prawy górny róg)", value=True
    )

    cols_t2 = st.columns(2)
    with cols_t2[0]:
        tournament_type = st.radio(
            "Rodzaj",
            ["Indywidualny", "Drużynowy 2-os.", "Drużynowy 3-os.", "Drużynowy 4-os."],
            index=0,
            help="Obecnie zaimplementowany tylko turniej indywidualny"
        )
    with cols_t2[1]:
        tournament_phase = st.radio(
            "Faza",
            ["Grupowa (2 sety)",
             "Pucharowa 1/64 (best of 3)",
             "Pucharowa 1/32 (best of 3)",
             "Pucharowa 1/16 (best of 3)",
             "Pucharowa 1/8 (best of 3)",
             "Pucharowa 1/4 (best of 5)",
             "Pucharowa 1/2 (best of 5)",
             "Mecz o 3. miejsce (best of 5)",
             "Finał (best of 5)"],
            index=0,
            help="Obecnie zaimplementowana tylko faza grupowa indywidualna"
        )

    if tournament_type != "Indywidualny" or "Grupowa" not in tournament_phase:
        st.warning("⚠️ Obecnie zaimplementowana jest tylko **faza grupowa turnieju indywidualnego**. "
                   "Pozostałe opcje będą dostępne w kolejnych wersjach.")

    # ─── 2. Link do arkusza ─────────────────────────────────────────────
    st.header("2. Link do arkusza Google Sheets")
    sheets_url = st.text_input("URL arkusza",
        placeholder="https://docs.google.com/spreadsheets/d/XXXX/edit...",
        help="Arkusz musi być publiczny. Zakładki grup: 'Gr. A', 'Gr. B', ...",
        label_visibility="collapsed")

    # ─── 3. Domyślne elementy ───────────────────────────────────────────
    st.header("3. Domyślne elementy")
    cols_dom = st.columns(2)
    with cols_dom[0]:
        include_qr = st.checkbox("Kod QR (link do arkusza)", value=True)
    with cols_dom[1]:
        include_pfm_logo = st.checkbox("Logo Polskiej Federacji Mölkky", value=True)

    # ─── 4. Dodatkowe grafiki ───────────────────────────────────────────
    st.header("4. Dodatkowe grafiki (max 4)")
    NUM_LOGOS = 4
    logo_files = []
    cols_log = st.columns(2)
    for i in range(NUM_LOGOS):
        with cols_log[i % 2]:
            f = st.file_uploader(f"Grafika {i+1}", type=["png","jpg","jpeg"], key=f"logo_{i}")
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

    # ─── 5. Pozycje elementów ──────────────────────────────────────────
    # KLUCZOWA ZMIANA: pozycje są ZAWSZE auto-obliczane na podstawie
    # aktywnych elementów. Jeśli user chce zmienić, robi to przez sliders.
    # Dzięki temu zaznacz/odznacz checkboxa zawsze daje porządny układ.
    
    # Zawsze obliczamy domyślne pozycje od zera
    default_pos = compute_default_positions([k for k, _ in elements_active])
    
    # active_keys_signature decyduje kiedy slidery dostają nowe klucze (= reset)
    active_keys_signature = "|".join(sorted(k for k, _ in elements_active))
    
    # Slidery używają kluczy z signature - dzięki temu Streamlit RESETUJE
    # ich stan przy każdej zmianie zestawu aktywnych elementów
    image_positions = {}
    for key in default_pos:
        image_positions[key] = dict(default_pos[key])

    if elements_active:
        st.header("5. Pozycja i rozmiar elementów")
        cols_h5 = st.columns([4, 1])
        with cols_h5[0]:
            st.caption("Elementy są ułożone automatycznie z odstępami 1 cm. "
                       "Rozwiń element żeby dostosować pozycję ręcznie.")
        with cols_h5[1]:
            # Reset = zmiana sig, czyli wymuszona przeładowywanie wszystkich kluczy
            if st.button("↻ Rozmieść równomiernie",
                         help="Przywróć automatyczne rozmieszczenie"):
                # Zwiększamy nonce żeby slidery dostały nowe klucze i się zresetowały
                st.session_state['reset_nonce'] = st.session_state.get('reset_nonce', 0) + 1
                st.rerun()

        # Nonce dodajemy do klucza sliderów żeby reset działał
        nonce = st.session_state.get('reset_nonce', 0)
        for key, label in elements_active:
            with st.expander(f"📍 {label}", expanded=False):
                cols = st.columns(4)
                pos = image_positions[key]
                # Klucze sliderów ZAWIERAJĄ signature + nonce
                # → reset przy zmianie zestawu LUB kliknięciu "Rozmieść"
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

    with st.expander("🔍 Debug – sprawdź zakładki"):
        if st.button("Sprawdź zakładki"):
            sid = extract_id(sheets_url.strip()) if sheets_url.strip() else None
            if not sid:
                st.error("Wklej najpierw poprawny link do arkusza.")
            else:
                with st.spinner("Sprawdzam zakładki Gr. A – Gr. Z..."):
                    info = generate_docx.get_sheet_names_debug(sid)
                st.code("\n".join(info))


# ─── PODGLĄD HTML/CSS po prawej (przesunięty 5cm w dół) ───────────────
with col_preview:
    # Dodatkowy odstęp od góry żeby przesunąć podgląd niżej
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

    # Tabela 1: przesunięta w prawo o 600/9690 = 6.2% (tblInd)
    TBL1_OFFSET_FRAC = 600 / 9690
    TBL1_OFFSET_PX = int(PAGE_W_PX * TBL1_OFFSET_FRAC)
    NAMES_FRAC = 3915 / 9090  # pierwsza komórka tabeli 1 (z nazwiskami) względem tabeli 1
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
      
      <!-- Tabela 1: pierwszy wiersz (Tor/Godzina/Grupa/Mecz) - przesunięty w prawo -->
      <div style="position:absolute; left:{TBL1_OFFSET_PX}px; top:30px; right:0; height:30px;
                  display:flex; align-items:center; padding-left:6px;
                  font-size:11px; gap:18px;">
        <span>Tor <b>1</b></span>
        <span>Godzina <b>09:30</b></span>
        <span>Grupa <b>A</b></span>
        <span>Mecz # <b>1</b></span>
      </div>
      
      <!-- Tabela 1: nagłówki kolumn -->
      <div style="position:absolute; left:{TBL1_OFFSET_PX + NAMES_PX}px; top:65px; right:0; height:32px;
                  background:#f0f0f0; border:1px solid #999;
                  display:flex; font-size:9px; font-weight:bold; text-align:center;
                  align-items:center; justify-content:space-around;">
        <div style="flex:1; border-right:1px solid #999;">Punkty<br>SET 1</div>
        <div style="flex:1; border-right:1px solid #999;">Punkty<br>SET 2</div>
        <div style="flex:1.15; border-right:1px solid #999;">Wygrane<br>sety</div>
        <div style="flex:1.95;">Podpis</div>
      </div>
      
      <!-- Wiersze zawodników -->
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
      
      <!-- Tabela 2 -->
      <div style="position:absolute; left:0; top:180px; right:0; bottom:30px;
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
            # Opcje PDF żeby zmniejszyć rozmiar:
            # - SelectPdfVersion=0 = standard PDF 1.4 (mniejszy niż PDF/A)
            # - EmbedStandardFonts=false = nie embeduj standardowych czcionek
            # - ReduceImageResolution=true + MaxImageResolution=150 = niższa rozdzielczość obrazów
            # - UseLosslessCompression=false + Quality=90 = JPEG z sensowną jakością
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

# Format wyboru: docx + pdf, oba domyślnie zaznaczone
cols_fmt = st.columns([1, 1, 4])
with cols_fmt[0]:
    fmt_docx = st.checkbox("📄 Word (.docx)", value=True)
with cols_fmt[1]:
    fmt_pdf = st.checkbox("📕 PDF (.pdf)", value=True)

# Główny przycisk wyśrodkowany (3 kolumny: empty | button | empty)
cols_main = st.columns([1, 2, 1])
with cols_main[1]:
    gen_clicked = st.button("🚀 Generuj protokoły z arkusza",
                            type="primary", use_container_width=True)

# ─── Dodatkowe opcje (zwijalne, dyskretne) ──────────────────────────────
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

    with st.spinner("Pobieram dane z grup..."):
        try:
            sheets_data = generate_docx.fetch_all_group_sheets(sid)
        except Exception as e:
            st.error(f"Błąd pobierania: {e}"); st.stop()

    total = sum(len(m) for _,m in sheets_data)
    if total == 0:
        st.error("0 meczów. Użyj Debug żeby sprawdzić zakładki."); st.stop()

    with st.spinner(f"Generuję {total} protokołów..."):
        logos_bytes, image_order, img_pos = build_image_args()
        show_name = tournament_name.strip() if show_header_on_protocol else ""
        show_date = tournament_date if show_header_on_protocol else ""
        docx_bytes = generate_docx.build_document(
            sid, sheets_url.strip(), sheets_data,
            logos=logos_bytes or None,
            tournament_name=show_name, tournament_date=show_date,
            include_qr=include_qr, include_pfm_logo=include_pfm_logo,
            image_order=image_order or None, image_positions=img_pos or None)

    safe_name = re.sub(r'[^\w\s-]','', tournament_name).strip().replace(' ','_') or "protokoly"
    
    # Konwersja do PDF od razu (jeśli wybrana)
    pdf_bytes, pdf_err = (None, None)
    if fmt_pdf:
        with st.spinner("Konwertuję do PDF..."):
            pdf_bytes, pdf_err = docx_to_pdf(docx_bytes, safe_name)

    # Zapisz w session_state żeby przeżyło rerun po download_button
    st.session_state['last_gen'] = {
        'docx': docx_bytes if fmt_docx else None,
        'pdf': pdf_bytes if fmt_pdf else None,
        'pdf_err': pdf_err,
        'name': safe_name,
        'total': total,
        'groups': len(sheets_data),
        'kind': 'full',
    }

if blank_clicked:
    if not fmt_docx and not fmt_pdf:
        st.error("Wybierz co najmniej jeden format pliku."); st.stop()
    with st.spinner("Generuję pusty formularz..."):
        logos_bytes, image_order, img_pos = build_image_args()
        show_name = tournament_name.strip() if show_header_on_protocol else ""
        show_date = tournament_date if show_header_on_protocol else ""
        docx_bytes = generate_docx.build_blank_document(
            num_pages=1, logos=logos_bytes or None,
            tournament_name=show_name, tournament_date=show_date,
            sheets_url=sheets_url.strip(),
            include_qr=include_qr, include_pfm_logo=include_pfm_logo,
            image_order=image_order or None, image_positions=img_pos or None)

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

# Wyświetl przyciski pobierania (poza if-blokiem żeby przeżyły rerun)
if 'last_gen' in st.session_state:
    gen = st.session_state['last_gen']
    if gen['kind'] == 'full':
        st.success(f"✅ Gotowe! {gen['total']} protokołów w {gen['groups']} grupach.")
    else:
        st.success("✅ Pusty formularz gotowy!")

    cols_dl = st.columns(2)
    if gen['docx']:
        with cols_dl[0]:
            st.download_button(f"⬇️ Pobierz {gen['name']}.docx",
                data=gen['docx'], file_name=f"{gen['name']}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True, key=f"dl_docx_{gen['name']}")
    if gen['pdf']:
        with cols_dl[1]:
            st.download_button(f"⬇️ Pobierz {gen['name']}.pdf",
                data=gen['pdf'], file_name=f"{gen['name']}.pdf",
                mime="application/pdf",
                use_container_width=True, key=f"dl_pdf_{gen['name']}")
    elif gen.get('pdf_err'):
        st.error(f"Konwersja PDF nie powiodła się: {gen['pdf_err']}")

st.divider()
st.caption("Polska Federacja Mölkky · github.com/polska-federacja-molkky/protocol")
