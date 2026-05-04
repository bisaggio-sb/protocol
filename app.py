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


@st.cache_data(show_spinner=False)
def _bytes_to_data_url(img_bytes: bytes) -> str:
    """Cache'owana wersja konwersji bajtów obrazu na data URL.
    Cache'owanie po hashu bytes — różne pliki dadzą różny cache."""
    img = Image.open(io.BytesIO(img_bytes)).convert("RGBA")
    buf = io.BytesIO()
    img.thumbnail((400, 400))
    img.save(buf, format="PNG")
    b64 = base64.b64encode(buf.getvalue()).decode()
    return f"data:image/png;base64,{b64}"


def img_to_data_url(file_or_bytes):
    if hasattr(file_or_bytes, 'seek'):
        file_or_bytes.seek(0)
        img_bytes = file_or_bytes.read()
        file_or_bytes.seek(0)
    else:
        img_bytes = file_or_bytes
    return _bytes_to_data_url(img_bytes)


@st.cache_data(show_spinner=False)
def _bytes_to_aspect(img_bytes: bytes) -> float:
    img = Image.open(io.BytesIO(img_bytes))
    return img.width / img.height


def get_image_aspect(file_or_bytes):
    """Zwraca aspect ratio (width/height) obrazu."""
    if hasattr(file_or_bytes, 'seek'):
        file_or_bytes.seek(0)
        img_bytes = file_or_bytes.read()
        file_or_bytes.seek(0)
    else:
        img_bytes = file_or_bytes
    return _bytes_to_aspect(img_bytes)


@st.cache_data(show_spinner=False)
def _load_pfm_assets():
    """Wczytuje PFM logo raz i cache'uje data url + aspect."""
    if os.path.exists(PFM_PATH):
        with open(PFM_PATH, 'rb') as fp:
            data = fp.read()
        return _bytes_to_data_url(data), _bytes_to_aspect(data)
    return "", 1.10


# Header z logo PFM (cache'owany)
pfm_data_url, pfm_aspect = _load_pfm_assets()

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
# Trójka: po rozszerzeniu tabeli 2, lewa kolumna R1.tc[0] = 2700 dxa = 4.76 cm.
# Używamy PEŁNEJ szerokości tej komórki do centrowania obrazów - obrazy są wtedy
# wycentrowane wizualnie w komórce, z naturalnymi odstępami od jej krawędzi
# (zarówno od lewego marginesu jak i od kolumny z numerkami 1-18).
TROJKA_LEFT_AREA_CM = 4.76
TROJKA_GRAPHIC_W_CM = 3.8   # max szerokość obrazu (z 0.5 cm marginesem od krawędzi 4.76)
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
    
    # Trójka ma węższy lewy pas niż indywidualny - mniejsze rozmiary domyślne
    # ale po rozszerzeniu tabeli 2 (4.76 cm) jest dość miejsca dla grafik
    if is_trojka:
        # ─── Preview TRÓJKOWY ───
        # WSZYSTKIE pozycje są PROPORCJONALNE do SCALE żeby skalowały się i nigdy
        # nie wystawały poza iframe (=brak scrollbara).
        TROJKA_TBL_W_PX = int((18.46 / PAGE_W_CM) * PAGE_W_PX)
        TROJKA_NAMES_PX = int((4541 / 10466) * TROJKA_TBL_W_PX)
        TROJKA_HEADER_W_PX = TROJKA_TBL_W_PX - TROJKA_NAMES_PX
        TROJKA_R1_W_PX = int((4.76 / PAGE_W_CM) * PAGE_W_PX)
        # Pozycje w cm × SCALE
        Y_HEADER_TOP = int(0.4 * SCALE)
        Y_TOR_ROW = int(1.1 * SCALE)
        H_TOR = int(0.6 * SCALE)
        Y_PUNKTY_TOP = int(2.4 * SCALE)
        H_PUNKTY = int(0.9 * SCALE)
        Y_TEAM_A = int(3.5 * SCALE)
        Y_TEAM_B = int(4.2 * SCALE)
        H_TEAM = int(0.7 * SCALE)
        Y_DISCLAIMER = int(5.1 * SCALE)
        Y_TBL_RESULTS = int(6.1 * SCALE)
        H_TBL_RESULTS = int((PAGE_H_CM - 6.1 - 0.5) * SCALE)
        H_TBL_HEADER = int(0.5 * SCALE)
        H_TBL_SUMA = int(1.2 * SCALE)
        H_PKT = int(0.5 * SCALE)
        
        html = f"""
        <div style="background:white; border:1px solid #ccc;
                    width:{PAGE_W_PX}px; height:{PAGE_H_PX}px;
                    position:relative; font-family:Arial, sans-serif;
                    box-shadow:0 2px 8px rgba(0,0,0,0.1); margin:0 auto;
                    overflow:hidden;">
          <div style="position:absolute; right:6px; top:{Y_HEADER_TOP}px;
                      font-size:7px; color:#666; font-style:italic;">
            {header_text}
          </div>
          <div style="position:absolute; left:6px; top:{Y_TOR_ROW}px; right:0;
                      height:{H_TOR}px; display:flex; align-items:center;
                      font-size:8px; gap:8px;">
            <span>Tor <b>1</b></span>
            <span>Godzina <b>09:00</b></span>
            {('' if is_pucharowa else '<span>Grupa <b>A</b></span>')}
            <span>Mecz # <b>1</b></span>
          </div>
          <div style="position:absolute; left:{TROJKA_NAMES_PX}px; top:{Y_PUNKTY_TOP}px;
                      width:{TROJKA_HEADER_W_PX}px; height:{H_PUNKTY}px;
                      background:#f0f0f0; border:1px solid #999;
                      display:flex; font-size:5px; font-weight:bold; text-align:center;
                      align-items:center; justify-content:space-around;">
            <div style="flex:1; border-right:1px solid #999;">Punkty<br>SET 1</div>
            <div style="flex:1; border-right:1px solid #999;">Punkty<br>SET 2</div>
            <div style="flex:1; border-right:1px solid #999;">Wygrane<br>sety</div>
            <div style="flex:1.7;">Podpis</div>
          </div>
          <div style="position:absolute; left:0; top:{Y_TEAM_A}px;
                      width:{TROJKA_TBL_W_PX}px; height:{H_TEAM}px;
                      border:1px solid #999; display:flex; font-size:7px;">
            <div style="flex:0 0 {TROJKA_NAMES_PX}px; padding-right:6px;
                        text-align:center; line-height:{H_TEAM}px; font-weight:bold;
                        border-right:1px solid #999;">KARMI</div>
            <div style="flex:1; border-right:1px solid #999;"></div>
            <div style="flex:1; border-right:1px solid #999;"></div>
            <div style="flex:1; border-right:1px solid #999;"></div>
            <div style="flex:1.7;"></div>
          </div>
          <div style="position:absolute; left:0; top:{Y_TEAM_B}px;
                      width:{TROJKA_TBL_W_PX}px; height:{H_TEAM}px;
                      border:1px solid #999; border-top:none; display:flex; font-size:7px;">
            <div style="flex:0 0 {TROJKA_NAMES_PX}px; padding-right:6px;
                        text-align:center; line-height:{H_TEAM}px; font-weight:bold;
                        border-right:1px solid #999;">Trzech Silnych Mężczyzn</div>
            <div style="flex:1; border-right:1px solid #999;"></div>
            <div style="flex:1; border-right:1px solid #999;"></div>
            <div style="flex:1; border-right:1px solid #999;"></div>
            <div style="flex:1.7;"></div>
          </div>
          <div style="position:absolute; left:0; right:0; top:{Y_DISCLAIMER}px;
                      text-align:center; font-size:5px; color:#555; font-style:italic;">
            Każda drużyna zaczyna po jednym secie (w dowolnej kolejności)<br>
            Set przegrany przez 3 kolejne chybienia oznacza wynik 0:50
          </div>
          <div style="position:absolute; left:0; top:{Y_TBL_RESULTS}px;
                      width:{TROJKA_TBL_W_PX}px; height:{H_TBL_RESULTS}px;
                      border:1px solid #999;">
            <div style="position:absolute; left:0; top:0; bottom:0;
                        width:{TROJKA_R1_W_PX}px; overflow:hidden;
                        border-right:1px solid #999;">
              {elements_html}
              {f'<div style="position:absolute; left:0; right:0; top:{label_y_px}px; text-align:center; font-size:5px; font-weight:bold;">Wyniki<br>turnieju</div>' if include_qr else ''}
            </div>
            <div style="position:absolute; left:{TROJKA_R1_W_PX}px; top:0; right:0; bottom:0;">
              <div style="position:absolute; left:0; top:0; right:0; height:{H_TBL_HEADER}px;
                          background:#f0f0f0; display:flex; font-size:7px; font-weight:bold;
                          align-items:center; text-align:center; border-bottom:1px solid #999;">
                <div style="flex:0.5; border-right:1px solid #999;"></div>
                <div style="flex:4; border-right:1px solid #999;">SET 1</div>
                <div style="flex:4;">SET 2</div>
              </div>
              <div style="position:absolute; left:0; top:{H_TBL_HEADER}px; right:0; height:{H_TBL_SUMA}px;
                          background:#f8f8f8; display:flex; font-size:5px; font-weight:bold;
                          align-items:center; text-align:center;
                          border-bottom:1px solid #999;">
                <div style="flex:0.5; border-right:1px solid #999;"></div>
                <div style="flex:1.5; border-right:1px solid #999;"></div>
                <div style="flex:0.5; border-right:1px solid #999; writing-mode:vertical-rl; transform:rotate(180deg); line-height:1.2;">SUMA</div>
                <div style="flex:1.5; border-right:1px solid #999;"></div>
                <div style="flex:0.5; border-right:1px solid #999; writing-mode:vertical-rl; transform:rotate(180deg); line-height:1.2;">SUMA</div>
                <div style="flex:1.5; border-right:1px solid #999;"></div>
                <div style="flex:0.5; border-right:1px solid #999; writing-mode:vertical-rl; transform:rotate(180deg); line-height:1.2;">SUMA</div>
                <div style="flex:1.5; border-right:1px solid #999;"></div>
                <div style="flex:0.5; writing-mode:vertical-rl; transform:rotate(180deg); line-height:1.2;">SUMA</div>
              </div>
              <div style="position:absolute; left:0; top:{H_TBL_HEADER + H_TBL_SUMA}px; right:0; bottom:{H_PKT}px;
                          background:repeating-linear-gradient(to bottom, transparent 0, transparent 10px, #ddd 10px, #ddd 11px);"></div>
              <div style="position:absolute; left:0; right:0; bottom:0; height:{H_PKT}px;
                          border-top:1px solid #999; background:#f0f0f0; display:flex;
                          font-size:7px; font-weight:bold; text-align:center; align-items:center;">
                <div style="flex:0.5; border-right:1px solid #999; line-height:{H_PKT}px;">PKT</div>
                <div style="flex:8;"></div>
              </div>
            </div>
          </div>
        </div>
        """
    else:
        # ─── Preview INDYWIDUALNY ───
        # Wszystkie pozycje proporcjonalne do SCALE.
        Y_HEADER_TOP = int(0.4 * SCALE)
        Y_TOR_ROW = int(1.1 * SCALE)
        H_TOR = int(0.7 * SCALE)
        Y_PUNKTY = int(2.6 * SCALE)
        H_PUNKTY = int(1.0 * SCALE)
        Y_NAME_A = int(3.7 * SCALE)
        Y_NAME_B = int(4.4 * SCALE)
        H_NAME = int(0.7 * SCALE)
        Y_DISCLAIMER = int(5.3 * SCALE)
        Y_TBL_RESULTS = int(6.3 * SCALE)
        H_TBL_RESULTS = int((PAGE_H_CM - 6.3 - 0.5) * SCALE)
        H_TBL_HEADER = int(0.5 * SCALE)
        H_TBL_SUMA = int(0.9 * SCALE)
        H_WYNIK = int(0.5 * SCALE)
        
        html = f"""
    <div style="background:white; border:1px solid #ccc; 
                width:{PAGE_W_PX}px; height:{PAGE_H_PX}px;
                position:relative; font-family:Arial, sans-serif;
                box-shadow:0 2px 8px rgba(0,0,0,0.1); margin:0 auto;
                overflow:hidden;">
      <div style="position:absolute; right:6px; top:{Y_HEADER_TOP}px; 
                  font-size:7px; color:#666; font-style:italic;">
        {header_text}
      </div>
      <div style="position:absolute; left:{TBL1_OFFSET_PX}px; top:{Y_TOR_ROW}px; right:0;
                  height:{H_TOR}px;
                  display:flex; align-items:center; padding-left:6px;
                  font-size:8px; gap:8px;">
        <span>Tor <b>1</b></span>
        <span>Godzina <b>09:30</b></span>
        {('' if is_pucharowa else '<span>Grupa <b>A</b></span>')}
        <span>Mecz # <b>1</b></span>
      </div>
      <div style="position:absolute; left:{TBL1_OFFSET_PX + NAMES_PX}px; top:{Y_PUNKTY}px;
                  right:0; height:{H_PUNKTY}px;
                  background:#f0f0f0; border:1px solid #999;
                  display:flex; font-size:6px; font-weight:bold; text-align:center;
                  align-items:center; justify-content:space-around;">
        <div style="flex:1; border-right:1px solid #999;">Punkty<br>SET 1</div>
        <div style="flex:1; border-right:1px solid #999;">Punkty<br>SET 2</div>
        <div style="flex:1.15; border-right:1px solid #999;">Wygrane<br>sety</div>
        <div style="flex:1.95;">Podpis</div>
      </div>
      <div style="position:absolute; left:{TBL1_OFFSET_PX}px; top:{Y_NAME_A}px;
                  right:0; height:{H_NAME}px;
                  border:1px solid #999; display:flex; font-size:7px;">
        <div style="flex:0 0 {NAMES_PX}px; padding-right:6px; 
                    text-align:right; line-height:{H_NAME}px; font-weight:bold;
                    border-right:1px solid #999;">Łukasz Szulc</div>
        <div style="flex:1; border-right:1px solid #999;"></div>
        <div style="flex:1; border-right:1px solid #999;"></div>
        <div style="flex:1.15; border-right:1px solid #999;"></div>
        <div style="flex:1.95;"></div>
      </div>
      <div style="position:absolute; left:{TBL1_OFFSET_PX}px; top:{Y_NAME_B}px;
                  right:0; height:{H_NAME}px;
                  border:1px solid #999; border-top:none; display:flex; font-size:7px;">
        <div style="flex:0 0 {NAMES_PX}px; padding-right:6px;
                    text-align:right; line-height:{H_NAME}px; font-weight:bold;
                    border-right:1px solid #999;">Anna Ściepuro</div>
        <div style="flex:1; border-right:1px solid #999;"></div>
        <div style="flex:1; border-right:1px solid #999;"></div>
        <div style="flex:1.15; border-right:1px solid #999;"></div>
        <div style="flex:1.95;"></div>
      </div>
      <div style="position:absolute; left:0; right:0; top:{Y_DISCLAIMER}px;
                  text-align:center; font-size:5px; color:#555; font-style:italic;">
        Każdy zawodnik zaczyna po jednym secie (w dowolnej kolejności)<br>
        Set przegrany przez 3 kolejne chybienia oznacza wynik 0:50
      </div>
      <div style="position:absolute; left:0; top:{Y_TBL_RESULTS}px; right:0;
                  height:{H_TBL_RESULTS}px; border:1px solid #999;">
        <div style="position:absolute; left:0; top:0; bottom:0;
                    width:{LEFT_AREA_PX}px; border-right:1px solid #999;
                    overflow:hidden;">
          {elements_html}
          {f'<div style="position:absolute; left:0; right:0; top:{label_y_px}px; text-align:center; font-size:7px; font-weight:bold;">Wyniki turnieju</div>' if include_qr else ''}
        </div>
        <div style="position:absolute; left:{LEFT_AREA_PX+1}px; top:0; right:0; bottom:0;">
          <div style="position:absolute; left:0; top:0; right:0; height:{H_TBL_HEADER}px;
                      background:#f0f0f0; display:flex; font-size:7px; font-weight:bold;
                      align-items:center; text-align:center; border-bottom:1px solid #999;">
            <div style="flex:0.84; border-right:1px solid #999;"></div>
            <div style="flex:2.94; border-right:1px solid #999;">SET 1</div>
            <div style="flex:2.94;">SET 2</div>
          </div>
          <div style="position:absolute; left:0; top:{H_TBL_HEADER}px; right:0; height:{H_TBL_SUMA}px;
                      background:#f8f8f8; display:flex; font-size:5px; font-weight:bold;
                      align-items:center; text-align:center; border-bottom:1px solid #999;">
            <div style="flex:0.84; border-right:1px solid #999; writing-mode:vertical-rl; transform:rotate(180deg); line-height:1.2;">IMIONA</div>
            <div style="flex:0.735; border-right:1px solid #999;"></div>
            <div style="flex:0.735; border-right:1px solid #999; writing-mode:vertical-rl; transform:rotate(180deg); line-height:1.2;">SUMA</div>
            <div style="flex:0.735; border-right:1px solid #999;"></div>
            <div style="flex:0.735; border-right:1px solid #999; writing-mode:vertical-rl; transform:rotate(180deg); line-height:1.2;">SUMA</div>
            <div style="flex:0.735; border-right:1px solid #999;"></div>
            <div style="flex:0.735; border-right:1px solid #999; writing-mode:vertical-rl; transform:rotate(180deg); line-height:1.2;">SUMA</div>
            <div style="flex:0.735; border-right:1px solid #999;"></div>
            <div style="flex:0.735; writing-mode:vertical-rl; transform:rotate(180deg); line-height:1.2;">SUMA</div>
          </div>
          <div style="position:absolute; left:0; top:{H_TBL_HEADER + H_TBL_SUMA}px; right:0; bottom:{H_WYNIK}px;
                      background:repeating-linear-gradient(to bottom, transparent 0, transparent 10px, #ddd 10px, #ddd 11px);"></div>
          <div style="position:absolute; left:0; right:0; bottom:0; height:{H_WYNIK}px;
                      border-top:1px solid #999; background:#f0f0f0; display:flex;
                      font-size:7px; font-weight:bold; text-align:center; align-items:center;">
            <div style="flex:0.84; border-right:1px solid #999; line-height:{H_WYNIK}px;">WYNIK</div>
            <div style="flex:7.46;"></div>
          </div>
        </div>
      </div>
    </div>
    """
    st.components.v1.html(html, height=PAGE_H_PX + 30, scrolling=False)
    st.caption("📝 Schemat. Dokładny wygląd w pobranym pliku.")
    
    # ─── Pozycje grafik ──────────────────────────────────────────────────
    # Sekcja edycji pozycji obrazów - bezpośrednio pod podglądem żeby
    # użytkownik widział od razu wpływ zmian.
    if elements_active:
        cols_pos = st.columns([3, 1])
        with cols_pos[0]:
            st.markdown("**📍 Pozycje grafik**")
        with cols_pos[1]:
            if st.button("↻ Reset",
                         help="Przywróć automatyczne rozmieszczenie",
                         use_container_width=True,
                         key="reset_positions_btn"):
                st.session_state['reset_nonce'] = st.session_state.get('reset_nonce', 0) + 1
                st.rerun()
        
        nonce = st.session_state.get('reset_nonce', 0)
        max_x = TROJKA_LEFT_AREA_CM if is_trojka else 5.0
        max_w = TROJKA_LEFT_AREA_CM if is_trojka else 5.0
        for key, label in elements_active:
            with st.expander(f"📍 {label}", expanded=False):
                cols = st.columns(4)
                pos = image_positions[key]
                slider_key_base = f"{key}_{active_keys_signature}_{nonce}"
                with cols[0]:
                    new_x = st.number_input("X (cm)", value=float(pos['x']),
                                           min_value=0.0, max_value=max_x, step=0.1,
                                           key=f"x_{slider_key_base}")
                with cols[1]:
                    new_y = st.number_input("Y (cm)", value=float(pos['y']),
                                           min_value=0.0, max_value=18.0, step=0.1,
                                           key=f"y_{slider_key_base}")
                with cols[2]:
                    new_w = st.number_input("Szer. (cm)", value=float(pos['w']),
                                           min_value=0.5, max_value=max_w, step=0.1,
                                           key=f"w_{slider_key_base}")
                with cols[3]:
                    new_h = st.number_input("Wys. (cm)", value=float(pos['h']),
                                           min_value=0.5, max_value=5.0, step=0.1,
                                           key=f"h_{slider_key_base}")
                image_positions[key] = {
                    'x': new_x, 'y': new_y, 'w': new_w, 'h': new_h
                }


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
st.header("4. Generuj")

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
        # Tekst fazy w prawym górnym rogu (obok nazwy/daty turnieju):
        # Grupowa → "Faza grupowa", pucharowe → "1/32 finału", "Półfinał", "Finał" itp.
        if show_header_on_protocol:
            if is_pucharowa:
                if phase_label_short in ("Półfinał", "3. miejsce", "Finał"):
                    show_phase = phase_label_short
                else:
                    show_phase = f"{phase_label_short} finału"
            else:
                show_phase = "Faza grupowa"
        else:
            show_phase = ""
        # Mapowanie typu turnieju → szablon docx
        template_type = 'TROJKA' if is_trojka else 'IND'
        docx_bytes = generate_docx.build_document(
            sid, sheets_url.strip(), sheets_data,
            logos=logos_bytes or None,
            tournament_name=show_name, tournament_date=show_date,
            tournament_phase_text=show_phase,
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
        if show_header_on_protocol:
            if is_pucharowa:
                if phase_label_short in ("Półfinał", "3. miejsce", "Finał"):
                    show_phase = phase_label_short
                else:
                    show_phase = f"{phase_label_short} finału"
            else:
                show_phase = "Faza grupowa"
        else:
            show_phase = ""
        template_type = 'TROJKA' if is_trojka else 'IND'
        docx_bytes = generate_docx.build_blank_document(
            num_pages=1, logos=logos_bytes or None,
            tournament_name=show_name, tournament_date=show_date,
            tournament_phase_text=show_phase,
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
        # Polski plural dla "protokół": 1 protokół, 2-4 protokoły, 5+ protokołów
        n_p = gen['total']
        p_last = n_p % 10
        p_last2 = n_p % 100
        if n_p == 1:
            p_word = "protokół"
        elif p_last in (2, 3, 4) and p_last2 not in (12, 13, 14):
            p_word = "protokoły"
        else:
            p_word = "protokołów"
        if gen.get('is_pucharowa'):
            st.success(f"✅ Gotowe! {n_p} {p_word} dla fazy {gen.get('phase_name','')}.")
        else:
            n_g = gen['groups']
            g_last = n_g % 10
            g_last2 = n_g % 100
            if n_g == 1:
                g_word = "grupie"
            elif g_last in (2, 3, 4) and g_last2 not in (12, 13, 14):
                g_word = "grupach"
            else:
                g_word = "grupach"
            st.success(f"✅ Gotowe! {n_p} {p_word} w {n_g} {g_word}.")
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
