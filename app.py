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


class _CachedUpload:
    """Wrapper na cache'owane bajty pliku - emuluje interfejs UploadedFile.
    Używany żeby zachować pliki między rerunami Streamlita gdy file_uploader
    traci stan (np. po download_button)."""
    def __init__(self, name, raw_bytes):
        self.name = name
        self._bytes = raw_bytes
        self._pos = 0
    def seek(self, pos):
        self._pos = pos
    def read(self, n=-1):
        if n < 0:
            data = self._bytes[self._pos:]
            self._pos = len(self._bytes)
        else:
            data = self._bytes[self._pos:self._pos+n]
            self._pos += len(data)
        return data
    def getvalue(self):
        return self._bytes


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
# To sporo miejsca na grafiki - wystarczy żeby grafiki były czytelnie widoczne.
# Zostawiamy 0.3 cm marginesu od prawej (kolumny numerków 1-18).
TROJKA_LEFT_AREA_CM = 4.76    # pełna szerokość komórki R1.tc[0] - obrazy idealnie wyśrodkowane
TROJKA_GRAPHIC_W_CM = 2.7     # max szer logo w trójce (zbliżona do PFM domyślnego ~2.18 - spójność wizualna)
TROJKA_AREA_HEIGHT_CM = 18.5  # Trójka: lewa kolumna tabeli wynikowej, z R1 aż do dolnej krawędzi (PKT)


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
        area_width_cm = TROJKA_LEFT_AREA_CM  # 4.76 cm - pełna komórka R1.tc[0]
        area_height_cm = TROJKA_AREA_HEIGHT_CM
        QR_W = 2.0  # mniej niż w IND (2.4) ale wystarczająco czytelne
        PFM_TARGET_W = QR_W * 1.09  # ~2.18 cm
        OTHER_MAX_W = TROJKA_GRAPHIC_W_CM  # 4.0
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
        # Trójka: kompensacja cellMargin (0.185 cm) - patrz komentarz przy logo poniżej
        if is_trojka:
            qr_x -= 0.185
        positions['qr'] = {'x': qr_x, 'y': 0.2, 'w': QR_W, 'h': QR_W}
        # Po QR: napis "Wyniki turnieju" + WYRAŹNY odstęp do pierwszego logo.
        # text_h to wysokość napisu, SPACE_AFTER_QR to odstęp od napisu do logo.
        text_h = 0.5
        first_logo_y = 0.2 + QR_W + text_h + SPACE_AFTER_QR
        pfm_w_target = PFM_TARGET_W
    else:
        first_logo_y = 0.3
        pfm_w_target = PFM_TARGET_W
    
    if not other_keys:
        return positions
    
    n = len(other_keys)
    available = area_height_cm - first_logo_y
    
    # SPACING — w trójce węższy pas, ale ZACHOWAĆ widoczne odstępy.
    if is_trojka:
        if n <= 2:
            SPACING = 1.2
        elif n <= 4:
            SPACING = 0.8
        else:
            SPACING = 0.5
        max_slot_h_default_max = 2.7
        max_slot_h_default_min = 1.4
    else:
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
        # OBLICZENIE z myślą o WYPEŁNIENIU CAŁEJ DOSTĘPNEJ WYSOKOŚCI:
        # available = n*slot_h + (n-1)*SPACING
        # → slot_h = (available - (n-1)*SPACING) / n
        max_slot_h = (available - (n - 1) * SPACING) / n
        # Min/max z clamp żeby nie były za małe ani za duże
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
        
        # Skalowanie do max_slot_h (jeśli za wysokie)
        if target_h > max_slot_h:
            target_h = max_slot_h
            target_w = target_h * aspect
            if target_w > area_width_cm - 0.2:
                target_w = area_width_cm - 0.2
                target_h = target_w / aspect
        
        # Centrujemy obraz w slocie pionowo (gdy slot wyższy niż obraz)
        img_y = cur_y + (max_slot_h - target_h) / 2
        img_x = (area_width_cm - target_w) / 2
        # Trójka: kompensacja cellMargin (105 dxa = 0.185 cm) przesuwającego paragraf
        # w prawo wewnątrz komórki R1.tc[0]. Bez tego grafiki są subtelnie offset w prawo
        # względem geometrycznego środka przestrzeni "margines lewy → widoczna krawędź tabeli".
        if is_trojka:
            img_x -= 0.185
        
        positions[key] = {'x': img_x, 'y': img_y, 'w': target_w, 'h': target_h}
        cur_y += max_slot_h + SPACING
    
    return positions


# ════════════════════════════════════════════════════════════════════════
col_form, col_preview = st.columns([3, 2])
# ════════════════════════════════════════════════════════════════════════

with col_form:

    # ─── 1. Turniej ──────────────────────────────────────────────────────
    st.header("1. Turniej")
    
    # ─── Wiersz 1: Nazwa | Data | Rodzaj ──────────────────────────────
    cols_t1 = st.columns([5, 3, 4])
    with cols_t1[0]:
        tournament_name = st.text_input("Nazwa", value="TDT 2026",
                                        placeholder="np. GP2 2026")
    with cols_t1[1]:
        next_sat = date.today() + timedelta(days=(5 - date.today().weekday()) % 7)
        tournament_date_d = st.date_input("Data", value=next_sat, format="DD.MM.YYYY")
    tournament_date = tournament_date_d.strftime("%d.%m.%Y") if tournament_date_d else ""
    with cols_t1[2]:
        rodzaj_options = {
            "Indywidualny":     "✅ Indywidualny",
            "Drużynowy 2-os.":  "🔴 Drużynowy 2-os. (wkrótce)",
            "Drużynowy 3-os.":  "✅ Drużynowy 3-os.",
            "Drużynowy 4-os.":  "🔴 Drużynowy 4-os. (wkrótce)",
        }
        tournament_type = st.selectbox(
            "Rodzaj",
            list(rodzaj_options.keys()),
            format_func=lambda k: rodzaj_options[k],
            index=0,
            help="✅ = w pełni zweryfikowane, 🟡 = zaimplementowane lecz niezweryfikowane, 🔴 = w trakcie przygotowywania"
        )
    
    is_trojka_pre = tournament_type == "Drużynowy 3-os."
    
    # ─── Wiersz 2: Drabinka | Faza | Format ────────────────────────────
    # Dwupoziomowy wybór: najpierw drabinka (Grupowa / Główna / Turniej odpadów),
    # potem konkretna faza w obrębie tej drabinki.
    cols_t2 = st.columns([3, 4, 3])
    with cols_t2[0]:
        if is_trojka_pre:
            drabinka_options = {
                "Grupowa": "✅ Faza grupowa",
                "Główna":  "✅ Drabinka główna",
                "Odpadów": "🟡 Turniej odpadów",
            }
        else:
            drabinka_options = {
                "Grupowa": "✅ Faza grupowa",
                "Główna":  "🔴 Drabinka główna (wkrótce)",
                "Odpadów": "🔴 Turniej odpadów (wkrótce)",
            }
        drabinka = st.selectbox(
            "Drabinka",
            list(drabinka_options.keys()),
            format_func=lambda k: drabinka_options[k],
            index=0,
            help=("Drabinka główna prowadzi do finału (1/8, 1/4, 1/2, Finał). "
                  "Turniej odpadów — mecze o miejsca dla zespołów które odpadły "
                  "(grane równolegle z drabinką główną).")
        )
    
    with cols_t2[1]:
        # Faza zależna od wybranej drabinki
        if drabinka == "Grupowa":
            tournament_phase = "Grupowa"
            st.selectbox("Faza", ["Faza grupowa"],
                         format_func=lambda k: f"✅ {k}",
                         index=0, disabled=True,
                         help="Wybrano fazę grupową — nie ma podfaz.")
        elif drabinka == "Główna":
            if is_trojka_pre:
                faza_options = {
                    "Pucharowa 1/16": "🔴 1/16 finału (niedostępne dla trójek)",
                    "Pucharowa 1/8":  "✅ 1/8 finału",
                    "Pucharowa 1/4":  "✅ 1/4 finału (Ćwierćfinał)",
                    "Pucharowa 1/2":  "✅ Półfinał",
                    "Finał":          "✅ Finał",
                }
                default_idx = 1  # 1/8 finału
            else:
                faza_options = {
                    "Pucharowa 1/64": "🔴 1/64 finału",
                    "Pucharowa 1/32": "🔴 1/32 finału",
                    "Pucharowa 1/16": "🔴 1/16 finału",
                    "Pucharowa 1/8":  "🔴 1/8 finału",
                    "Pucharowa 1/4":  "🔴 1/4 finału (Ćwierćfinał)",
                    "Pucharowa 1/2":  "🔴 Półfinał",
                    "Finał":          "🔴 Finał",
                }
                default_idx = 0
            tournament_phase = st.selectbox(
                "Faza",
                list(faza_options.keys()),
                format_func=lambda k: faza_options[k],
                index=default_idx,
            )
        else:  # Odpadów
            # Pełna lista możliwych meczów drabinki przegranych:
            # • Mecze o pojedyncze miejsce (3, 5, 7, ..., 31)
            # • Mini-turnieje (Miejsca X-Y)
            if is_trojka_pre:
                # Mecz o N (single match) — typowo 3, 5, 7, 9, 11, 13, 15
                # Miejsca X-Y (mini-bracket) — typowo 5-8, 9-12, 9-16, 13-16
                faza_options = {
                    "Mecz o 3. miejsce":  "✅ Mecz o 3. miejsce",
                    "Mecz o 5. miejsce":  "🟡 Mecz o 5. miejsce",
                    "Mecz o 7. miejsce":  "🟡 Mecz o 7. miejsce",
                    "Mecz o 9. miejsce":  "🟡 Mecz o 9. miejsce",
                    "Mecz o 11. miejsce": "🟡 Mecz o 11. miejsce",
                    "Mecz o 13. miejsce": "🟡 Mecz o 13. miejsce",
                    "Mecz o 15. miejsce": "🟡 Mecz o 15. miejsce",
                    "Miejsca 5-8":   "🟡 Miejsca 5-8",
                    "Miejsca 9-12":  "🟡 Miejsca 9-12",
                    "Miejsca 9-16":  "🟡 Miejsca 9-16",
                    "Miejsca 13-16": "🟡 Miejsca 13-16",
                }
            else:
                faza_options = {
                    "Mecz o 3. miejsce":  "🔴 Mecz o 3. miejsce",
                    "Mecz o 5. miejsce":  "🔴 Mecz o 5. miejsce",
                    "Mecz o 7. miejsce":  "🔴 Mecz o 7. miejsce",
                    "Mecz o 9. miejsce":  "🔴 Mecz o 9. miejsce",
                    "Mecz o 11. miejsce": "🔴 Mecz o 11. miejsce",
                    "Mecz o 13. miejsce": "🔴 Mecz o 13. miejsce",
                    "Mecz o 15. miejsce": "🔴 Mecz o 15. miejsce",
                    "Mecz o 17. miejsce": "🔴 Mecz o 17. miejsce",
                    "Mecz o 19. miejsce": "🔴 Mecz o 19. miejsce",
                    "Mecz o 21. miejsce": "🔴 Mecz o 21. miejsce",
                    "Mecz o 23. miejsce": "🔴 Mecz o 23. miejsce",
                    "Mecz o 25. miejsce": "🔴 Mecz o 25. miejsce",
                    "Mecz o 27. miejsce": "🔴 Mecz o 27. miejsce",
                    "Mecz o 29. miejsce": "🔴 Mecz o 29. miejsce",
                    "Mecz o 31. miejsce": "🔴 Mecz o 31. miejsce",
                    "Miejsca 5-8":   "🔴 Miejsca 5-8",
                    "Miejsca 9-12":  "🔴 Miejsca 9-12",
                    "Miejsca 9-16":  "🔴 Miejsca 9-16",
                    "Miejsca 13-16": "🔴 Miejsca 13-16",
                    "Miejsca 17-20": "🔴 Miejsca 17-20",
                    "Miejsca 17-24": "🔴 Miejsca 17-24",
                    "Miejsca 17-32": "🔴 Miejsca 17-32",
                    "Miejsca 21-24": "🔴 Miejsca 21-24",
                    "Miejsca 25-28": "🔴 Miejsca 25-28",
                    "Miejsca 25-32": "🔴 Miejsca 25-32",
                    "Miejsca 29-32": "🔴 Miejsca 29-32",
                }
            tournament_phase = st.selectbox(
                "Faza",
                list(faza_options.keys()),
                format_func=lambda k: faza_options[k],
                index=0,
            )
    
    with cols_t2[2]:
        # Format setów zależny od drabinki:
        # Grupowa → tylko "2 sety"
        # Główna / Odpadów (pucharowe) → tylko Bo3 / Bo5 (2 sety NIE jest opcją)
        if drabinka == "Grupowa":
            sets_format = st.selectbox("Format", ["2 sety"],
                format_func=lambda k: f"✅ {k}",
                index=0,
                help="Faza grupowa zawsze 2 sety")
        else:
            is_trojka_now = tournament_type == "Drużynowy 3-os."
            if is_trojka_now:
                bo_options = {
                    "Best of 3": "✅ Best of 3",
                    "Best of 5": "✅ Best of 5",
                }
            else:
                bo_options = {
                    "Best of 3": "🔴 Best of 3 (wkrótce)",
                    "Best of 5": "🔴 Best of 5 (wkrótce)",
                }
            # Default: Bo3 dla 1/8, Bo5 dla 1/4 i wyżej (zgodnie z FAQ)
            default_idx = 0
            if "1/4" in tournament_phase or "1/2" in tournament_phase or \
               tournament_phase == "Finał" or tournament_phase == "Mecz o 3. miejsce":
                default_idx = 1
            sets_format = st.selectbox("Format",
                list(bo_options.keys()),
                format_func=lambda k: bo_options[k],
                index=default_idx,
                help="Faza pucharowa zawsze Best of 3 lub Best of 5."
            )

    # Walidacja - obecnie sprawdzone i zweryfikowane:
    # 1. Indywidualny + Grupowa + 2 sety  → ✅
    # 2. Drużynowy 3-os. + Grupowa + 2 sety  → ✅
    # Wszystko inne — eksperymentalne lub w trakcie przygotowywania.
    is_individual = tournament_type == "Indywidualny"
    is_trojka = tournament_type == "Drużynowy 3-os."
    is_supported_type = is_individual or is_trojka
    is_2_sety = sets_format == "2 sety"
    is_grupowa = tournament_phase == "Grupowa"
    is_pucharowa = (
        "Pucharowa" in tournament_phase or
        "Miejsca" in tournament_phase or
        tournament_phase in ("Mecz o 3. miejsce", "Finał")
    )
    # is_drabinka_b: faza z meczów o miejsca (drabinka przegranych) — w UI
    # wybierana z sekcji "Drabinka B".
    is_drabinka_b = (
        "Miejsca" in tournament_phase or
        tournament_phase == "Mecz o 3. miejsce"
    )
    
    is_fully_tested = is_supported_type and is_2_sety and is_grupowa
    
    # Pełna czytelna etykieta fazy do wyświetlenia w prawym górnym rogu protokołu.
    # • "Pucharowa 1/8" → "1/8 finału"
    # • "Pucharowa 1/2" → "Półfinał"
    # • "Mecz o 9. miejsce" → "Mecz o 9. miejsce"
    # • "Miejsca 5-8" → "Miejsca 5-8" (NIE "5-8 finału" ani podobne)
    phase_label_short = ""
    if is_pucharowa:
        # "Mecz o N miejsce" — zachowaj pełną nazwę
        m = re.search(r'[Mm]ecz\s+o\s+(\d+)\.?\s*miejsc', tournament_phase)
        if m:
            phase_label_short = f"Mecz o {m.group(1)}. miejsce"
        # "Miejsca X-Y" — zachowaj pełną nazwę
        elif "Miejsca" in tournament_phase:
            phase_label_short = tournament_phase
        # "1/N finału" — z regex, ALE tylko gdy ma slash (nie zmatchuje "5-8")
        elif re.search(r'\b1/\d+\b', tournament_phase):
            m = re.search(r'(1/(\d+))', tournament_phase)
            if m.group(2) == '2':
                phase_label_short = "Półfinał"
            else:
                phase_label_short = f"{m.group(1)} finału"
        elif tournament_phase == "Finał":
            phase_label_short = "Finał"
    
    if not is_supported_type:
        st.warning(f"⚠️ Format **{tournament_type}** nie jest jeszcze obsługiwany. "
                   "Obecnie dostępne: Indywidualny, Drużynowy 3-os.")
    elif is_pucharowa:
        # Pucharowa wymaga best-of-3/5 - jeszcze nie ma. Można użyć "2 sety" jako
        # zastępstwo, ale nie jest to oficjalny format pucharowej.
        st.warning(f"⚠️ Faza **{tournament_phase}** wymaga formatu **Best of 3/5** "
                   "który jest jeszcze w przygotowaniu. Generator wczyta mecze z "
                   "zakładki **Drabinka** używając formatu 2 sety jako tymczasowego "
                   "rozwiązania — protokół może być nieoficjalny. **Zweryfikowana jest "
                   "obecnie tylko faza grupowa.**")
    elif not is_2_sety:
        st.warning("⚠️ Obecnie zaimplementowane są tylko mecze na **2 sety**. "
                   "Best of 3/5 będą dostępne w kolejnej wersji.")
    
    if is_pucharowa and is_supported_type:
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
    # Każda faza pucharowa NIE używa grafik (Bo3/Bo5 i 2 sety) — całe miejsce
    # potrzebne na rozszerzoną tabelę wyników. Tylko faza GRUPOWA ma grafiki.
    is_no_graphics = is_pucharowa
    if is_no_graphics:
        st.header("3. Domyślne elementy")
        st.info("ℹ️ **Faza pucharowa nie używa grafik** — "
                "całe miejsce jest potrzebne na rozszerzoną tabelę wyników. "
                "Ta sekcja oraz Grafiki sponsorów są pominięte przy generowaniu.")
        # Wyłącz wszystko poza nagłówkiem (ten i tak chcemy w prawym górnym rogu)
        include_qr = False
        show_header_on_protocol = True
        include_pfm_logo = False
    else:
        st.header("3. Domyślne elementy")
        cols_dom = st.columns(2)
        with cols_dom[0]:
            include_qr = st.checkbox("Kod QR (link do arkusza)", value=True)
            show_header_on_protocol = st.checkbox(
                "Nazwa, data i faza turnieju w prawym górnym rogu", value=True)
        with cols_dom[1]:
            include_pfm_logo = st.checkbox("Logo Polskiej Federacji Mölkky", value=True)

    # ─── 4. Grafiki (zwijane) - łączy dodawanie + pozycjonowanie ─────────
    NUM_LOGOS = 4
    logo_files = []
    
    if is_no_graphics:
        # Pomijamy całą sekcję grafik — pucharowa ich nie wspiera
        for _ in range(NUM_LOGOS):
            logo_files.append(None)
        logos_aspect = {}
    else:
        # Cache uploadowanych bajtów w session_state. Dzięki temu pliki "pamiętają się"
        # nawet gdy expander zostanie zwinięty albo widget straci state po download_button.
        # Klucze: 'logo_bytes_0' ... 'logo_bytes_3' przechowują (filename, bytes).
        
        # Sprawdź ile grafik już dodanych żeby pokazać liczbę (z widget LUB cache bajtów)
        n_uploaded = sum(
            1 for i in range(NUM_LOGOS)
            if (st.session_state.get(f'logo_{i}') is not None) or
               (st.session_state.get(f'logo_bytes_{i}') is not None)
        )
        
        grafiki_label = "Grafiki"
        if n_uploaded > 0:
            grafiki_label += f" – dodano {n_uploaded}/{NUM_LOGOS}"
        
        with st.expander(grafiki_label, expanded=(n_uploaded > 0)):
            st.caption("Dodaj do 4 dodatkowych grafik (np. logo sponsora, miasta, klubu).")
            cols_log = st.columns(2)
            # logo_files: lista (filename, bytes) lub None — taka żeby logika niżej działała.
            # Korzystamy z prostego wrappera _CachedUpload.
            for i in range(NUM_LOGOS):
                with cols_log[i % 2]:
                    f = st.file_uploader(f"Grafika {i+1}", type=["png","jpg","jpeg"],
                                         key=f"logo_{i}", label_visibility="visible")
                    if f is not None:
                        # Świeży upload — cache bajtów do session_state
                        f.seek(0)
                        raw_bytes = f.read()
                        f.seek(0)
                        st.session_state[f'logo_bytes_{i}'] = (f.name, raw_bytes)
                        logo_files.append(_CachedUpload(f.name, raw_bytes))
                    elif st.session_state.get(f'logo_bytes_{i}') is not None:
                        # Widget zwrócił None (np. po download_button rerun), ale mamy cache
                        name, raw_bytes = st.session_state[f'logo_bytes_{i}']
                        logo_files.append(_CachedUpload(name, raw_bytes))
                    else:
                        logo_files.append(None)
            
            # Aspect ratios uploadowanych grafik (potrzebne wcześniej do compute_default_positions)
            logos_aspect = {}
            for i, f in enumerate(logo_files):
                if f is not None:
                    try:
                        f.seek(0)
                        logos_aspect[f'logo{i+1}'] = get_image_aspect(f.read())
                        f.seek(0)
                    except Exception:
                        logos_aspect[f'logo{i+1}'] = 1.5
        
        # ─── Aktywne elementy ────────────────────────────────────────────
        elements_active = []
        if include_qr:
            elements_active.append(('qr', 'Kod QR'))
        if include_pfm_logo:
            elements_active.append(('pfm', 'Logo PFM'))
        for i, f in enumerate(logo_files):
            if f is not None:
                elements_active.append((f'logo{i+1}', f'Grafika {i+1}'))
        
        # Pozycje (uwzględniając typ szablonu - trójka ma węższy lewy pasek)
        default_pos = compute_default_positions([k for k, _ in elements_active],
                                                logos_aspect=logos_aspect,
                                                template_type=('TROJKA' if is_trojka else 'IND'))
        active_keys_signature = "|".join(sorted(k for k, _ in elements_active))
        
        image_positions = {}
        for key in default_pos:
            image_positions[key] = dict(default_pos[key])
        
        if elements_active:
            st.caption("📍 **Pozycje grafik** możesz dostosować pod podglądem po prawej.")
    
    # Po expanderze: dla bezpieczeństwa fallback gdyby któraś zmienna nie była ustawiona
    if 'logo_files' not in dir():
        logo_files = [None] * NUM_LOGOS
    if 'logos_aspect' not in dir():
        logos_aspect = {}
    if 'elements_active' not in dir():
        elements_active = []
        if include_qr:
            elements_active.append(('qr', 'Kod QR'))
        if include_pfm_logo:
            elements_active.append(('pfm', 'Logo PFM'))
    if 'image_positions' not in dir():
        image_positions = compute_default_positions(
            [k for k, _ in elements_active],
            logos_aspect=logos_aspect,
            template_type=('TROJKA' if is_trojka else 'IND'))
    if 'active_keys_signature' not in dir():
        active_keys_signature = "|".join(sorted(k for k, _ in elements_active))


# ─── PODGLĄD HTML/CSS po prawej (przesunięty 5cm w dół) ───────────────
with col_preview:
    st.markdown("<div style='height:64px;'></div>", unsafe_allow_html=True)
    st.markdown("##### 📄 Podgląd")

    SCALE = 16   # Większy podgląd dla lepszej czytelności (kolumna prawa 2/5 szerokości)
    PAGE_W_CM = 18.46
    PAGE_H_CM = 27.16
    PAGE_W_PX = int(PAGE_W_CM * SCALE)
    PAGE_H_PX = int(PAGE_H_CM * SCALE)
    # Dla trójki lewa kolumna jest węższa (2.09 cm w docu, ~1.59 cm na grafiki + 0.5 cm gap)
    LEFT_AREA_PX = int((TROJKA_LEFT_AREA_CM if is_trojka else LEFT_AREA_CM) * SCALE)
    # Trójka ma większą tabelę wynikową (więcej kolumn) - widać że tabela zaczyna się
    # zaraz po lewym pasku grafik (2.09 cm)
    TBL2_LEFT_PX = int((2.09 if is_trojka else LEFT_AREA_CM) * SCALE)

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
    if show_header_on_protocol:
        parts = []
        if tournament_name: parts.append(tournament_name)
        if tournament_date: parts.append(tournament_date)
        # Faza w prawym górnym rogu — phase_label_short zawiera już pełną
        # czytelną nazwę ("1/8 finału", "Półfinał", "Miejsca 5-8", "Mecz o 9. miejsce" itd.)
        if is_pucharowa:
            parts.append(phase_label_short)
        else:
            parts.append("Faza grupowa")
        header_text = " · ".join(parts)
    else:
        header_text = ""

    TBL1_OFFSET_FRAC = 600 / 9690
    TBL1_OFFSET_PX = int(PAGE_W_PX * TBL1_OFFSET_FRAC)
    NAMES_FRAC = 3915 / 9090
    NAMES_PX = int((PAGE_W_PX - TBL1_OFFSET_PX) * NAMES_FRAC)
    
    if is_trojka:
        # ─── Preview TRÓJKOWY ───
        # Bo2 (grupowa lub puchar 2-set): lewa kolumna R1.tc[0] poszerzona do 4.76 cm
        # na grafiki (QR + loga). Pucharowa (Bo3/Bo5/2sety): bez grafik — kolumna
        # IMIONA/numerów wierszy zostaje wąska (~0.7 cm).
        is_no_graphics_preview = is_pucharowa
        
        TROJKA_TBL_W_PX = int((18.46 / PAGE_W_CM) * PAGE_W_PX)  # pełna szerokość strony
        if is_no_graphics_preview:
            # Bez grafik: kolumna IMIONA/numerów rzędów (~0.7 cm)
            TROJKA_R1_W_PX = int((0.7 / PAGE_W_CM) * PAGE_W_PX)
            # Tabela 1 nazwy: dla Bo3/Bo5 cell tc[0] ma 5.24 cm (2970 dxa scaled to ~3358)
            TROJKA_NAMES_PX = int((5.24 / PAGE_W_CM) * PAGE_W_PX)
        else:
            # Bo2: tabela 1 nazwa po skalowaniu × scale_t1 = 10466/9026 ≈ 1.16: tc[0] ≈ 4541 dxa
            TROJKA_NAMES_PX = int((4541 / 10466) * TROJKA_TBL_W_PX)
            # Lewa kolumna tabeli wynikowej (R1.tc[0] = 2700 dxa = 4.76 cm)
            TROJKA_R1_W_PX = int((4.76 / PAGE_W_CM) * PAGE_W_PX)
        TROJKA_HEADER_W_PX = TROJKA_TBL_W_PX - TROJKA_NAMES_PX
        
        html = f"""
        <div style="background:white; border:1px solid #ccc;
                    width:{PAGE_W_PX}px; height:{PAGE_H_PX}px;
                    position:relative; font-family:Arial, sans-serif;
                    box-shadow:0 2px 8px rgba(0,0,0,0.1); margin:0 auto;">
          
          <div style="position:absolute; right:8px; top:6px;
                      font-size:9px; color:#666; font-style:italic;">
            {header_text}
          </div>
          
          <!-- Header: Tor / Godzina / Grupa (jeśli grupowa) / Mecz # -->
          <div style="position:absolute; left:6px; top:30px; right:0; height:24px;
                      display:flex; align-items:center;
                      font-size:11px; gap:18px;">
            <span>Tor <b>1</b></span>
            <span>Godzina <b>09:00</b></span>
            {('' if is_pucharowa else '<span>Grupa <b>A</b></span>')}
            <span>Mecz # <b>1</b></span>
          </div>
          
          <!-- Nagłówki Punkty/Wygrane/Podpis (prawa strona, nad drużynami) -->
          <div style="position:absolute; left:{TROJKA_NAMES_PX}px; top:62px;
                      width:{TROJKA_HEADER_W_PX}px; height:34px;
                      background:#f0f0f0; border:1px solid #999;
                      display:flex; font-size:8px; font-weight:bold; text-align:center;
                      align-items:center; justify-content:space-around;">
            <div style="flex:1; border-right:1px solid #999;">Punkty<br>SET 1</div>
            <div style="flex:1; border-right:1px solid #999;">Punkty<br>SET 2</div>
            <div style="flex:1; border-right:1px solid #999;">Wygrane<br>sety</div>
            <div style="flex:1.7;">Podpis</div>
          </div>
          
          <!-- Drużyna A -->
          <div style="position:absolute; left:0; top:96px;
                      width:{TROJKA_TBL_W_PX}px; height:24px;
                      border:1px solid #999; display:flex; font-size:9px;">
            <div style="flex:0 0 {TROJKA_NAMES_PX}px; padding-right:2px;
                        text-align:right; line-height:24px; font-weight:bold;
                        white-space:nowrap; overflow:hidden; text-overflow:ellipsis;
                        border-right:1px solid #999;">
              KARMI
            </div>
            <div style="flex:1; border-right:1px solid #999;"></div>
            <div style="flex:1; border-right:1px solid #999;"></div>
            <div style="flex:1; border-right:1px solid #999;"></div>
            <div style="flex:1.7;"></div>
          </div>
          <!-- Drużyna B -->
          <div style="position:absolute; left:0; top:120px;
                      width:{TROJKA_TBL_W_PX}px; height:24px;
                      border:1px solid #999; border-top:none; display:flex; font-size:9px;">
            <div style="flex:0 0 {TROJKA_NAMES_PX}px; padding-right:2px;
                        text-align:right; line-height:24px; font-weight:bold;
                        white-space:nowrap; overflow:hidden; text-overflow:ellipsis;
                        border-right:1px solid #999;">
              Trzech Silnych Mężczyzn
            </div>
            <div style="flex:1; border-right:1px solid #999;"></div>
            <div style="flex:1; border-right:1px solid #999;"></div>
            <div style="flex:1; border-right:1px solid #999;"></div>
            <div style="flex:1.7;"></div>
          </div>
          
          <!-- Disclaimer -->
          <div style="position:absolute; left:0; right:0; top:148px;
                      text-align:center; font-size:8px; color:#555; font-style:italic;">
            Każda drużyna zaczyna po jednym secie (w dowolnej kolejności)<br>
            Set przegrany przez 3 kolejne chybienia oznacza wynik 0:50
          </div>
          
          <!-- Tabela wynikowa (od lewego brzegu) — wysokość dopasowana
               żeby nie wystawała poza stronę: PAGE_H_PX - top(180px) - margines 4px -->
          <div style="position:absolute; left:0; top:180px;
                      width:{TROJKA_TBL_W_PX}px; height:{PAGE_H_PX - 184}px;
                      border:1px solid #999;">
            <!-- Lewa kolumna R1.tc[0] na grafiki (2.09 cm) -->
            <div style="position:absolute; left:0; top:0; bottom:0;
                        width:{TROJKA_R1_W_PX}px; overflow:hidden;
                        border-right:1px solid #999;">
              {elements_html}
              {f'''<div style="position:absolute; left:0; right:0; top:{label_y_px}px;
                          text-align:center; font-size:7px; font-weight:bold;">
                Wyniki<br>turnieju
              </div>''' if include_qr else ''}
            </div>
            <!-- Prawa część tabeli wyników (kolumna numerków + 4 grupy SUMA) -->
            <div style="position:absolute; left:{TROJKA_R1_W_PX}px; top:0; right:0; bottom:0;">
              <!-- SET 1 / SET 2 nagłówki -->
              <div style="position:absolute; left:0; top:0; right:0; height:22px;
                          background:#f0f0f0; display:flex; font-size:10px; font-weight:bold;
                          align-items:center; text-align:center; border-bottom:1px solid #999;">
                <div style="flex:0.5; border-right:1px solid #999;"></div>
                <div style="flex:4; border-right:1px solid #999;">SET 1</div>
                <div style="flex:4;">SET 2</div>
              </div>
              <!-- 4 SUMA nagłówki -->
              <div style="position:absolute; left:0; top:22px; right:0; height:50px;
                          background:#f8f8f8; display:flex; font-size:7px; font-weight:bold;
                          align-items:center; text-align:center;
                          border-bottom:1px solid #999;">
                <div style="flex:0.5; border-right:1px solid #999;"></div>
                <div style="flex:1.5; border-right:1px solid #999;"></div>
                <div style="flex:0.5; border-right:1px solid #999;
                            writing-mode:vertical-rl; transform:rotate(180deg);
                            line-height:1.2;">SUMA</div>
                <div style="flex:1.5; border-right:1px solid #999;"></div>
                <div style="flex:0.5; border-right:1px solid #999;
                            writing-mode:vertical-rl; transform:rotate(180deg);
                            line-height:1.2;">SUMA</div>
                <div style="flex:1.5; border-right:1px solid #999;"></div>
                <div style="flex:0.5; border-right:1px solid #999;
                            writing-mode:vertical-rl; transform:rotate(180deg);
                            line-height:1.2;">SUMA</div>
                <div style="flex:1.5; border-right:1px solid #999;"></div>
                <div style="flex:0.5;
                            writing-mode:vertical-rl; transform:rotate(180deg);
                            line-height:1.2;">SUMA</div>
              </div>
              <!-- Wiersze 1-18 -->
              <div style="position:absolute; left:0; top:72px; right:0; bottom:24px;
                          background:repeating-linear-gradient(
                            to bottom, transparent 0, transparent 13px,
                            #ddd 13px, #ddd 14px);"></div>
              <!-- Wiersz PKT na dole -->
              <div style="position:absolute; left:0; right:0; bottom:0; height:24px;
                          border-top:1px solid #999;
                          background:#f0f0f0; display:flex;
                          font-size:10px; font-weight:bold; text-align:center;
                          align-items:center;">
                <div style="flex:0.5; border-right:1px solid #999; line-height:24px;">PKT</div>
                <div style="flex:8;"></div>
              </div>
            </div>
          </div>
        </div>
        """
    else:
        # ─── Preview INDYWIDUALNY (oryginalny) ───
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
        {('' if is_pucharowa else '<span>Grupa <b>A</b></span>')}
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
                  border:1px solid #999; display:flex; font-size:9px;">
        <div style="flex:0 0 {NAMES_PX}px; padding-right:2px; 
                    text-align:right; line-height:24px; font-weight:bold;
                    white-space:nowrap; overflow:hidden; text-overflow:ellipsis;
                    border-right:1px solid #999;">
          Łukasz Szulc
        </div>
        <div style="flex:1; border-right:1px solid #999;"></div>
        <div style="flex:1; border-right:1px solid #999;"></div>
        <div style="flex:1.15; border-right:1px solid #999;"></div>
        <div style="flex:1.95;"></div>
      </div>
      <div style="position:absolute; left:{TBL1_OFFSET_PX}px; top:121px; right:0; height:24px;
                  border:1px solid #999; border-top:none; display:flex; font-size:9px;">
        <div style="flex:0 0 {NAMES_PX}px; padding-right:2px;
                    text-align:right; line-height:24px; font-weight:bold;
                    white-space:nowrap; overflow:hidden; text-overflow:ellipsis;
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
        # Lewa kolumna lub cały lewy obszar (granica do której obraz może sięgać).
        # Trójka: kolumna R1.tc[0] = 4.76 cm (TROJKA_LEFT_AREA_CM).
        # IND: lewy obszar 5.24 cm.
        # Z bezpiecznym marginesem 0.1 cm żeby obraz nie dotykał krawędzi kolumny.
        col_width = TROJKA_LEFT_AREA_CM if is_trojka else 5.24
        SAFE_MARGIN = 0.1
        max_x_abs = col_width - 0.5  # zostaw min. 0.5 cm na obraz
        max_h_abs = 5.0
        for key, label in elements_active:
            with st.expander(f"📍 {label}", expanded=False):
                cols = st.columns(4)
                pos = image_positions[key]
                slider_key_base = f"{key}_{active_keys_signature}_{nonce}"
                # Aktualne wartości (z możliwymi już zmodyfikowanymi w session_state)
                cur_x = float(pos['x'])
                cur_w = float(pos['w'])
                # Dynamiczne ograniczenia: x + w nie może przekroczyć szer. kolumny.
                # max_x ogranicza X tak, żeby przy obecnym W nadal się zmieścił.
                # max_w ogranicza W tak, żeby przy obecnym X nadal się zmieścił.
                max_x_dyn = max(0.0, col_width - max(0.5, cur_w) - SAFE_MARGIN)
                max_w_dyn = max(0.5, col_width - max(0.0, cur_x) - SAFE_MARGIN)
                # Clamp aktualnych wartości na wypadek gdyby były spoza zakresu
                # (np. zapisane z poprzedniej wersji z innymi stałymi).
                cur_x = min(cur_x, max_x_dyn)
                cur_w = min(cur_w, max_w_dyn)
                with cols[0]:
                    new_x = st.number_input("X (cm)", value=cur_x,
                                           min_value=0.0, max_value=max_x_dyn, step=0.1,
                                           key=f"x_{slider_key_base}")
                with cols[1]:
                    new_y = st.number_input("Y (cm)", value=float(pos['y']),
                                           min_value=0.0, max_value=18.0, step=0.1,
                                           key=f"y_{slider_key_base}")
                with cols[2]:
                    # max_w przeliczone na podstawie świeżego new_x z tej samej tury renderu
                    max_w_for_input = max(0.5, col_width - max(0.0, new_x) - SAFE_MARGIN)
                    new_w = st.number_input("Szer. (cm)", value=min(cur_w, max_w_for_input),
                                           min_value=0.5, max_value=max_w_for_input, step=0.1,
                                           key=f"w_{slider_key_base}")
                with cols[3]:
                    new_h = st.number_input("Wys. (cm)", value=float(pos['h']),
                                           min_value=0.5, max_value=max_h_abs, step=0.1,
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

with st.expander("➕ Pobierz pusty formularz"):
    st.caption("Pusty formularz przyda się gdy chcesz wydrukować protokół do ręcznego wypełnienia "
               "(bez pobierania danych z arkusza). Wybierz typ + format setów.")
    
    cols_blank = st.columns([2, 2, 3])
    with cols_blank[0]:
        blank_type = st.selectbox(
            "Typ protokołu",
            ["Indywidualny", "Drużynowy 3-os."],
            index=0, key="blank_type",
            help="Określa szablon docx (Grupa_IND.docx / Grupa_TROJKA.docx itd.)."
        )
    with cols_blank[1]:
        if blank_type == "Drużynowy 3-os.":
            blank_format = st.selectbox(
                "Format setów",
                ["2 sety (grupowa)", "Best of 3 (puchar)", "Best of 5 (puchar)"],
                index=0, key="blank_format"
            )
        else:
            blank_format = st.selectbox(
                "Format setów",
                ["2 sety (grupowa)"],
                index=0, key="blank_format",
                help="Best of 3/5 dla typu Indywidualny — wkrótce."
            )
    with cols_blank[2]:
        blank_prefill = st.checkbox(
            "Wypełnij nagłówek z ustawień powyżej",
            value=False, key="blank_prefill",
            help="Jeśli zaznaczone — formularz dostanie nazwę turnieju, datę i fazę z sekcji 1. "
                 "Jeśli nie — totalnie czysty (do ręcznego wpisania)."
        )
    
    blank_clicked = st.button("📝 Pobierz pusty formularz",
                              use_container_width=False, key="blank_btn")

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
        # Tekst fazy w prawym górnym rogu — phase_label_short zawiera już pełną
        # czytelną nazwę ("1/8 finału", "Półfinał", "Miejsca 5-8", "Mecz o 9. miejsce").
        if show_header_on_protocol:
            show_phase = phase_label_short if is_pucharowa else "Faza grupowa"
        else:
            show_phase = ""
        # Mapowanie typu turnieju → szablon docx
        # Trójka pucharowa: Best of 3 → Bo3_TROJKA.docx, Best of 5 → Bo5_TROJKA.docx
        if is_trojka and is_pucharowa and sets_format == "Best of 3":
            template_type = 'TROJKA_Bo3'
        elif is_trojka and is_pucharowa and sets_format == "Best of 5":
            template_type = 'TROJKA_Bo5'
        else:
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
        # Dodaj fazę + format setów do nazwy pliku.
        # np. "TDT_2026_pucharowa_1_4_Bo5", "TDT_2026_finał_Bo3", "TDT_2026_miejsca_9_16_Bo3"
        phase_suffix = re.sub(r'[^\w]', '_', tournament_phase).strip('_').lower()
        # Format: dla pucharowej zawsze Bo3 lub Bo5 (2 sety wyłączone w UI dla puchar).
        format_suffix = ''
        if sets_format == 'Best of 3':
            format_suffix = '_Bo3'
        elif sets_format == 'Best of 5':
            format_suffix = '_Bo5'
        safe_name = f"{safe_name}_{phase_suffix}{format_suffix}"
    
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
        # Mapowanie blank_type + blank_format → template_type
        if blank_type == "Drużynowy 3-os.":
            if "Best of 3" in blank_format:
                blank_template = 'TROJKA_Bo3'
            elif "Best of 5" in blank_format:
                blank_template = 'TROJKA_Bo5'
            else:
                blank_template = 'TROJKA'
        else:
            blank_template = 'IND'
        
        # Pre-fill: jeśli zaznaczone, użyj wartości z sekcji 1 (nazwa, data, faza).
        # Inaczej totalnie czysty (wszystko puste).
        if blank_prefill:
            blank_show_name = tournament_name.strip()
            blank_show_date = tournament_date
            blank_show_phase = phase_label_short if is_pucharowa else "Faza grupowa"
            logos_bytes, image_order, img_pos = build_image_args()
            blank_include_qr = include_qr
            blank_include_pfm = include_pfm_logo
        else:
            blank_show_name = ""
            blank_show_date = ""
            blank_show_phase = ""
            logos_bytes, image_order, img_pos = ([], None, {})
            blank_include_qr = False
            blank_include_pfm = False
        
        docx_bytes = generate_docx.build_blank_document(
            num_pages=1, logos=logos_bytes or None,
            tournament_name=blank_show_name, tournament_date=blank_show_date,
            tournament_phase_text=blank_show_phase,
            sheets_url=sheets_url.strip() if blank_prefill else "",
            include_qr=blank_include_qr, include_pfm_logo=blank_include_pfm,
            image_order=image_order or None, image_positions=img_pos or None,
            template_type=blank_template)

    # Nazwa pliku odpowiednia do typu szablonu i ewentualnego pre-fill
    template_label = {
        'IND': 'indywidualny_2sety',
        'TROJKA': 'trojka_2sety',
        'TROJKA_Bo3': 'trojka_Bo3',
        'TROJKA_Bo5': 'trojka_Bo5',
    }.get(blank_template, blank_template.lower())
    if blank_prefill and tournament_name.strip():
        prefix = re.sub(r'[^\w\s-]','', tournament_name).strip().replace(' ','_')
        safe_name = f"{prefix}_pusty_{template_label}"
    else:
        safe_name = f"pusty_formularz_{template_label}"
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
