#!/usr/bin/env python3
"""
◈ AERO·ROSE  —  Wind Rose Diagram Generator  v7
WEB by ABDUL SAMAD | Run: python windrose_final.py
"""
import sys, subprocess, io

def _in_streamlit():
    try:
        from streamlit.runtime.scriptrunner import get_script_run_ctx
        return get_script_run_ctx() is not None
    except Exception:
        return False

if __name__ == "__main__" and not _in_streamlit():
    DEPS = [("streamlit","streamlit"),("matplotlib","matplotlib"),
            ("numpy","numpy"),("pandas","pandas"),
            ("reportlab","reportlab"),("openpyxl","openpyxl"),("Pillow","PIL")]
    print("\n  ◈  AERO·ROSE  —  Wind Rose Generator\n  " + "─"*40)
    for pkg, mod in DEPS:
        try: __import__(mod)
        except ImportError:
            print(f"  Installing {pkg}…")
            subprocess.check_call(
                [sys.executable,"-m","pip","install",pkg,"-q","--disable-pip-version-check"],
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    print("  Launching →  http://localhost:8501\n")
    subprocess.run([sys.executable,"-m","streamlit","run",__file__,
        "--server.port=8501","--server.headless=false",
        "--browser.gatherUsageStats=false"])
    sys.exit(0)

# ══════════════════════════════════════════════════════════════════════
import matplotlib; matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np
import pandas as pd
import time
import streamlit as st
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors as RC
from reportlab.lib.units import cm
from reportlab.lib.utils import ImageReader
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                 Image as RLImage, PageBreak, HRFlowable)
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER

st.set_page_config(page_title="◈ AERO·ROSE", page_icon="🧭",
                   layout="wide", initial_sidebar_state="collapsed")

# ══════════════════════════════════════════════════════════════════════
#  CONSTANTS
# ══════════════════════════════════════════════════════════════════════
DIRS_16    = ["N","NNE","NE","ENE","E","ESE","SE","SSE",
              "S","SSW","SW","WSW","W","WNW","NW","NNW"]
C2D        = {d: i*22.5 for i,d in enumerate(DIRS_16)}

# Speed bins km/h  (5 classes matching the provided image)
SPD_BINS   = [-0.001, 6, 25, 40, 60, 9999]
SPD_LABELS = ["<6 km/h", "6–25 km/h", "25–40 km/h", "40–60 km/h", ">60 km/h"]
TBL_COLS   = ["6.0–25 km/h", "25–40 km/h", "40–60 km/h"]
TBL_IDX    = [1, 2, 3]

# Very distinct colors for Type II bars (work well on white diagram background)
T2_COLORS  = ["#b0bec5", "#1976d2", "#43a047", "#fb8c00", "#e53935"]
# Names for legend
T2_NAMES   = ["< 6 km/h (Calm)", "6–25 km/h", "25–40 km/h", "40–60 km/h", "> 60 km/h"]

# ── Theme definitions ──────────────────────────────────────────────────────────
# RULE: UI colors can be CSS (rgba/hex). Matplotlib m_ keys MUST be HEX only.
TH = {
    "dark": {
        # UI palette
        "bg":       "#030810",
        "card_css": "rgba(5,14,28,0.93)",
        "brd_css":  "rgba(0,200,255,0.18)",
        "brd2_css": "rgba(255,255,255,0.07)",
        "acc":  "#00d8ff", "acc2": "#005f9e",
        "gold": "#f5a623", "suc":  "#00e5a0", "dng": "#ff4d6d",
        "txt":  "#ddeeff", "mut":  "#5a7a9a",
        "ibg":  "#0d1e35", "itxt": "#ddeeff",   # solid input bg/text
        "pbg":  "#0d1e35", "ptxt": "#ddeeff",   # popup bg/text
        "psel": "#0a2a44", "phov": "#0f2840",
        "ebg":  "#08142a", "etxt": "#ddeeff",   # expander bg/text
        "shd":  "0 8px 40px rgba(0,0,0,0.60)",
        "blur": "blur(18px)",
        # Matplotlib HEX only
        "m_bg":    "#030810", "m_card":  "#07101e",
        "m_grid":  "#0f1d30", "m_tick":  "#8aaccc", "m_title": "#00d8ff",
        "m_poly":  "#00d8ff", "m_pfill": "#003850",
    },
    "light": {
        "bg":       "#e4eeff",
        "card_css": "rgba(255,255,255,0.96)",
        "brd_css":  "rgba(0,70,180,0.20)",
        "brd2_css": "rgba(0,0,0,0.08)",
        "acc":  "#004db3", "acc2": "#002d80",
        "gold": "#b06800", "suc":  "#005533", "dng": "#bb1122",
        "txt":  "#07192e", "mut":  "#3a5880",
        "ibg":  "#ffffff", "itxt": "#07192e",
        "pbg":  "#ffffff", "ptxt": "#07192e",
        "psel": "#deeaff", "phov": "#eef3ff",
        "ebg":  "#f0f5ff", "etxt": "#07192e",
        "shd":  "0 4px 22px rgba(0,50,160,0.14)",
        "blur": "blur(10px)",
        "m_bg":    "#f8f8ff", "m_card":  "#ffffff",
        "m_grid":  "#ccd4ee", "m_tick":  "#0d2244", "m_title": "#004db3",
        "m_poly":  "#004db3", "m_pfill": "#ccddf8",
    },
}

# Session state
_SS = dict(theme="dark", diagrams={}, freq=None, rwy1=None, rwy2=None,
           stats=None, cxlim=19.4, ready=False, show_table=False,
           _file_bytes=None, _file_name=None, _cols=None,
           _file_rows=0, _file_loaded=False,
           _pdf_name="", _pdf_roll="", _pdf_site="", _pdf_logo=None,
           _processing=False)
for k, v in _SS.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ══════════════════════════════════════════════════════════════════════
#  CSS  — fully theme-reactive, no expander hacks
# ══════════════════════════════════════════════════════════════════════
def inject_css():
    T  = TH[st.session_state.theme]
    dk = st.session_state.theme == "dark"
    g  = f"linear-gradient(135deg,{T['acc']},{T['acc2']})"
    BG = T["bg"]; TXT = T["txt"]; MUT = T["mut"]
    IBG = T["ibg"]; ITXT = T["itxt"]
    EBG = T["ebg"]; ETXT = T["etxt"]
    PBG = T["pbg"]; PTXT = T["ptxt"]

    bg_layers = (
        "radial-gradient(ellipse 70% 38% at 50% 100%,"
        f"{'rgba(255,150,30,.07)' if dk else 'rgba(0,80,200,.06)'} 0%,transparent 55%),"
        f"{'linear-gradient(180deg,#020509,#030810 30%,#040b18 70%,#020509)' if dk else 'linear-gradient(180deg,#dce8ff,#e4eeff 50%,#d8e4ff)'}"
    )
    overlay    = "rgba(2,6,16,0.68)" if dk else "rgba(215,228,255,0.56)"
    rwy_stripe = "rgba(255,255,255,0.038)" if dk else "rgba(0,50,180,0.034)"

    st.markdown(f"""<style>
@import url('https://fonts.googleapis.com/css2?family=Bebas+Neue&family=IBM+Plex+Sans:wght@300;400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap');

/* ── Variables ──────────────────────────────── */
:root{{
  --acc:{T['acc']};--acc2:{T['acc2']};--gold:{T['gold']};
  --suc:{T['suc']};--dng:{T['dng']};--txt:{TXT};--mut:{MUT};
  --card:{T['card_css']};--brd:{T['brd_css']};--brd2:{T['brd2_css']};
  --inb:{T['brd_css']};--grd:{g};
  --shd:{T['shd']};--blur:{T['blur']};--rad:12px;
}}

/* ── Global ─────────────────────────────────── */
html,body,.stApp,[data-testid="stApp"],
[data-testid="stAppViewContainer"]{{
  font-family:'IBM Plex Sans',sans-serif!important;
  color:{TXT}!important;
}}
[data-testid="stAppViewContainer"]{{
  background:{bg_layers};background-attachment:fixed;
}}
[data-testid="stAppViewContainer"]::after{{
  content:'';position:fixed;inset:0;pointer-events:none;z-index:0;
  background:{overlay};
}}
[data-testid="stAppViewContainer"]::before{{
  content:'';position:fixed;inset:0;pointer-events:none;z-index:0;
  background-image:repeating-linear-gradient(180deg,
    transparent 0,transparent 58px,{rwy_stripe} 58px,{rwy_stripe} 60px);
}}
.block-container{{
  position:relative;z-index:1;max-width:960px!important;
  margin:0 auto!important;padding:1.2rem 1.8rem 4rem!important;
}}

/* ── Hide sidebar ───────────────────────────── */
[data-testid="stSidebar"],section[data-testid="stSidebar"],
[data-testid="collapsedControl"],[data-testid="stSidebarNav"],
button[title="Open sidebar"],button[title="Close sidebar"],
[data-testid="stSidebarCollapseButton"]{{
  display:none!important;width:0!important;min-width:0!important;
}}
#MainMenu,header,footer{{visibility:hidden!important;}}

/* ── UNIVERSAL TEXT — all elements ─────────── */
*, *::before, *::after{{
  color:{TXT}!important;
  -webkit-text-fill-color:{TXT}!important;
}}
/* Exceptions for white text on colored backgrounds */
.stButton>button, .stDownloadButton>button,
button[data-testid], .ar-step-num,
.ar-hero .ar-eyebrow{{
  color:#ffffff!important;
  -webkit-text-fill-color:#ffffff!important;
}}
/* Accent colored elements */
.ar-lbl, .ar-sv, .ar-step-title, .ar-type-code,
.ar-title em, .ar-eyebrow, .ar-chip,
.ar-footer-brand, .ar-file-banner b,
.ar-cov b, .ar-pass, .ar-dlbl, .ar-sub,
.ar-tbl th, .ar-tbl td.dir-cell, .ar-tbl tr.trow td,
.ar-freq-hdr, .ar-freq-subhdr{{
  -webkit-text-fill-color:unset!important;
}}

/* ══ TEXT INPUTS  — guaranteed visible ══════════
   Use solid (opaque) background + explicit text color.
   -webkit-text-fill-color overrides WebKit's own fill.
══════════════════════════════════════════════ */
.stTextInput>div>div>input,
.stTextInput input,
.stNumberInput>div>div>input,
.stNumberInput input{{
  background:{IBG}!important;
  color:{ITXT}!important;
  -webkit-text-fill-color:{ITXT}!important;
  caret-color:{ITXT}!important;
  border:1.5px solid {T['brd_css']}!important;
  border-radius:8px!important;
  font-family:'IBM Plex Mono',monospace!important;
  font-size:.88rem!important;
  padding:.5rem .85rem!important;
}}
.stTextInput>div>div>input:focus,
.stNumberInput>div>div>input:focus{{
  border-color:{T['acc']}!important;
  box-shadow:0 0 0 3px {'rgba(0,210,255,0.15)' if dk else 'rgba(0,60,200,0.15)'}!important;
  outline:none!important;
}}
.stTextInput>div>div>input::placeholder,
.stNumberInput>div>div>input::placeholder{{
  color:{MUT}!important;
  -webkit-text-fill-color:{MUT}!important;
  opacity:0.75!important;
}}

/* ══ SELECTBOX ══════════════════════════════ */
.stSelectbox>div>div{{
  background:{IBG}!important;
  border:1.5px solid {T['brd_css']}!important;
  border-radius:8px!important;
}}
.stSelectbox [class*="singleValue"],
.stSelectbox [class*="placeholder"],
.stSelectbox [data-baseweb="select"] span,
.stSelectbox [data-baseweb="select"] div{{
  color:{ITXT}!important;
  -webkit-text-fill-color:{ITXT}!important;
  background:transparent!important;
}}
/* Dropdown popover */
[data-baseweb="popover"],[data-baseweb="popover"]>div,
div[data-baseweb="menu"],ul[data-baseweb="menu"],
[data-baseweb="popover"] ul{{
  background-color:{PBG}!important;
  border:1px solid {T['brd_css']}!important;
  border-radius:8px!important;
  box-shadow:0 8px 32px {'rgba(0,0,0,0.55)' if dk else 'rgba(0,40,160,0.18)'}!important;
}}
[data-baseweb="popover"] li,[data-baseweb="popover"] [role="option"],
div[data-baseweb="option"],li[data-baseweb="option"]{{
  background:{PBG}!important;
  color:{PTXT}!important;
  -webkit-text-fill-color:{PTXT}!important;
  font-family:'IBM Plex Mono',monospace!important;
  font-size:.86rem!important;
}}
[data-baseweb="popover"] li:hover,div[data-baseweb="option"]:hover{{
  background:{T['phov']}!important;
}}
[data-baseweb="popover"] [aria-selected="true"],li[aria-selected="true"]{{
  background:{T['psel']}!important;
  color:{T['acc']}!important;
  -webkit-text-fill-color:{T['acc']}!important;
  font-weight:600!important;
}}

/* ══ LABELS ═════════════════════════════════ */
.stTextInput label,.stSelectbox label,.stNumberInput label,
.stFileUploader label,.stCheckbox label,.stRadio label{{
  color:{MUT}!important;
  -webkit-text-fill-color:{MUT}!important;
  font-family:'IBM Plex Mono',monospace!important;
  font-size:.7rem!important;font-weight:500!important;
  letter-spacing:.1em!important;text-transform:uppercase!important;
}}
div[data-testid="stCheckbox"] label span,
div[data-testid="stRadio"] label span,
div[data-testid="stCheckbox"] p,
div[data-testid="stRadio"] p{{
  color:{TXT}!important;
  -webkit-text-fill-color:{TXT}!important;
  font-family:'IBM Plex Sans',sans-serif!important;
  font-size:.9rem!important;
  text-transform:none!important;letter-spacing:0!important;
}}

/* ══ BUTTONS ════════════════════════════════ */
.stButton>button{{
  background:{g}!important;color:#ffffff!important;
  -webkit-text-fill-color:#ffffff!important;
  border:none!important;border-radius:8px!important;
  font-family:'IBM Plex Mono',monospace!important;
  font-weight:600!important;font-size:.8rem!important;
  letter-spacing:.08em!important;text-transform:uppercase!important;
  padding:.5rem 1.2rem!important;width:100%!important;
  transition:opacity .18s,transform .15s!important;
}}
.stButton>button:hover{{opacity:.82!important;transform:translateY(-1px)!important;}}
.stDownloadButton>button{{
  background:{g}!important;color:#ffffff!important;
  -webkit-text-fill-color:#ffffff!important;
  border:none!important;border-radius:9px!important;
  font-family:'IBM Plex Mono',monospace!important;
  font-weight:700!important;font-size:.85rem!important;
  letter-spacing:.07em!important;text-transform:uppercase!important;
  padding:.62rem 1.4rem!important;width:100%!important;
  box-shadow:0 3px 14px {'rgba(0,200,255,0.28)' if dk else 'rgba(0,60,200,0.22)'}!important;
  transition:all .2s!important;
}}
.stDownloadButton>button:hover{{
  opacity:.84!important;transform:translateY(-1px)!important;
}}

/* ══ FILE UPLOADER ══════════════════════════ */
[data-testid="stFileUploader"]{{
  background:{'rgba(0,160,220,0.07)' if dk else 'rgba(0,60,180,0.05)'}!important;
  border:2px dashed {T['brd_css']}!important;
  border-radius:var(--rad)!important;transition:border .2s!important;
}}
[data-testid="stFileUploader"]:hover{{border-color:{T['acc']}!important;}}
[data-testid="stFileDropzoneInstructions"] p,
[data-testid="stFileDropzoneInstructions"] span{{
  color:{MUT}!important;-webkit-text-fill-color:{MUT}!important;
  font-size:.85rem!important;
}}

/* ══ ALERTS ══════════════════════════════════ */
.stSuccess{{background:{'rgba(0,229,160,0.09)' if dk else 'rgba(0,85,51,0.09)'}!important;
  border:1px solid {T['suc']}!important;border-radius:var(--rad)!important;}}
.stError{{background:{'rgba(255,77,109,0.09)' if dk else 'rgba(170,0,30,0.09)'}!important;
  border:1px solid {T['dng']}!important;border-radius:var(--rad)!important;}}
.stWarning{{background:{'rgba(245,166,35,0.09)' if dk else 'rgba(160,100,0,0.09)'}!important;
  border:1px solid {T['gold']}!important;border-radius:var(--rad)!important;}}
.stInfo{{background:{'rgba(0,200,255,0.07)' if dk else 'rgba(0,60,180,0.07)'}!important;
  border:1px solid {T['acc']}!important;border-radius:var(--rad)!important;}}
.stSuccess p,.stError p,.stWarning p,.stInfo p,
div[data-testid="stAlert"] p{{
  color:{TXT}!important;-webkit-text-fill-color:{TXT}!important;
}}

/* ══ IMAGES / PROGRESS ═══════════════════════ */
.stImage img,[data-testid="stImage"] img{{
  border-radius:var(--rad)!important;
  border:1px solid var(--brd)!important;box-shadow:var(--shd)!important;
}}
.stProgress>div>div>div{{background:{g}!important;border-radius:99px!important;}}
.stProgress>div>div{{background:var(--brd2)!important;border-radius:99px!important;}}

/* ══ EXPANDER  — override all inner elements ══
   We hide the native expander and use our own
   custom toggle div instead.
══════════════════════════════════════════════ */
[data-testid="stExpander"]{{
  background:{EBG}!important;
  border:1px solid {T['brd_css']}!important;
  border-radius:var(--rad)!important;
  overflow:hidden!important;
}}
[data-testid="stExpander"] details,
[data-testid="stExpander"] summary{{
  background:{EBG}!important;
  color:{ETXT}!important;
  -webkit-text-fill-color:{ETXT}!important;
  list-style:none!important;
}}
[data-testid="stExpander"] summary *,
[data-testid="stExpander"] summary p{{
  color:{ETXT}!important;
  -webkit-text-fill-color:{ETXT}!important;
  font-family:'IBM Plex Mono',monospace!important;
  font-size:.86rem!important;font-weight:600!important;
}}
[data-testid="stExpanderDetails"]{{
  background:{EBG}!important;
  padding:.8rem 1rem!important;
}}
/* Everything inside expander details */
[data-testid="stExpanderDetails"] *{{
  color:{ETXT}!important;
  -webkit-text-fill-color:{ETXT}!important;
}}
/* But download buttons inside get white text */
[data-testid="stExpanderDetails"] .stDownloadButton>button,
[data-testid="stExpanderDetails"] .stButton>button{{
  color:#ffffff!important;
  -webkit-text-fill-color:#ffffff!important;
}}

/* ══ THEME TOGGLE BUTTON ══════════════════════ */
.theme-toggle-btn{{
  display:inline-flex;align-items:center;gap:.4rem;
  background:{T['card_css']};border:1.5px solid {T['brd_css']};
  border-radius:99px;padding:.35rem .9rem;cursor:pointer;
  font-family:'IBM Plex Mono',monospace;font-size:.72rem;
  font-weight:600;letter-spacing:.08em;text-transform:uppercase;
  color:{TXT};transition:border-color .2s,background .2s;
  text-decoration:none;
}}
.theme-toggle-btn:hover{{border-color:{T['acc']};}}

/* ══ CUSTOM COMPONENTS ═══════════════════════ */

/* Hero */
.ar-hero{{
  position:relative;overflow:hidden;
  border:1px solid var(--brd);border-radius:20px;
  padding:2.5rem 2.4rem 2.2rem;margin-bottom:1.4rem;
  background:{'linear-gradient(150deg,rgba(3,10,25,0.97),rgba(5,16,36,0.97))' if dk
              else 'linear-gradient(150deg,rgba(216,230,255,0.97),rgba(228,240,255,0.97))'};
  backdrop-filter:var(--blur);box-shadow:var(--shd);
}}
.ar-hero::after{{
  content:'';position:absolute;top:-65px;right:-65px;
  width:280px;height:280px;border-radius:50%;pointer-events:none;
  background:radial-gradient(circle,
    {'rgba(0,200,255,0.09)' if dk else 'rgba(0,70,200,0.07)'} 0%,transparent 68%);
}}
.ar-eyebrow{{
  font-family:'IBM Plex Mono',monospace;font-size:.66rem;font-weight:600;
  letter-spacing:.2em;text-transform:uppercase;color:{T['acc']}!important;
  -webkit-text-fill-color:{T['acc']}!important;
  display:flex;align-items:center;gap:.6rem;margin-bottom:.6rem;
}}
.ar-dot{{width:5px;height:5px;background:{T['gold']};border-radius:50%;display:inline-block;}}
.ar-title{{
  font-family:'Bebas Neue',cursive;font-size:clamp(2.5rem,5vw,3.7rem);
  letter-spacing:.06em;line-height:1;color:{TXT}!important;
  -webkit-text-fill-color:{TXT}!important;margin:0 0 .42rem 0;
}}
.ar-title em{{
  font-style:normal;color:{T['acc']}!important;
  -webkit-text-fill-color:{T['acc']}!important;
}}
.ar-tagline{{
  font-size:.9rem;color:{MUT}!important;-webkit-text-fill-color:{MUT}!important;
  max-width:500px;line-height:1.62;margin:0 0 1.1rem;
}}
.ar-chips{{display:flex;flex-wrap:wrap;gap:.36rem;}}
.ar-chip{{
  font-family:'IBM Plex Mono',monospace;font-size:.62rem;font-weight:500;
  letter-spacing:.07em;text-transform:uppercase;padding:.22rem .66rem;
  border-radius:99px;border:1px solid var(--brd);
  background:{'rgba(0,200,255,0.08)' if dk else 'rgba(0,60,200,0.07)'};
  color:{T['acc']}!important;-webkit-text-fill-color:{T['acc']}!important;
}}
.ar-compass{{position:absolute;right:2rem;top:50%;transform:translateY(-50%);
  opacity:{'0.14' if dk else '0.10'};}}

/* Section label */
.ar-lbl{{
  font-family:'IBM Plex Mono',monospace;font-size:.65rem;font-weight:600;
  letter-spacing:.15em;text-transform:uppercase;
  color:{T['acc']}!important;-webkit-text-fill-color:{T['acc']}!important;
  display:flex;align-items:center;gap:.5rem;margin-bottom:.85rem;
}}
.ar-lbl::before{{content:'';display:inline-block;width:15px;height:2px;
  background:{T['acc']};border-radius:2px;}}
.ar-lbl::after{{content:'';flex:1;height:1px;background:var(--brd2);margin-left:.3rem;}}
.ar-hr{{border:none;height:1px;background:var(--brd2);margin:1.5rem 0;}}

/* Cards */
.ar-card{{
  background:var(--card);border:1px solid var(--brd);border-radius:var(--rad);
  padding:1.4rem 1.6rem;margin-bottom:1rem;
  backdrop-filter:var(--blur);box-shadow:var(--shd);
}}
.ar-info-card{{
  background:{'rgba(0,200,255,0.06)' if dk else 'rgba(0,60,200,0.05)'};
  border:1px solid {T['acc']}44;border-radius:var(--rad);
  padding:1.2rem 1.5rem;margin-bottom:1rem;
  backdrop-filter:var(--blur);
}}

/* Steps */
.ar-step{{display:flex;align-items:flex-start;gap:.9rem;padding:.66rem 0;
  border-bottom:1px solid var(--brd2);}}
.ar-step:last-child{{border-bottom:none;}}
.ar-step-num{{
  min-width:26px;height:26px;background:{g};border-radius:7px;
  font-family:'Bebas Neue',cursive;font-size:.88rem;
  color:#ffffff!important;-webkit-text-fill-color:#ffffff!important;
  display:flex;align-items:center;justify-content:center;flex-shrink:0;
}}
.ar-step-title{{
  font-family:'IBM Plex Mono',monospace;font-size:.7rem;font-weight:600;
  letter-spacing:.06em;text-transform:uppercase;
  color:{T['acc']}!important;-webkit-text-fill-color:{T['acc']}!important;
  margin-bottom:.08rem;
}}
.ar-step-desc{{font-size:.82rem;color:{MUT}!important;
  -webkit-text-fill-color:{MUT}!important;line-height:1.43;}}

/* Diagram type cards */
.ar-type-grid{{display:grid;grid-template-columns:1fr 1fr;gap:.7rem;margin:.4rem 0;}}
.ar-type-card{{
  background:{'rgba(0,180,255,0.04)' if dk else 'rgba(0,60,200,0.04)'};
  border:1px solid var(--brd);border-radius:10px;padding:.85rem .95rem;
}}
.ar-type-code{{
  font-family:'Bebas Neue',cursive;font-size:1.42rem;letter-spacing:.06em;
  color:{T['acc']}!important;-webkit-text-fill-color:{T['acc']}!important;line-height:1;
}}
.ar-type-name{{
  font-family:'IBM Plex Mono',monospace;font-size:.62rem;letter-spacing:.07em;
  text-transform:uppercase;color:{MUT}!important;
  -webkit-text-fill-color:{MUT}!important;margin-top:.12rem;
}}
.ar-type-desc{{
  font-size:.76rem;color:{MUT}!important;
  -webkit-text-fill-color:{MUT}!important;margin-top:.26rem;line-height:1.38;
}}
.ar-type-badge{{
  display:inline-block;font-family:'IBM Plex Mono',monospace;
  font-size:.6rem;font-weight:700;letter-spacing:.06em;text-transform:uppercase;
  padding:.15rem .55rem;border-radius:4px;margin-top:.32rem;
}}
.ar-badge-t1{{
  background:{'rgba(0,200,255,0.14)' if dk else 'rgba(0,60,200,0.11)'};
  color:{T['acc']}!important;-webkit-text-fill-color:{T['acc']}!important;
}}
.ar-badge-t2{{
  background:{'rgba(0,229,160,0.14)' if dk else 'rgba(0,85,51,0.11)'};
  color:{T['suc']}!important;-webkit-text-fill-color:{T['suc']}!important;
}}

/* Stats */
.ar-stats{{display:grid;grid-template-columns:repeat(6,1fr);gap:.56rem;margin:.9rem 0;}}
.ar-stat{{
  background:var(--card);border:1px solid var(--brd);border-radius:10px;
  padding:.76rem .48rem;text-align:center;backdrop-filter:var(--blur);
}}
.ar-sv{{
  font-family:'Bebas Neue',cursive;font-size:1.32rem;letter-spacing:.04em;
  color:{T['acc']}!important;-webkit-text-fill-color:{T['acc']}!important;line-height:1;
}}
.ar-sl{{
  font-family:'IBM Plex Mono',monospace;font-size:.54rem;letter-spacing:.09em;
  text-transform:uppercase;color:{MUT}!important;
  -webkit-text-fill-color:{MUT}!important;margin-top:.22rem;
}}

/* Coverage bar */
.ar-cov{{
  background:var(--card);border:1px solid var(--brd);border-radius:10px;
  padding:.68rem 1rem;font-family:'IBM Plex Mono',monospace;font-size:.74rem;
  color:{TXT}!important;display:flex;flex-wrap:wrap;gap:.8rem;
  align-items:center;margin:.7rem 0;
}}
.ar-cov b{{color:{T['acc']}!important;-webkit-text-fill-color:{T['acc']}!important;}}
.ar-pass{{
  display:inline-block;
  background:{'rgba(0,229,160,0.12)' if dk else 'rgba(0,85,51,0.12)'};
  border:1px solid {T['suc']};
  color:{T['suc']}!important;-webkit-text-fill-color:{T['suc']}!important;
  font-family:'IBM Plex Mono',monospace;font-size:.65rem;font-weight:700;
  letter-spacing:.1em;text-transform:uppercase;padding:1px 9px;border-radius:99px;
}}
.ar-fail{{
  display:inline-block;
  background:{'rgba(255,77,109,0.12)' if dk else 'rgba(170,0,30,0.12)'};
  border:1px solid {T['dng']};
  color:{T['dng']}!important;-webkit-text-fill-color:{T['dng']}!important;
  font-family:'IBM Plex Mono',monospace;font-size:.65rem;font-weight:700;
  letter-spacing:.1em;text-transform:uppercase;padding:1px 9px;border-radius:99px;
}}

/* Diagram label + white wrapper */
.ar-dlbl{{
  font-family:'IBM Plex Mono',monospace;font-size:.66rem;font-weight:600;
  letter-spacing:.1em;text-transform:uppercase;
  color:{T['acc']}!important;-webkit-text-fill-color:{T['acc']}!important;
  text-align:center;padding:.4rem .2rem .2rem;
}}
.ar-diag-white{{
  background:#ffffff;border-radius:12px;padding:6px;
  box-shadow:0 4px 24px {'rgba(0,0,0,0.45)' if dk else 'rgba(0,40,160,0.14)'};
  border:1px solid {'rgba(0,200,255,0.18)' if dk else 'rgba(0,60,200,0.15)'};
}}

/* ══ FREQUENCY TABLE ══════════════════════════
   Rendered as plain HTML div — no expander CSS fight
══════════════════════════════════════════════ */
.ar-freq-box{{
  background:{EBG};border:1px solid {T['brd_css']};
  border-radius:var(--rad);padding:1rem 1.2rem;margin:.5rem 0;
}}
.ar-freq-hdr{{
  font-family:'Bebas Neue',cursive;font-size:1.15rem;letter-spacing:.12em;
  color:{TXT}!important;-webkit-text-fill-color:{TXT}!important;
  margin-bottom:.1rem;
}}
.ar-freq-note{{
  font-family:'IBM Plex Mono',monospace;font-size:.66rem;
  color:{MUT}!important;-webkit-text-fill-color:{MUT}!important;
  margin-bottom:.7rem;line-height:1.6;
}}
.ar-tbl{{
  width:100%;border-collapse:collapse;
  font-family:'IBM Plex Mono',monospace;font-size:.72rem;
  border-radius:8px;overflow:hidden;
}}
.ar-tbl th{{
  background:{'#0a2244' if dk else '#1a3a6a'}!important;
  color:#ffffff!important;-webkit-text-fill-color:#ffffff!important;
  padding:7px 10px;letter-spacing:.05em;font-size:.7rem;
  text-align:center;font-weight:700;
  border-bottom:2px solid {'rgba(0,200,255,0.3)' if dk else 'rgba(0,80,200,0.3)'};
}}
.ar-tbl th.dh{{
  background:{'#061828' if dk else '#0d2244'}!important;
  color:#ffffff!important;-webkit-text-fill-color:#ffffff!important;
  text-align:left;
}}
.ar-tbl td{{
  padding:5px 10px;
  border-bottom:1px solid {T['brd2_css']};
  text-align:center;
  color:{ETXT}!important;-webkit-text-fill-color:{ETXT}!important;
  font-size:.73rem;
}}
.ar-tbl td.dc{{
  font-weight:700;text-align:left;
  color:{T['acc']}!important;-webkit-text-fill-color:{T['acc']}!important;
  background:{'rgba(0,200,255,0.04)' if dk else 'rgba(0,60,200,0.04)'}!important;
}}
.ar-tbl tr:nth-child(even) td{{
  background:{'rgba(255,255,255,0.02)' if dk else 'rgba(0,60,200,0.025)'}!important;
}}
.ar-tbl tr.trow td{{
  font-weight:700;
  border-top:2px solid {'rgba(0,200,255,0.3)' if dk else 'rgba(0,60,200,0.3)'};
  background:{'rgba(0,200,255,0.08)' if dk else 'rgba(0,60,200,0.08)'}!important;
  border-bottom:none;
  color:{T['acc']}!important;-webkit-text-fill-color:{T['acc']}!important;
}}

/* Toggle freq table button */
.ar-tog-btn{{
  display:inline-flex;align-items:center;gap:.4rem;cursor:pointer;
  font-family:'IBM Plex Mono',monospace;font-size:.72rem;font-weight:600;
  letter-spacing:.1em;text-transform:uppercase;padding:.42rem 1rem;
  border-radius:8px;border:1.5px solid {T['brd_css']};
  background:var(--card);color:{TXT}!important;
  -webkit-text-fill-color:{TXT}!important;
  transition:border-color .2s,background .2s;
}}
.ar-tog-btn:hover{{
  border-color:{T['acc']};
  background:{'rgba(0,200,255,0.08)' if dk else 'rgba(0,60,200,0.07)'};
}}

/* Sub-label / hint */
.ar-sub{{
  font-family:'IBM Plex Mono',monospace;font-size:.67rem;letter-spacing:.12em;
  text-transform:uppercase;color:{MUT}!important;
  -webkit-text-fill-color:{MUT}!important;margin-bottom:.36rem;
}}
.ar-hint{{
  font-family:'IBM Plex Mono',monospace;font-size:.65rem;
  color:{MUT}!important;-webkit-text-fill-color:{MUT}!important;
  line-height:1.7;margin-top:.28rem;
}}
.ar-file-banner{{
  background:{'rgba(0,229,160,0.09)' if dk else 'rgba(0,100,60,0.08)'};
  border:1px solid {T['suc']};border-radius:10px;
  padding:.68rem 1rem;font-family:'IBM Plex Mono',monospace;font-size:.77rem;
  color:{TXT}!important;display:flex;align-items:center;gap:.8rem;margin-bottom:.6rem;
}}
.ar-file-banner b{{
  color:{T['suc']}!important;-webkit-text-fill-color:{T['suc']}!important;
}}

/* Generate button */
.gen-wrap .stButton>button{{
  background:{g}!important;color:#ffffff!important;
  -webkit-text-fill-color:#ffffff!important;
  font-family:'Bebas Neue',cursive!important;font-size:1.3rem!important;
  letter-spacing:.18em!important;text-transform:uppercase!important;
  padding:1rem 2rem!important;border-radius:12px!important;border:none!important;
  box-shadow:0 4px 28px {'rgba(0,200,255,0.30)' if dk else 'rgba(0,60,200,0.24)'}!important;
  animation:pulse 2.8s ease-in-out infinite;
}}
.gen-wrap .stButton>button:hover{{
  transform:translateY(-3px)!important;
  box-shadow:0 8px 36px {'rgba(0,200,255,0.50)' if dk else 'rgba(0,60,200,0.40)'}!important;
}}
@keyframes pulse{{
  0%,100%{{box-shadow:0 4px 28px {'rgba(0,200,255,0.26)' if dk else 'rgba(0,60,200,0.20)'};}}
  50%{{box-shadow:0 4px 40px {'rgba(0,200,255,0.52)' if dk else 'rgba(0,60,200,0.42)'};}}
}}

/* ══ LOADING SPINNER ══════════════════════════ */
.ar-loading{{
  display:flex;flex-direction:column;align-items:center;
  justify-content:center;padding:2.5rem 1rem;gap:1rem;
}}
.ar-spinner{{
  width:72px;height:72px;position:relative;
}}
.ar-spinner-ring{{
  position:absolute;inset:0;border-radius:50%;
  border:3px solid transparent;
  border-top-color:{T['acc']};
  animation:spin 1.1s linear infinite;
}}
.ar-spinner-ring:nth-child(2){{
  inset:10px;border-top-color:{T['gold']};
  animation:spin 0.8s linear infinite reverse;
}}
.ar-spinner-ring:nth-child(3){{
  inset:20px;border-top-color:{T['suc']};
  animation:spin 1.4s linear infinite;
}}
@keyframes spin{{0%{{transform:rotate(0deg);}}100%{{transform:rotate(360deg);}}}}
.ar-load-txt{{
  font-family:'IBM Plex Mono',monospace;font-size:.78rem;letter-spacing:.12em;
  text-transform:uppercase;color:{T['acc']}!important;
  -webkit-text-fill-color:{T['acc']}!important;
}}

/* Runway progress */
.rwy-wrap{{margin:1.3rem 0;}}
.rwy-hdr{{
  font-family:'IBM Plex Mono',monospace;font-size:.7rem;letter-spacing:.1em;
  text-transform:uppercase;color:{MUT}!important;-webkit-text-fill-color:{MUT}!important;
  display:flex;justify-content:space-between;margin-bottom:.42rem;
}}
.rwy-pct{{
  font-family:'Bebas Neue',cursive;font-size:.95rem;
  color:{T['acc']}!important;-webkit-text-fill-color:{T['acc']}!important;letter-spacing:.1em;
}}
.rwy-outer{{
  position:relative;width:100%;height:56px;
  background:{'#050c18' if dk else '#b0c4e4'};
  border:1px solid var(--brd);border-radius:6px;overflow:hidden;
}}
.rwy-outer::before{{content:'';position:absolute;inset:0;
  background:repeating-linear-gradient(90deg,transparent 0,transparent 27px,
    {'rgba(255,255,255,0.012)' if dk else 'rgba(0,0,0,0.04)'} 27px,
    {'rgba(255,255,255,0.012)' if dk else 'rgba(0,0,0,0.04)'} 28px);}}
.rwy-et,.rwy-eb{{position:absolute;left:0;right:0;height:4px;
  background:repeating-linear-gradient(90deg,{T['gold']} 0,{T['gold']} 15px,transparent 15px,transparent 27px);
  opacity:.58;}}
.rwy-et{{top:0;}}.rwy-eb{{bottom:0;}}
.rwy-cl{{position:absolute;top:50%;transform:translateY(-50%);left:0;right:0;height:2px;
  background:repeating-linear-gradient(90deg,
    {'rgba(255,255,255,0.35)' if dk else 'rgba(255,255,255,0.70)'} 0,
    {'rgba(255,255,255,0.35)' if dk else 'rgba(255,255,255,0.70)'} 22px,
    transparent 22px,transparent 37px);}}
.rwy-fill{{position:absolute;top:0;left:0;bottom:0;
  transition:width .3s cubic-bezier(.4,0,.2,1);
  background:{'linear-gradient(90deg,rgba(0,200,255,0.07),rgba(0,200,255,0.20))' if dk
              else 'linear-gradient(90deg,rgba(0,60,200,0.08),rgba(0,60,200,0.22))'};
  border-right:2px solid {'rgba(0,200,255,0.60)' if dk else 'rgba(0,60,200,0.55)'};}}
.rwy-plane{{position:absolute;top:50%;transform:translateY(-50%);font-size:22px;
  filter:{'drop-shadow(0 0 6px rgba(0,220,255,0.85))' if dk else 'drop-shadow(0 0 4px rgba(0,60,200,0.55))'};
  transition:left .3s cubic-bezier(.4,0,.2,1);user-select:none;}}
.rwy-marks{{display:flex;justify-content:space-between;margin-top:.26rem;padding:0 2px;}}
.rwy-mark{{
  font-family:'IBM Plex Mono',monospace;font-size:.55rem;
  color:{MUT}!important;-webkit-text-fill-color:{MUT}!important;letter-spacing:.05em;
}}

/* Footer */
.ar-footer{{
  background:var(--card);border:1px solid var(--brd);border-radius:14px;
  padding:2rem 1.5rem;text-align:center;margin-top:2.5rem;backdrop-filter:var(--blur);
}}
.ar-footer-brand{{
  font-family:'Bebas Neue',cursive;font-size:1rem;letter-spacing:.22em;
  text-transform:uppercase;color:{T['acc']}!important;
  -webkit-text-fill-color:{T['acc']}!important;
}}
.ar-footer-line{{
  font-family:'IBM Plex Mono',monospace;font-size:.75rem;
  color:{MUT}!important;-webkit-text-fill-color:{MUT}!important;margin:.17rem 0;
}}
.ar-footer-link{{color:{T['acc']}!important;-webkit-text-fill-color:{T['acc']}!important;text-decoration:none;}}

</style>""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════
#  COMPASS SVG
# ══════════════════════════════════════════════════════════════════════
def compass_svg(accent, size=145):
    c=size/2; r=c*0.9; ticks=""
    for i in range(16):
        a=np.radians(i*22.5); r1=r*0.76; r2=r*0.90
        x1=c+r1*np.sin(a); y1=c-r1*np.cos(a)
        x2=c+r2*np.sin(a); y2=c-r2*np.cos(a)
        w="1.2" if i%4==0 else "0.55"
        ticks+=(f'<line x1="{x1:.1f}" y1="{y1:.1f}" x2="{x2:.1f}" y2="{y2:.1f}"'
                f' stroke="{accent}" stroke-width="{w}" opacity="0.65"/>')
    return (f'<svg width="{size}" height="{size}" viewBox="0 0 {size} {size}"'
            f' xmlns="http://www.w3.org/2000/svg" class="ar-compass">'
            f'<circle cx="{c}" cy="{c}" r="{r*0.93:.1f}" fill="none"'
            f' stroke="{accent}" stroke-width="0.8" opacity="0.38"/>'
            f'<circle cx="{c}" cy="{c}" r="{r*0.76:.1f}" fill="none"'
            f' stroke="{accent}" stroke-width="0.5" opacity="0.22"/>'
            f'{ticks}'
            f'<polygon points="{c},{c-r*0.62:.1f} {c-5.5},{c:.1f}'
            f' {c},{c+r*0.17:.1f} {c+5.5},{c:.1f}" fill="{accent}" opacity="0.95"/>'
            f'<polygon points="{c},{c+r*0.62:.1f} {c-5.5},{c:.1f}'
            f' {c},{c-r*0.17:.1f} {c+5.5},{c:.1f}" fill="{accent}" opacity="0.38"/>'
            f'<circle cx="{c}" cy="{c}" r="3.8" fill="{accent}" opacity="0.85"/>'
            f'<circle cx="{c}" cy="{c}" r="1.8" fill="white" opacity="0.55"/></svg>')

# ══════════════════════════════════════════════════════════════════════
#  RUNWAY PROGRESS BAR
# ══════════════════════════════════════════════════════════════════════
def rwy_progress(pct, msg=""):
    pct=max(0.,min(100.,pct)); left=2+pct*0.91; tilt=min(pct*0.10,9)
    return (f'<div class="rwy-wrap">'
            f'<div class="rwy-hdr"><span>&#9658; {msg}</span>'
            f'<span class="rwy-pct">{pct:.0f}%</span></div>'
            f'<div class="rwy-outer"><div class="rwy-et"></div><div class="rwy-eb"></div>'
            f'<div class="rwy-cl"></div><div class="rwy-fill" style="width:{pct}%;"></div>'
            f'<div class="rwy-plane" style="left:{left:.1f}%;'
            f'transform:translateY(-50%) rotate(-{tilt:.1f}deg);">&#9992;</div></div>'
            f'<div class="rwy-marks">'
            f'<span class="rwy-mark">THR</span><span class="rwy-mark">25%</span>'
            f'<span class="rwy-mark">MID</span><span class="rwy-mark">75%</span>'
            f'<span class="rwy-mark">TKOF</span></div></div>')

# ══════════════════════════════════════════════════════════════════════
#  DATA PROCESSING
# ══════════════════════════════════════════════════════════════════════
def load_file(f):
    name=f.name.lower(); raw=f.read(); f.seek(0)
    if name.endswith((".xlsx",".xls")):
        try: return pd.read_excel(io.BytesIO(raw),engine="openpyxl"), None
        except Exception as e: return None, str(e)
    for enc in ("utf-8","utf-8-sig","latin-1","cp1252"):
        try: return pd.read_csv(io.BytesIO(raw),encoding=enc,low_memory=False), None
        except Exception: pass
    return None, "Cannot decode file."

@st.cache_data(show_spinner=False)
def process_data(fb, fname, dc, sc, dfmt, su):
    name=fname.lower()
    if name.endswith((".xlsx",".xls")): df=pd.read_excel(io.BytesIO(fb),engine="openpyxl")
    else:
        df=None
        for enc in ("utf-8","utf-8-sig","latin-1","cp1252"):
            try: df=pd.read_csv(io.BytesIO(fb),encoding=enc,low_memory=False); break
            except Exception: pass
        if df is None: raise ValueError("Cannot decode file.")
    df.columns=df.columns.str.strip()
    for col in (dc,sc):
        if col not in df.columns:
            raise ValueError(f"Column '{col}' not found. Available: {list(df.columns)}")
    w=df[[dc,sc]].copy(); w.columns=["dir","spd"]
    if dfmt=="Compass (N, NNE …)":
        w["dg"]=w["dir"].astype(str).str.strip().str.upper().map(C2D)
    else:
        w["dg"]=pd.to_numeric(w["dir"],errors="coerce")%360
    w["kmh"]=pd.to_numeric(w["spd"],errors="coerce")
    if su=="knots": w["kmh"]*=1.852
    elif su=="m/s":  w["kmh"]*=3.6
    w=w.dropna(subset=["dg","kmh"])
    if len(w)==0: raise ValueError("No valid rows after processing.")
    w["sec"]=(((w["dg"]+11.25)%360)//22.5).astype(int).clip(0,15)
    w["sc"]=pd.cut(w["kmh"],bins=SPD_BINS,labels=list(range(5))).astype(float).astype("Int64")
    total=len(w); freq=np.zeros((16,5))
    for s in range(16):
        for c in range(5):
            freq[s,c]=((w.sec==s)&(w.sc==c)).sum()/total*100
    op=freq[:,1]+freq[:,2]+freq[:,3]
    return freq,{"total":total,"calm":round(freq[:,0].sum(),1),
                 "op":round(op.sum(),1),"avg":round(w.kmh.mean(),1),
                 "max":round(w.kmh.max(),1),"dom":DIRS_16[int(op.argmax())]}

# ══════════════════════════════════════════════════════════════════════
#  RUNWAY ANALYSIS
# ══════════════════════════════════════════════════════════════════════
def ha(cx): return np.degrees(np.arcsin(min(cx/24.1,1.)))
def rwy_cov(freq,hdg,cx):
    h=ha(cx); t=0.
    for i in range(16):
        d=abs(((i*22.5-hdg+180)%360)-180)
        if d<=h or d>=180-h: t+=freq[i,1:4].sum()
    return min(t,100.)
def best_rwy(freq,cx,excl=None):
    bh,bc=0.,0.
    for hdg in np.arange(0,180,5):
        if excl is not None and abs(((hdg-excl+90)%180)-90)<20: continue
        c=rwy_cov(freq,hdg,cx)
        if c>bc: bc,bh=c,hdg
    return float(bh)
def comb_cov(freq,r1,r2,cx):
    h=ha(cx); t=0.
    for i in range(16):
        a=i*22.5; d1=abs(((a-r1+180)%360)-180); d2=abs(((a-r2+180)%360)-180)
        if d1<=h or d1>=180-h or d2<=h or d2>=180-h: t+=freq[i,1:4].sum()
    return min(t,100.)
def rwy_lbl(hdg):
    e1=int(round(hdg/10))%36 or 36; e2=int(round((hdg+180)/10))%36 or 36
    return f"Runway {e1:02d}/{e2:02d}"

# ══════════════════════════════════════════════════════════════════════
#  FREQUENCY TABLE HTML  — standalone div, no expander
# ══════════════════════════════════════════════════════════════════════
def freq_table_html(freq, T):
    note = (f'Calm (&lt;6 km/h) = {freq[:,0].sum():.1f}%  ·  '
            f'Strong (&gt;60 km/h) = {freq[:,4].sum():.1f}%  ·  '
            f'Operational (6–60 km/h) = {sum(freq[:,j].sum() for j in TBL_IDX):.1f}%')
    hdr = (f'<tr><th class="dh" rowspan="2">Direction</th>'
           f'<th colspan="3">Duration of Wind (%)</th>'
           f'<th rowspan="2">Total % of time wind blew<br>between 6.0 to 60 km/h</th></tr>'
           f'<tr>'
           f'<th>{TBL_COLS[0]}</th>'
           f'<th>{TBL_COLS[1]}</th>'
           f'<th>{TBL_COLS[2]}</th>'
           f'</tr>')
    rows=""
    for i,d in enumerate(DIRS_16):
        c1=freq[i,1]; c2=freq[i,2]; c3=freq[i,3]; tot=c1+c2+c3
        rows+=(f'<tr><td class="dc">{d}</td>'
               f'<td>{c1:.1f}</td><td>{c2:.1f}</td><td>{c3:.1f}</td>'
               f'<td>{tot:.1f}</td></tr>')
    t1=freq[:,1].sum(); t2=freq[:,2].sum(); t3=freq[:,3].sum(); tt=t1+t2+t3
    rows+=(f'<tr class="trow"><td class="dc">TOTAL</td>'
           f'<td>{t1:.1f}</td><td>{t2:.1f}</td><td>{t3:.1f}</td>'
           f'<td>{tt:.1f}</td></tr>')
    return (f'<div class="ar-freq-box">'
            f'<div class="ar-freq-hdr">Frequency Table</div>'
            f'<div class="ar-freq-note">{note}</div>'
            f'<table class="ar-tbl">{hdr}{rows}</table>'
            f'</div>')

def freq_to_csv(freq):
    rows=[]
    for i,d in enumerate(DIRS_16):
        r={"Direction":d}
        for j,lbl in enumerate(TBL_COLS): r[lbl]=round(freq[i,TBL_IDX[j]],4)
        r["Total % (6.0-60 km/h)"]=round(sum(freq[i,j] for j in TBL_IDX),4)
        rows.append(r)
    t={"Direction":"TOTAL"}
    for j,lbl in enumerate(TBL_COLS): t[lbl]=round(freq[:,TBL_IDX[j]].sum(),4)
    t["Total % (6.0-60 km/h)"]=round(sum(freq[:,j].sum() for j in TBL_IDX),4)
    rows.append(t)
    return pd.DataFrame(rows).to_csv(index=False).encode("utf-8")

# ══════════════════════════════════════════════════════════════════════
#  DIAGRAM RENDERERS  — white bg, HEX only, NO runway lines
# ══════════════════════════════════════════════════════════════════════
def _polar(title, theme):
    T=TH[theme]
    fig,ax=plt.subplots(figsize=(7.5,7.5),subplot_kw=dict(polar=True),facecolor="#ffffff")
    ax.set_facecolor("#f8faff"); ax.set_theta_zero_location("N"); ax.set_theta_direction(-1)
    ax.grid(color="#dde4f0",linestyle="--",lw=0.55,alpha=0.8)
    ax.spines["polar"].set_color("#dde4f0")
    ax.set_xticks(np.linspace(0,2*np.pi,16,endpoint=False))
    ax.set_xticklabels(DIRS_16,fontsize=9,fontweight="bold",color="#0d2244",fontfamily="monospace")
    ax.tick_params(axis="y",labelsize=7.5,labelcolor="#3d5a80")
    ax.set_title(title,fontsize=10.5,fontweight="bold",pad=26,color=T["m_title"],wrap=True)
    return fig,ax

def _leg(ax, handles):
    ax.legend(handles=handles,loc="lower left",bbox_to_anchor=(-0.22,-0.28),
              fontsize=8,framealpha=0.97,facecolor="#ffffff",
              edgecolor="#ccd4ee",labelcolor="#0d2244")

def _png(fig):
    buf=io.BytesIO()
    fig.savefig(buf,format="png",dpi=160,bbox_inches="tight",facecolor="#ffffff")
    plt.close(fig); buf.seek(0); return buf.getvalue()

def _refcircles(ax, mv):
    for frac in [.25,.5,.75,1.]:
        rv=mv*frac
        ax.plot(np.linspace(0,2*np.pi,200),[rv]*200,color="#ccd4ee",lw=0.6,alpha=0.7,zorder=1)
        ax.text(np.radians(10),rv,f"{rv:.1f}%",fontsize=6,color="#6688aa",ha="left",va="bottom")

# Type I — polygon, uses 6-60 km/h total per direction
def render_t1s(freq, theme):
    T=TH[theme]; dp=freq[:,1]+freq[:,2]+freq[:,3]; N=16
    th=np.linspace(0,2*np.pi,N,endpoint=False); mv=max(dp.max(),1.)
    fig,ax=_polar("TYPE I  —  SINGLE RUNWAY\n",theme)
    _refcircles(ax,mv)
    ax.fill(th,dp,color=T["m_pfill"],alpha=0.35,zorder=2)
    ax.plot(np.append(th,th[0]),np.append(dp,dp[0]),color=T["m_poly"],lw=2.4,alpha=0.95,zorder=3)
    sz=[100 if dp[i]>=mv*.90 else 60 if dp[i]>=mv*.60 else 30 for i in range(N)]
    ax.scatter(th,dp,s=sz,color=T["m_poly"],zorder=4,edgecolors="#ffffff",linewidths=0.9)
    dom=int(np.argmax(dp))
    ax.text(th[dom],dp[dom]*1.16,f"{dp[dom]:.1f}%",fontsize=8.5,color=T["m_poly"],
            ha="center",va="bottom",fontweight="bold")
    ax.set_ylim(0,mv*1.30)
    _leg(ax,[
        mpatches.Patch(color=T["m_pfill"],alpha=0.5,label="Wind Rose Diagram"),
        plt.Line2D([0],[0],color=T["m_poly"],lw=2.4,label="Polygon outline"),
        plt.Line2D([0],[0],marker='o',color=T["m_poly"],lw=0,markersize=5.5,label="Direction value"),
    ])
    plt.tight_layout(rect=[0,.07,1,.97]); return _png(fig)

def render_t1m(freq, theme):
    T=TH[theme]; dp=freq[:,1]+freq[:,2]+freq[:,3]; N=16
    th=np.linspace(0,2*np.pi,N,endpoint=False); mv=max(dp.max(),1.)
    fig,ax=_polar("TYPE I  —  MULTI RUNWAY\n",theme)
    _refcircles(ax,mv)
    ax.fill(th,dp,color=T["m_pfill"],alpha=0.32,zorder=2)
    ax.plot(np.append(th,th[0]),np.append(dp,dp[0]),color=T["m_poly"],lw=2.4,alpha=0.92,zorder=3)
    ax.scatter(th,dp,s=32,color=T["m_poly"],zorder=4,edgecolors="#ffffff",linewidths=0.8)
    dom=int(np.argmax(dp))
    ax.text(th[dom],dp[dom]*1.16,f"{dp[dom]:.1f}%",fontsize=8,color=T["m_poly"],
            ha="center",va="bottom",fontweight="bold")
    ax.set_ylim(0,mv*1.32)
    _leg(ax,[
        mpatches.Patch(color=T["m_pfill"],alpha=0.5,label="Wind Rose Diagram"),
        plt.Line2D([0],[0],color=T["m_poly"],lw=2.4,label="Polygon outline"),
        plt.Line2D([0],[0],marker='o',color=T["m_poly"],lw=0,markersize=5.5,label="Direction value"),
    ])
    plt.tight_layout(rect=[0,.07,1,.97]); return _png(fig)

# Type II — stacked MULTI-COLOR speed class bars
def render_t2s(freq, theme):
    T=TH[theme]; N=16
    th=np.linspace(0,2*np.pi,N,endpoint=False); w=2*np.pi/N*.80
    fig,ax=_polar("TYPE II  —  SINGLE RUNWAY\n",theme)
    bot=np.zeros(N)
    for s in range(5):
        ax.bar(th,freq[:,s],width=w,bottom=bot,
               color=T2_COLORS[s],edgecolor="#ffffff",lw=0.45,alpha=0.93,
               label=T2_NAMES[s],zorder=3)
        bot+=freq[:,s]
    dom=int(bot.argmax())
    ax.text(th[dom],bot[dom]*1.10,f"{bot[dom]:.1f}%",fontsize=8,
            color="#0d2244",ha="center",va="bottom",fontweight="bold")
    ax.set_ylim(0,bot.max()*1.22 or 1)
    _leg(ax,[mpatches.Patch(color=T2_COLORS[i],label=T2_NAMES[i]) for i in range(5)])
    plt.tight_layout(rect=[0,.09,1,.97]); return _png(fig)

def render_t2m(freq, theme):
    T=TH[theme]; N=16
    th=np.linspace(0,2*np.pi,N,endpoint=False); w=2*np.pi/N*.80
    fig,ax=_polar("TYPE II  —  MULTI RUNWAY\n" ,theme)
    bot=np.zeros(N)
    for s in range(5):
        ax.bar(th,freq[:,s],width=w,bottom=bot,
               color=T2_COLORS[s],edgecolor="#ffffff",lw=0.45,alpha=0.93,
               label=T2_NAMES[s],zorder=3)
        bot+=freq[:,s]
    dom=int(bot.argmax())
    ax.text(th[dom],bot[dom]*1.10,f"{bot[dom]:.1f}%",fontsize=8,
            color="#0d2244",ha="center",va="bottom",fontweight="bold")
    ax.set_ylim(0,bot.max()*1.22 or 1)
    _leg(ax,[mpatches.Patch(color=T2_COLORS[i],label=T2_NAMES[i]) for i in range(5)])
    plt.tight_layout(rect=[0,.10,1,.97]); return _png(fig)

# ══════════════════════════════════════════════════════════════════════
#  PDF BUILDER
# ══════════════════════════════════════════════════════════════════════
def build_pdf(diagrams, sname, roll, site, logo_b=None):
    buf=io.BytesIO(); PW,PH=A4; MG=1.8*cm
    lr=None
    if logo_b:
        try: lr=ImageReader(io.BytesIO(logo_b))
        except Exception: pass
    def pg(cvs,doc):
        cvs.saveState()
        cvs.setFont("Times-Bold",9.5); cvs.setFillColor(RC.HexColor("#0d1829"))
        cvs.drawCentredString(PW/2,PH-1.0*cm,"WIND ROSE DIAGRAM REPORT")
        cvs.setLineWidth(0.7); cvs.setStrokeColor(RC.HexColor("#0d1829"))
        cvs.line(MG,PH-1.35*cm,PW-MG,PH-1.35*cm)
        if lr:
            ls=1.4*cm
            try: cvs.drawImage(lr,PW-MG-ls,PH-1.32*cm,width=ls,height=ls,
                               preserveAspectRatio=True,mask="auto")
            except Exception: pass
        cvs.line(MG,1.4*cm,PW-MG,1.4*cm)
        cvs.setFont("Times-Roman",8.5); cvs.setFillColor(RC.black)
        parts=[]
        if sname and sname.strip(): parts.append(sname.strip())
        if roll  and roll.strip():  parts.append(f"Roll No: {roll.strip()}")
        if site  and site.strip():  parts.append(f"Site: {site.strip()}")
        if parts: cvs.drawString(MG,0.85*cm,"  |  ".join(parts))
        cvs.drawRightString(PW-MG,0.85*cm,f"Page {doc.page}")
        cvs.restoreState()
    doc=SimpleDocTemplate(buf,pagesize=A4,leftMargin=MG,rightMargin=MG,
                          topMargin=1.9*cm,bottomMargin=1.8*cm)
    sty_t=ParagraphStyle("t",fontName="Times-Bold",fontSize=15,
                          textColor=RC.HexColor("#0d1829"),alignment=TA_CENTER,spaceAfter=4)
    sty_c=ParagraphStyle("c",fontName="Times-Italic",fontSize=9,
                          textColor=RC.HexColor("#444"),alignment=TA_CENTER,spaceAfter=4)
    ORD=["t1s","t1m","t2s","t2m"]
    LBL={"t1s":"Type I – Single Runway  (Polygon)",
         "t1m":"Type I – Multi Runway  (Polygon)",
         "t2s":"Type II – Single Runway  (Speed Bars)",
         "t2m":"Type II – Multi Runway  (Speed Bars)"}
    story=[]; first=True
    for key in ORD:
        if key not in diagrams: continue
        if not first: story.append(PageBreak())
        first=False
        story.append(Paragraph(f"Wind Rose Diagram — {LBL[key]}",sty_t))
        story.append(HRFlowable(width="100%",thickness=0.8,
                                color=RC.HexColor("#0d1829"),spaceAfter=10))
        story.append(Spacer(1,0.4*cm))
        story.append(RLImage(io.BytesIO(diagrams[key]),width=14.5*cm,height=14.5*cm,kind="proportional"))
        story.append(Spacer(1,0.25*cm))
        story.append(Paragraph(f"Figure: {LBL[key]}",sty_c))
    if not story: story.append(Paragraph("No diagrams selected.",sty_t))
    doc.build(story,onFirstPage=pg,onLaterPages=pg)
    buf.seek(0); return buf.getvalue()

# ══════════════════════════════════════════════════════════════════════
#  HELPERS
# ══════════════════════════════════════════════════════════════════════
def sc(v,l):
    return (f'<div class="ar-stat"><div class="ar-sv">{v}</div>'
            f'<div class="ar-sl">{l}</div></div>')

# ══════════════════════════════════════════════════════════════════════
#  MAIN UI
# ══════════════════════════════════════════════════════════════════════
def main():
    inject_css()
    T  = TH[st.session_state.theme]
    dk = st.session_state.theme == "dark"

    # ── HERO ─────────────────────────────────────────────────────
    st.markdown(f"""
    <div class="ar-hero">
      {compass_svg(T['acc'],145)}
      <div class="ar-eyebrow">
        <span class="ar-dot"></span>
        TRANSPORTATION ENGINEERING II  ·  RUNWAY ANALYSIS TOOL
        <span class="ar-dot"></span>
      </div>
      <div class="ar-title">&#9672;&nbsp;AERO <em>ROSE</em></div>
      <p class="ar-tagline">
        One Stop to make Wind Rose Diagram with Precision.
      </p>
      <div class="ar-chips">
        <span class="ar-chip">Type I Polygon</span>
        <span class="ar-chip">Type II Speed Bars</span>
        <span class="ar-chip">6-25 · 25-40 · 40-60 km/h</span>
        <span class="ar-chip">ICAO Annex 14</span>
        <span class="ar-chip">CSV / Excel</span>
        <span class="ar-chip">A4 PDF Export</span>
        <span class="ar-chip">Freq CSV</span>
      </div>
    </div>""", unsafe_allow_html=True)

    # ── THEME TOGGLE  — single button ────────────────────────────
    icon  = "☀️ Light Mode" if dk else "🌑 Dark Mode"
    togg  = st.columns([2,5])
    with togg[0]:
        if st.button(icon, key="theme_btn"):
            st.session_state.theme = "light" if dk else "dark"
            st.rerun()

    st.markdown('<div class="ar-hr"></div>',unsafe_allow_html=True)

    # ── HOW IT WORKS ─────────────────────────────────────────────
    st.markdown('<div class="ar-lbl">How It Works</div>',unsafe_allow_html=True)
    hc1,hc2=st.columns(2,gap="large")
    steps=[("Upload","CSV or Excel with wind direction + speed data"),
           ("Map","Select direction and speed columns"),
           ("Configure","Speed unit and direction format"),
           ("Details","Name, roll no, site — appear in PDF footer"),
           ("Generate","Animated runway progress bar while processing"),
           ("Download","A4 PDF + Frequency Table CSV export")]
    with hc1:
        st.markdown('<div class="ar-card">',unsafe_allow_html=True)
        for i,(t,d) in enumerate(steps[:3],1):
            st.markdown(f'<div class="ar-step"><div class="ar-step-num">{i}</div>'
                        f'<div><div class="ar-step-title">{t}</div>'
                        f'<div class="ar-step-desc">{d}</div></div></div>',
                        unsafe_allow_html=True)
        st.markdown('</div>',unsafe_allow_html=True)
    with hc2:
        st.markdown('<div class="ar-card">',unsafe_allow_html=True)
        for i,(t,d) in enumerate(steps[3:],4):
            st.markdown(f'<div class="ar-step"><div class="ar-step-num">{i}</div>'
                        f'<div><div class="ar-step-title">{t}</div>'
                        f'<div class="ar-step-desc">{d}</div></div></div>',
                        unsafe_allow_html=True)
        st.markdown('</div>',unsafe_allow_html=True)

    # ── DIAGRAM TYPES ─────────────────────────────────────────────
    st.markdown('<div class="ar-hr"></div>',unsafe_allow_html=True)
    st.markdown('<div class="ar-lbl">Type I vs Type II</div>',unsafe_allow_html=True)
    st.markdown("""<div class="ar-type-grid">
      <div class="ar-type-card">
        <div class="ar-type-code">I · S</div>
        <div class="ar-type-name">Type I — Single Runway</div>
        <span class="ar-type-badge ar-badge-t1">Polygon / Line Method</span>
        <div class="ar-type-desc">Direction totals (6–60 km/h) joined as a <b>closed polygon</b>.
        Longest spoke = best runway heading. No speed breakdown.</div>
      </div>
      <div class="ar-type-card">
        <div class="ar-type-code">I · M</div>
        <div class="ar-type-name">Type I — Multi Runway</div>
        <span class="ar-type-badge ar-badge-t1">Polygon / Line Method</span>
        <div class="ar-type-desc">Same polygon for two-runway layouts.
        Mark runway axes manually on the printout along the two longest spokes.</div>
      </div>
      <div class="ar-type-card">
        <div class="ar-type-code">II · S</div>
        <div class="ar-type-name">Type II — Single Runway</div>
        <span class="ar-type-badge ar-badge-t2">Multi-Color Speed Bars</span>
        <div class="ar-type-desc">Each direction = <b>stacked bar in 5 distinct colors</b>
        (&lt;6 grey · 6-25 blue · 25-40 green · 40-60 orange · &gt;60 red).</div>
      </div>
      <div class="ar-type-card">
        <div class="ar-type-code">II · M</div>
        <div class="ar-type-name">Type II — Multi Runway</div>
        <span class="ar-type-badge ar-badge-t2">Multi-Color Speed Bars</span>
        <div class="ar-type-desc">Same color-coded speed bars for two-runway layout.
        Mark runway axes manually — highest stacked bars = dominant directions.</div>
      </div>
    </div>""",unsafe_allow_html=True)

    # ── UPLOAD & CONFIG ───────────────────────────────────────────
    st.markdown('<div class="ar-hr"></div>',unsafe_allow_html=True)
    st.markdown('<div class="ar-lbl">Upload &amp; Configure</div>',unsafe_allow_html=True)
    st.markdown('<div class="ar-card">',unsafe_allow_html=True)
    st.markdown('<div class="ar-sub">&#9312; Wind Data File</div>',unsafe_allow_html=True)

    uploaded=st.file_uploader("Upload file",type=["csv","xlsx","xls"],
                               label_visibility="collapsed",key="wind_file")
    st.markdown('<div class="ar-hint">'
                'Required: wind_direction + wind_speed columns  ·  '
                'Speed auto-converted to km/h  ·  '
                'Day/month/year columns ignored automatically  ·  '
                '1–20+ years of data supported'
                '</div>',unsafe_allow_html=True)

    # File processing — show spinner while loading
    if uploaded is not None:
        load_ph = st.empty()
        load_ph.markdown(
            '<div class="ar-loading">'
            '<div class="ar-spinner">'
            '<div class="ar-spinner-ring"></div>'
            '<div class="ar-spinner-ring"></div>'
            '<div class="ar-spinner-ring"></div>'
            '</div>'
            '<div class="ar-load-txt">Loading wind data…</div>'
            '</div>',
            unsafe_allow_html=True)
        df_tmp,err=load_file(uploaded); uploaded.seek(0)
        load_ph.empty()
        if err or df_tmp is None:
            st.error(f"Cannot read file: {err}")
            st.session_state._file_loaded=False
        else:
            raw=uploaded.read(); uploaded.seek(0)
            st.session_state._file_bytes  = raw
            st.session_state._file_name   = uploaded.name
            st.session_state._cols        = list(df_tmp.columns)
            st.session_state._file_rows   = len(df_tmp)
            st.session_state._file_loaded = True

    fl   = st.session_state._file_loaded
    cols = st.session_state._cols or []
    fn   = st.session_state._file_name or ""
    fr   = st.session_state._file_rows or 0

    if fl:
        st.markdown(f'<div class="ar-file-banner">&#10003;&nbsp;'
                    f'<b>{fn}</b>&nbsp;&nbsp;·&nbsp;&nbsp;{fr:,} rows&nbsp;&nbsp;'
                    f'·&nbsp;&nbsp;{len(cols)} columns&nbsp;&nbsp;'
                    f'<span style="opacity:.65;">· Config persists across theme changes</span>'
                    f'</div>',unsafe_allow_html=True)

    if fl and cols:
        st.markdown("<br>",unsafe_allow_html=True)
        st.markdown('<div class="ar-sub">&#9313; Column Mapping</div>',unsafe_allow_html=True)
        def _g(opts,kws):
            for kw in kws:
                for i,c in enumerate(opts):
                    if kw in c.lower(): return i
            return 0
        di=_g(cols,["dir","wd","wind_d"]); si=_g(cols,["spee","ws","wind_s","vel"])
        if si==di: si=min(di+1,len(cols)-1)
        m1,m2,m3,m4=st.columns(4)
        with m1: dir_col =st.selectbox("Direction Column",cols,index=di,key="dcol")
        with m2: spd_col =st.selectbox("Speed Column",    cols,index=si,key="scol")
        with m3: dir_fmt =st.selectbox("Direction Format",["Degrees (0–360)","Compass (N, NNE …)"],key="dfmt")
        with m4: spd_unit=st.selectbox("Input Speed Unit",["km/h","knots","m/s"],key="sunit")

        st.markdown("<br>",unsafe_allow_html=True)
        st.markdown('<div class="ar-sub">&#9314; Runway Config</div>',unsafe_allow_html=True)
        rw1,rw2,rw3,rw4=st.columns(4)
        with rw1:
            cx_s=st.selectbox("Crosswind Limit",
                ["10.5 kt (19.4 km/h) — Light","13 kt (24.1 km/h) — Medium",
                 "20 kt (37.0 km/h) — Heavy"],key="cxs")
            cxlim=float(cx_s.split("(")[1].split()[0])
        with rw2: auto=st.checkbox("Auto-detect runways",value=True,key="auto")
        with rw3: r1_in=st.number_input("Runway 1 heading (°)",0,179,0,5,disabled=auto,key="r1i")
        with rw4: r2_in=st.number_input("Runway 2 heading (°)",0,179,45,5,disabled=auto,key="r2i")

        st.markdown("<br>",unsafe_allow_html=True)
        st.markdown('<div class="ar-sub">&#9315; Select Diagrams</div>',unsafe_allow_html=True)
        d1,d2,d3,d4=st.columns(4)
        with d1: s_t1s=st.checkbox("Type I  — Single",  value=True,key="ct1s")
        with d2: s_t1m=st.checkbox("Type I  — Multi",   value=True,key="ct1m")
        with d3: s_t2s=st.checkbox("Type II — Single",  value=True,key="ct2s")
        with d4: s_t2m=st.checkbox("Type II — Multi",   value=True,key="ct2m")
        sel={"t1s":s_t1s,"t1m":s_t1m,"t2s":s_t2s,"t2m":s_t2m}

    st.markdown('</div>',unsafe_allow_html=True)

    # ── STUDENT DETAILS ───────────────────────────────────────────
    st.markdown('<div class="ar-hr"></div>',unsafe_allow_html=True)
    st.markdown('<div class="ar-lbl">Student / Report Details  (Optional)</div>',
                unsafe_allow_html=True)
    st.markdown(f'<div class="ar-info-card">'
                f'<span style="font-family:\'IBM Plex Mono\',monospace;font-size:.7rem;'
                f'letter-spacing:.08em;color:{T["mut"]};">'
                f'&#9432;&nbsp;Name, Roll No and Site appear in the '
                f'<b style="color:{T["acc"]}">PDF footer</b>. '
                f'Logo (PNG/JPG) appears in the '
                f'<b style="color:{T["acc"]}">top-right header</b>. All optional.'
                f'</span></div>',unsafe_allow_html=True)

    ui1,ui2,ui3,ui4=st.columns(4)
    with ui1: stu_name=st.text_input("👤 Student Name",    key="sname")
    with ui2: stu_roll=st.text_input("🎓 Roll Number",     key="sroll")
    with ui3: stu_site=st.text_input("📍 Site / Location", key="ssite")
    with ui4:
        logo_file=st.file_uploader("🖼 Your Logo",
                                    type=["png","jpg","jpeg"],key="logo_up",
                                    label_visibility="visible")

    # ── GENERATE BUTTON ───────────────────────────────────────────
    gen_btn=False
    if fl:
        st.markdown("<br>",unsafe_allow_html=True)
        _,gc,_=st.columns([1,3,1])
        with gc:
            st.markdown('<div class="gen-wrap">',unsafe_allow_html=True)
            gen_btn=st.button("&#9889;  GENERATE WIND ROSE DIAGRAMS  &#9889;",
                              use_container_width=True)
            st.markdown('</div>',unsafe_allow_html=True)

    # ── GENERATE LOGIC ────────────────────────────────────────────
    if gen_btn:
        if not fl:
            st.warning("Please upload a wind data file first.")
        elif not any(sel.values()):
            st.warning("Please select at least one diagram type.")
        else:
            # Save student info
            st.session_state['_pdf_name'] = stu_name or ""
            st.session_state['_pdf_roll'] = stu_roll or ""
            st.session_state['_pdf_site'] = stu_site or ""
            if logo_file is not None:
                try: st.session_state['_pdf_logo']=logo_file.read(); logo_file.seek(0)
                except: st.session_state['_pdf_logo']=None
            else: st.session_state['_pdf_logo']=None

            ph=st.empty()
            ph.markdown(rwy_progress(0,"LOADING DATA"),unsafe_allow_html=True)
            try:
                freq,stats=process_data(
                    st.session_state._file_bytes,st.session_state._file_name,
                    dir_col,spd_col,dir_fmt,spd_unit)
            except ValueError as e:
                ph.empty(); st.error(f"Data error: {e}"); st.stop()
            except Exception as e:
                ph.empty(); st.error(f"Error: {e}"); st.stop()

            ph.markdown(rwy_progress(15,"ANALYSING WIND DATA"),unsafe_allow_html=True)
            if auto:
                r1=best_rwy(freq,cxlim); r2=best_rwy(freq,cxlim,excl=r1)
            else:
                r1,r2=float(r1_in),float(r2_in)

            ph.markdown(rwy_progress(25,"RUNWAY HEADING RESOLVED"),unsafe_allow_html=True)
            tnow=st.session_state.theme
            diags={}
            rmap={"t1s":lambda:render_t1s(freq,tnow),
                  "t1m":lambda:render_t1m(freq,tnow),
                  "t2s":lambda:render_t2s(freq,tnow),
                  "t2m":lambda:render_t2m(freq,tnow)}
            lmap={"t1s":"TYPE I SINGLE","t1m":"TYPE I MULTI",
                  "t2s":"TYPE II SINGLE","t2m":"TYPE II MULTI"}
            ts=sum(sel.values()); done=0
            for key,fn in rmap.items():
                if not sel[key]: continue
                try: diags[key]=fn()
                except Exception as e: st.warning(f"Cannot render {key}: {e}")
                done+=1
                ph.markdown(rwy_progress(25+done/ts*72,f"RENDERING {lmap[key]}"),
                             unsafe_allow_html=True)

            ph.markdown(rwy_progress(100,"CLEARED FOR TAKEOFF"),unsafe_allow_html=True)
            time.sleep(0.5); ph.empty()

            st.session_state.diagrams=diags; st.session_state.freq=freq
            st.session_state.rwy1=r1; st.session_state.rwy2=r2
            st.session_state.stats=stats; st.session_state.cxlim=cxlim
            st.session_state.ready=True
            st.success(f" {len(diags)} Diagram(s) Generated!")

    # ── RESULTS ───────────────────────────────────────────────────
    if st.session_state.ready and st.session_state.diagrams:
        freq=st.session_state.freq; r1=st.session_state.rwy1
        r2=st.session_state.rwy2;   stats=st.session_state.stats
        cx=st.session_state.cxlim;  diags=st.session_state.diagrams
        c1=rwy_cov(freq,r1,cx); c2=rwy_cov(freq,r2,cx)
        cc=comb_cov(freq,r1,r2,cx); icao=cc>=95.

        st.markdown('<div class="ar-hr"></div>',unsafe_allow_html=True)
        st.markdown('<div class="ar-lbl">Analysis Results</div>',unsafe_allow_html=True)

        st.markdown(
            '<div class="ar-stats">'
            +sc(f"{stats['total']:,}","Total Obs.")
            +sc(f"{stats['calm']:.1f}%","Calm <6")
            +sc(f"{stats['op']:.1f}%","6–60 km/h")
            +sc(f"{stats['avg']:.0f}","Avg km/h")
            +sc(stats['dom'],"Dominant Dir.")
            +sc(f"{cc:.1f}%","Coverage")
            +'</div>',unsafe_allow_html=True)

        badge=(f'<span class="ar-pass">&#10003; ICAO PASS</span>'
               if icao else f'<span class="ar-fail">&#10007; ICAO FAIL</span>')
        st.markdown(f"""<div class="ar-cov">
          &#9992; <b>{rwy_lbl(r1)}</b>: {c1:.1f}%
          &nbsp;&middot;&nbsp; &#9992; <b>{rwy_lbl(r2)}</b>: {c2:.1f}%
          &nbsp;&middot;&nbsp; Combined: <b>{cc:.1f}%</b>
          &nbsp;&middot;&nbsp; CW Limit: <b>{cx:.1f} km/h</b>
          &nbsp;&middot;&nbsp; ICAO &ge;95%: {badge}
        </div>""",unsafe_allow_html=True)

        # ── FREQUENCY TABLE — custom toggle, no expander ──────────
        tog_lbl = "▲ Hide Frequency Table" if st.session_state.show_table else "▼ Show Frequency Table"
        st.markdown(f'<div style="margin:.5rem 0 .3rem;">',unsafe_allow_html=True)
        if st.button(tog_lbl, key="tog_tbl"):
            st.session_state.show_table = not st.session_state.show_table
            st.rerun()
        st.markdown('</div>',unsafe_allow_html=True)

        if st.session_state.show_table:
            st.markdown(freq_table_html(freq, T), unsafe_allow_html=True)
            # Export CSV button — always visible below table
            ec1,ec2=st.columns([2,3])
            with ec1:
                st.download_button(
                    label="⬇  Export Frequency Table CSV",
                    data=freq_to_csv(freq),
                    file_name="wind_frequency_table.csv",
                    mime="text/csv",
                    key="csv_dl")

        # ── DIAGRAM PREVIEWS — always white ───────────────────────
        st.markdown("<br>",unsafe_allow_html=True)
        st.markdown('<div class="ar-lbl">Diagram Previews</div>',unsafe_allow_html=True)
        DLBL={"t1s":"Type I — Single","t1m":"Type I — Multi",
              "t2s":"Type II — Single","t2m":"Type II — Multi"}
        vis=[k for k in ["t1s","t1m","t2s","t2m"] if k in diags]
        for ri in range(0,len(vis),2):
            row_k=vis[ri:ri+2]; rcols=st.columns(len(row_k),gap="large")
            for ci,key in enumerate(row_k):
                with rcols[ci]:
                    st.markdown(f'<div class="ar-dlbl">{DLBL[key]}</div>',
                                unsafe_allow_html=True)
                    st.markdown('<div class="ar-diag-white">',unsafe_allow_html=True)
                    st.image(diags[key],use_container_width=True)
                    st.markdown('</div>',unsafe_allow_html=True)

        # ── PDF DOWNLOAD ──────────────────────────────────────────
        st.markdown('<div class="ar-hr"></div>',unsafe_allow_html=True)
        st.markdown('<div class="ar-lbl">Export PDF Report</div>',unsafe_allow_html=True)

        _pn=st.session_state.get('_pdf_name','')
        _pr=st.session_state.get('_pdf_roll','')
        _ps=st.session_state.get('_pdf_site','')
        _pl=st.session_state.get('_pdf_logo',None)


        st.info("Thank  you! Now you can Download you Report.")

        _,dc,_=st.columns([1,2,1])
        with dc:
            with st.spinner("Preparing PDF…"):
                pdf_b=build_pdf(diags,_pn,_pr,_ps,_pl)
            st.download_button("&#11015;  DOWNLOAD  PDF  REPORT  ",
                               data=pdf_b,file_name="WindRose_Diagram.pdf",
                               mime="application/pdf",use_container_width=True)

    # ── FOOTER ────────────────────────────────────────────────────
    st.markdown(f"""
    <div class="ar-footer">
      <div class="ar-footer-brand">&#9672; AERO&middot;ROSE</div>
      <div class="ar-footer-line">Wind Rose Diagram Generator &nbsp;&middot;&nbsp; Transportation Engineering II</div>
      <div style="font-family:'Bebas Neue',cursive;font-size:.88rem;
           letter-spacing:.2em;color:{T['acc']};margin:.7rem 0 .3rem;">WEB BY ABDUL SAMAD</div>
      <div class="ar-footer-line">
        Contact: <a class="ar-footer-link" href="mailto:abdulsamad@gmail.com">abdulsamad@gmail.com</a>
      
    </div>""",unsafe_allow_html=True)

main()
