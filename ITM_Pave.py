# ╔══════════════════════════════════════════════════════════════════╗
# ║  ITM Pave Pro — AASHTO 1993 Pavement Design System             ║
# ║  Single-file Streamlit App | Engineering Green Theme            ║
# ║  พัฒนาโดย รศ.ดร.อิทธิพล มีผล | ภาควิชาครุศาสตร์โยธา มจพ.    ║
# ╠══════════════════════════════════════════════════════════════════╣
# ║  TAB 1 │ ESAL Calculator (Upload Excel / Manual + Growth Rate) ║
# ║  TAB 2 │ CBR Analysis (Percentile)                             ║
# ║  TAB 3 │ Flexible Pavement Design                              ║
# ║  TAB 4 │ K-Value Nomograph (k∞ + Loss of Support)             ║
# ║  TAB 5 │ Rigid Design (JPCP / JRCP / CRCP)                    ║
# ║  TAB 6 │ Report & Save (Word + JSON)                           ║
# ╚══════════════════════════════════════════════════════════════════╝

# ─────────────────────────────────────────────
#  SEC 1: IMPORTS
# ─────────────────────────────────────────────
import streamlit as st
import pandas as pd
import numpy as np
import math
import json
import io
import base64
from datetime import datetime

try:
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt
    import matplotlib.patches as patches
    MPL_OK = True
except ImportError:
    MPL_OK = False

try:
    import plotly.graph_objects as go
    PLOTLY_OK = True
except ImportError:
    PLOTLY_OK = False

try:
    from scipy.optimize import brentq as _brentq
    SCIPY_OK = True
except ImportError:
    SCIPY_OK = False
    def _brentq(f, a, b, xtol=1e-6, maxiter=500):
        fa, fb = f(a), f(b)
        if fa * fb > 0:
            raise ValueError("No sign change")
        for _ in range(maxiter):
            mid = (a + b) / 2.0
            fm = f(mid)
            if abs(fm) < xtol or (b - a) / 2.0 < xtol:
                return mid
            if fa * fm < 0:
                b, fb = mid, fm
            else:
                a, fa = mid, fm
        return (a + b) / 2.0

# openpyxl — required for pd.read_excel
try:
    import openpyxl
    from openpyxl import load_workbook
    OPENPYXL_OK = True
except ModuleNotFoundError:
    OPENPYXL_OK = False
    openpyxl = None
    load_workbook = None

try:
    from docx import Document as DocxDoc
    from docx.shared import Inches, Pt, Cm, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    DOCX_OK = True
except ImportError:
    DOCX_OK = False

# ─────────────────────────────────────────────
#  SEC 2: PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="ITM Pave Pro – AASHTO 1993",
    page_icon="🛣️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
#  SEC 3: CSS — Engineering Green Theme
# ─────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;600;700&family=IBM+Plex+Mono:wght@400;600&display=swap');

html, body, [class*="css"] {
    font-family: 'Sarabun', sans-serif;
}

/* ── Header ── */
.main-header {
    background: linear-gradient(135deg, #1B5E20 0%, #388E3C 60%, #43A047 100%);
    color: white;
    padding: 1.4rem 2rem;
    border-radius: 14px;
    margin-bottom: 1.5rem;
    box-shadow: 0 6px 20px rgba(27,94,32,0.35);
    border-left: 6px solid #A5D6A7;
}
.main-header h1 { margin:0; font-size:1.75rem; font-weight:700; letter-spacing:-0.5px; }
.main-header p  { margin:0.3rem 0 0; font-size:0.9rem; opacity:0.88; }

/* ── Section Cards ── */
.card {
    background: #fff;
    border: 1px solid #C8E6C9;
    border-left: 5px solid #2E7D32;
    border-radius: 10px;
    padding: 1rem 1.3rem;
    margin-bottom: 1rem;
    box-shadow: 0 2px 8px rgba(46,125,50,0.08);
}
.card h4 { color:#1B5E20; margin:0 0 0.8rem; font-size:1rem; font-weight:700; }

/* ── Status Badges ── */
.badge-ready  { background:#E8F5E9; color:#2E7D32; border:1px solid #A5D6A7;
                border-radius:20px; padding:0.25rem 0.8rem; font-size:0.82rem; font-weight:600; display:inline-block; }
.badge-wait   { background:#FFF8E1; color:#E65100; border:1px solid #FFE082;
                border-radius:20px; padding:0.25rem 0.8rem; font-size:0.82rem; font-weight:600; display:inline-block; }
.badge-na     { background:#F5F5F5; color:#757575; border:1px solid #E0E0E0;
                border-radius:20px; padding:0.25rem 0.8rem; font-size:0.82rem; font-weight:600; display:inline-block; }

/* ── Result Boxes ── */
.result-pass { background:#E8F5E9; border:1px solid #A5D6A7; border-radius:8px;
               padding:0.8rem 1rem; color:#1B5E20; font-weight:600; margin:0.3rem 0; }
.result-fail { background:#FFEBEE; border:1px solid #EF9A9A; border-radius:8px;
               padding:0.8rem 1rem; color:#B71C1C; font-weight:600; margin:0.3rem 0; }
.result-info { background:#E3F2FD; border:1px solid #90CAF9; border-radius:8px;
               padding:0.8rem 1rem; color:#0D47A1; font-weight:600; margin:0.3rem 0; }
.result-warn { background:#FFF8E1; border:1px solid #FFE082; border-radius:8px;
               padding:0.8rem 1rem; color:#E65100; font-weight:600; margin:0.3rem 0; }

/* ── Metric Box ── */
.metric-box {
    background:#fff; border:1px solid #C8E6C9; border-radius:12px;
    padding:1rem; text-align:center;
    box-shadow:0 2px 8px rgba(46,125,50,0.10);
}
.metric-box .val { font-size:1.5rem; font-weight:700; color:#1B5E20;
                   font-family:'IBM Plex Mono', monospace; }
.metric-box .lbl { font-size:0.78rem; color:#558B2F; margin-top:0.2rem; }

/* ── Tabs ── */
[data-baseweb="tab-list"] { gap:3px; }
[data-baseweb="tab"] {
    background:#E8F5E9 !important; border-radius:8px 8px 0 0 !important;
    font-weight:600 !important; color:#1B5E20 !important;
    padding:0.45rem 0.9rem !important;
}
[aria-selected="true"][data-baseweb="tab"] {
    background:#2E7D32 !important; color:white !important;
}

/* ── Sidebar ── */
[data-testid="stSidebar"] { background:#1B5E20; }
[data-testid="stSidebar"] * { color:#E8F5E9 !important; }
[data-testid="stSidebar"] hr { border-color:#2E7D32; }

/* ── Buttons ── */
button[kind="primary"] {
    background:#2E7D32 !important; border-radius:8px !important;
    font-weight:700 !important;
}

/* ── Number inputs ── */
.stNumberInput > div > div > input {
    font-family:'IBM Plex Mono', monospace; font-weight:600;
}

/* ── DataFrames ── */
.stDataFrame { border-radius:8px; overflow:hidden; }

/* ── Flow Arrow ── */
.flow-arrow {
    text-align:center; font-size:1.5rem; color:#43A047;
    margin:0.3rem 0; line-height:1;
}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
#  SEC 4: CONSTANTS & LOOKUP TABLES
# ─────────────────────────────────────────────
TON_TO_KIP = 2.2046

VEHICLE_AXLES = {
    "MB":  [(4,  1, 1), (11, 1, 1)],
    "HB":  [(5,  1, 1), (20, 2, 1)],
    "MT":  [(4,  1, 1), (11, 1, 1)],
    "HT":  [(5,  1, 1), (20, 2, 1)],
    "TR":  [(5,  1, 1), (20, 2, 1), (11, 1, 2)],
    "STR": [(5,  1, 1), (20, 2, 2)],
}
VEHICLE_LABELS = {
    "MB": "Medium Bus",  "HB": "Heavy Bus",
    "MT": "Medium Truck","HT": "Heavy Truck",
    "TR": "Trailer",     "STR": "Semi Trailer",
}
VEHICLE_COLS = ["MB", "HB", "MT", "HT", "TR", "STR"]

SLAB_THICKNESSES = [25, 28, 30, 32, 35]
SN_DEFAULTS      = [6.5, 7.1, 7.5, 8.0]

ZR_MAP = {
    50:0.000,60:-0.253,70:-0.524,75:-0.674,80:-0.841,85:-1.037,
    90:-1.282,91:-1.340,92:-1.405,93:-1.476,94:-1.555,95:-1.645,
    96:-1.751,97:-1.881,98:-2.054,99:-2.327,
}

J_VALUES = {"JPCP": 2.8, "JRCP": 2.8, "CRCP": 2.6}

FLEX_LAYER_MATERIALS = {
    "None":                                       (None, None),
    "Asphalt Concrete (AC)":                      (0.42, 1.0),
    "Cement Treated Base UCS 24.5 ksc":           (0.15, 1.0),
    "Cement Treated Base UCS 17.5 ksc":           (0.13, 1.0),
    "Crushed Rock Base CBR 80%":                  (0.14, 1.0),
    "Soil Aggregate Subbase CBR 25%":             (0.10, 1.0),
    "Soil Aggregate Subbase CBR 20%":             (0.09, 1.0),
    "Soil Aggregate Subbase CBR 15%":             (0.08, 1.0),
    "Sand Embankment CBR 10%":                    (0.08, 1.0),
}

RIGID_LAYER_MATERIALS = {
    "None":                                       None,
    "AC Interlayer":                              2500,
    "Lean Concrete Base (LCB)":                  5000,
    "Cement Treated Base UCS 24.5 ksc":          1200,
    "Cement Treated Base UCS 17.5 ksc":           850,
    "Crushed Rock Base CBR 80%":                  350,
    "Soil Aggregate Subbase CBR 25%":             150,
    "Soil Aggregate Subbase CBR 20%":             120,
    "Soil Aggregate Subbase CBR 15%":             100,
    "Sand Embankment CBR 10%":                    100,
}

SAMPLE_CBR = [14.8,14.37,5.31,17.37,5.48,18.46,4.85,6.23,
              5.02,10.78,10.52,14,15.5,8.7,12.93,8.19,
              8.1,15.56,16.88,20.75,20.3,8,7.84,7.48,
              23.55,8.92,13.3,13.5,13.86,7.18,6.95,5.8,
              6,11.18,9.69,7.48]

# ─────────────────────────────────────────────
#  SEC 5: ENGINE FUNCTIONS
# ─────────────────────────────────────────────

def ealf_flex(L1_ton, L2, SN, Pt):
    L1  = L1_ton * TON_TO_KIP
    Gt  = math.log10((4.2 - Pt) / (4.2 - 1.5))
    Bx  = 0.40 + 0.081*(L1+L2)**3.23 / ((SN+1)**5.19 * L2**3.23)
    B18 = 0.40 + 0.081*(18+1)**3.23  / ((SN+1)**5.19 * 1.0**3.23)
    return 10**(4.79*math.log10(L1+L2) - 4.33*math.log10(L2)
                - 4.79*math.log10(19) + Gt*(1/B18 - 1/Bx))

def ealf_rigid(L1_ton, L2, D_cm, Pt):
    L1  = L1_ton * TON_TO_KIP
    D   = D_cm / 2.54
    Gt  = math.log10((4.5 - Pt) / (4.5 - 1.5))
    Bx  = 1.0 + 3.63*(L1+L2)**5.20 / ((D+1)**8.46 * L2**3.52)
    B18 = 1.0 + 3.63*(18+1)**5.20  / ((D+1)**8.46 * 1.0**3.52)
    return 10**(4.62*math.log10(L1+L2) - 3.28*math.log10(L2)
                - 4.62*math.log10(19) + Gt*(1/B18 - 1/Bx))

def truck_factor_flex(vtype, SN, Pt):
    return sum(ealf_flex(L1,L2,SN,Pt)*cnt for L1,L2,cnt in VEHICLE_AXLES[vtype])

def truck_factor_rigid(vtype, D_cm, Pt):
    return sum(ealf_rigid(L1,L2,D_cm,Pt)*cnt for L1,L2,cnt in VEHICLE_AXLES[vtype])

def compute_esal_from_df(traffic_df, ldf, ddf, Pt, mode="rigid", sn_list=None):
    """
    traffic_df: DataFrame with columns Year, MB, HB, MT, HT, TR, STR
    Returns dict: {D_cm: esal} for rigid, {SN: esal} for flexible
    """
    if mode == "rigid":
        keys = SLAB_THICKNESSES
        results = {k: 0.0 for k in keys}
        for _, row in traffic_df.iterrows():
            for vtype in VEHICLE_COLS:
                cnt = float(row.get(vtype, 0) or 0)
                if cnt <= 0: continue
                for D in keys:
                    tf = truck_factor_rigid(vtype, D, Pt)
                    results[D] += cnt * tf * ldf * ddf
        return results
    else:
        keys = sn_list or SN_DEFAULTS
        results = {k: 0.0 for k in keys}
        for _, row in traffic_df.iterrows():
            for vtype in VEHICLE_COLS:
                cnt = float(row.get(vtype, 0) or 0)
                if cnt <= 0: continue
                for SN in keys:
                    tf = truck_factor_flex(vtype, SN, Pt)
                    results[SN] += cnt * tf * ldf * ddf
        return results

def aashto_sn_required(esal, zr, so, pi, pt, mr_psi):
    delta_psi = pi - pt
    logW18 = math.log10(max(esal, 1))
    def eq(SN):
        if SN <= 0: return -1e10
        t1 = zr * so
        t2 = 9.36*math.log10(SN+1) - 0.20
        t3 = math.log10(delta_psi/2.7) / (0.40 + 1094/(SN+1)**5.19)
        t4 = 2.32*math.log10(mr_psi) - 8.07
        return t1+t2+t3+t4 - logW18
    try:
        return _brentq(eq, 0.1, 30, xtol=1e-4)
    except:
        return None

def aashto_rigid_w18(d_cm, pi, pt, zr, so, sc_psi, cd, j, ec_psi, k_pci):
    d_in = d_cm / 2.54
    delta_psi = pi - pt
    t1 = zr * so
    t2 = 7.35*math.log10(d_in+1) - 0.06
    t3 = math.log10(delta_psi/3.0) / (1 + 1.624e7/(d_in+1)**8.46)
    num4 = sc_psi * cd * (d_in**0.75 - 1.132)
    den4 = 215.63 * j * (d_in**0.75 - 18.42/(ec_psi/k_pci)**0.25)
    if num4 <= 0 or den4 <= 0:
        return None
    inner = num4 / den4
    if inner <= 0:
        return None
    t4 = (4.22 - 0.32*pt) * math.log10(inner)
    return 10**(t1+t2+t3+t4)

def cbr_to_mr(cbr):
    return 1500.0 * cbr

def mr_to_k(mr_psi):
    return mr_psi / 19.4

def calc_percentile_cbr(cbr_values):
    arr = np.sort(np.array(cbr_values, dtype=float))
    n   = len(arr)
    unique_cbr = np.unique(arr)
    unique_pct = np.array([np.sum(arr >= v)/n*100 for v in unique_cbr])
    return arr, n, unique_cbr, unique_pct

def grow_traffic(base_row, growth_rate_pct, years):
    """Generate DataFrame with yearly traffic from base year + growth rate."""
    r = growth_rate_pct / 100.0
    rows = []
    for y in range(1, years+1):
        factor = (1 + r)**(y - 1)
        row = {"Year": y}
        for v in VEHICLE_COLS:
            row[v] = int(round(base_row.get(v, 0) * factor))
        rows.append(row)
    return pd.DataFrame(rows)

# ─────────────────────────────────────────────
#  SEC 6: NOMOGRAPH DRAW FUNCTIONS (Matplotlib)
# ─────────────────────────────────────────────

def draw_k_infinity_nomograph(esb_psi, dsb_in, k_sub_pci):
    """
    Composite k∞ Nomograph (AASHTO 1993 Fig. 3.3 approximation)
    3 axes: Esb (left), DSB (center), k∞ (right)
    """
    fig, ax = plt.subplots(figsize=(8, 9))
    ax.set_xlim(0, 10); ax.set_ylim(0, 10)
    ax.axis('off')
    ax.set_facecolor('#F1F8E9')
    fig.patch.set_facecolor('#F1F8E9')

    # Axis x-positions
    x_esb, x_dsb, x_kinf = 1.5, 5.0, 8.5

    # ── Axis lines ──
    for x in [x_esb, x_dsb, x_kinf]:
        ax.plot([x, x], [0.5, 9.5], color='#1B5E20', lw=2.5)

    # ── Esb axis (left): 5,000 – 100,000 psi ──
    esb_range = [5000, 10000, 20000, 30000, 50000, 100000]
    esb_log_min = math.log10(5000)
    esb_log_max = math.log10(100000)
    def esb_to_y(v):
        return 0.5 + 9.0*(math.log10(v)-esb_log_min)/(esb_log_max-esb_log_min)
    for v in esb_range:
        y = esb_to_y(v)
        ax.plot([x_esb-0.15, x_esb+0.15], [y, y], color='#1B5E20', lw=1.5)
        ax.text(x_esb-0.25, y, f"{v:,}", ha='right', va='center', fontsize=8, color='#1B5E20')
    ax.text(x_esb, 9.8, "Esb (psi)", ha='center', va='bottom', fontsize=9, fontweight='bold', color='#1B5E20')

    # ── DSB axis (center): 0 – 36 in ──
    dsb_range = [0, 4, 8, 12, 16, 20, 24, 28, 32, 36]
    def dsb_to_y(v):
        return 0.5 + 9.0*(v/36.0)
    for v in dsb_range:
        y = dsb_to_y(v)
        ax.plot([x_dsb-0.15, x_dsb+0.15], [y, y], color='#2E7D32', lw=1.5)
        ax.text(x_dsb+0.25, y, f"{v}", ha='left', va='center', fontsize=8, color='#2E7D32')
    ax.text(x_dsb, 9.8, "DSB (in)", ha='center', va='bottom', fontsize=9, fontweight='bold', color='#2E7D32')

    # ── k∞ axis (right): 50 – 1000 pci ──
    kinf_range = [50, 100, 150, 200, 300, 500, 700, 1000]
    kinf_log_min = math.log10(50)
    kinf_log_max = math.log10(1000)
    def kinf_to_y(v):
        return 0.5 + 9.0*(math.log10(v)-kinf_log_min)/(kinf_log_max-kinf_log_min)
    for v in kinf_range:
        y = kinf_to_y(v)
        ax.plot([x_kinf-0.15, x_kinf+0.15], [y, y], color='#388E3C', lw=1.5)
        ax.text(x_kinf+0.25, y, f"{v}", ha='left', va='center', fontsize=8, color='#388E3C')
    ax.text(x_kinf, 9.8, "k∞ (pci)", ha='center', va='bottom', fontsize=9, fontweight='bold', color='#388E3C')

    # ── Compute k∞ from inputs via AASHTO approximation ──
    # Odemark: k∞ = k_sub * (1 + Esb*DSB/(k_sub*19.4*DSB^0.5 + 1))^0.4  (simplified)
    if esb_psi > 0 and dsb_in >= 0 and k_sub_pci > 0:
        if dsb_in == 0:
            k_inf_calc = k_sub_pci
        else:
            h_eq = dsb_in * (esb_psi / (k_sub_pci * 19.4))**(1/3)
            k_inf_calc = min(k_sub_pci * (1 + 0.55*h_eq**0.45), 1000)
        k_inf_calc = max(50, min(1000, k_inf_calc))
    else:
        k_inf_calc = k_sub_pci

    # ── Draw reading lines ──
    y_esb  = esb_to_y(max(5000, min(100000, esb_psi)))
    y_dsb  = dsb_to_y(max(0, min(36, dsb_in)))
    y_kinf = kinf_to_y(max(50, min(1000, k_inf_calc)))

    ax.annotate("", xy=(x_dsb, y_dsb), xytext=(x_esb, y_esb),
                arrowprops=dict(arrowstyle="-", color='red', lw=2, linestyle='dashed'))
    ax.annotate("", xy=(x_kinf, y_kinf), xytext=(x_dsb, y_dsb),
                arrowprops=dict(arrowstyle="->", color='red', lw=2, linestyle='dashed'))

    # Dots at reading points
    for (xp, yp) in [(x_esb, y_esb), (x_dsb, y_dsb), (x_kinf, y_kinf)]:
        ax.plot(xp, yp, 'o', color='red', markersize=8, zorder=5)

    # k∞ result label
    ax.text(x_kinf+1.0, y_kinf, f"k∞ = {k_inf_calc:.0f} pci",
            ha='left', va='center', fontsize=11, fontweight='bold',
            color='red', bbox=dict(boxstyle='round,pad=0.3', facecolor='white', edgecolor='red', alpha=0.9))

    ax.set_title("Composite k∞ Nomograph (AASHTO 1993 Fig.3.3)", fontsize=11,
                 fontweight='bold', color='#1B5E20', pad=15)
    plt.tight_layout()
    return fig, k_inf_calc


def draw_loss_of_support_nomograph(k_inf_pci, ls_value):
    """
    Loss of Support Nomograph (AASHTO 1993 Fig. 3.7)
    k_corrected = k_inf / 10^(LS*0.5) approximation
    """
    fig, ax = plt.subplots(figsize=(7, 8))
    ax.set_xlim(0, 10); ax.set_ylim(0, 10)
    ax.axis('off')
    ax.set_facecolor('#F1F8E9')
    fig.patch.set_facecolor('#F1F8E9')

    x_kinf, x_kcorr = 2.5, 7.5
    ls_colors = {0.0:'#1B5E20', 0.5:'#2E7D32', 1.0:'#43A047',
                 1.5:'#66BB6A', 2.0:'#EF6C00', 3.0:'#B71C1C'}

    k_range = [10, 20, 50, 100, 200, 300, 500, 700, 1000, 1500, 2000, 3000]
    k_log_min = math.log10(10)
    k_log_max = math.log10(3000)
    def k_to_y(v):
        return 0.5 + 9.0*(math.log10(max(v,10))-k_log_min)/(k_log_max-k_log_min)

    for x in [x_kinf, x_kcorr]:
        ax.plot([x,x],[0.5,9.5], color='#1B5E20', lw=2.5)

    for v in k_range:
        for x in [x_kinf, x_kcorr]:
            y = k_to_y(v)
            ax.plot([x-0.15, x+0.15],[y,y], color='#1B5E20', lw=1.5)
        ax.text(x_kinf-0.25, k_to_y(v), f"{v}", ha='right', va='center', fontsize=8, color='#1B5E20')
        ax.text(x_kcorr+0.25, k_to_y(v), f"{v}", ha='left', va='center', fontsize=8, color='#388E3C')

    ax.text(x_kinf, 9.8, "k∞ (pci)", ha='center', va='bottom', fontsize=9, fontweight='bold', color='#1B5E20')
    ax.text(x_kcorr, 9.8, "k_eff (pci)", ha='center', va='bottom', fontsize=9, fontweight='bold', color='#388E3C')

    # Draw LS lines for all LS values
    for ls, lc in ls_colors.items():
        for k_val in [20, 50, 100, 200, 500, 1000, 2000]:
            k_corr_ls = k_val / (10**(ls * 0.5))
            k_corr_ls = max(10, min(3000, k_corr_ls))
            y1 = k_to_y(k_val)
            y2 = k_to_y(k_corr_ls)
            ax.plot([x_kinf, x_kcorr],[y1, y2], color=lc, lw=0.8, alpha=0.4)
        ax.text(5.0, k_to_y(50/(10**(ls*0.5))) + ls*0.3,
                f"LS={ls}", ha='center', va='center', fontsize=7, color=lc, alpha=0.8)

    # ── Compute k_corrected ──
    k_corr_calc = k_inf_pci / (10**(ls_value * 0.5))
    k_corr_calc = max(10, min(3000, k_corr_calc))

    # Draw user's reading line
    y1 = k_to_y(max(10, min(3000, k_inf_pci)))
    y2 = k_to_y(k_corr_calc)
    ax.annotate("", xy=(x_kcorr, y2), xytext=(x_kinf, y1),
                arrowprops=dict(arrowstyle="->", color='red', lw=2.5))
    for (xp, yp) in [(x_kinf, y1), (x_kcorr, y2)]:
        ax.plot(xp, yp, 'o', color='red', markersize=9, zorder=5)

    ax.text(x_kcorr+1.2, y2,
            f"k_eff =\n{k_corr_calc:.0f} pci",
            ha='left', va='center', fontsize=10, fontweight='bold',
            color='red', bbox=dict(boxstyle='round,pad=0.3', facecolor='white', edgecolor='red', alpha=0.9))

    ax.set_title(f"Loss of Support Nomograph  (LS = {ls_value})", fontsize=11,
                 fontweight='bold', color='#1B5E20', pad=15)

    # Legend
    legend_x, legend_y = 3.5, 1.5
    for i, (ls, lc) in enumerate(ls_colors.items()):
        ax.plot(legend_x, legend_y - i*0.3, 's', color=lc, markersize=7)
        ax.text(legend_x+0.2, legend_y - i*0.3, f"LS = {ls}", va='center', fontsize=7, color=lc)

    plt.tight_layout()
    return fig, k_corr_calc

def fig_to_bytes(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=150, bbox_inches='tight',
                facecolor=fig.get_facecolor())
    buf.seek(0)
    return buf.read()

# ─────────────────────────────────────────────
#  SEC 7: WORD REPORT FUNCTIONS
# ─────────────────────────────────────────────

def _new_doc():
    if not DOCX_OK:
        return None
    doc = DocxDoc()
    style = doc.styles['Normal']
    style.font.name = 'TH SarabunPSK'
    style.font.size = Pt(15)
    try:
        style._element.rPr.rFonts.set(qn('w:eastAsia'), 'TH SarabunPSK')
    except:
        pass
    return doc

def _th_run(para, text, bold=False, size=15, color=None):
    run = para.add_run(text)
    run.font.name = 'TH SarabunPSK'
    run.font.size = Pt(size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = RGBColor(*color)
    try:
        run._element.rPr.rFonts.set(qn('w:eastAsia'), 'TH SarabunPSK')
    except:
        pass
    return run

def _set_cell_bg(cell, hex_color):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

def _doc_to_bytes(doc):
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()

def _add_section_heading(doc, text, level=1):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    r = _th_run(p, text, bold=True, size=16 if level==1 else 14)
    return p

def _add_simple_table(doc, headers, rows, hdr_bg='C8E6C9'):
    tbl = doc.add_table(rows=1+len(rows), cols=len(headers))
    tbl.style = 'Table Grid'
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    for j, h in enumerate(headers):
        cell = tbl.rows[0].cells[j]
        p = cell.paragraphs[0]
        _th_run(p, h, bold=True, size=13)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _set_cell_bg(cell, hdr_bg)
    for i, row_data in enumerate(rows):
        for j, val in enumerate(row_data):
            cell = tbl.rows[i+1].cells[j]
            p = cell.paragraphs[0]
            _th_run(p, str(val), size=13)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    return tbl

def build_report_esal(ss):
    doc = _new_doc()
    if not doc: return None
    _add_section_heading(doc, "1. ผลการคำนวณ ESAL (Equivalent Single Axle Load)")
    p = doc.add_paragraph()
    _th_run(p, f"วิธีคำนวณ AASHTO 1993  |  LDF = {ss.get('ldf',0.9)}  |  DDF = {ss.get('ddf',0.5)}", size=13)
    doc.add_paragraph()

    if ss.get('esal_rigid'):
        _add_section_heading(doc, "1.1 ESAL – ผิวทางคอนกรีต", level=2)
        hdr = ["Slab Thickness (cm)", "ESAL (Design Lane)"]
        rows = [[f"{d} cm", f"{v:,.0f}"] for d,v in ss['esal_rigid'].items()]
        _add_simple_table(doc, hdr, rows)
        doc.add_paragraph()

    if ss.get('esal_flex'):
        _add_section_heading(doc, "1.2 ESAL – ผิวทางลาดยาง", level=2)
        hdr = ["Structure Number (SN)", "ESAL (Design Lane)"]
        rows = [[f"SN = {sn}", f"{v:,.0f}"] for sn,v in ss['esal_flex'].items()]
        _add_simple_table(doc, hdr, rows)

    _add_footer(doc)
    return _doc_to_bytes(doc)

def build_report_cbr(ss):
    doc = _new_doc()
    if not doc: return None
    _add_section_heading(doc, "2. ผลการวิเคราะห์ค่า CBR")
    cbr_vals = ss.get('cbr_values', [])
    if cbr_vals:
        arr, n, u_cbr, u_pct = calc_percentile_cbr(cbr_vals)
        pct = ss.get('cbr_percentile', 90)
        cbr_d = float(np.interp(pct, u_pct[::-1], u_cbr[::-1]))
        p = doc.add_paragraph()
        _th_run(p, f"จำนวนตัวอย่าง: {n}  |  Percentile: {pct}%  |  CBR ออกแบบ: {cbr_d:.2f}%", size=13)
        doc.add_paragraph()
        hdr = ["ลำดับ","CBR (%)","จำนวน ≥","Percentile (%)"]
        rows = []
        for i,(v,p2) in enumerate(zip(u_cbr, u_pct)):
            cnt = int(np.sum(arr >= v))
            rows.append([str(i+1), f"{v:.2f}", str(cnt), f"{p2:.1f}"])
        _add_simple_table(doc, hdr, rows)
    _add_footer(doc)
    return _doc_to_bytes(doc)

def build_report_flexible(ss):
    doc = _new_doc()
    if not doc: return None
    _add_section_heading(doc, "3. ผลการออกแบบโครงสร้างชั้นทางลาดยาง (AASHTO 1993)")
    res = ss.get('flex_results', {})
    if res:
        p = doc.add_paragraph()
        _th_run(p, f"Design ESAL = {res.get('esal',0):,.0f}  |  SN Required = {res.get('sn_req',0):.3f}  |  SN Provided = {res.get('sn_prov',0):.3f}", size=13)
        p2 = doc.add_paragraph()
        status = "✅ PASS" if res.get('pass') else "❌ FAIL"
        _th_run(p2, f"ผลการตรวจสอบ: {status}", bold=True, size=14)
        doc.add_paragraph()
        layers = res.get('layers', [])
        if layers:
            hdr = ["ชั้นที่","วัสดุ","หนา (cm)","ai","mi","SNi","ΣSNi"]
            rows = [[str(l['layer']),l['material'],str(l['h_cm']),
                     f"{l['ai']:.2f}",f"{l['mi']:.1f}",
                     f"{l['sni']:.3f}",f"{l['cum_sn']:.3f}"] for l in layers]
            _add_simple_table(doc, hdr, rows)
    _add_footer(doc)
    return _doc_to_bytes(doc)

def build_report_kvalue(ss):
    doc = _new_doc()
    if not doc: return None
    _add_section_heading(doc, "4. ค่า k_eff (Effective Modulus of Subgrade Reaction)")
    p = doc.add_paragraph()
    _th_run(p, f"k∞ = {ss.get('k_inf',0):.1f} pci  |  LS = {ss.get('ls_value',0)}  |  k_eff = {ss.get('k_corrected',0):.1f} pci", size=13)
    imgs = ss.get('nomograph_img_k')
    if imgs:
        doc.add_paragraph()
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p_img.add_run()
        r.add_picture(io.BytesIO(imgs), width=Cm(12))
    imgs2 = ss.get('nomograph_img_ls')
    if imgs2:
        doc.add_paragraph()
        p_img2 = doc.add_paragraph()
        p_img2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r2 = p_img2.add_run()
        r2.add_picture(io.BytesIO(imgs2), width=Cm(12))
    _add_footer(doc)
    return _doc_to_bytes(doc)

def build_report_rigid(ss):
    doc = _new_doc()
    if not doc: return None
    _add_section_heading(doc, "5. ผลการออกแบบความหนาถนนคอนกรีต (AASHTO 1993)")
    rigid_res = ss.get('rigid_results', {})
    for ptype, res in rigid_res.items():
        if not res: continue
        _add_section_heading(doc, f"5.{list(rigid_res.keys()).index(ptype)+1} {ptype}", level=2)
        p = doc.add_paragraph()
        _th_run(p, f"Slab = {res.get('d_cm',0)} cm  |  k_eff = {res.get('k_eff',0):.1f} pci  |  f'c = {res.get('fc',0)} ksc", size=13)
        p2 = doc.add_paragraph()
        status = "✅ PASS" if res.get('pass') else "❌ FAIL"
        _th_run(p2, f"W18 Capacity = {res.get('w18_cap',0):,.0f}  |  W18 Required = {res.get('w18_req',0):,.0f}  |  {status}", bold=True, size=13)
        doc.add_paragraph()
    _add_footer(doc)
    return _doc_to_bytes(doc)

def build_report_full(ss):
    doc = _new_doc()
    if not doc: return None
    # Cover
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _th_run(p, "รายการคำนวณออกแบบโครงสร้างชั้นทาง", bold=True, size=20)
    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _th_run(p2, "ตามวิธี AASHTO 1993", size=16)
    p3 = doc.add_paragraph()
    p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _th_run(p3, f"วันที่: {datetime.now().strftime('%d/%m/%Y %H:%M')}", size=13)
    doc.add_page_break()

    sections = [
        ('esal_rigid', build_report_esal),
        ('cbr_values', build_report_cbr),
        ('flex_results', build_report_flexible),
        ('k_corrected', build_report_kvalue),
        ('rigid_results', build_report_rigid),
    ]
    for key, fn in sections:
        if ss.get(key):
            sub_bytes = fn(ss)
            if sub_bytes:
                sub_doc = DocxDoc(io.BytesIO(sub_bytes))
                for elem in sub_doc.element.body:
                    doc.element.body.append(elem)
    _add_footer(doc)
    return _doc_to_bytes(doc)

def _add_footer(doc):
    doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _th_run(p, "พัฒนาโดย รศ.ดร.อิทธิพล มีผล  |  ภาควิชาครุศาสตร์โยธา  |  มจพ.", italic=False, size=12, color=(80,80,80))

# patch _th_run for italic param
_th_run_orig = _th_run
def _th_run(para, text, bold=False, size=15, color=None, italic=False):
    run = para.add_run(text)
    run.font.name = 'TH SarabunPSK'
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    if color:
        run.font.color.rgb = RGBColor(*color)
    try:
        run._element.rPr.rFonts.set(qn('w:eastAsia'), 'TH SarabunPSK')
    except:
        pass
    return run

# ─────────────────────────────────────────────
#  SEC 8: SESSION STATE INIT
# ─────────────────────────────────────────────
def ss_init():
    defaults = {
        # Traffic & ESAL
        'traffic_df':      None,
        'esal_rigid':      {},
        'esal_flex':       {},
        'ldf':             0.9,
        'ddf':             0.5,
        'pt_rigid':        2.5,
        'pt_flex':         2.5,
        'sn_list':         [6.5, 7.1, 7.5, 8.0],
        # CBR
        'cbr_values':      [],
        'cbr_percentile':  90.0,
        'cbr_design':      3.0,
        'mr_subgrade_psi': 4500.0,
        'k_subgrade_pci':  231.9,
        # Flexible
        'flex_results':    {},
        # K-Value
        'k_inf':           0.0,
        'k_corrected':     0.0,
        'ls_value':        1.0,
        'nomograph_img_k': None,
        'nomograph_img_ls':None,
        # Rigid
        'rigid_results':   {},
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

ss_init()
ss = st.session_state

# ─────────────────────────────────────────────
#  HELPER: Status badge
# ─────────────────────────────────────────────
def status_badge(key, label=None):
    val = ss.get(key)
    has = (val is not None and val != {} and val != [] and val != 0.0)
    cls = "badge-ready" if has else "badge-wait"
    icon = "✅" if has else "⚠️"
    lbl = label or key
    return f'<span class="{cls}">{icon} {lbl}</span>'

# ─────────────────────────────────────────────
#  SIDEBAR
# ─────────────────────────────────────────────
with st.sidebar:
    # ── Dependency warnings ──
    if not OPENPYXL_OK:
        st.warning("⚠️ openpyxl ไม่พร้อม — Upload Excel จะไม่ทำงาน")
    if not DOCX_OK:
        st.warning("⚠️ python-docx ไม่พร้อม — Word report จะไม่ทำงาน")

    st.markdown("""
    <div style='text-align:center;padding:1rem 0 0.5rem;'>
        <div style='font-size:2.2rem;'>🛣️</div>
        <div style='font-weight:700;font-size:1.1rem;color:#A5D6A7;'>ITM Pave Pro</div>
        <div style='font-size:0.78rem;color:#81C784;'>AASHTO 1993</div>
    </div>
    """, unsafe_allow_html=True)
    st.divider()

    st.markdown("**📊 สถานะข้อมูล**")
    st.markdown(status_badge('esal_rigid','ESAL Rigid'), unsafe_allow_html=True)
    st.markdown(status_badge('esal_flex','ESAL Flexible'), unsafe_allow_html=True)
    st.markdown(status_badge('cbr_values','CBR Data'), unsafe_allow_html=True)
    st.markdown(status_badge('flex_results','Flex Design'), unsafe_allow_html=True)
    st.markdown(status_badge('k_corrected','K-Value'), unsafe_allow_html=True)
    st.markdown(status_badge('rigid_results','Rigid Design'), unsafe_allow_html=True)
    st.divider()

    st.markdown("**💾 Save / Load Project**")
    if st.button("💾 Save JSON", use_container_width=True):
        save_data = {
            'esal_rigid':     ss.esal_rigid,
            'esal_flex':      {str(k):v for k,v in ss.esal_flex.items()},
            'ldf':            ss.ldf, 'ddf': ss.ddf,
            'pt_rigid':       ss.pt_rigid, 'pt_flex': ss.pt_flex,
            'sn_list':        ss.sn_list,
            'cbr_values':     ss.cbr_values,
            'cbr_percentile': ss.cbr_percentile,
            'cbr_design':     ss.cbr_design,
            'mr_subgrade_psi':ss.mr_subgrade_psi,
            'k_subgrade_pci': ss.k_subgrade_pci,
            'flex_results':   ss.flex_results,
            'k_inf':          ss.k_inf,
            'k_corrected':    ss.k_corrected,
            'ls_value':       ss.ls_value,
            'rigid_results':  ss.rigid_results,
            'traffic_df':     ss.traffic_df.to_dict('records') if ss.traffic_df is not None else None,
        }
        json_bytes = json.dumps(save_data, ensure_ascii=False, indent=2).encode('utf-8')
        st.download_button("📥 Download JSON", json_bytes,
                           file_name=f"itm_pave_{datetime.now().strftime('%Y%m%d_%H%M')}.json",
                           mime="application/json", use_container_width=True)

    uploaded_json = st.file_uploader("📂 Load JSON", type=['json'])
    if uploaded_json:
        try:
            data = json.loads(uploaded_json.read().decode('utf-8'))
            for k in ['esal_rigid','ldf','ddf','pt_rigid','pt_flex','sn_list',
                      'cbr_values','cbr_percentile','cbr_design','mr_subgrade_psi',
                      'k_subgrade_pci','flex_results','k_inf','k_corrected',
                      'ls_value','rigid_results']:
                if k in data:
                    ss[k] = data[k]
            if 'esal_flex' in data:
                ss.esal_flex = {float(k):v for k,v in data['esal_flex'].items()}
            if data.get('traffic_df'):
                ss.traffic_df = pd.DataFrame(data['traffic_df'])
            st.success("✅ โหลดข้อมูลสำเร็จ!")
            st.rerun()
        except Exception as e:
            st.error(f"❌ โหลดไม่สำเร็จ: {e}")

    st.divider()
    st.markdown("""
    <div style='font-size:0.72rem;color:#81C784;text-align:center;line-height:1.8;'>
    รศ.ดร.อิทธิพล มีผล<br>
    ภาควิชาครุศาสตร์โยธา<br>
    คณะครุศาสตร์อุตสาหกรรม มจพ.<br>
    ITM Pave Pro v1.0
    </div>
    """, unsafe_allow_html=True)

# ─────────────────────────────────────────────
#  HEADER
# ─────────────────────────────────────────────
st.markdown("""
<div class="main-header">
    <h1>🛣️ ITM Pave Pro — ระบบออกแบบโครงสร้างชั้นทาง AASHTO 1993</h1>
    <p>ESAL Calculator · CBR Analysis · Flexible Design · K-Value Nomograph · Rigid Design · Report</p>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
#  MAIN TABS
# ─────────────────────────────────────────────
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "🚛 ESAL Calculator",
    "📊 CBR Analysis",
    "🔧 Flexible Design",
    "📐 K-Value Nomograph",
    "🏗️ Rigid Design",
    "📄 Report & Save",
])

# ══════════════════════════════════════════════
#  TAB 1: ESAL CALCULATOR
# ══════════════════════════════════════════════
with tab1:
    st.markdown("### 🚛 ESAL Calculator — AASHTO 1993")

    sub_esal_rigid, sub_esal_flex = st.tabs(["🔴 Rigid Pavement", "🟢 Flexible Pavement"])

    # ── Shared Traffic Input (above sub-tabs) ──
    with st.expander("📋 ข้อมูลปริมาณจราจร (ใช้ร่วมกันทั้ง Rigid & Flexible)", expanded=True):
        col_inp1, col_inp2 = st.columns([1, 1])

        with col_inp1:
            st.markdown('<div class="card"><h4>📁 Upload Excel / กรอกมือ</h4>', unsafe_allow_html=True)
            input_mode = st.radio("วิธีกรอกข้อมูล", ["📁 Upload Excel", "✏️ กรอกมือ + Growth Rate"], horizontal=True)

            if input_mode == "📁 Upload Excel":
                uploaded_xl = st.file_uploader("เลือกไฟล์ Excel (.xlsx)", type=['xlsx'])
                st.caption("รูปแบบ: คอลัมน์ Year, MB, HB, MT, HT, TR, STR")
                if uploaded_xl:
                    if not OPENPYXL_OK:
                        st.error("❌ openpyxl ไม่ได้ติดตั้ง — กรุณาใช้วิธี 'กรอกมือ + Growth Rate' แทน")
                    else:
                        try:
                            df_up = pd.read_excel(uploaded_xl, engine='openpyxl')
                            df_up.columns = [c.strip() for c in df_up.columns]
                            col_map = {}
                            for c in df_up.columns:
                                for vc in ['Year']+VEHICLE_COLS:
                                    if c.upper() == vc.upper():
                                        col_map[c] = vc
                            df_up = df_up.rename(columns=col_map)
                            for vc in VEHICLE_COLS:
                                if vc not in df_up.columns:
                                    df_up[vc] = 0
                            ss.traffic_df = df_up[['Year']+VEHICLE_COLS].fillna(0)
                            st.success(f"✅ อ่านข้อมูล {len(df_up)} ปีสำเร็จ")
                        except Exception as e:
                            st.error(f"❌ {e}")
            else:
                st.markdown("**ปริมาณจราจรปีแรก (คัน/วัน)**")
                base_cols = st.columns(6)
                base_row = {}
                defaults_base = {"MB":120,"HB":60,"MT":250,"HT":180,"TR":100,"STR":120}
                for i, vc in enumerate(VEHICLE_COLS):
                    with base_cols[i]:
                        base_row[vc] = st.number_input(vc, value=defaults_base[vc],
                                                       min_value=0, step=10, key=f"base_{vc}")
                gc1, gc2 = st.columns(2)
                with gc1:
                    growth_rate = st.number_input("Growth Rate (%/ปี)", value=4.5, step=0.5, min_value=0.0, max_value=20.0)
                with gc2:
                    design_years = st.number_input("Design Period (ปี)", value=20, min_value=1, max_value=40, step=1)
                if st.button("🔄 สร้างตารางจราจร", type="primary"):
                    ss.traffic_df = grow_traffic(base_row, growth_rate, int(design_years))
                    st.success(f"✅ สร้างตาราง {int(design_years)} ปีสำเร็จ")

            st.markdown('</div>', unsafe_allow_html=True)

        with col_inp2:
            if ss.traffic_df is not None:
                st.markdown('<div class="card"><h4>📊 ตารางปริมาณจราจร</h4>', unsafe_allow_html=True)
                st.dataframe(ss.traffic_df.style.format({c: "{:,.0f}" for c in VEHICLE_COLS}),
                             use_container_width=True, height=280)
                total_row = {vc: ss.traffic_df[vc].sum() for vc in VEHICLE_COLS}
                st.markdown(f'<div class="result-info">📊 รวมตลอดอายุออกแบบ: '
                            + " | ".join(f"<b>{vc}</b>: {total_row[vc]:,.0f}" for vc in VEHICLE_COLS)
                            + '</div>', unsafe_allow_html=True)
                st.markdown('</div>', unsafe_allow_html=True)
            else:
                st.info("⬅️ กรอกหรือ Upload ข้อมูลจราจรก่อน")

    # ─── Sub-tab: Rigid ───
    with sub_esal_rigid:
        st.markdown('<div class="card"><h4>⚙️ พารามิเตอร์ – Rigid Pavement</h4>', unsafe_allow_html=True)
        c1,c2,c3,c4 = st.columns(4)
        with c1: ldf_r = st.number_input("Lane Distribution Factor", value=0.9, step=0.05, min_value=0.1, max_value=1.0, key="ldf_r")
        with c2: ddf_r = st.number_input("Directional Dist. Factor", value=0.5, step=0.05, min_value=0.1, max_value=1.0, key="ddf_r")
        with c3: pt_r  = st.number_input("Terminal Serviceability Pt", value=2.5, step=0.1, min_value=1.5, max_value=3.5, key="pt_r")
        with c4: st.markdown(f"<br><small>Pi (Rigid) = 4.5</small>", unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

        # EALF Table
        st.markdown('<div class="card"><h4>📋 Truck Factor (EALF/คัน) ตาม Slab Thickness</h4>', unsafe_allow_html=True)
        tf_rows = []
        for vt in VEHICLE_COLS:
            row = {"ประเภทรถ": f"{VEHICLE_LABELS[vt]} ({vt})"}
            for D in SLAB_THICKNESSES:
                row[f"{D} cm"] = f"{truck_factor_rigid(vt,D,pt_r):.3f}"
            tf_rows.append(row)
        st.dataframe(pd.DataFrame(tf_rows), use_container_width=True, hide_index=True)
        st.markdown('</div>', unsafe_allow_html=True)

        if st.button("🔄 คำนวณ ESAL Rigid", type="primary", key="calc_r"):
            if ss.traffic_df is None:
                st.warning("⚠️ กรุณากรอกข้อมูลจราจรก่อน")
            else:
                esal_r = compute_esal_from_df(ss.traffic_df, ldf_r, ddf_r, pt_r, mode="rigid")
                ss.esal_rigid = esal_r
                ss.ldf = ldf_r; ss.ddf = ddf_r; ss.pt_rigid = pt_r

                st.markdown("---")
                st.markdown("#### 📊 ผลการคำนวณ ESAL – Rigid Pavement")
                cols_m = st.columns(len(SLAB_THICKNESSES))
                for i, D in enumerate(SLAB_THICKNESSES):
                    with cols_m[i]:
                        st.markdown(f"""
                        <div class="metric-box">
                            <div class="val">{esal_r[D]:,.0f}</div>
                            <div class="lbl">ESAL – Slab {D} cm</div>
                        </div>""", unsafe_allow_html=True)
                st.markdown('<div class="result-info">✅ ค่า ESAL บันทึกแล้ว → ใช้ได้ใน Tab K-Value และ Rigid Design</div>', unsafe_allow_html=True)

        if ss.esal_rigid:
            st.markdown("**ค่า ESAL Rigid ปัจจุบัน:**")
            df_er = pd.DataFrame({"Slab (cm)": list(ss.esal_rigid.keys()),
                                   "ESAL": [f"{v:,.0f}" for v in ss.esal_rigid.values()]})
            st.dataframe(df_er, use_container_width=True, hide_index=True)

    # ─── Sub-tab: Flexible ───
    with sub_esal_flex:
        st.markdown('<div class="card"><h4>⚙️ พารามิเตอร์ – Flexible Pavement</h4>', unsafe_allow_html=True)
        c1,c2,c3,c4 = st.columns(4)
        with c1: ldf_f = st.number_input("Lane Distribution Factor", value=0.9, step=0.05, min_value=0.1, max_value=1.0, key="ldf_f")
        with c2: ddf_f = st.number_input("Directional Dist. Factor", value=0.5, step=0.05, min_value=0.1, max_value=1.0, key="ddf_f")
        with c3: pt_f  = st.number_input("Terminal Serviceability Pt", value=2.5, step=0.1, min_value=1.5, max_value=3.5, key="pt_f")
        with c4: st.markdown(f"<br><small>Pi (Flexible) = 4.2</small>", unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown('<div class="card"><h4>📐 กำหนด Structure Number (SN)</h4>', unsafe_allow_html=True)
        sn_cols = st.columns(4)
        user_sn = []
        sn_defs = [6.5, 7.1, 7.5, 8.0]
        for i, col in enumerate(sn_cols):
            with col:
                user_sn.append(round(st.number_input(f"SN {i+1}", value=sn_defs[i],
                                min_value=1.0, max_value=20.0, step=0.1,
                                key=f"sn_{i}", format="%.1f"), 2))
        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown('<div class="card"><h4>📋 Truck Factor (EALF/คัน) ตาม SN</h4>', unsafe_allow_html=True)
        tf_rows_f = []
        for vt in VEHICLE_COLS:
            row = {"ประเภทรถ": f"{VEHICLE_LABELS[vt]} ({vt})"}
            for sn in user_sn:
                row[f"SN={sn}"] = f"{truck_factor_flex(vt,sn,pt_f):.3f}"
            tf_rows_f.append(row)
        st.dataframe(pd.DataFrame(tf_rows_f), use_container_width=True, hide_index=True)
        st.markdown('</div>', unsafe_allow_html=True)

        if st.button("🔄 คำนวณ ESAL Flexible", type="primary", key="calc_f"):
            if ss.traffic_df is None:
                st.warning("⚠️ กรุณากรอกข้อมูลจราจรก่อน")
            else:
                esal_fv = compute_esal_from_df(ss.traffic_df, ldf_f, ddf_f, pt_f, mode="flex", sn_list=user_sn)
                ss.esal_flex = esal_fv
                ss.sn_list = user_sn
                ss.pt_flex = pt_f

                st.markdown("---")
                st.markdown("#### 📊 ผลการคำนวณ ESAL – Flexible Pavement")
                cols_m2 = st.columns(len(user_sn))
                for i, sn in enumerate(user_sn):
                    with cols_m2[i]:
                        st.markdown(f"""
                        <div class="metric-box">
                            <div class="val">{esal_fv[sn]:,.0f}</div>
                            <div class="lbl">ESAL – SN {sn}</div>
                        </div>""", unsafe_allow_html=True)
                st.markdown('<div class="result-info">✅ ค่า ESAL บันทึกแล้ว → ใช้ได้ใน Tab Flexible Design</div>', unsafe_allow_html=True)

        if ss.esal_flex:
            st.markdown("**ค่า ESAL Flexible ปัจจุบัน:**")
            df_ef = pd.DataFrame({"SN": [f"SN {k}" for k in ss.esal_flex.keys()],
                                   "ESAL": [f"{v:,.0f}" for v in ss.esal_flex.values()]})
            st.dataframe(df_ef, use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════
#  TAB 2: CBR ANALYSIS
# ══════════════════════════════════════════════
with tab2:
    st.markdown("### 📊 CBR Analysis — Percentile Method")

    col_cbr_l, col_cbr_r = st.columns([1, 1])

    with col_cbr_l:
        st.markdown('<div class="card"><h4>📁 ข้อมูล CBR</h4>', unsafe_allow_html=True)
        cbr_mode = st.radio("แหล่งข้อมูล", ["📁 Upload Excel", "✏️ กรอกค่า", "📌 ใช้ข้อมูลตัวอย่าง"], horizontal=True)

        cbr_vals_input = None
        if cbr_mode == "📁 Upload Excel":
            cbr_xl = st.file_uploader("ไฟล์ Excel (คอลัมน์ CBR)", type=['xlsx'], key="cbr_xl")
            if cbr_xl:
                try:
                    df_cbr = pd.read_excel(cbr_xl, engine='openpyxl')
                    col_cbr = next((c for c in df_cbr.columns if 'cbr' in c.lower()), df_cbr.columns[0])
                    cbr_vals_input = pd.to_numeric(df_cbr[col_cbr], errors='coerce').dropna().tolist()
                    st.success(f"✅ {len(cbr_vals_input)} ตัวอย่าง")
                except Exception as e:
                    st.error(str(e))
        elif cbr_mode == "✏️ กรอกค่า":
            cbr_txt = st.text_area("กรอกค่า CBR (%) คั่นด้วย , หรือ Enter",
                                   placeholder="6.5, 7.2, 8.1, 5.3, ...",
                                   height=120)
            if cbr_txt.strip():
                import re
                parts = re.split(r'[,\n\r\s]+', cbr_txt.strip())
                try:
                    cbr_vals_input = [float(x) for x in parts if x]
                    st.success(f"✅ {len(cbr_vals_input)} ค่า")
                except:
                    st.error("กรุณากรอกตัวเลขเท่านั้น")
        else:
            cbr_vals_input = SAMPLE_CBR
            st.info(f"📌 ใช้ข้อมูลตัวอย่าง {len(SAMPLE_CBR)} ค่า")

        if cbr_vals_input:
            ss.cbr_values = cbr_vals_input

        target_pct = st.slider("Percentile ที่ต้องการ (%)", 50, 99, int(ss.cbr_percentile), step=1, key="pct_slider")
        ss.cbr_percentile = float(target_pct)
        st.markdown('</div>', unsafe_allow_html=True)

        if ss.cbr_values:
            arr, n, u_cbr, u_pct = calc_percentile_cbr(ss.cbr_values)
            cbr_at_pct = float(np.interp(target_pct, u_pct[::-1], u_cbr[::-1]))

            st.markdown('<div class="card"><h4>🎯 ค่า CBR ที่ใช้ออกแบบ</h4>', unsafe_allow_html=True)
            mr_auto = cbr_to_mr(cbr_at_pct)
            k_auto  = mr_to_k(mr_auto)

            c_m1, c_m2, c_m3 = st.columns(3)
            with c_m1:
                st.markdown(f"""<div class="metric-box">
                    <div class="val">{cbr_at_pct:.2f}</div>
                    <div class="lbl">CBR @ P{target_pct:.0f} (%)</div>
                </div>""", unsafe_allow_html=True)
            with c_m2:
                st.markdown(f"""<div class="metric-box">
                    <div class="val">{mr_auto:,.0f}</div>
                    <div class="lbl">Mr (psi) = 1500×CBR</div>
                </div>""", unsafe_allow_html=True)
            with c_m3:
                st.markdown(f"""<div class="metric-box">
                    <div class="val">{k_auto:.1f}</div>
                    <div class="lbl">k subgrade (pci)</div>
                </div>""", unsafe_allow_html=True)

            design_cbr_input = st.number_input(
                "CBR ที่ใช้ออกแบบจริง (ปรับได้)",
                value=float(round(cbr_at_pct, 1)),
                min_value=0.5, max_value=100.0, step=0.5,
                key="design_cbr_input"
            )
            mr_design = cbr_to_mr(design_cbr_input)
            k_design  = mr_to_k(mr_design)
            st.markdown(f"""
            <div class="result-info">
                CBR ออกแบบ = <b>{design_cbr_input:.1f}%</b> →
                Mr = <b>{mr_design:,.0f} psi</b> →
                k_subgrade = <b>{k_design:.1f} pci</b>
            </div>""", unsafe_allow_html=True)

            if st.button("✅ ใช้ค่านี้", type="primary", key="use_cbr"):
                ss.cbr_design      = design_cbr_input
                ss.mr_subgrade_psi = mr_design
                ss.k_subgrade_pci  = k_design
                st.success("✅ บันทึกค่า CBR/Mr/k แล้ว → ใช้ได้ใน Tab Flexible, K-Value, Rigid Design")
            st.markdown('</div>', unsafe_allow_html=True)

    with col_cbr_r:
        if ss.cbr_values:
            arr, n, u_cbr, u_pct = calc_percentile_cbr(ss.cbr_values)
            cbr_at_pct = float(np.interp(target_pct, u_pct[::-1], u_cbr[::-1]))

            st.markdown('<div class="card"><h4>📈 กราฟ Percentile vs CBR</h4>', unsafe_allow_html=True)
            if PLOTLY_OK:
                fig_cbr = go.Figure()
                fig_cbr.add_trace(go.Scatter(
                    x=u_cbr, y=u_pct, mode='lines+markers', name='CBR Distribution',
                    line=dict(color='#2E7D32', width=2.5),
                    marker=dict(size=7, symbol='x', color='#1B5E20')
                ))
                fig_cbr.add_trace(go.Scatter(
                    x=[0, cbr_at_pct], y=[target_pct, target_pct],
                    mode='lines', name=f'P{target_pct:.0f}%',
                    line=dict(color='red', width=2, dash='dash')
                ))
                fig_cbr.add_trace(go.Scatter(
                    x=[cbr_at_pct, cbr_at_pct], y=[0, target_pct],
                    mode='lines', name=f'CBR={cbr_at_pct:.2f}%',
                    line=dict(color='red', width=2, dash='dash')
                ))
                fig_cbr.add_annotation(
                    x=cbr_at_pct, y=0,
                    text=f"<b>{cbr_at_pct:.2f}%</b>",
                    showarrow=True, arrowhead=2, arrowcolor='red',
                    font=dict(size=14, color='red'), ay=40
                )
                fig_cbr.update_layout(
                    xaxis_title="CBR (%)", yaxis_title="Percentile (%)",
                    plot_bgcolor='white', height=380,
                    xaxis=dict(range=[0, max(u_cbr)*1.1], gridcolor='#E8F5E9'),
                    yaxis=dict(range=[0, 100], gridcolor='#E8F5E9'),
                    legend=dict(bgcolor='rgba(255,255,255,0.8)', bordercolor='#C8E6C9'),
                    margin=dict(l=50, r=30, t=30, b=50)
                )
                st.plotly_chart(fig_cbr, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)

            # Stats
            st.markdown('<div class="card"><h4>📋 สถิติ CBR</h4>', unsafe_allow_html=True)
            c1,c2,c3,c4 = st.columns(4)
            with c1: st.metric("n", n)
            with c2: st.metric("Min", f"{np.min(ss.cbr_values):.2f}%")
            with c3: st.metric("Max", f"{np.max(ss.cbr_values):.2f}%")
            with c4: st.metric("Mean", f"{np.mean(ss.cbr_values):.2f}%")
            st.markdown('</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════
#  TAB 3: FLEXIBLE DESIGN
# ══════════════════════════════════════════════
with tab3:
    st.markdown("### 🔧 Flexible Pavement Design — AASHTO 1993")

    col_fl, col_fr = st.columns([1, 1])

    with col_fl:
        # Auto-fill from session
        st.markdown('<div class="card"><h4>📥 Design ESAL</h4>', unsafe_allow_html=True)
        if ss.esal_flex:
            sn_keys = list(ss.esal_flex.keys())
            sel_idx = st.selectbox("เลือก SN", range(len(sn_keys)),
                                   format_func=lambda i: f"SN {sn_keys[i]}  →  ESAL = {ss.esal_flex[sn_keys[i]]:,.0f}",
                                   key="flex_sn_sel")
            design_esal_f = ss.esal_flex[sn_keys[sel_idx]]
            st.markdown(f'<div class="result-info">📊 Design ESAL = <b>{design_esal_f:,.0f}</b></div>', unsafe_allow_html=True)
        else:
            st.warning("⚠️ ยังไม่มีค่า ESAL — คำนวณใน Tab 1 ก่อน หรือกรอกเอง")
            design_esal_f = st.number_input("Design ESAL (กรอกเอง)", value=0, step=100000, key="flex_esal_manual")
        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown('<div class="card"><h4>🌍 Subgrade</h4>', unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        with c1:
            cbr_fl = st.number_input("CBR (%)", value=float(ss.cbr_design) if ss.cbr_design else 3.0,
                                     step=0.5, min_value=0.5, key="cbr_fl")
        with c2:
            mr_fl_auto = cbr_to_mr(cbr_fl)
            mr_fl = st.number_input("Mr (psi) [อัตโนมัติ]", value=float(ss.mr_subgrade_psi) if ss.mr_subgrade_psi else mr_fl_auto,
                                    step=500.0, min_value=500.0, key="mr_fl")
        st.markdown(f"Mr = **{mr_fl:,.0f} psi**  ({mr_fl/145.038:.1f} MPa)")
        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown('<div class="card"><h4>⚙️ Design Parameters</h4>', unsafe_allow_html=True)
        c1,c2,c3,c4 = st.columns(4)
        with c1:
            r0_fl = st.selectbox("Reliability R0 (%)", list(ZR_MAP.keys()), index=10, key="r0_fl")
            st.caption(f"ZR = {ZR_MAP[r0_fl]}")
        with c2: so_fl = st.number_input("So", value=0.45, step=0.01, min_value=0.3, max_value=0.6, key="so_fl")
        with c3: pi_fl = st.number_input("Pi", value=4.2, step=0.1, key="pi_fl")
        with c4: pt_fl2 = st.number_input("Pt", value=2.5, step=0.1, key="pt_fl2")
        st.markdown('</div>', unsafe_allow_html=True)

    with col_fr:
        st.markdown('<div class="card"><h4>🔩 Layer Design</h4>', unsafe_allow_html=True)
        hcols = st.columns([3, 1.5, 0.8, 0.8, 0.9, 1.1])
        for txt, col in zip(["**วัสดุ**","**หนา(cm)**","**ai**","**mi**","**SNi**","**ΣSNi**"], hcols):
            col.markdown(txt)

        mat_options = list(FLEX_LAYER_MATERIALS.keys())
        layer_results = []
        cum_sn = 0.0

        for li in range(6):
            lc0, lc1 = st.columns([3, 1.5])
            with lc0:
                mat_f = st.selectbox(f"L{li+1}", mat_options, key=f"fmat_{li}", label_visibility="collapsed")
            with lc1:
                h_f = st.number_input("cm", value=0, step=1, min_value=0, key=f"fh_{li}", label_visibility="collapsed")

            if mat_f != "None" and h_f > 0:
                ai, mi = FLEX_LAYER_MATERIALS[mat_f]
                h_in   = h_f / 2.54
                sn_i   = ai * h_in * mi
                cum_sn += sn_i
                layer_results.append({
                    'layer': li+1, 'material': mat_f,
                    'h_cm': h_f, 'ai': ai, 'mi': mi,
                    'sni': round(sn_i,3), 'cum_sn': round(cum_sn,3)
                })
                _, d1,d2,d3,d4,d5 = st.columns([3, 1.5, 0.8, 0.8, 0.9, 1.1])
                d1.markdown(f"`{h_f} cm`"); d2.markdown(f"`{ai:.2f}`")
                d3.markdown(f"`{mi:.1f}`"); d4.markdown(f"`{sn_i:.3f}`")
                d5.markdown(f"**`{cum_sn:.3f}`**")

        st.markdown(f"""<div class="result-info">
            ΣSN Provided = <b>{cum_sn:.3f}</b>
        </div>""", unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

        if st.button("✅ Design Check", type="primary", key="flex_check"):
            if design_esal_f <= 0:
                st.warning("⚠️ กรุณาใส่ Design ESAL")
            else:
                sn_req = aashto_sn_required(design_esal_f, ZR_MAP[r0_fl], so_fl, pi_fl, pt_fl2, mr_fl)
                if sn_req:
                    passed = cum_sn >= sn_req
                    margin = cum_sn - sn_req
                    css    = "result-pass" if passed else "result-fail"
                    chk    = "✅ PASS" if passed else "❌ FAIL"

                    c1,c2,c3 = st.columns(3)
                    with c1: st.markdown(f"""<div class="metric-box"><div class="val">{cum_sn:.3f}</div><div class="lbl">SN Provided</div></div>""", unsafe_allow_html=True)
                    with c2: st.markdown(f"""<div class="metric-box"><div class="val">{sn_req:.3f}</div><div class="lbl">SN Required</div></div>""", unsafe_allow_html=True)
                    with c3: st.markdown(f"""<div class="metric-box"><div class="val" style="color:{'#1B5E20' if passed else '#B71C1C'}">{margin:+.3f}</div><div class="lbl">Safety Margin</div></div>""", unsafe_allow_html=True)

                    st.markdown(f'<div class="{css}" style="margin-top:0.8rem;font-size:1.05rem">{chk} — SN Required = {sn_req:.3f} | SN Provided = {cum_sn:.3f}</div>', unsafe_allow_html=True)

                    ss.flex_results = {
                        'esal': design_esal_f, 'sn_req': sn_req,
                        'sn_prov': cum_sn, 'pass': passed,
                        'layers': layer_results,
                        'mr_psi': mr_fl, 'cbr': cbr_fl,
                    }
                    if layer_results:
                        st.dataframe(pd.DataFrame(layer_results), use_container_width=True, hide_index=True)
                else:
                    st.error("ไม่สามารถคำนวณ SN Required ได้ — ตรวจสอบ ESAL และ Mr")

# ══════════════════════════════════════════════
#  TAB 4: K-VALUE NOMOGRAPH
# ══════════════════════════════════════════════
with tab4:
    st.markdown("### 📐 K-Value Nomograph — AASHTO 1993")

    sub_kinf, sub_ls = st.tabs(["📊 Composite k∞", "📉 Loss of Support"])

    with sub_kinf:
        col_k1, col_k2 = st.columns([1, 1.4])
        with col_k1:
            st.markdown('<div class="card"><h4>📥 Input Parameters</h4>', unsafe_allow_html=True)
            # Auto-fill k_subgrade from CBR tab
            k_sub_auto = ss.k_subgrade_pci if ss.k_subgrade_pci else 231.9
            mr_auto_k  = ss.mr_subgrade_psi if ss.mr_subgrade_psi else 4500.0

            cbr_k = st.number_input("Subgrade CBR (%)", value=float(ss.cbr_design) if ss.cbr_design else 3.0,
                                    step=0.5, min_value=0.5, key="cbr_k")
            mr_k  = st.number_input("Mr Subgrade (psi)", value=float(mr_auto_k), step=500.0, min_value=100.0, key="mr_k")
            k_sub = st.number_input("k Subgrade (pci) = Mr/19.4",
                                    value=float(round(mr_k/19.4, 1)), step=5.0, min_value=1.0, key="k_sub_input")

            st.divider()
            esb_psi = st.number_input("Esb – Subbase Modulus (psi)", value=15000, step=1000, min_value=1000, max_value=100000, key="esb_psi")
            dsb_in  = st.number_input("DSB – Subbase Thickness (in)", value=6.0, step=0.5, min_value=0.0, max_value=36.0, key="dsb_in")
            st.markdown('</div>', unsafe_allow_html=True)

            if st.button("🔄 คำนวณ k∞", type="primary", key="calc_kinf"):
                fig_k, k_inf_val = draw_k_infinity_nomograph(esb_psi, dsb_in, k_sub)
                ss.k_inf = k_inf_val
                ss.nomograph_img_k = fig_to_bytes(fig_k)
                st.markdown(f"""<div class="result-pass">
                    k∞ = <b>{k_inf_val:.1f} pci</b><br>
                    ✅ บันทึกแล้ว → ใช้ใน Tab Loss of Support
                </div>""", unsafe_allow_html=True)
                plt.close(fig_k)

        with col_k2:
            if ss.nomograph_img_k:
                st.image(ss.nomograph_img_k, use_container_width=True)
                st.markdown(f'<div class="result-info">k∞ = <b>{ss.k_inf:.1f} pci</b></div>', unsafe_allow_html=True)
            else:
                st.info("กด 'คำนวณ k∞' เพื่อแสดง Nomograph")

    with sub_ls:
        col_ls1, col_ls2 = st.columns([1, 1.4])
        with col_ls1:
            st.markdown('<div class="card"><h4>📥 Loss of Support</h4>', unsafe_allow_html=True)
            k_inf_input = st.number_input("k∞ (pci) [อัตโนมัติจาก k∞ Tab]",
                                          value=float(ss.k_inf) if ss.k_inf else 200.0,
                                          step=10.0, min_value=10.0, max_value=3000.0, key="k_inf_ls")
            ls_opts = [0.0, 0.5, 1.0, 1.5, 2.0, 3.0]
            ls_sel  = st.select_slider("Loss of Support (LS)", options=ls_opts,
                                       value=ss.ls_value if ss.ls_value in ls_opts else 1.0,
                                       key="ls_sel")
            ls_desc = {0.0:"LS=0 (ไม่มีการสูญเสีย)", 0.5:"LS=0.5", 1.0:"LS=1.0 (ทั่วไป)",
                       1.5:"LS=1.5", 2.0:"LS=2.0 (ไม่มีแรงยึดประสาน)", 3.0:"LS=3.0 (สูงสุด)"}
            st.caption(ls_desc.get(ls_sel,""))
            st.markdown('</div>', unsafe_allow_html=True)

            if st.button("🔄 คำนวณ k_eff", type="primary", key="calc_ls"):
                fig_ls, k_corr = draw_loss_of_support_nomograph(k_inf_input, ls_sel)
                ss.k_corrected     = k_corr
                ss.ls_value        = ls_sel
                ss.nomograph_img_ls= fig_to_bytes(fig_ls)
                st.markdown(f"""<div class="result-pass">
                    k_eff = <b>{k_corr:.1f} pci</b><br>
                    (k∞={k_inf_input:.0f} pci, LS={ls_sel})<br>
                    ✅ บันทึกแล้ว → ใช้ใน Tab Rigid Design
                </div>""", unsafe_allow_html=True)
                plt.close(fig_ls)

        with col_ls2:
            if ss.nomograph_img_ls:
                st.image(ss.nomograph_img_ls, use_container_width=True)
                st.markdown(f'<div class="result-info">k_eff = <b>{ss.k_corrected:.1f} pci</b></div>', unsafe_allow_html=True)
            else:
                st.info("กด 'คำนวณ k_eff' เพื่อแสดง Nomograph")

# ══════════════════════════════════════════════
#  TAB 5: RIGID DESIGN
# ══════════════════════════════════════════════
with tab5:
    st.markdown("### 🏗️ Rigid Pavement Design — AASHTO 1993")

    # Auto-fill info
    k_eff_display = ss.k_corrected if ss.k_corrected else 0.0
    st.markdown(f"""
    <div class="result-info" style="margin-bottom:1rem;">
        📐 k_eff (จาก Tab K-Value) = <b>{k_eff_display:.1f} pci</b> &nbsp;|&nbsp;
        📊 ESAL Rigid พร้อม: {'✅' if ss.esal_rigid else '⚠️ ยังไม่มี'}
    </div>""", unsafe_allow_html=True)

    # Shared design parameters
    st.markdown('<div class="card"><h4>⚙️ พารามิเตอร์ร่วม (ใช้กับทุก Type)</h4>', unsafe_allow_html=True)
    rp1,rp2,rp3,rp4,rp5,rp6 = st.columns(6)
    with rp1: fc_cube = st.number_input("f'c (ksc)", value=350, step=10, min_value=200, key="fc_cube")
    with rp2:
        fc_cyl  = 0.8 * fc_cube
        fc_psi  = fc_cyl * 14.223
        ec_psi  = 57000 * math.sqrt(fc_psi)
        sc_auto = min(600, 10.0 * math.sqrt(fc_psi))
        sc_inp  = st.number_input("Sc (psi)", value=int(sc_auto), step=10, min_value=100, max_value=700, key="sc_inp")
    with rp3:
        r0_rig = st.selectbox("Reliability R0 (%)", list(ZR_MAP.keys()), index=10, key="r0_rig")
        zr_rig = ZR_MAP[r0_rig]
    with rp4: so_rig = st.number_input("So", value=0.35, step=0.01, min_value=0.2, max_value=0.5, key="so_rig")
    with rp5: pi_rig = st.number_input("Pi", value=4.5, step=0.1, key="pi_rig")
    with rp6: pt_rig2= st.number_input("Pt", value=2.5, step=0.1, key="pt_rig2")
    st.markdown(f"Ec = **{ec_psi:,.0f} psi** | f'c cylinder = **{fc_cyl:.0f} ksc** | ZR = **{zr_rig}**")
    st.markdown('</div>', unsafe_allow_html=True)

    sub_jpcp, sub_jrcp, sub_crcp = st.tabs(["🟦 JPCP", "🟧 JRCP", "🟥 CRCP"])

    def rigid_design_panel(ptype, tab_key):
        j_default = J_VALUES[ptype]
        cd_default = 1.0

        col_rd_l, col_rd_r = st.columns([1, 1])
        with col_rd_l:
            st.markdown(f'<div class="card"><h4>🔩 ชั้นโครงสร้าง – {ptype}</h4>', unsafe_allow_html=True)
            mat_opts_r = list(RIGID_LAYER_MATERIALS.keys())
            layer_r = []
            for li in range(4):
                lc_a, lc_b = st.columns([3, 1])
                with lc_a:
                    mat_r = st.selectbox(f"ชั้น {li+1}", mat_opts_r, key=f"rmat_{tab_key}_{li}")
                with lc_b:
                    if mat_r != "None":
                        h_r = st.number_input("cm", value=20, step=1, min_value=1, key=f"rh_{tab_key}_{li}", label_visibility="visible")
                        layer_r.append((mat_r, h_r, RIGID_LAYER_MATERIALS[mat_r]))
            st.markdown('</div>', unsafe_allow_html=True)

        with col_rd_r:
            st.markdown(f'<div class="card"><h4>⚙️ พารามิเตอร์ – {ptype}</h4>', unsafe_allow_html=True)
            c1,c2 = st.columns(2)
            with c1:
                j_val  = st.number_input(f"J ({ptype})", value=j_default, step=0.1, key=f"j_{tab_key}")
                cd_val = st.number_input("Cd (Drainage)", value=cd_default, step=0.05, min_value=0.5, max_value=1.25, key=f"cd_{tab_key}")
            with c2:
                d_sel  = st.selectbox("Slab Thickness (cm)", SLAB_THICKNESSES, index=1, key=f"d_{tab_key}")
                # Auto ESAL
                esal_auto = ss.esal_rigid.get(d_sel, 0)
                w18_req = st.number_input("W18 Design (ESAL)", value=int(esal_auto), step=100000, min_value=0, key=f"w18_{tab_key}")

            # k_eff input
            k_eff_inp = st.number_input("k_eff (pci) [อัตโนมัติจาก Tab K-Value]",
                                        value=float(k_eff_display) if k_eff_display > 0 else 200.0,
                                        step=10.0, min_value=10.0, key=f"keff_{tab_key}")
            st.markdown('</div>', unsafe_allow_html=True)

            if st.button(f"✅ Design Check – {ptype}", type="primary", key=f"dc_{tab_key}"):
                w18_cap = aashto_rigid_w18(d_sel, pi_rig, pt_rig2, zr_rig, so_rig,
                                           sc_inp, cd_val, j_val, ec_psi, k_eff_inp)
                if w18_cap is None:
                    st.error("ไม่สามารถคำนวณได้ — ตรวจสอบพารามิเตอร์")
                else:
                    passed  = w18_cap >= w18_req
                    margin  = (w18_cap/w18_req - 1)*100 if w18_req > 0 else float('inf')
                    css     = "result-pass" if passed else "result-fail"
                    chk     = "✅ PASS" if passed else "❌ FAIL"

                    c1,c2,c3 = st.columns(3)
                    with c1: st.markdown(f"""<div class="metric-box"><div class="val">{w18_cap:,.0f}</div><div class="lbl">W18 Capacity</div></div>""", unsafe_allow_html=True)
                    with c2: st.markdown(f"""<div class="metric-box"><div class="val">{w18_req:,.0f}</div><div class="lbl">W18 Required</div></div>""", unsafe_allow_html=True)
                    with c3: st.markdown(f"""<div class="metric-box"><div class="val" style="color:{'#1B5E20' if passed else '#B71C1C'}">{margin:+.1f}%</div><div class="lbl">Safety Margin</div></div>""", unsafe_allow_html=True)

                    st.markdown(f'<div class="{css}" style="margin-top:0.8rem;font-size:1.05rem">{chk}<br>Slab {d_sel} cm | k_eff = {k_eff_inp:.0f} pci | J = {j_val}</div>', unsafe_allow_html=True)

                    if 'rigid_results' not in ss or not isinstance(ss.rigid_results, dict):
                        ss.rigid_results = {}
                    ss.rigid_results[ptype] = {
                        'd_cm': d_sel, 'k_eff': k_eff_inp, 'fc': fc_cube,
                        'sc': sc_inp, 'j': j_val, 'cd': cd_val,
                        'w18_cap': w18_cap, 'w18_req': w18_req,
                        'pass': passed, 'margin_pct': margin,
                    }

    with sub_jpcp: rigid_design_panel("JPCP", "jpcp")
    with sub_jrcp: rigid_design_panel("JRCP", "jrcp")
    with sub_crcp: rigid_design_panel("CRCP", "crcp")

# ══════════════════════════════════════════════
#  TAB 6: REPORT & SAVE
# ══════════════════════════════════════════════
with tab6:
    st.markdown("### 📄 Report & Save — Word / JSON")

    if not DOCX_OK:
        st.warning("⚠️ ไม่พบ python-docx — ติดตั้งด้วย `pip install python-docx`")

    st.markdown('<div class="card"><h4>📋 สถานะข้อมูลแต่ละส่วน</h4>', unsafe_allow_html=True)
    sc1,sc2,sc3,sc4,sc5 = st.columns(5)
    with sc1: st.markdown(status_badge('esal_rigid','ESAL Rigid'), unsafe_allow_html=True)
    with sc2: st.markdown(status_badge('cbr_values','CBR Analysis'), unsafe_allow_html=True)
    with sc3: st.markdown(status_badge('flex_results','Flexible Design'), unsafe_allow_html=True)
    with sc4: st.markdown(status_badge('k_corrected','K-Value'), unsafe_allow_html=True)
    with sc5: st.markdown(status_badge('rigid_results','Rigid Design'), unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("#### 📥 Download รายงาน Word")

    # Project info
    col_pi1, col_pi2 = st.columns(2)
    with col_pi1: proj_name = st.text_input("ชื่อโครงการ", value="", placeholder="กรอกชื่อโครงการ...", key="proj_name")
    with col_pi2: eng_name  = st.text_input("ผู้ออกแบบ", value="รศ.ดร.อิทธิพล มีผล", key="eng_name")

    st.markdown("**เลือกส่วนที่ต้องการ Report:**")
    r1,r2,r3,r4,r5 = st.columns(5)
    with r1: chk_esal  = st.checkbox("🚛 ESAL", value=True)
    with r2: chk_cbr   = st.checkbox("📊 CBR Analysis", value=True)
    with r3: chk_flex  = st.checkbox("🔧 Flexible Design", value=True)
    with r4: chk_kval  = st.checkbox("📐 K-Value", value=True)
    with r5: chk_rigid = st.checkbox("🏗️ Rigid Design", value=True)

    col_dl1, col_dl2 = st.columns(2)

    def get_ss_dict():
        return {k: ss.get(k) for k in [
            'esal_rigid','esal_flex','ldf','ddf','pt_rigid','pt_flex',
            'cbr_values','cbr_percentile','cbr_design','mr_subgrade_psi',
            'k_subgrade_pci','flex_results','k_inf','k_corrected','ls_value',
            'nomograph_img_k','nomograph_img_ls','rigid_results'
        ]}

    with col_dl1:
        st.markdown("**แยกส่วน:**")
        ss_d = get_ss_dict()

        if chk_esal and ss.esal_rigid:
            b = build_report_esal(ss_d)
            if b: st.download_button("📥 ESAL Report", b, "ESAL_Report.docx",
                                     mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                     use_container_width=True)

        if chk_cbr and ss.cbr_values:
            b = build_report_cbr(ss_d)
            if b: st.download_button("📥 CBR Report", b, "CBR_Report.docx",
                                     mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                     use_container_width=True)

        if chk_flex and ss.flex_results:
            b = build_report_flexible(ss_d)
            if b: st.download_button("📥 Flexible Report", b, "Flexible_Report.docx",
                                     mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                     use_container_width=True)

        if chk_kval and ss.k_corrected:
            b = build_report_kvalue(ss_d)
            if b: st.download_button("📥 K-Value Report", b, "KValue_Report.docx",
                                     mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                     use_container_width=True)

        if chk_rigid and ss.rigid_results:
            b = build_report_rigid(ss_d)
            if b: st.download_button("📥 Rigid Report", b, "Rigid_Report.docx",
                                     mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                     use_container_width=True)

    with col_dl2:
        st.markdown("**รวมทุกส่วน:**")
        if st.button("🗂️ สร้างรายงานรวม", type="primary", use_container_width=True):
            ss_d = get_ss_dict()
            b_full = build_report_full(ss_d)
            if b_full:
                st.download_button("📥 Download Full Report", b_full,
                                   f"ITM_Pave_Full_Report_{datetime.now().strftime('%Y%m%d')}.docx",
                                   mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                   use_container_width=True)
            else:
                st.warning("ไม่มีข้อมูลสำหรับสร้างรายงาน หรือ python-docx ไม่พร้อม")

    st.markdown("---")
    st.markdown("#### 💾 JSON Save / Load (Sidebar)")
    st.info("💡 ใช้ปุ่ม Save/Load JSON ใน Sidebar ด้านซ้ายเพื่อบันทึกและโหลดโปรเจกต์ทั้งหมด")

    st.divider()
    st.markdown("""
    <div style='text-align:center;color:#558B2F;font-size:0.85rem;padding:0.5rem;'>
        🛣️ <b>ITM Pave Pro v1.0</b> — AASHTO 1993 Pavement Design System<br>
        พัฒนาโดย รศ.ดร.อิทธิพล มีผล | ภาควิชาครุศาสตร์โยธา | คณะครุศาสตร์อุตสาหกรรม | มจพ.
    </div>
    """, unsafe_allow_html=True)
