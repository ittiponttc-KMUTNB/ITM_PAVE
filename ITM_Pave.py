import streamlit as st
import pandas as pd
import numpy as np
import math

# scipy brentq — with pure-Python fallback (bisection) for Streamlit Cloud
try:
    from scipy.optimize import brentq as _brentq
except ImportError:
    def _brentq(f, a, b, xtol=1e-6, maxiter=500):
        """Pure-Python bisection fallback when scipy is unavailable."""
        fa, fb = f(a), f(b)
        if fa * fb > 0:
            raise ValueError("f(a) and f(b) must have opposite signs")
        for _ in range(maxiter):
            mid = (a + b) / 2.0
            fm  = f(mid)
            if abs(fm) < xtol or (b - a) / 2.0 < xtol:
                return mid
            if fa * fm < 0:
                b, fb = mid, fm
            else:
                a, fa = mid, fm
        return (a + b) / 2.0

# ─────────────────────────────────────────────
#  PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="FR Pave – AASHTO 1993 Pavement Design",
    page_icon="🛣️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
#  GLOBAL CSS
# ─────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;600;700&display=swap');

html, body, [class*="css"] {
    font-family: 'Sarabun', sans-serif;
}

/* Header */
.main-header {
    background: linear-gradient(135deg, #1a3a5c 0%, #2e6da4 100%);
    color: white;
    padding: 1.2rem 2rem;
    border-radius: 12px;
    margin-bottom: 1.5rem;
    box-shadow: 0 4px 15px rgba(26,58,92,0.3);
}
.main-header h1 { margin: 0; font-size: 1.8rem; font-weight: 700; }
.main-header p  { margin: 0.2rem 0 0; font-size: 0.95rem; opacity: 0.85; }

/* Section cards */
.section-card {
    background: #f7fafd;
    border: 1px solid #d0e4f5;
    border-left: 4px solid #2e6da4;
    border-radius: 8px;
    padding: 1rem 1.2rem;
    margin-bottom: 1rem;
}
.section-card h4 { color: #1a3a5c; margin: 0 0 0.8rem; font-size: 1rem; font-weight: 600; }

/* Result cards */
.result-pass {
    background: #e8f5e9; border: 1px solid #a5d6a7;
    border-radius: 8px; padding: 0.7rem 1rem; margin: 0.3rem 0;
    color: #2e7d32; font-weight: 600;
}
.result-fail {
    background: #ffebee; border: 1px solid #ef9a9a;
    border-radius: 8px; padding: 0.7rem 1rem; margin: 0.3rem 0;
    color: #c62828; font-weight: 600;
}
.result-info {
    background: #e3f2fd; border: 1px solid #90caf9;
    border-radius: 8px; padding: 0.7rem 1rem; margin: 0.3rem 0;
    color: #1565c0; font-weight: 600;
}
.result-warn {
    background: #fff8e1; border: 1px solid #ffe082;
    border-radius: 8px; padding: 0.7rem 1rem; margin: 0.3rem 0;
    color: #e65100; font-weight: 600;
}

/* Metric box */
.metric-box {
    background: white;
    border: 1px solid #d0e4f5;
    border-radius: 10px;
    padding: 1rem;
    text-align: center;
    box-shadow: 0 2px 8px rgba(0,0,0,0.06);
}
.metric-box .val { font-size: 1.6rem; font-weight: 700; color: #1a3a5c; }
.metric-box .lbl { font-size: 0.8rem; color: #5c7a99; margin-top: 0.2rem; }

/* Tab style */
[data-baseweb="tab-list"] { gap: 4px; }
[data-baseweb="tab"] {
    background: #e8f0f8 !important;
    border-radius: 8px 8px 0 0 !important;
    font-weight: 600 !important;
    color: #1a3a5c !important;
    padding: 0.5rem 1rem !important;
}
[aria-selected="true"][data-baseweb="tab"] {
    background: #2e6da4 !important;
    color: white !important;
}

/* Sidebar */
[data-testid="stSidebar"] { background: #1a3a5c; }
[data-testid="stSidebar"] * { color: #e8f0f8 !important; }
[data-testid="stSidebar"] .stSelectbox label,
[data-testid="stSidebar"] .stRadio label { color: #b0c8e0 !important; }

/* Divider */
hr { border-color: #d0e4f5; }

/* Number input */
.stNumberInput > div > div > input {
    font-family: 'Courier New', monospace;
    font-weight: 600;
}

/* DataFrame */
.stDataFrame { border-radius: 8px; overflow: hidden; }

button[kind="primary"] {
    background: #2e6da4 !important;
    border-radius: 8px !important;
    font-weight: 700 !important;
}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
#  LOOKUP DATA  (AASHTO 1993)
# ─────────────────────────────────────────────

# ─────────────────────────────────────────────
#  VEHICLE AXLE CONFIGURATIONS
#  Each axle group: (L1_ton, L2_axle_code, count)
#  L2: 1=single, 2=tandem, 3=tridem
#  Loads from Thai DOH standard / FR Pave V1.0
# ─────────────────────────────────────────────
VEHICLE_AXLES = {
    #         axle groups: [(L1_ton, L2, count), ...]
    "MB":  [(4,  1, 1), (11, 1, 1)],          # steer 4t + rear single 11t
    "HB":  [(5,  1, 1), (20, 2, 1)],          # steer 5t + drive tandem 20t
    "MT":  [(4,  1, 1), (11, 1, 1)],          # same as MB
    "HT":  [(5,  1, 1), (20, 2, 1)],          # same as HB
    "TR":  [(5,  1, 1), (20, 2, 1), (11, 1, 2)],  # steer 5t + drive tandem 20t + 2 trailer singles 11t
    "STR": [(5,  1, 1), (20, 2, 2)],          # steer + 2 drive tandems
}

VEHICLE_LABELS = {
    "MB":  "Medium Bus (MB)",
    "HB":  "Heavy Bus (HB)",
    "MT":  "Medium Truck (MT)",
    "HT":  "Heavy Truck (HT)",
    "TR":  "Trailer (TR)",
    "STR": "Semi Trailer (STR)",
}

# Lookup Table at Pt=2.5 (FR Pave V1.0 reference) — used only for display comparison
RIGID_TF_REF = {
    "MB":  {25:3.63, 28:3.70, 30:3.72, 32:3.73, 35:3.74},
    "HB":  {25:5.86, 28:6.07, 30:6.15, 32:6.20, 35:6.24},
    "MT":  {25:3.63, 28:3.70, 30:3.72, 32:3.73, 35:3.74},
    "HT":  {25:5.86, 28:6.07, 30:6.15, 32:6.20, 35:6.24},
    "TR":  {25:13.03, 28:13.37, 30:13.50, 32:13.57, 35:13.63},
    "STR": {25:11.60, 28:12.02, 30:12.18, 32:12.27, 35:12.34},
}
FLEX_TF_REF = {
    "MB":  {6.5:3.55, 7.1:3.65, 7.5:3.70, 8.0:3.75},
    "HB":  {6.5:3.36, 7.1:3.42, 7.5:3.45, 8.0:3.47},
    "MT":  {6.5:3.55, 7.1:3.65, 7.5:3.70, 8.0:3.75},
    "HT":  {6.5:3.36, 7.1:3.42, 7.5:3.45, 8.0:3.47},
    "TR":  {6.5:10.36, 7.1:10.62, 7.5:10.76, 8.0:10.89},
    "STR": {6.5:6.60, 7.1:6.72, 7.5:6.78, 8.0:6.83},
}

# ─────────────────────────────────────────────
#  AASHTO 1993 EALF ENGINE
# ─────────────────────────────────────────────
TON_TO_KIP = 2.2046

def ealf_flex_axle(L1_ton: float, L2: int, SN: float, Pt: float) -> float:
    """
    EALF for one axle group — Flexible pavement (AASHTO 1993 App. D).

    สมการ: log(Wtx/Wt18) = 4.79*log(19) - 4.79*log(Lx+L2) + 4.33*log(L2) + Gt/Bx - Gt/B18
    → log(EALF) = 4.79*log(Lx+L2) - 4.33*log(L2) - 4.79*log(19) + Gt*(1/B18 - 1/Bx)

    L1_ton : axle group load (metric ton)
    L2     : 1=single, 2=tandem, 3=tridem
    SN     : Structural Number
    Pt     : Terminal Serviceability
    """
    L1  = L1_ton * TON_TO_KIP                          # ton → kips
    Gt  = math.log10((4.2 - Pt) / (4.2 - 1.5))        # = log10(ΔPSI/2.7)
    Bx  = 0.40 + 0.081*(L1 + L2)**3.23 / ((SN+1)**5.19 * L2**3.23)
    B18 = 0.40 + 0.081*(18  + 1)**3.23 / ((SN+1)**5.19 * 1.0**3.23)
    log_ealf = (4.79*math.log10(L1 + L2)
                - 4.33*math.log10(L2)
                - 4.79*math.log10(19)
                + Gt * (1/B18 - 1/Bx))
    return 10**log_ealf

def ealf_rigid_axle(L1_ton: float, L2: int, D_cm: float, Pt: float) -> float:
    """
    EALF for one axle group — Rigid pavement (AASHTO 1993 App. D).

    สมการจากภาพ:
    log(Wtx/Wt18) = 4.62*log(19) - 4.62*log(Lx+L2) + 3.28*log(L2) + Gt/Bx - Gt/B18
    → log(EALF) = 4.62*log(Lx+L2) - 3.28*log(L2) - 4.62*log(19) + Gt*(1/B18 - 1/Bx)

    L1_ton : axle group load (metric ton)
    L2     : 1=single, 2=tandem, 3=tridem
    D_cm   : slab thickness (cm)
    Pt     : Terminal Serviceability
    """
    L1 = L1_ton * TON_TO_KIP                           # ton → kips
    D  = D_cm / 2.54                                   # cm → inches
    Gt  = math.log10((4.5 - Pt) / (4.5 - 1.5))        # = log10(ΔPSI/3.0)
    Bx  = 1.0 + 3.63*(L1 + L2)**5.20 / ((D+1)**8.46 * L2**3.52)
    B18 = 1.0 + 3.63*(18  + 1)**5.20 / ((D+1)**8.46 * 1.0**3.52)
    log_ealf = (4.62*math.log10(L1 + L2)
                - 3.28*math.log10(L2)
                - 4.62*math.log10(19)
                + Gt * (1/B18 - 1/Bx))
    return 10**log_ealf

def truck_factor_flex(vtype: str, SN: float, Pt: float) -> float:
    """Truck Factor = sum of EALF over all axle groups for one vehicle."""
    return sum(
        ealf_flex_axle(L1, L2, SN, Pt) * cnt
        for L1, L2, cnt in VEHICLE_AXLES[vtype]
    )

def truck_factor_rigid(vtype: str, D_cm: float, Pt: float) -> float:
    """Truck Factor = sum of EALF over all axle groups for one vehicle."""
    return sum(
        ealf_rigid_axle(L1, L2, D_cm, Pt) * cnt
        for L1, L2, cnt in VEHICLE_AXLES[vtype]
    )

def ealf_table_flex(vtype: str, sn_list: list, Pt: float) -> dict:
    """Return {SN: TF} for display."""
    return {sn: truck_factor_flex(vtype, sn, Pt) for sn in sn_list}

def ealf_table_rigid(vtype: str, thick_list: list, Pt: float) -> dict:
    """Return {D_cm: TF} for display."""
    return {d: truck_factor_rigid(vtype, d, Pt) for d in thick_list}

# Layer material library
RIGID_LAYER_MATERIALS = {
    "None": None,
    "AC under Concrete Pavement": 360000,
    "Lean Concrete Base (LCB)": 700000,
    "Cement Modified Crush Rock Base, UCS 24.5 ksc (min)": 120000,
    "Cement Modified Crush Rock Base, UCS 17.5 ksc (min)": 80000,
    "Crush Rock Base, CBR 80% (min)": 50750,
    "Soil Aggregate Subbase, CBR 25% (min)": 21750,
    "Soil Aggregate Subbase, CBR 20% (min)": 17400,
    "Soil Aggregate Subbase, CBR 15% (min)": 13050,
    "Sand Embankment, CBR 10% (min)": 14500,
    "Cruched Rock Under Concrete Pavement, CBR 80%": 50750,
}

FLEX_LAYER_MATERIALS = {
    "None": (None, None),
    "Asphalt Concrete": (0.40, 1),
    "Cement Modified Crush Rock Base, UCS 24.5 ksc (min)": (0.15, 1),
    "Cement Modified Crush Rock Base, UCS 17.5 ksc (min)": (0.13, 1),
    "Crush Rock Base, CBR 80% (min)": (0.14, 1),
    "Soil Aggregate Subbase, CBR 25% (min)": (0.10, 1),
    "Soil Aggregate Subbase, CBR 20% (min)": (0.09, 1),
    "Soil Aggregate Subbase, CBR 15% (min)": (0.08, 1),
    "Sand Embankment, CBR 10% (min)": (0.08, 1),
}

SLAB_THICKNESSES = [25, 28, 30, 32, 35]
SN_VALUES        = [6.5, 7.1, 7.5, 8.0]


# ─────────────────────────────────────────────
#  ENGINE FUNCTIONS
# ─────────────────────────────────────────────

def cbr_to_mr(cbr: float) -> float:
    """Convert CBR to Resilient Modulus (psi) — AASHTO 1993"""
    return 1500.0 * cbr

def compute_esal_rigid(vehicles: dict, ldf: float, ddf: float, Pt: float) -> dict:
    """Compute ESAL for rigid pavement per slab thickness using AASHTO equation."""
    results = {}
    for thick in SLAB_THICKNESSES:
        esal = 0.0
        for vtype, count in vehicles.items():
            if count <= 0:
                continue
            tf = truck_factor_rigid(vtype, thick, Pt)
            esal += count * tf * ldf * ddf
        results[thick] = esal
    return results

def compute_tf_rigid(vehicles: dict, Pt: float) -> dict:
    """Weighted average Truck Factor per slab thickness."""
    results = {}
    total_veh = sum(v for v in vehicles.values() if v > 0)
    for thick in SLAB_THICKNESSES:
        total_tf = sum(
            count * truck_factor_rigid(vtype, thick, Pt)
            for vtype, count in vehicles.items() if count > 0
        )
        results[thick] = total_tf / total_veh if total_veh > 0 else 0.0
    return results

def compute_esal_flexible(vehicles: dict, ldf: float, ddf: float,
                          sn_list: list, Pt: float) -> dict:
    """Compute ESAL for flexible pavement per user-defined SN list."""
    results = {}
    for sn in sn_list:
        esal = 0.0
        for vtype, count in vehicles.items():
            if count <= 0:
                continue
            tf = truck_factor_flex(vtype, sn, Pt)
            esal += count * tf * ldf * ddf
        results[sn] = esal
    return results

def compute_tf_flexible(vehicles: dict, sn_list: list, Pt: float) -> dict:
    """Weighted average Truck Factor per SN."""
    results = {}
    total_veh = sum(v for v in vehicles.values() if v > 0)
    for sn in sn_list:
        total_tf = sum(
            count * truck_factor_flex(vtype, sn, Pt)
            for vtype, count in vehicles.items() if count > 0
        )
        results[sn] = total_tf / total_veh if total_veh > 0 else 0.0
    return results


def aashto_lef_single(W, pt=2.5, sn_or_d=None, pave_type="flexible"):
    """AASHTO 1993 Load Equivalency Factor via equation (single axle)."""
    L1 = W  # axle load in kips (single)
    L2 = 1  # axle code
    if pave_type == "flexible":
        SN = sn_or_d
        Gt = math.log10((4.2 - pt) / (4.2 - 1.5))
        beta = 0.4 + (0.081*(L1+L2)**3.23) / ((SN+1)**5.19 * L2**3.23)
        Lx = (L1/(L2**0.5))
        lef = (10**(beta * Gt / (10**(-0.255)))) * (L1/(18))**4.79 * L2**(-4.33)
        # Simplified AASHTO equation
        lef = (L1/18)**4 * 10**(4.79*math.log10(L1/18))
        lef = max(lef, 0.0001)
    else:
        D = sn_or_d  # slab thickness inches
        Gt = math.log10((4.5 - pt) / (4.5 - 1.5))
        delta_psi = 4.5 - pt
        beta18 = 1.0 + (3.63*(18+L2)**5.20) / ((D+1)**8.46 * L2**3.52)
        betax  = 1.0 + (3.63*(L1+L2)**5.20) / ((D+1)**8.46 * L2**3.52)
        lef = 10**(Gt*(1/betax - 1/beta18))
        lef = max(lef, 0.0001)
    return lef

def compute_keff_odemark(layer_stack, mr_subgrade):
    """
    Odemark equivalent thickness method to compute keff.
    layer_stack: list of (thickness_cm, Mr_psi)
    Returns keff in pci.
    """
    # Convert cm to inches
    Es = mr_subgrade  # subgrade Mr (psi)
    # Compute equivalent subbase thickness (inches)
    h_eq = 0.0
    for (h_cm, mr_layer) in layer_stack:
        h_in = h_cm / 2.54
        h_eq += h_in * (mr_layer / Es) ** (1/3)

    # Compute k from Mr using AASHTO correlation
    k_subgrade = Es / 19.4  # pci (approximate)

    # Correction for equivalent thickness
    # Using simplified Westergaard/AASHTO nomograph approximation
    if h_eq <= 0:
        return k_subgrade

    # Interpolation from AASHTO Figure 3.3 approximation
    k_corr = k_subgrade * (1 + 0.64 * h_eq**0.5)
    return min(k_corr, 3000)

def compute_keff_from_subbase_mr(esb_mr, k_subgrade):
    """
    Compute keff from equivalent subbase modulus using AASHTO Figure 3.3
    approximation.
    """
    if esb_mr <= 0:
        return k_subgrade
    ratio = esb_mr / 19.4  # pci
    k_corr = k_subgrade * (esb_mr / (19.4 * k_subgrade)) ** 0.33
    k_corr = min(k_corr, 3000)
    return k_corr

def compute_sn_required(esal, r0_pct, so, pi, pt):
    """Helper – returns ZR for a given reliability (kept for backward compat)."""
    ZR_map = {50: 0.0, 60: -0.253, 70: -0.524, 75: -0.674,
              80: -0.841, 85: -1.037, 90: -1.282,
              91: -1.340, 92: -1.405, 93: -1.476,
              94: -1.555, 95: -1.645, 96: -1.751,
              97: -1.881, 98: -2.054, 99: -2.327}
    return ZR_map.get(int(r0_pct), -1.282)

def aashto_sn_required_flex(esal, zr, so, pi, pt, mr_psi):
    """Return required SN given design inputs (AASHTO 1993 flexible).
    Ref: AASHTO Guide for Design of Pavement Structures 1993, Eq. 3.1
    log10(W18) = ZR*So + 9.36*log10(SN+1) - 0.20
               + log10(ΔPSI/2.7) / (0.40 + 1094/(SN+1)^5.19)
               + 2.32*log10(Mr) - 8.07
    where ΔPSI = Pi - Pt, and 2.7 = (4.2 - 1.5)
    """
    delta_psi = pi - pt
    logW18 = math.log10(max(esal, 1))

    def equation(SN):
        if SN <= 0:
            return -1e10
        term1 = zr * so
        term2 = 9.36 * math.log10(SN + 1) - 0.20
        term3 = math.log10(delta_psi / 2.7) / (0.40 + 1094 / (SN + 1)**5.19)
        term4 = 2.32 * math.log10(mr_psi) - 8.07
        return term1 + term2 + term3 + term4 - logW18

    try:
        sn = _brentq(equation, 0.1, 30, xtol=1e-4)
    except Exception:
        sn = None
    return sn

def aashto_keff_rigid(esal, zr, so, pi, pt, ec_psi, sc_psi, j, cd, d_in):
    """
    AASHTO 1993 Rigid pavement: solve for keff given D (slab thickness).
    log10(W18) = ZR*So + 7.35*log10(D+1) - 0.06
                 + log10(ΔPSI/4.5-1.5)/(1+1.624e7/(D+1)^8.46)
                 + (4.22-0.32pt)*log10(Sc*Cd*(D^0.75 - 1.132) /
                   (215.63*J*(D^0.75 - 18.42/(Ec/k)^0.25)))
    Solve for k (keff).
    """
    delta_psi = pi - pt
    logW18 = math.log10(max(esal, 1))

    def equation(k):
        if k <= 0:
            return -1e10
        try:
            term1 = zr * so
            term2 = 7.35 * math.log10(d_in + 1) - 0.06
            # AASHTO 1993 Eq.3.28: denominator = (4.5-1.5) = 3.0
            term3 = math.log10(delta_psi / 3.0) / (1 + 1.624e7 / (d_in + 1)**8.46)
            inner = (sc_psi * cd * (d_in**0.75 - 1.132) /
                     (215.63 * j * (d_in**0.75 - 18.42 / (ec_psi / k)**0.25)))
            if inner <= 0:
                return -1e10
            term4 = (4.22 - 0.32 * pt) * math.log10(inner)
            return term1 + term2 + term3 + term4 - logW18
        except Exception:
            return -1e10

    try:
        k = _brentq(equation, 1, 3000, xtol=0.1)
    except Exception:
        k = 3000  # cap
    return k

ZR_MAP = {
    50: 0.0, 60: -0.253, 70: -0.524, 75: -0.674,
    80: -0.841, 85: -1.037, 90: -1.282,
    91: -1.340, 92: -1.405, 93: -1.476,
    94: -1.555, 95: -1.645, 96: -1.751,
    97: -1.881, 98: -2.054, 99: -2.327
}

# ─────────────────────────────────────────────
#  SESSION STATE INIT
# ─────────────────────────────────────────────
def init_state():
    defaults = {
        "esal_rigid": {25: 0, 28: 0, 30: 0, 32: 0, 35: 0},
        "esal_flex":  {},
        "design_pt":  2.5,
        "lef_mode":   "Lookup Table",
        "user_sn_values": [6.5, 7.1, 7.5, 8.0],
        "keff_jpcp":  {25: 0, 28: 0, 30: 0, 32: 0, 35: 0},
        "keff_crcp":  {25: 0, 28: 0, 30: 0, 32: 0, 35: 0},
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

init_state()

# ─────────────────────────────────────────────
#  SIDEBAR
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown("### ⚙️ ตั้งค่าทั่วไป")
    lef_mode = st.radio(
        "วิธีคำนวณ Truck Factor / LEF",
        ["Lookup Table", "AASHTO Equation"],
        index=0 if st.session_state["lef_mode"] == "Lookup Table" else 1,
        help="Lookup Table: ค่าสำเร็จรูปจากตาราง AASHTO\nAASHTO Equation: คำนวณจากสมการจริง"
    )
    st.session_state["lef_mode"] = lef_mode

    st.divider()
    st.markdown("### 📊 ESAL ที่คำนวณได้")

    er = st.session_state["esal_rigid"]
    ef = st.session_state["esal_flex"]

    st.markdown("**Rigid Pavement**")
    for t, v in er.items():
        st.markdown(f"- Slab {t} cm: **{v:,.0f}**")

    st.markdown("**Flexible Pavement**")
    user_sn_saved = st.session_state.get("user_sn_values", SN_VALUES)
    for sn, v in ef.items():
        st.markdown(f"- SN {sn}: **{v:,.0f}**")

    st.divider()
    st.markdown("""
    <small style='opacity:0.7'>
    FR Pave Web v1.0<br>
    AASHTO 1993 Pavement Design<br>
    Developed with Streamlit
    </small>
    """, unsafe_allow_html=True)

# ─────────────────────────────────────────────
#  HEADER
# ─────────────────────────────────────────────
st.markdown("""
<div class="main-header">
    <h1>🛣️ FR Pave – ระบบออกแบบโครงสร้างชั้นทาง AASHTO 1993</h1>
    <p>Flexible & Rigid Pavement Design | Traffic ESAL | Structural Layer Analysis</p>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
#  TABS
# ─────────────────────────────────────────────
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "🚛 ESAL – Concrete",
    "🛣️ ESAL – Flexible",
    "📐 Design K Value",
    "🏗️ Concrete Thickness",
    "🔧 Flexible Design",
])

# ══════════════════════════════════════════════
#  TAB 1: ESAL for Rigid Pavement
# ══════════════════════════════════════════════
with tab1:
    st.markdown("### 📊 ปริมาณเพลาเดี่ยวมาตรฐานออกแบบ – ผิวทางคอนกรีต")
    st.markdown('<div class="section-card"><h4>🔧 พารามิเตอร์การออกแบบ</h4>', unsafe_allow_html=True)

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        pt_r = st.number_input("Terminal Serviceability, Pt", value=2.5, step=0.1, key="pt_r")
    with c2:
        dp_r = st.number_input("Design Period (Year)", value=20, step=1, key="dp_r")
    with c3:
        ldf_r = st.number_input("Lane Distribution Factor", value=0.9, step=0.05, key="ldf_r")
    with c4:
        ddf_r = st.number_input("Directional Distribution Factor", value=0.5, step=0.05, key="ddf_r")
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="section-card"><h4>🚛 ประเภทและจำนวนยานพาหนะ (2 ทิศทาง ตลอดอายุออกแบบ)</h4>', unsafe_allow_html=True)

    veh_data_r = {}
    vehicle_order = ["MB", "HB", "MT", "HT", "TR", "STR"]

    # Dynamic column widths: [type, single, tandem, tridam, count] + N slab cols
    _slab_n   = len(SLAB_THICKNESSES)
    _col_w    = [2, 1.5, 1.5, 1.5, 2] + [1.8] * _slab_n

    hdr = st.columns(_col_w)
    hdr[0].markdown("**ประเภทรถ**")
    hdr[1].markdown("**Single Axle**")
    hdr[2].markdown("**Tandem Axle**")
    hdr[3].markdown("**Tridam Axle**")
    hdr[4].markdown("**จำนวนรถ (คัน)**")
    for i, t in enumerate(SLAB_THICKNESSES):
        hdr[5+i].markdown(f"**EALF Slab {t}cm**")

    for vtype in vehicle_order:
        cols = st.columns(_col_w)
        axles = VEHICLE_AXLES[vtype]
        # Display axle config summary
        singles = [(L1,cnt) for L1,L2,cnt in axles if L2==1]
        tandems = [(L1,cnt) for L1,L2,cnt in axles if L2==2]
        trimems = [(L1,cnt) for L1,L2,cnt in axles if L2==3]
        s_str = "+".join(f"{cnt}×{L1}t" for L1,cnt in singles) if singles else "-"
        t_str = "+".join(f"{cnt}×{L1}t" for L1,cnt in tandems) if tandems else "-"
        r_str = "+".join(f"{cnt}×{L1}t" for L1,cnt in trimems) if trimems else "-"

        cols[0].markdown(f"**{VEHICLE_LABELS[vtype]}**")
        cols[1].markdown(f"`{s_str}`")
        cols[2].markdown(f"`{t_str}`")
        cols[3].markdown(f"`{r_str}`")
        count = cols[4].number_input("", min_value=0, value=0, step=10000,
                                      key=f"r_count_{vtype}", label_visibility="collapsed")
        veh_data_r[vtype] = count
        # EALF computed from AASHTO equation with current Pt
        for i, t in enumerate(SLAB_THICKNESSES):
            ealf_val = truck_factor_rigid(vtype, t, pt_r)
            cols[5+i].markdown(f"`{ealf_val:.2f}`")

    st.markdown('</div>', unsafe_allow_html=True)

    if st.button("🔄 คำนวณ ESAL (Rigid)", type="primary", key="calc_rigid"):
        esal_r = compute_esal_rigid(veh_data_r, ldf_r, ddf_r, pt_r)
        tf_r   = compute_tf_rigid(veh_data_r, pt_r)
        st.session_state["esal_rigid"] = esal_r

        st.markdown("---")
        st.markdown(f"### 📋 ผลการคำนวณ ESAL – ผิวทางคอนกรีต  `Pt = {pt_r}`")
        st.markdown(f'<div class="result-info">🧮 EALF คำนวณจากสมการ AASHTO 1993 App.D  |  Pt = <b>{pt_r}</b>  |  Pi = 4.5</div>',
                    unsafe_allow_html=True)

        # Metric cards
        cols = st.columns(len(SLAB_THICKNESSES))
        for i, t in enumerate(SLAB_THICKNESSES):
            with cols[i]:
                st.markdown(f"""
                <div class="metric-box">
                    <div class="val">{esal_r[t]:,.0f}</div>
                    <div class="lbl">ESAL – Slab {t} cm</div>
                    <div style="margin-top:0.5rem;font-size:0.85rem;color:#2e6da4;">
                        Total TF = {tf_r[t]:.3f}
                    </div>
                </div>""", unsafe_allow_html=True)

        # Summary table
        df_r = pd.DataFrame({
            "Slab (cm)":         SLAB_THICKNESSES,
            "ESAL in Design Lane": [f"{esal_r[t]:,.0f}" for t in SLAB_THICKNESSES],
            "Total Truck Factor":  [f"{tf_r[t]:.3f}"    for t in SLAB_THICKNESSES],
        })
        st.dataframe(df_r, use_container_width=True, hide_index=True)

        # EALF per vehicle breakdown
        with st.expander("📊 EALF ต่อยานพาหนะ (แยกตาม Slab Thickness)"):
            veh_active = {v: c for v, c in veh_data_r.items() if c > 0}
            if veh_active:
                rows = []
                for vtype, count in veh_active.items():
                    row = {"ประเภทรถ": VEHICLE_LABELS[vtype], "จำนวน (คัน)": f"{count:,}"}
                    for t in SLAB_THICKNESSES:
                        row[f"EALF {t}cm"] = f"{truck_factor_rigid(vtype, t, pt_r):.3f}"
                    rows.append(row)
                st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
            else:
                st.info("ยังไม่มีรถที่กำหนดจำนวน")

        st.markdown('<div class="result-info">✅ ค่า ESAL ถูกบันทึกเข้า Session State แล้ว → ใช้ได้ใน Tab Design K Value</div>',
                    unsafe_allow_html=True)

# ══════════════════════════════════════════════
#  TAB 2: ESAL for Flexible Pavement
# ══════════════════════════════════════════════
with tab2:
    st.markdown("### 📊 ปริมาณเพลาเดี่ยวมาตรฐานออกแบบ – ผิวทางลาดยาง")
    st.markdown('<div class="section-card"><h4>🔧 พารามิเตอร์การออกแบบ</h4>', unsafe_allow_html=True)

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        pt_f = st.number_input("Terminal Serviceability, Pt", value=2.5, step=0.1, key="pt_f")
    with c2:
        dp_f = st.number_input("Design Period (Year)", value=20, step=1, key="dp_f")
    with c3:
        ldf_f = st.number_input("Lane Distribution Factor", value=0.9, step=0.05, key="ldf_f")
    with c4:
        ddf_f = st.number_input("Directional Distribution Factor", value=0.5, step=0.05, key="ddf_f")
    st.markdown('</div>', unsafe_allow_html=True)

    # ── User-defined SN values ──────────────────────────────────────────
    st.markdown('<div class="section-card"><h4>📐 กำหนดค่า Structure Number (SN) สำหรับการคำนวณ</h4>', unsafe_allow_html=True)
    sn_cols = st.columns(4)
    user_sn_values = []
    sn_defaults = [6.5, 7.1, 7.5, 8.0]
    for i, col in enumerate(sn_cols):
        with col:
            sn_val = st.number_input(f"SN ที่ {i+1}", value=sn_defaults[i],
                                     min_value=1.0, max_value=20.0, step=0.1,
                                     key=f"user_sn_{i}", format="%.1f")
            user_sn_values.append(round(sn_val, 2))
    st.markdown('</div>', unsafe_allow_html=True)

    # ── Vehicle input ───────────────────────────────────────────────────
    st.markdown('<div class="section-card"><h4>🚛 ประเภทและจำนวนยานพาหนะ (2 ทิศทาง ตลอดอายุออกแบบ)</h4>', unsafe_allow_html=True)

    veh_data_f = {}
    _sn_n  = len(user_sn_values)
    _col_w2 = [2, 1.5, 1.5, 1.5, 2] + [1.8] * _sn_n

    hdr2 = st.columns(_col_w2)
    hdr2[0].markdown("**ประเภทรถ**")
    hdr2[1].markdown("**Single Axle**")
    hdr2[2].markdown("**Tandem Axle**")
    hdr2[3].markdown("**Tridam Axle**")
    hdr2[4].markdown("**จำนวนรถ (คัน)**")
    for i, sn in enumerate(user_sn_values):
        hdr2[5+i].markdown(f"**EALF SN={sn}**")

    for vtype in vehicle_order:
        cols2 = st.columns(_col_w2)
        axles = VEHICLE_AXLES[vtype]
        singles = [(L1,cnt) for L1,L2,cnt in axles if L2==1]
        tandems = [(L1,cnt) for L1,L2,cnt in axles if L2==2]
        trimems = [(L1,cnt) for L1,L2,cnt in axles if L2==3]
        s_str = "+".join(f"{cnt}×{L1}t" for L1,cnt in singles) if singles else "-"
        t_str = "+".join(f"{cnt}×{L1}t" for L1,cnt in tandems) if tandems else "-"
        r_str = "+".join(f"{cnt}×{L1}t" for L1,cnt in trimems) if trimems else "-"

        cols2[0].markdown(f"**{VEHICLE_LABELS[vtype]}**")
        cols2[1].markdown(f"`{s_str}`")
        cols2[2].markdown(f"`{t_str}`")
        cols2[3].markdown(f"`{r_str}`")
        count_f = cols2[4].number_input("", min_value=0, value=0, step=10000,
                                         key=f"f_count_{vtype}", label_visibility="collapsed")
        veh_data_f[vtype] = count_f

        # EALF from AASHTO equation with current Pt
        for i, sn in enumerate(user_sn_values):
            ealf = truck_factor_flex(vtype, sn, pt_f)
            cols2[5+i].markdown(f"`{ealf:.2f}`")

    st.markdown('</div>', unsafe_allow_html=True)

    if st.button("🔄 คำนวณ ESAL (Flexible)", type="primary", key="calc_flex"):
        esal_f_custom = compute_esal_flexible(veh_data_f, ldf_f, ddf_f, user_sn_values, pt_f)
        tf_f_custom   = compute_tf_flexible(veh_data_f, user_sn_values, pt_f)

        st.session_state["esal_flex"]      = esal_f_custom
        st.session_state["user_sn_values"] = user_sn_values

        st.markdown("---")
        st.markdown(f"### 📋 ผลการคำนวณ ESAL – ผิวทางลาดยาง  `Pt = {pt_f}`")
        st.markdown(f'<div class="result-info">🧮 EALF คำนวณจากสมการ AASHTO 1993 App.D  |  Pt = <b>{pt_f}</b>  |  Pi = 4.2</div>',
                    unsafe_allow_html=True)

        cols = st.columns(len(user_sn_values))
        for i, sn in enumerate(user_sn_values):
            with cols[i]:
                st.markdown(f"""
                <div class="metric-box">
                    <div class="val">{esal_f_custom[sn]:,.0f}</div>
                    <div class="lbl">ESAL – SN {sn}</div>
                    <div style="margin-top:0.5rem;font-size:0.85rem;color:#2e6da4;">
                        Total TF = {tf_f_custom[sn]:.3f}
                    </div>
                </div>""", unsafe_allow_html=True)

        df_f = pd.DataFrame({
            "SN":                  user_sn_values,
            "ESAL in Design Lane": [f"{esal_f_custom[sn]:,.0f}" for sn in user_sn_values],
            "Total Truck Factor":  [f"{tf_f_custom[sn]:.3f}"    for sn in user_sn_values],
        })
        st.dataframe(df_f, use_container_width=True, hide_index=True)

        # EALF per vehicle breakdown
        with st.expander("📊 EALF ต่อยานพาหนะ (แยกตาม SN)"):
            veh_active_f = {v: c for v, c in veh_data_f.items() if c > 0}
            if veh_active_f:
                rows_f = []
                for vtype, count in veh_active_f.items():
                    row_f = {"ประเภทรถ": VEHICLE_LABELS[vtype], "จำนวน (คัน)": f"{count:,}"}
                    for sn in user_sn_values:
                        row_f[f"EALF SN={sn}"] = f"{truck_factor_flex(vtype, sn, pt_f):.3f}"
                    rows_f.append(row_f)
                st.dataframe(pd.DataFrame(rows_f), use_container_width=True, hide_index=True)

                # Pt sensitivity note
                st.markdown(f"""
                <div class="result-warn">
                ⚡ <b>ค่า EALF ข้างต้นคำนวณที่ Pt = {pt_f}</b><br>
                หากเปลี่ยน Pt → EALF จะเปลี่ยนตาม Gt = log₁₀((4.2−Pt)/2.7)<br>
                Gt ปัจจุบัน = {math.log10((4.2-pt_f)/2.7):.4f}
                </div>""", unsafe_allow_html=True)
            else:
                st.info("ยังไม่มีรถที่กำหนดจำนวน")

        st.markdown('<div class="result-info">✅ ค่า ESAL ถูกบันทึกเข้า Session State แล้ว → ใช้ได้ใน Tab Flexible Design</div>',
                    unsafe_allow_html=True)


# ══════════════════════════════════════════════
#  TAB 3: Design K Value
# ══════════════════════════════════════════════
with tab3:
    st.markdown("### 📐 Effective Modulus of Subgrade Reaction, k_eff")

    col_left, col_right = st.columns([1, 1])

    with col_left:
        st.markdown('<div class="section-card"><h4>📥 Input Design ESAL (ดึงจาก Tab 1 อัตโนมัติ)</h4>', unsafe_allow_html=True)
        esal_prev = st.session_state["esal_rigid"]
        has_rigid_esal = any(v > 0 for v in esal_prev.values())

        if has_rigid_esal:
            st.markdown('<div class="result-info">✅ พบค่า ESAL จาก Tab 1 — แสดงอัตโนมัติ (แก้ไขได้)</div>',
                        unsafe_allow_html=True)
        else:
            st.info("ยังไม่มีค่า ESAL จาก Tab 1 — กรอกค่าเองได้")

        esal_k = {}
        for t in SLAB_THICKNESSES:
            auto_val = int(esal_prev.get(t, 0))
            esal_k[t] = st.number_input(
                f"Design ESAL – Slab {t} cm",
                value=auto_val, step=100000, key=f"esal_k_{t}"
            )
        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown('<div class="section-card"><h4>🏗️ คุณสมบัติคอนกรีต</h4>', unsafe_allow_html=True)
        fc_ksc = st.number_input("f'c ที่ 28 วัน (ksc)", value=350, step=10, key="fc_ksc")
        fc_psi = fc_ksc * 14.223
        ec_psi = 57000 * math.sqrt(fc_psi)
        sc_psi = min(600, 8.3 * fc_ksc**0.5 * 14.223**0.5)
        st.markdown(f"Ec = `{ec_psi:,.0f}` psi")
        sc_input = st.number_input("Sc – Modulus of Rupture (psi) max 600",
                                   value=min(600, int(sc_psi)), step=10, key="sc_psi")
        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown('<div class="section-card"><h4>⚙️ พารามิเตอร์ออกแบบ</h4>', unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        with c1:
            r0_k = st.selectbox("Reliability, R0 (%)",
                                [50,60,70,75,80,85,90,91,92,93,94,95,96,97,98,99],
                                index=10, key="r0_k")
            zr_k = ZR_MAP[r0_k]
            st.markdown(f"ZR = `{zr_k}`")
            so_k = st.number_input("So (0.3–0.4)", value=0.35, step=0.01,
                                   min_value=0.3, max_value=0.4, key="so_k")
        with c2:
            pi_k = st.number_input("Initial Serviceability, Pi", value=4.5, step=0.1, key="pi_k")
            pt_k = st.number_input("Terminal Serviceability, Pt", value=2.5, step=0.1, key="pt_k")
        c3, c4 = st.columns(2)
        with c3:
            j_jpcp = st.number_input("J – JRCP/JPCP", value=2.5, step=0.1, key="j_jpcp")
            j_crcp = st.number_input("J – CRCP", value=2.3, step=0.1, key="j_crcp")
        with c4:
            cd_k = st.number_input("Drainage Coeff., Cd", value=1.1, step=0.05, key="cd_k")
        st.markdown('</div>', unsafe_allow_html=True)

    with col_right:
        st.markdown('<div class="section-card"><h4>📊 ผลการคำนวณ keff</h4>', unsafe_allow_html=True)
        if st.button("🔄 คำนวณ keff", type="primary", key="calc_keff"):
            keff_jpcp_res = {}
            keff_crcp_res = {}
            for t in SLAB_THICKNESSES:
                d_in = t / 2.54
                esal_val = max(esal_k[t], 1)

                keff_j = aashto_keff_rigid(
                    esal_val, zr_k, so_k, pi_k, pt_k,
                    ec_psi, sc_input, j_jpcp, cd_k, d_in
                )
                keff_c = aashto_keff_rigid(
                    esal_val, zr_k, so_k, pi_k, pt_k,
                    ec_psi, sc_input, j_crcp, cd_k, d_in
                )
                keff_jpcp_res[t] = round(keff_j, 3)
                keff_crcp_res[t] = round(keff_c, 3)

            st.session_state["keff_jpcp"] = keff_jpcp_res
            st.session_state["keff_crcp"] = keff_crcp_res

            st.markdown("#### JRCP / JPCP")
            df_kj = pd.DataFrame({
                "Slab Thickness (cm)": SLAB_THICKNESSES,
                "keff (pci)": [f"{keff_jpcp_res[t]:.3f}" for t in SLAB_THICKNESSES],
                "สถานะ": ["✅ OK" if keff_jpcp_res[t] <= 3000 else "⚠️ Cap 3000" for t in SLAB_THICKNESSES],
            })
            st.dataframe(df_kj, use_container_width=True, hide_index=True)

            st.markdown("#### CRCP")
            df_kc = pd.DataFrame({
                "Slab Thickness (cm)": SLAB_THICKNESSES,
                "keff (pci)": [f"{keff_crcp_res[t]:.3f}" for t in SLAB_THICKNESSES],
                "สถานะ": ["✅ OK" if keff_crcp_res[t] <= 3000 else "⚠️ Cap 3000" for t in SLAB_THICKNESSES],
            })
            st.dataframe(df_kc, use_container_width=True, hide_index=True)

            st.markdown('<div class="result-info">✅ ค่า keff บันทึกแล้ว → ใช้ใน Tab Concrete Thickness</div>',
                        unsafe_allow_html=True)
        else:
            st.info("กด 'คำนวณ keff' เพื่อแสดงผล")
        st.markdown('</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════
#  TAB 4: Concrete Thickness Design
# ══════════════════════════════════════════════
with tab4:
    st.markdown("### 🏗️ Concrete Pavement Thickness Design")

    # ── Shared Roadbed Mr (top) ────────────────────────────────────────
    st.markdown('<div class="section-card"><h4>🌍 Roadbed Resilient Modulus (ใช้ร่วมกันทุก Type)</h4>', unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        cbr_rd = st.number_input("Subgrade CBR (%)", value=3.0, step=0.5,
                                  min_value=0.5, max_value=100.0, key="cbr_rd")
    with c2:
        mr_auto_psi = 1500.0 * cbr_rd
        mr_auto_mpa = mr_auto_psi / 145.038
        # Store auto value; only override widget default when CBR actually changes
        prev_cbr = st.session_state.get("_cbr_rd_prev", None)
        if prev_cbr != cbr_rd:
            st.session_state["_mr_sub_mpa_val"] = round(mr_auto_mpa, 2)
            st.session_state["_cbr_rd_prev"]    = cbr_rd
        mr_sub_mpa = st.number_input(
            "Mr of Subgrade (MPa) [อัตโนมัติจาก CBR — แก้ไขได้]",
            value=st.session_state.get("_mr_sub_mpa_val", round(mr_auto_mpa, 2)),
            step=0.5, min_value=1.0, key="mr_sub_mpa"
        )
        st.session_state["_mr_sub_mpa_val"] = mr_sub_mpa
    with c3:
        mr_sub_psi = mr_sub_mpa * 145.038
        mr_sub_pci = mr_sub_psi / 19.4
        st.markdown(f"**Mr = {mr_sub_psi:,.0f} psi**")
        st.markdown(f"**k (subgrade) ≈ {mr_sub_pci:.1f} pci**")
    with c4:
        st.markdown(f"""
        <div style="background:#e3f2fd;border-radius:8px;padding:0.7rem;font-size:0.85rem;color:#1565c0;">
        สูตรอ้างอิง:<br>
        Mr (psi) = 1,500 × CBR<br>
        Mr อัตโนมัติ = <b>{mr_auto_psi:,.0f} psi</b><br>
        = <b>{mr_auto_mpa:.2f} MPa</b>
        </div>""", unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # ── Three pavement types ────────────────────────────────────────────
    type_names = ["Type I – JPCP", "Type II – JPCP (ลดชั้น AC)", "Type III – CRCP"]
    type_keys  = ["I", "II", "III"]

    for t_idx, (tname, tkey) in enumerate(zip(type_names, type_keys)):
        st.markdown(f'<div class="section-card"><h4>🔩 {tname}</h4>', unsafe_allow_html=True)

        col_left4, col_right4 = st.columns([3, 2])

        with col_left4:
            # Layer inputs (left side)
            layers = []
            for li in range(4):
                lc1, lc2 = st.columns([3, 1])
                with lc1:
                    mat = st.selectbox(f"ชั้นที่ {li+1}",
                                       list(RIGID_LAYER_MATERIALS.keys()),
                                       key=f"mat_{tkey}_{li}")
                with lc2:
                    if mat != "None":
                        h = st.number_input("หนา (cm)", value=20, step=1,
                                            min_value=1, key=f"h_{tkey}_{li}",
                                            label_visibility="visible")
                        layers.append((h, RIGID_LAYER_MATERIALS[mat]))
                    else:
                        st.markdown("<br>", unsafe_allow_html=True)

        with col_right4:
            # ── Slab Thickness (right side) ──
            st.markdown("**Slab Thickness**")
            slab_t = st.selectbox("ความหนาแผ่นคอนกรีต (cm)",
                                  SLAB_THICKNESSES, index=1,
                                  key=f"slab_t_{tkey}")

            # ── Auto-pull ESAL from session state when slab changes ──
            keff_dict = (st.session_state["keff_crcp"] if tkey == "III"
                         else st.session_state["keff_jpcp"])
            keff_min_val = keff_dict.get(slab_t, 0)

            esal_from_tab1 = st.session_state["esal_rigid"].get(slab_t, 0)
            st.markdown(f"""
            <div style="background:#f0f4f8;border-radius:6px;padding:0.5rem 0.8rem;
                        font-size:0.85rem;margin:0.4rem 0;">
                📊 Design ESAL (Slab {slab_t} cm) = <b>{esal_from_tab1:,.0f}</b><br>
                📐 keff (min) required = <b>{keff_min_val:.3f} pci</b>
            </div>""", unsafe_allow_html=True)

            st.divider()

            if st.button(f"✅ Design Check – {tname}", key=f"dc_{tkey}"):
                if mr_sub_psi <= 0:
                    st.error("กรุณาระบุ Mr of Subgrade ก่อน")
                else:
                    # Odemark equivalent subbase thickness
                    if layers:
                        h_eq_in = sum(
                            (h / 2.54) * (mr / mr_sub_psi) ** (1/3)
                            for h, mr in layers
                        )
                    else:
                        h_eq_in = 0.0

                    # Compute keff from equivalent subbase thickness
                    k_input = mr_sub_pci * (1 + 0.5 * h_eq_in ** 0.4)
                    k_input = min(k_input, 3000.0)

                    # Equivalent subbase Mr (for display)
                    esb_mr = mr_sub_psi * (h_eq_in ** 3 + 1) * 0.9 if h_eq_in > 0 else mr_sub_psi

                    check     = k_input >= keff_min_val
                    chk_txt   = "✅ PASS" if check else "❌ FAIL"
                    css_class = "result-pass" if check else "result-fail"

                    st.markdown(f"""
                    <div class="{css_class}">
                        Subbase Equiv. Thick. = <b>{h_eq_in:.2f} in</b><br>
                        Esb (equiv.) = <b>{esb_mr:,.0f} psi</b><br>
                        keff (input) = <b>{k_input:.1f} pci</b><br>
                        keff (min) = <b>{keff_min_val:.3f} pci</b><br>
                        <strong style="font-size:1.05rem">{chk_txt}</strong>
                    </div>""", unsafe_allow_html=True)

        st.markdown('</div>', unsafe_allow_html=True)


# ══════════════════════════════════════════════
#  TAB 5: Flexible Pavement Design
# ══════════════════════════════════════════════
with tab5:
    st.markdown("### 🔧 Flexible Pavement Design – AASHTO 1993")

    col_l, col_r = st.columns([1, 1])

    with col_l:
        st.markdown('<div class="section-card"><h4>📥 Design ESAL</h4>', unsafe_allow_html=True)
        esal_prev_f = st.session_state["esal_flex"]
        has_esal_f  = any(v > 0 for v in esal_prev_f.values())

        if has_esal_f:
            # Build SN list from computed session state
            avail_sn = list(esal_prev_f.keys())
            sn_labels = [f"SN {sn}  →  ESAL = {esal_prev_f[sn]:,.0f}" for sn in avail_sn]
            sel_idx = st.selectbox("เลือก SN สำหรับออกแบบ (ดึงจาก Tab 2 อัตโนมัติ)",
                                   range(len(avail_sn)),
                                   format_func=lambda i: sn_labels[i],
                                   key="sel_sn_idx")
            sel_sn = avail_sn[sel_idx]
            design_esal_f = esal_prev_f[sel_sn]
            st.markdown(f"""
            <div class="result-info">
                📊 ใช้ SN = <b>{sel_sn}</b> → Design ESAL = <b>{design_esal_f:,.0f}</b>
            </div>""", unsafe_allow_html=True)
        else:
            st.info("ยังไม่มีค่า ESAL จาก Tab 2 — กรอกค่าด้วยตนเอง")
            sel_sn = st.number_input("SN ที่ใช้ออกแบบ", value=7.1, step=0.1, key="sel_sn_manual")
            design_esal_f = st.number_input("Design ESAL", value=0, step=100000, key="design_esal_manual")

        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown('<div class="section-card"><h4>🌍 Roadbed Resilient Modulus</h4>', unsafe_allow_html=True)
        c1, c2, c3 = st.columns([1, 1, 1])
        with c1:
            cbr_f = st.number_input("Subgrade CBR (%)", value=3.0, step=0.5,
                                    min_value=0.5, max_value=100.0, key="cbr_f")
        with c2:
            mr_f_auto_psi = 1500.0 * cbr_f
            mr_f_auto_mpa = mr_f_auto_psi / 145.038
            prev_cbr_f = st.session_state.get("_cbr_f_prev", None)
            if prev_cbr_f != cbr_f:
                st.session_state["_mr_sub_f_mpa_val"] = round(mr_f_auto_mpa, 2)
                st.session_state["_cbr_f_prev"]       = cbr_f
            mr_sub_f_mpa = st.number_input(
                "Mr of Subgrade (MPa) [อัตโนมัติจาก CBR — แก้ไขได้]",
                value=st.session_state.get("_mr_sub_f_mpa_val", round(mr_f_auto_mpa, 2)),
                step=0.5, min_value=1.0, key="mr_sub_f_mpa"
            )
            st.session_state["_mr_sub_f_mpa_val"] = mr_sub_f_mpa
            mr_sub_f_psi = mr_sub_f_mpa * 145.038
        with c3:
            st.markdown(f"""
            <div style="background:#e3f2fd;border-radius:8px;padding:0.7rem;
                        font-size:0.85rem;color:#1565c0;margin-top:1.6rem;">
                Mr = <b>{mr_sub_f_psi:,.0f} psi</b><br>
                (CBR×1,500 = {mr_f_auto_psi:,.0f} psi)
            </div>""", unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown('<div class="section-card"><h4>⚙️ Design Parameters</h4>', unsafe_allow_html=True)
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            r0_f = st.selectbox("Reliability, R0 (%)",
                                [50,60,70,75,80,85,90,91,92,93,94,95,96,97,98,99],
                                index=10, key="r0_f")
            zr_f = ZR_MAP[r0_f]
            st.markdown(f"ZR = `{zr_f}`")
        with c2:
            so_f = st.number_input("So (0.4–0.5)", value=0.45, step=0.01,
                                   min_value=0.4, max_value=0.5, key="so_f")
        with c3:
            pi_f = st.number_input("Initial Serviceability, Pi", value=4.2,
                                   step=0.1, key="pi_f")
        with c4:
            pt_f2 = st.number_input("Terminal Serviceability, Pt", value=2.5,
                                    step=0.1, key="pt_f2")
        st.markdown('</div>', unsafe_allow_html=True)

    with col_r:
        st.markdown('<div class="section-card"><h4>🔩 Flexible Pavement Design Layers</h4>', unsafe_allow_html=True)

        # Header row
        hf0, hf1, hf2, hf3, hf4, hf5 = st.columns([3, 1.2, 1, 1, 1, 1.2])
        hf0.markdown("**วัสดุ**")
        hf1.markdown("**หนา (cm)**")
        hf2.markdown("**ai**")
        hf3.markdown("**mi**")
        hf4.markdown("**SNi**")
        hf5.markdown("**ΣSNi**")

        layer_results = []
        cum_sn = 0.0
        mat_options = list(FLEX_LAYER_MATERIALS.keys())

        for li in range(5):
            lf0, lf1 = st.columns([3, 1.2])
            with lf0:
                mat_f = st.selectbox(f"ชั้น {li+1}", mat_options,
                                     key=f"fmat_{li}", label_visibility="collapsed")
            with lf1:
                h_f = st.number_input("cm", value=0, step=1, min_value=0,
                                      key=f"fh_{li}", label_visibility="collapsed")

            if mat_f != "None" and h_f > 0:
                ai, mi = FLEX_LAYER_MATERIALS[mat_f]
                h_in = h_f / 2.54
                sn_i = ai * h_in * mi
                cum_sn += sn_i
                layer_results.append({
                    "ชั้น": li+1, "วัสดุ": mat_f,
                    "h (cm)": h_f, "ai": ai, "mi": mi,
                    "SNi": round(sn_i, 3),
                    "ΣSNi": round(cum_sn, 3),
                })
                # Inline display
                _, d1, d2, d3, d4, d5 = st.columns([3, 1.2, 1, 1, 1, 1.2])
                d1.markdown(f"`{h_f}`")
                d2.markdown(f"`{ai:.2f}`")
                d3.markdown(f"`{mi}`")
                d4.markdown(f"`{sn_i:.3f}`")
                d5.markdown(f"**`{cum_sn:.3f}`**")

        st.markdown(f"""
        <div style="background:#e8f5e9;border-radius:6px;padding:0.6rem 1rem;
                    margin-top:0.5rem;font-size:0.9rem;color:#2e7d32;">
            ΣSN ที่ออกแบบ = <b>{cum_sn:.3f}</b>
        </div>""", unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

        # ── Design Check ──────────────────────────────────────────────
        if st.button("✅ Design Check (Flexible)", type="primary", key="flex_check"):
            if design_esal_f <= 0:
                st.warning("กรุณาใส่ Design ESAL ก่อน")
            else:
                try:
                    sn_req = aashto_sn_required_flex(
                        design_esal_f, zr_f, so_f, pi_f, pt_f2, mr_sub_f_psi
                    )
                except Exception:
                    sn_req = None

                if sn_req:
                    passed = cum_sn >= sn_req
                    margin = cum_sn - sn_req
                    css    = "result-pass" if passed else "result-fail"
                    chk    = "✅ PASS" if passed else "❌ FAIL"

                    # Summary metrics
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.markdown(f"""
                        <div class="metric-box">
                            <div class="val">{cum_sn:.3f}</div>
                            <div class="lbl">SN Provided</div>
                        </div>""", unsafe_allow_html=True)
                    with col2:
                        st.markdown(f"""
                        <div class="metric-box">
                            <div class="val">{sn_req:.3f}</div>
                            <div class="lbl">SN Required</div>
                        </div>""", unsafe_allow_html=True)
                    with col3:
                        st.markdown(f"""
                        <div class="metric-box">
                            <div class="val" style="color:{'#2e7d32' if passed else '#c62828'}">
                                {margin:+.3f}
                            </div>
                            <div class="lbl">Safety Margin</div>
                        </div>""", unsafe_allow_html=True)

                    st.markdown(f"""
                    <div class="{css}" style="margin-top:0.8rem;">
                        Design ESAL = <b>{design_esal_f:,.0f}</b> &nbsp;|&nbsp;
                        SN Required = <b>{sn_req:.3f}</b> &nbsp;|&nbsp;
                        SN Provided = <b>{cum_sn:.3f}</b><br>
                        Require SN on Subgrade = <b>{sn_req:.3f}</b><br>
                        <strong style="font-size:1.1rem">{chk}</strong>
                    </div>""", unsafe_allow_html=True)

                    if layer_results:
                        st.markdown("#### 📋 ตารางชั้นทาง")
                        df_layer = pd.DataFrame(layer_results)
                        st.dataframe(df_layer, use_container_width=True, hide_index=True)
                else:
                    st.error("ไม่สามารถคำนวณ SN Required ได้ — กรุณาตรวจสอบค่า Mr และ ESAL")
