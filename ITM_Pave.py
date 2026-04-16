# ╔══════════════════════════════════════════════════════════════════╗
# ║  ITM Pave Pro v2.0 — AASHTO 1993 Pavement Design System        ║
# ║  พัฒนาโดย รศ.ดร.อิทธิพล มีผล | ภาควิชาครุศาสตร์โยธา มจพ.    ║
# ╠══════════════════════════════════════════════════════════════════╣
# ║  TAB 1 │ ESAL Calculator  (LDF/DDF/Pt/R0 ร่วมกัน)             ║
# ║  TAB 2 │ CBR Analysis                                          ║
# ║  TAB 3 │ Flexible Design  (Pi=4.2, So=0.45)                   ║
# ║  TAB 4 │ K-Value Nomograph                                     ║
# ║  TAB 5 │ Rigid Design     (JPCP/JRCP + CRCP)                  ║
# ║  TAB 6 │ Report & Save                                         ║
# ╚══════════════════════════════════════════════════════════════════╝

# ─────────────────────────────────────────────
#  SEC 1: IMPORTS
# ─────────────────────────────────────────────
import streamlit as st
import pandas as pd
import numpy as np
import math, json, io, base64
from datetime import datetime

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

import plotly.graph_objects as go

try:
    from scipy.optimize import brentq as _brentq
except ImportError:
    def _brentq(f, a, b, xtol=1e-6, maxiter=500):
        fa, fb = f(a), f(b)
        if fa * fb > 0: raise ValueError("No sign change")
        for _ in range(maxiter):
            mid = (a+b)/2.0; fm = f(mid)
            if abs(fm) < xtol or (b-a)/2.0 < xtol: return mid
            if fa*fm < 0: b,fb = mid,fm
            else: a,fa = mid,fm
        return (a+b)/2.0

import openpyxl
from openpyxl import load_workbook
OPENPYXL_OK = True

from docx import Document as DocxDoc
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
DOCX_OK = True

from PIL import Image as PILImage, ImageDraw as PILDraw

# ─────────────────────────────────────────────
#  SEC 2: PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="ITM Pave Pro v2.0",
    page_icon="🛣️", layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
#  SEC 3: CSS — Engineering Green Theme
# ─────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;600;700&family=IBM+Plex+Mono:wght@400;600&display=swap');
html,body,[class*="css"]{font-family:'Sarabun',sans-serif;}
.main-header{background:linear-gradient(135deg,#1B5E20 0%,#388E3C 60%,#43A047 100%);
  color:white;padding:1.2rem 2rem;border-radius:14px;margin-bottom:1.2rem;
  box-shadow:0 6px 20px rgba(27,94,32,0.35);border-left:6px solid #A5D6A7;}
.main-header h1{margin:0;font-size:1.65rem;font-weight:700;}
.main-header p{margin:0.2rem 0 0;font-size:0.88rem;opacity:0.88;}
.card{background:#fff;border:1px solid #C8E6C9;border-left:5px solid #2E7D32;
  border-radius:10px;padding:0.9rem 1.2rem;margin-bottom:0.9rem;
  box-shadow:0 2px 8px rgba(46,125,50,0.08);}
.card h4{color:#1B5E20;margin:0 0 0.7rem;font-size:0.97rem;font-weight:700;}
.badge-ready{background:#E8F5E9;color:#2E7D32;border:1px solid #A5D6A7;
  border-radius:20px;padding:0.2rem 0.75rem;font-size:0.8rem;font-weight:600;display:inline-block;}
.badge-wait{background:#FFF8E1;color:#E65100;border:1px solid #FFE082;
  border-radius:20px;padding:0.2rem 0.75rem;font-size:0.8rem;font-weight:600;display:inline-block;}
.result-pass{background:#E8F5E9;border:1px solid #A5D6A7;border-radius:8px;
  padding:0.7rem 1rem;color:#1B5E20;font-weight:600;margin:0.3rem 0;}
.result-fail{background:#FFEBEE;border:1px solid #EF9A9A;border-radius:8px;
  padding:0.7rem 1rem;color:#B71C1C;font-weight:600;margin:0.3rem 0;}
.result-info{background:#E3F2FD;border:1px solid #90CAF9;border-radius:8px;
  padding:0.7rem 1rem;color:#0D47A1;font-weight:600;margin:0.3rem 0;}
.result-warn{background:#FFF8E1;border:1px solid #FFE082;border-radius:8px;
  padding:0.7rem 1rem;color:#E65100;font-weight:600;margin:0.3rem 0;}
.metric-box{background:#fff;border:1px solid #C8E6C9;border-radius:12px;
  padding:0.9rem;text-align:center;box-shadow:0 2px 8px rgba(46,125,50,0.10);}
.metric-box .val{font-size:1.4rem;font-weight:700;color:#1B5E20;font-family:'IBM Plex Mono',monospace;}
.metric-box .lbl{font-size:0.75rem;color:#558B2F;margin-top:0.15rem;}
.layer-row{background:#F9FBE7;border:1px solid #DCEDC8;border-radius:6px;
  padding:0.3rem 0.8rem;margin:0.15rem 0;font-size:0.85rem;color:#33691E;
  font-family:'IBM Plex Mono',monospace;}
[data-baseweb="tab-list"]{gap:3px;}
[data-baseweb="tab"]{background:#E8F5E9!important;border-radius:8px 8px 0 0!important;
  font-weight:600!important;color:#1B5E20!important;padding:0.4rem 0.85rem!important;}
[aria-selected="true"][data-baseweb="tab"]{background:#2E7D32!important;color:white!important;}
[data-testid="stSidebar"]{background:#1B5E20;}
[data-testid="stSidebar"] *{color:#E8F5E9!important;}
button[kind="primary"]{background:#2E7D32!important;border-radius:8px!important;font-weight:700!important;}
.stNumberInput>div>div>input{font-family:'IBM Plex Mono',monospace;font-weight:600;}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
#  SEC 4: CONSTANTS
# ─────────────────────────────────────────────
TON_TO_KIP = 2.2046

VEHICLE_AXLES = {
    "MB":  [(4,1,1),(11,1,1)],
    "HB":  [(5,1,1),(20,2,1)],
    "MT":  [(4,1,1),(11,1,1)],
    "HT":  [(5,1,1),(20,2,1)],
    "TR":  [(5,1,1),(20,2,1),(11,1,2)],
    "STR": [(5,1,1),(20,2,2)],
}
VEHICLE_LABELS = {"MB":"Medium Bus","HB":"Heavy Bus","MT":"Medium Truck",
                  "HT":"Heavy Truck","TR":"Trailer","STR":"Semi Trailer"}
VEHICLE_COLS   = ["MB","HB","MT","HT","TR","STR"]

SLAB_THICKNESSES = [25, 28, 30, 32, 35]
SLAB_LABELS      = ["25 cm (10 in)","28 cm (11 in)","30 cm (12 in)",
                     "32 cm (13 in)","35 cm (14 in)"]

ZR_MAP = {50:0.000,60:-0.253,70:-0.524,75:-0.674,80:-0.841,85:-1.037,
          90:-1.282,91:-1.340,92:-1.405,93:-1.476,94:-1.555,95:-1.645,
          96:-1.751,97:-1.881,98:-2.054,99:-2.327}

# Flexible layer materials: {name: (ai, mi)}
FLEX_MATERIALS = {
    "None":                              (None, None),
    "Asphalt Concrete (AC)":             (0.42, 1.0),
    "Cement Treated Base UCS 24.5 ksc":  (0.15, 1.0),
    "Cement Treated Base UCS 17.5 ksc":  (0.13, 1.0),
    "Crushed Rock Base CBR 80%":         (0.14, 1.0),
    "Soil Aggregate Subbase CBR 25%":    (0.10, 1.0),
    "Soil Aggregate Subbase CBR 20%":    (0.09, 1.0),
    "Sand Embankment CBR 10%":           (0.08, 1.0),
}

# Rigid layer materials: {name: {e_mpa, e_psi}}
RIGID_MATERIALS = {
    "AC Interlayer":                       {"e_mpa":2500,"e_psi":362500},
    "Cement Treated Base UCS 40 ksc":      {"e_mpa":1200,"e_psi":174000},
    "MOD. Crushed Rock Base UCS 24.5 ksc": {"e_mpa": 850,"e_psi":123250},
    "Soil Cement UCS 17.5 ksc":            {"e_mpa": 350,"e_psi": 50750},
    "Crushed Rock Base CBR 80%":           {"e_mpa": 350,"e_psi": 50750},
    "Soil Aggregate Subbase CBR 25%":      {"e_mpa": 150,"e_psi": 21750},
    "Embankment":                          {"e_mpa": 100,"e_psi": 14500},
}
RIGID_MAT_NAMES = ["None"] + list(RIGID_MATERIALS.keys())

SAMPLE_CBR = [14.8,14.37,5.31,17.37,5.48,18.46,4.85,6.23,
              5.02,10.78,10.52,14,15.5,8.7,12.93,8.19,
              8.1,15.56,16.88,20.75,20.3,8,7.84,7.48,
              23.55,8.92,13.3,13.5,13.86,7.18,6.95,5.8,
              6,11.18,9.69,7.48]

# ─────────────────────────────────────────────
#  SEC 5: ENGINE FUNCTIONS
# ─────────────────────────────────────────────
def ealf_flex(L1t, L2, SN, Pt):
    L1=L1t*TON_TO_KIP
    Gt=math.log10((4.2-Pt)/(4.2-1.5))
    Bx=0.40+0.081*(L1+L2)**3.23/((SN+1)**5.19*L2**3.23)
    B18=0.40+0.081*(18+1)**3.23/((SN+1)**5.19*1.0**3.23)
    return 10**(4.79*math.log10(L1+L2)-4.33*math.log10(L2)
                -4.79*math.log10(19)+Gt*(1/B18-1/Bx))

def ealf_rigid(L1t, L2, D_cm, Pt):
    L1=L1t*TON_TO_KIP; D=D_cm/2.54
    Gt=math.log10((4.5-Pt)/(4.5-1.5))
    Bx=1.0+3.63*(L1+L2)**5.20/((D+1)**8.46*L2**3.52)
    B18=1.0+3.63*(18+1)**5.20/((D+1)**8.46*1.0**3.52)
    return 10**(4.62*math.log10(L1+L2)-3.28*math.log10(L2)
                -4.62*math.log10(19)+Gt*(1/B18-1/Bx))

def tf_flex(vt, SN, Pt):
    return sum(ealf_flex(L1,L2,SN,Pt)*c for L1,L2,c in VEHICLE_AXLES[vt])

def tf_rigid(vt, D_cm, Pt):
    return sum(ealf_rigid(L1,L2,D_cm,Pt)*c for L1,L2,c in VEHICLE_AXLES[vt])

def compute_esal(traffic_df, ldf, ddf, Pt, mode="rigid", sn_list=None):
    """ESAL = Σ_year [ AADT × 365 × DDF × LDF × TF ]"""
    DAYS = 365
    if mode == "rigid":
        res = {D: 0.0 for D in SLAB_THICKNESSES}
        for _, row in traffic_df.iterrows():
            for vt in VEHICLE_COLS:
                cnt = float(row.get(vt, 0) or 0)
                if cnt <= 0: continue
                for D in SLAB_THICKNESSES:
                    res[D] += cnt * DAYS * ddf * ldf * tf_rigid(vt, D, Pt)
        return res
    else:
        keys = sn_list or [6.5, 7.1, 7.5, 8.0]
        res  = {k: 0.0 for k in keys}
        for _, row in traffic_df.iterrows():
            for vt in VEHICLE_COLS:
                cnt = float(row.get(vt, 0) or 0)
                if cnt <= 0: continue
                for SN in keys:
                    res[SN] += cnt * DAYS * ddf * ldf * tf_flex(vt, SN, Pt)
        return res

def sn_required(esal, zr, so, pi, pt, mr_psi):
    logW = math.log10(max(esal,1))
    def eq(SN):
        if SN<=0: return -1e10
        return (zr*so + 9.36*math.log10(SN+1) - 0.20
                + math.log10((pi-pt)/2.7)/(0.40+1094/(SN+1)**5.19)
                + 2.32*math.log10(mr_psi) - 8.07 - logW)
    try: return _brentq(eq, 0.1, 30, xtol=1e-4)
    except: return None

def w18_rigid(d_cm, pi, pt, zr, so, sc_psi, cd, j, ec_psi, k_pci):
    d=d_cm/2.54; dp=pi-pt
    t1=zr*so; t2=7.35*math.log10(d+1)-0.06
    t3=math.log10(dp/3.0)/(1+1.624e7/(d+1)**8.46)
    n=sc_psi*cd*(d**0.75-1.132)
    ek=(ec_psi/k_pci)**0.25
    dv=215.63*j*(d**0.75-18.42/ek)
    if n<=0 or dv<=0: return None
    inner=n/dv
    if inner<=0: return None
    t4=(4.22-0.32*pt)*math.log10(inner)
    return 10**(t1+t2+t3+t4)

def cbr_to_mr(cbr): return 1500.0*cbr
def mr_to_k(mr):    return mr/19.4

def calc_percentile(vals):
    arr=np.sort(np.array(vals,dtype=float)); n=len(arr)
    u=np.unique(arr)
    pct=np.array([np.sum(arr>=v)/n*100 for v in u])
    return arr,n,u,pct

def grow_traffic(base, rate_pct, years):
    r=rate_pct/100.0
    rows=[]
    for y in range(1,years+1):
        f=(1+r)**(y-1)
        row={"Year":y}
        for v in VEHICLE_COLS: row[v]=int(round(base.get(v,0)*f))
        rows.append(row)
    return pd.DataFrame(rows)

# ─────────────────────────────────────────────
#  SEC 6: NOMOGRAPH FUNCTIONS
# ─────────────────────────────────────────────
def fig_to_bytes(fig):
    buf=io.BytesIO()
    fig.savefig(buf,format='png',dpi=150,bbox_inches='tight',
                facecolor=fig.get_facecolor())
    buf.seek(0); return buf.read()

def draw_arrow(drw, s, e, color, lw=4, aw=14):
    drw.line([s,e], fill=color, width=lw)
    dx=e[0]-s[0]; dy=e[1]-s[1]
    L=math.sqrt(dx*dx+dy*dy)
    if L>0:
        dx/=L; dy/=L; px=-dy; py=dx
        bx=e[0]-aw*dx; by=e[1]-aw*dy
        drw.polygon([(e[0],e[1]),(int(bx+aw*0.5*px),int(by+aw*0.5*py)),
                     (int(bx-aw*0.5*px),int(by-aw*0.5*py))],fill=color)

# ─────────────────────────────────────────────
#  SEC 7: WORD REPORT FUNCTIONS
# ─────────────────────────────────────────────
def _doc():
    doc=DocxDoc()
    s=doc.styles['Normal']; s.font.name='TH SarabunPSK'; s.font.size=Pt(15)
    try: s._element.rPr.rFonts.set(qn('w:eastAsia'),'TH SarabunPSK')
    except: pass
    return doc

def _run(p, txt, bold=False, size=15, color=None, italic=False):
    r=p.add_run(txt); r.font.name='TH SarabunPSK'; r.font.size=Pt(size)
    r.font.bold=bold; r.font.italic=italic
    if color: r.font.color.rgb=RGBColor(*color)
    try: r._element.rPr.rFonts.set(qn('w:eastAsia'),'TH SarabunPSK')
    except: pass
    return r

def _bg(cell, hex_):
    tc=cell._tc; tcPr=tc.get_or_add_tcPr()
    shd=OxmlElement('w:shd')
    shd.set(qn('w:val'),'clear'); shd.set(qn('w:color'),'auto')
    shd.set(qn('w:fill'),hex_); tcPr.append(shd)

def _tbl(doc, headers, rows, bg='C8E6C9'):
    t=doc.add_table(rows=1+len(rows),cols=len(headers))
    t.style='Table Grid'; t.alignment=WD_TABLE_ALIGNMENT.CENTER
    for j,h in enumerate(headers):
        c=t.rows[0].cells[j]; p=c.paragraphs[0]
        _run(p,h,bold=True,size=13); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
        _bg(c,bg)
    for i,row in enumerate(rows):
        for j,v in enumerate(row):
            c=t.rows[i+1].cells[j]; p=c.paragraphs[0]
            _run(p,str(v),size=13); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    return t

def _footer(doc):
    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    _run(p,"พัฒนาโดย รศ.ดร.อิทธิพล มีผล  |  ภาควิชาครุศาสตร์โยธา  |  มจพ.",
         size=12,color=(80,80,80))

def _bytes(doc):
    buf=io.BytesIO(); doc.save(buf); buf.seek(0); return buf.read()

def report_esal(ss):
    doc=_doc()
    p=doc.add_paragraph(); _run(p,"1. ผลการคำนวณ ESAL",bold=True,size=16)
    p2=doc.add_paragraph()
    _run(p2,f"LDF={ss.get('ldf',0.9)} | DDF={ss.get('ddf',0.5)} | "
             f"Pt={ss.get('pt',2.5)} | R0={ss.get('r0',90)}%",size=13)
    doc.add_paragraph()
    if ss.get('esal_rigid'):
        p3=doc.add_paragraph(); _run(p3,"1.1 ESAL – Rigid Pavement",bold=True,size=14)
        _tbl(doc,["Slab","ESAL (Design Lane)"],
             [[lbl,f"{ss['esal_rigid'].get(D,0):,.0f}"]
              for D,lbl in zip(SLAB_THICKNESSES,SLAB_LABELS)])
        doc.add_paragraph()
    if ss.get('esal_flex'):
        p4=doc.add_paragraph(); _run(p4,"1.2 ESAL – Flexible Pavement",bold=True,size=14)
        _tbl(doc,["SN","ESAL (Design Lane)"],
             [[f"SN {sn}",f"{v:,.0f}"] for sn,v in ss['esal_flex'].items()])
    _footer(doc); return _bytes(doc)

def report_cbr(ss):
    doc=_doc()
    p=doc.add_paragraph(); _run(p,"2. ผลการวิเคราะห์ค่า CBR",bold=True,size=16)
    vals=ss.get('cbr_values',[])
    if vals:
        arr,n,u,pct=calc_percentile(vals)
        pt_=ss.get('cbr_percentile',90)
        cbd=float(np.interp(pt_,pct[::-1],u[::-1]))
        p2=doc.add_paragraph()
        _run(p2,f"n={n} | Percentile={pt_}% | CBR={cbd:.2f}% | "
                 f"Mr={cbr_to_mr(cbd):,.0f} psi | k={mr_to_k(cbr_to_mr(cbd)):.1f} pci",size=13)
    _footer(doc); return _bytes(doc)

def report_flex(ss):
    doc=_doc()
    p=doc.add_paragraph(); _run(p,"3. ผลการออกแบบ Flexible Pavement",bold=True,size=16)
    res=ss.get('flex_results',{})
    if res:
        p2=doc.add_paragraph()
        status="✅ PASS" if res.get('pass') else "❌ FAIL"
        _run(p2,f"SN Required={res.get('sn_req',0):.3f} | "
                 f"SN Provided={res.get('sn_prov',0):.3f} | {status}",
             bold=True,size=14)
        doc.add_paragraph()
        layers=res.get('layers',[])
        if layers:
            _tbl(doc,["ชั้น","วัสดุ","หนา(cm)","ai","mi","SNi","ΣSNi"],
                 [[str(l['layer']),l['mat'],str(l['h']),
                   f"{l['ai']:.2f}",f"{l['mi']:.1f}",
                   f"{l['sni']:.3f}",f"{l['cum_sn']:.3f}"] for l in layers])
    _footer(doc); return _bytes(doc)

def report_kvalue(ss):
    doc=_doc()
    p=doc.add_paragraph(); _run(p,"4. ค่า k_eff (Effective Modulus of Subgrade Reaction)",bold=True,size=16)
    p2=doc.add_paragraph()
    _run(p2,f"k∞={ss.get('k_inf',0):.1f} pci | LS={ss.get('ls_value',0)} | "
             f"k_eff={ss.get('k_corrected',0):.1f} pci",size=13)
    for key,cap in [('nomograph_img_k','รูปที่ 1 Composite k∞ Nomograph'),
                    ('nomograph_img_ls','รูปที่ 2 Loss of Support Nomograph')]:
        img_b=ss.get(key)
        if img_b:
            doc.add_paragraph()
            pi_=doc.add_paragraph(); pi_.alignment=WD_ALIGN_PARAGRAPH.CENTER
            pi_.add_run().add_picture(io.BytesIO(img_b),width=Cm(12))
            pc=doc.add_paragraph(); pc.alignment=WD_ALIGN_PARAGRAPH.CENTER
            _run(pc,cap,bold=True,size=14)
    _footer(doc); return _bytes(doc)

def report_rigid(ss):
    doc=_doc()
    p=doc.add_paragraph(); _run(p,"5. ผลการออกแบบ Rigid Pavement",bold=True,size=16)
    rr=ss.get('rigid_results',{})
    if not rr or 'results' not in rr:
        _run(doc.add_paragraph(),"ยังไม่มีผลการคำนวณ",size=13)
        _footer(doc); return _bytes(doc)
    results=rr['results']; min_slab=rr['min_slab']
    de=rr.get('design_esal',0); kv=rr.get('k_eff',0)
    p2=doc.add_paragraph()
    _run(p2,f"W18 Required={de:,.0f} | k_eff={kv:.0f} pci | "
             f"f'c={rr.get('fc',0)} ksc | Sc={rr.get('sc',0)} psi | Cd={rr.get('cd',0)}",size=13)
    doc.add_paragraph()
    # ตารางผล
    hdr=["Slab","JPCP/JRCP (J={:.1f})".format(rr.get('j_jr',2.8)),
         "W18 Cap","CRCP (J={:.1f})".format(rr.get('j_crcp',2.6)),"W18 Cap"]
    rows=[]
    for D,lbl in zip(SLAB_THICKNESSES,SLAB_LABELS):
        r_jr=results[D]['JPCP_JRCP']
        r_cr=results[D]['CRCP']
        s_jr="★PASS" if (r_jr['pass'] and min_slab.get('JPCP_JRCP')==D) else ("PASS" if r_jr['pass'] else "FAIL")
        s_cr="★PASS" if (r_cr['pass'] and min_slab.get('CRCP')==D) else ("PASS" if r_cr['pass'] else "FAIL")
        rows.append([lbl,s_jr,f"{r_jr['w18_cap']/1e6:.2f}M",
                     s_cr,f"{r_cr['w18_cap']/1e6:.2f}M"])
    _tbl(doc,hdr,rows)
    doc.add_paragraph()
    # Design thickness summary
    p3=doc.add_paragraph(); _run(p3,"ความหนาออกแบบที่เลือก:",bold=True,size=14)
    for pt_,key in [("JPCP / JRCP","JPCP_JRCP"),("CRCP","CRCP")]:
        d_=min_slab.get(key)
        lbl_=SLAB_LABELS[SLAB_THICKNESSES.index(d_)] if d_ else "ไม่มี Slab ที่ผ่าน"
        p4=doc.add_paragraph()
        _run(p4,f"  {pt_} = {lbl_}",size=13)
    # Layer table
    layers=rr.get('layers',[])
    if layers:
        doc.add_paragraph()
        p5=doc.add_paragraph(); _run(p5,"ชั้นโครงสร้างทาง:",bold=True,size=14)
        _tbl(doc,["วัสดุ","หนา (cm)","E (MPa)"],
             [[l['name'],str(l['h_cm']),str(l['e_mpa'])] for l in layers])
        p6=doc.add_paragraph()
        _run(p6,f"E_equivalent={rr.get('e_eq_psi',0):,.0f} psi | "
                 f"DSB={rr.get('dsb_in',0):.1f} in",size=13)
    _footer(doc); return _bytes(doc)

def report_full(ss):
    doc=_doc()
    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    _run(p,"รายการคำนวณออกแบบโครงสร้างชั้นทาง",bold=True,size=20)
    p2=doc.add_paragraph(); p2.alignment=WD_ALIGN_PARAGRAPH.CENTER
    _run(p2,"ตามวิธี AASHTO 1993",size=16)
    p3=doc.add_paragraph(); p3.alignment=WD_ALIGN_PARAGRAPH.CENTER
    _run(p3,f"วันที่: {datetime.now().strftime('%d/%m/%Y %H:%M')}",size=13)
    doc.add_page_break()
    for fn,key in [(report_esal,'esal_rigid'),(report_cbr,'cbr_values'),
                   (report_flex,'flex_results'),(report_kvalue,'k_corrected'),
                   (report_rigid,'rigid_results')]:
        if ss.get(key):
            b=fn(ss)
            if b:
                sub=DocxDoc(io.BytesIO(b))
                for el in sub.element.body: doc.element.body.append(el)
                doc.add_page_break()
    _footer(doc); return _bytes(doc)

# ─────────────────────────────────────────────
#  SEC 8: SESSION STATE
# ─────────────────────────────────────────────
def init_ss():
    defs = {
        # Tab 1 — shared params
        'ldf':0.9,'ddf':0.5,'pt':2.5,'r0':90,
        'traffic_df':None,
        'esal_rigid':{D:0 for D in SLAB_THICKNESSES},
        'esal_flex':{},
        'sn_list':[6.5,7.1,7.5,8.0],
        # Tab 2
        'cbr_values':[],'cbr_percentile':90.0,
        'cbr_design':3.0,'mr_psi':4500.0,'k_sub_pci':231.9,
        # Tab 3
        'flex_results':{},
        # Tab 4
        'k_inf':0.0,'k_corrected':0.0,'ls_value':1.0,
        'nomograph_img_k':None,'nomograph_img_ls':None,
        'nomo_esb':50000,'nomo_dsb':6.0,
        'img1_original':None,'img2_original':None,
        # Tab 5
        'rigid_results':{},
    }
    for k,v in defs.items():
        if k not in st.session_state:
            st.session_state[k]=v

init_ss()
ss = st.session_state

def badge(key, label):
    val=ss.get(key)
    has=(val is not None and val!={}and val!=[]and val!=0.0 and val!={D:0 for D in SLAB_THICKNESSES})
    cls="badge-ready" if has else "badge-wait"
    icon="✅" if has else "⚠️"
    return f'<span class="{cls}">{icon} {label}</span>'

# ─────────────────────────────────────────────
#  SIDEBAR
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown("""<div style='text-align:center;padding:0.8rem 0 0.4rem;'>
        <div style='font-size:2rem;'>🛣️</div>
        <div style='font-weight:700;font-size:1rem;color:#A5D6A7;'>ITM Pave Pro</div>
        <div style='font-size:0.75rem;color:#81C784;'>v2.0 | AASHTO 1993</div>
    </div>""",unsafe_allow_html=True)
    st.divider()
    st.markdown("**📊 สถานะ**")
    for k,lbl in [('esal_rigid','ESAL'),('cbr_values','CBR'),
                   ('flex_results','Flexible'),('k_corrected','K-Value'),
                   ('rigid_results','Rigid')]:
        st.markdown(badge(k,lbl),unsafe_allow_html=True)
    st.divider()
    st.markdown("**💾 Save / Load**")
    if st.button("💾 Save JSON",use_container_width=True):
        save={
            'ldf':ss.ldf,'ddf':ss.ddf,'pt':ss.pt,'r0':ss.r0,
            'esal_rigid':{str(k):v for k,v in ss.esal_rigid.items()},
            'esal_flex':{str(k):v for k,v in ss.esal_flex.items()},
            'sn_list':ss.sn_list,
            'cbr_values':ss.cbr_values,'cbr_percentile':ss.cbr_percentile,
            'cbr_design':ss.cbr_design,'mr_psi':ss.mr_psi,'k_sub_pci':ss.k_sub_pci,
            'flex_results':ss.flex_results,
            'k_inf':ss.k_inf,'k_corrected':ss.k_corrected,'ls_value':ss.ls_value,
            'nomo_esb':ss.nomo_esb,'nomo_dsb':ss.nomo_dsb,
            'rigid_results':{str(k):v for k,v in ss.rigid_results.items()} if isinstance(ss.rigid_results,dict) and 'results' not in ss.rigid_results else {},
            'traffic_df':ss.traffic_df.to_dict('records') if ss.traffic_df is not None else None,
        }
        jb=json.dumps(save,ensure_ascii=False,indent=2).encode('utf-8')
        st.download_button("📥 Download",jb,
            f"itm_pave_{datetime.now().strftime('%Y%m%d_%H%M')}.json",
            "application/json",use_container_width=True)
    uj=st.file_uploader("📂 Load JSON",type=['json'])
    if uj:
        try:
            d=json.loads(uj.read().decode('utf-8'))
            for k in ['ldf','ddf','pt','r0','sn_list','cbr_values',
                       'cbr_percentile','cbr_design','mr_psi','k_sub_pci',
                       'flex_results','k_inf','k_corrected','ls_value',
                       'nomo_esb','nomo_dsb']:
                if k in d: ss[k]=d[k]
            if 'esal_rigid' in d:
                ss.esal_rigid={int(k):v for k,v in d['esal_rigid'].items()}
            if 'esal_flex' in d:
                ss.esal_flex={float(k):v for k,v in d['esal_flex'].items()}
            if d.get('traffic_df'):
                ss.traffic_df=pd.DataFrame(d['traffic_df'])
            st.success("✅ โหลดสำเร็จ!"); st.rerun()
        except Exception as e:
            st.error(f"❌ {e}")
    st.divider()
    st.markdown("""<div style='font-size:0.7rem;color:#81C784;text-align:center;line-height:1.8;'>
    รศ.ดร.อิทธิพล มีผล<br>ภาควิชาครุศาสตร์โยธา<br>คณะครุศาสตร์อุตสาหกรรม มจพ.
    </div>""",unsafe_allow_html=True)

# ─────────────────────────────────────────────
#  HEADER
# ─────────────────────────────────────────────
st.markdown("""<div class="main-header">
    <h1>🛣️ ITM Pave Pro v2.0 — ระบบออกแบบโครงสร้างชั้นทาง AASHTO 1993</h1>
    <p>ESAL · CBR · Flexible Design · K-Value Nomograph · Rigid Design (JPCP/JRCP/CRCP) · Report</p>
</div>""",unsafe_allow_html=True)

# ─────────────────────────────────────────────
#  MAIN TABS
# ─────────────────────────────────────────────
tab1,tab2,tab3,tab4,tab5,tab6 = st.tabs([
    "🚛 ESAL","📊 CBR","🔧 Flexible Design",
    "📐 K-Value Nomograph","🏗️ Rigid Design","📄 Report & Save"
])

# ══════════════════════════════════════════════
#  TAB 1: ESAL CALCULATOR
# ══════════════════════════════════════════════
with tab1:
    st.markdown("### 🚛 ESAL Calculator — AASHTO 1993")

    # ── พารามิเตอร์ร่วม ──
    st.markdown('<div class="card"><h4>⚙️ พารามิเตอร์ร่วม (ใช้กับ Flexible และ Rigid)</h4>',unsafe_allow_html=True)
    pc1,pc2,pc3,pc4 = st.columns(4)
    with pc1:
        ldf=st.number_input("Lane Distribution Factor",value=ss.ldf,
                            step=0.05,min_value=0.1,max_value=1.0,key="t1_ldf")
        ss.ldf=ldf
    with pc2:
        ddf=st.number_input("Directional Dist. Factor",value=ss.ddf,
                            step=0.05,min_value=0.1,max_value=1.0,key="t1_ddf")
        ss.ddf=ddf
    with pc3:
        pt=st.number_input("Terminal Serviceability Pt",value=ss.pt,
                           step=0.1,min_value=1.5,max_value=3.5,key="t1_pt")
        ss.pt=pt
    with pc4:
        r0=st.selectbox("Reliability R0 (%)",list(ZR_MAP.keys()),
                        index=list(ZR_MAP.keys()).index(ss.r0) if ss.r0 in ZR_MAP else 10,
                        key="t1_r0")
        ss.r0=r0
        zr=ZR_MAP[r0]
        st.caption(f"ZR = {zr}")
    st.markdown('</div>',unsafe_allow_html=True)

    # ── Traffic Data ──
    st.markdown('<div class="card"><h4>📁 ข้อมูลปริมาณจราจร (คัน/วัน สองทิศทาง)</h4>',unsafe_allow_html=True)
    col_l,col_r=st.columns([1,1])
    with col_l:
        mode=st.radio("วิธีกรอกข้อมูล",["📁 Upload Excel","✏️ กรอกมือ + Growth Rate"],horizontal=True)
        if mode=="📁 Upload Excel":
            xl=st.file_uploader("ไฟล์ Excel (Year,MB,HB,MT,HT,TR,STR)",type=['xlsx'])
            st.caption("คอลัมน์: Year | MB | HB | MT | HT | TR | STR")
            if xl:
                try:
                    df_up=pd.read_excel(xl,engine='openpyxl')
                    df_up.columns=[c.strip() for c in df_up.columns]
                    cm_={c:vc for c in df_up.columns for vc in ['Year']+VEHICLE_COLS if c.upper()==vc.upper()}
                    df_up=df_up.rename(columns=cm_)
                    for vc in VEHICLE_COLS:
                        if vc not in df_up.columns: df_up[vc]=0
                    ss.traffic_df=df_up[['Year']+VEHICLE_COLS].fillna(0)
                    st.success(f"✅ {len(df_up)} ปี")
                except Exception as e:
                    st.error(f"❌ {e}")
        else:
            st.markdown("**ปีแรก (คัน/วัน)**")
            bc=st.columns(6)
            base_={vc:bc[i].number_input(vc,value={"MB":120,"HB":60,"MT":250,"HT":180,"TR":100,"STR":120}[vc],
                                         min_value=0,step=10,key=f"b_{vc}") for i,vc in enumerate(VEHICLE_COLS)}
            gc1,gc2=st.columns(2)
            with gc1: gr=st.number_input("Growth Rate (%/ปี)",value=4.5,step=0.5,min_value=0.0,max_value=20.0)
            with gc2: yr=st.number_input("Design Period (ปี)",value=20,min_value=1,max_value=40,step=1)
            if st.button("🔄 สร้างตาราง",type="primary"):
                ss.traffic_df=grow_traffic(base_,gr,int(yr))
                st.success(f"✅ {int(yr)} ปี")
    with col_r:
        if ss.traffic_df is not None:
            st.markdown("**ตารางจราจร:**")
            st.dataframe(ss.traffic_df.style.format({c:"{:,.0f}" for c in VEHICLE_COLS}),
                        use_container_width=True,height=280)
        else:
            st.info("⬅️ กรอกหรือ Upload ข้อมูลก่อน")
    st.markdown('</div>',unsafe_allow_html=True)

    # ── SN List สำหรับ Flexible ──
    st.markdown('<div class="card"><h4>📐 Structure Number (SN) สำหรับ Flexible</h4>',unsafe_allow_html=True)
    sn_cols=st.columns(4)
    sn_list=[round(sn_cols[i].number_input(f"SN {i+1}",value=ss.sn_list[i] if i<len(ss.sn_list) else 6.5,
                   min_value=1.0,max_value=20.0,step=0.1,key=f"sn_{i}",format="%.1f"),2) for i in range(4)]
    st.markdown('</div>',unsafe_allow_html=True)

    # ── คำนวณปุ่มเดียว ──
    if st.button("🔄 คำนวณ ESAL (Rigid + Flexible พร้อมกัน)",type="primary",use_container_width=True):
        if ss.traffic_df is None:
            st.warning("⚠️ กรุณากรอกข้อมูลจราจรก่อน")
        else:
            er=compute_esal(ss.traffic_df,ldf,ddf,pt,"rigid")
            ef=compute_esal(ss.traffic_df,ldf,ddf,pt,"flex",sn_list)
            ss.esal_rigid=er; ss.esal_flex=ef; ss.sn_list=sn_list
            st.rerun()

    # ── แสดงผล ──
    if any(v>0 for v in ss.esal_rigid.values()):
        st.markdown("---")
        col_r1,col_r2=st.columns(2)
        with col_r1:
            st.markdown("**📊 ESAL — Rigid Pavement**")
            cr=st.columns(len(SLAB_THICKNESSES))
            for i,(D,lbl) in enumerate(zip(SLAB_THICKNESSES,SLAB_LABELS)):
                with cr[i]:
                    st.markdown(f"""<div class="metric-box">
                        <div class="val">{ss.esal_rigid[D]:,.0f}</div>
                        <div class="lbl">{lbl}</div>
                    </div>""",unsafe_allow_html=True)
        with col_r2:
            st.markdown("**📊 ESAL — Flexible Pavement**")
            cf=st.columns(len(sn_list))
            for i,(sn,v) in enumerate(ss.esal_flex.items()):
                with cf[i]:
                    st.markdown(f"""<div class="metric-box">
                        <div class="val">{v:,.0f}</div>
                        <div class="lbl">SN {sn}</div>
                    </div>""",unsafe_allow_html=True)

        # Truck Factor table
        with st.expander("📋 Truck Factor (EALF/คัน) ที่ Pt={:.1f}".format(pt)):
            tf_rows=[]
            for vt in VEHICLE_COLS:
                row={"ประเภทรถ":f"{VEHICLE_LABELS[vt]} ({vt})"}
                for D,lbl in zip(SLAB_THICKNESSES,SLAB_LABELS):
                    row[f"Rigid {lbl}"]=f"{tf_rigid(vt,D,pt):.3f}"
                for sn in sn_list:
                    row[f"Flex SN={sn}"]=f"{tf_flex(vt,sn,pt):.3f}"
                tf_rows.append(row)
            st.dataframe(pd.DataFrame(tf_rows),use_container_width=True,hide_index=True)

# ══════════════════════════════════════════════
#  TAB 2: CBR ANALYSIS
# ══════════════════════════════════════════════
with tab2:
    st.markdown("### 📊 CBR Analysis — Percentile Method")
    cl,cr=st.columns([1,1])
    with cl:
        st.markdown('<div class="card"><h4>📁 ข้อมูล CBR</h4>',unsafe_allow_html=True)
        cbr_mode=st.radio("แหล่งข้อมูล",["📁 Upload Excel","✏️ กรอกค่า","📌 ตัวอย่าง"],horizontal=True)
        cbr_in=None
        if cbr_mode=="📁 Upload Excel":
            cxl=st.file_uploader("ไฟล์ Excel (คอลัมน์ CBR)",type=['xlsx'],key="cbr_xl")
            if cxl:
                try:
                    dc=pd.read_excel(cxl,engine='openpyxl')
                    col_=next((c for c in dc.columns if 'cbr' in c.lower()),dc.columns[0])
                    cbr_in=pd.to_numeric(dc[col_],errors='coerce').dropna().tolist()
                    st.success(f"✅ {len(cbr_in)} ตัวอย่าง")
                except Exception as e: st.error(str(e))
        elif cbr_mode=="✏️ กรอกค่า":
            import re
            txt=st.text_area("ค่า CBR (%) คั่นด้วย , หรือ Enter",height=100)
            if txt.strip():
                try:
                    cbr_in=[float(x) for x in re.split(r'[,\n\r\s]+',txt.strip()) if x]
                    st.success(f"✅ {len(cbr_in)} ค่า")
                except: st.error("กรุณากรอกตัวเลขเท่านั้น")
        else:
            cbr_in=SAMPLE_CBR; st.info(f"📌 {len(SAMPLE_CBR)} ตัวอย่าง")
        if cbr_in: ss.cbr_values=cbr_in
        pct_=st.slider("Percentile (%)",50,99,int(ss.cbr_percentile),key="cbr_pct")
        ss.cbr_percentile=float(pct_)
        st.markdown('</div>',unsafe_allow_html=True)

        if ss.cbr_values:
            arr,n,u,upct=calc_percentile(ss.cbr_values)
            cbd=float(np.interp(pct_,upct[::-1],u[::-1]))
            st.markdown('<div class="card"><h4>🎯 ผลการวิเคราะห์</h4>',unsafe_allow_html=True)
            m1,m2,m3=st.columns(3)
            mr_auto=cbr_to_mr(cbd); k_auto=mr_to_k(mr_auto)
            with m1: st.markdown(f"""<div class="metric-box">
                <div class="val">{cbd:.2f}</div><div class="lbl">CBR @ P{pct_:.0f} (%)</div>
            </div>""",unsafe_allow_html=True)
            with m2: st.markdown(f"""<div class="metric-box">
                <div class="val">{mr_auto:,.0f}</div><div class="lbl">Mr (psi)</div>
            </div>""",unsafe_allow_html=True)
            with m3: st.markdown(f"""<div class="metric-box">
                <div class="val">{k_auto:.1f}</div><div class="lbl">k_sub (pci)</div>
            </div>""",unsafe_allow_html=True)
            d_cbr=st.number_input("CBR ออกแบบ (ปรับได้)",value=float(round(cbd,1)),
                                   min_value=0.5,max_value=100.0,step=0.5,key="d_cbr")
            mr_d=cbr_to_mr(d_cbr); k_d=mr_to_k(mr_d)
            st.markdown(f'<div class="result-info">CBR={d_cbr:.1f}% → Mr=<b>{mr_d:,.0f} psi</b> → k=<b>{k_d:.1f} pci</b></div>',
                       unsafe_allow_html=True)
            if st.button("✅ ใช้ค่านี้",type="primary",key="use_cbr"):
                ss.cbr_design=d_cbr; ss.mr_psi=mr_d; ss.k_sub_pci=k_d
                st.success("✅ บันทึกแล้ว → Tab Flexible, K-Value, Rigid")
            st.markdown('</div>',unsafe_allow_html=True)

    with cr:
        if ss.cbr_values:
            arr,n,u,upct=calc_percentile(ss.cbr_values)
            cbd=float(np.interp(pct_,upct[::-1],u[::-1]))
            st.markdown('<div class="card"><h4>📈 กราฟ Percentile vs CBR</h4>',unsafe_allow_html=True)
            fig_c=go.Figure()
            fig_c.add_trace(go.Scatter(x=u,y=upct,mode='lines+markers',name='CBR',
                line=dict(color='#2E7D32',width=2.5),marker=dict(size=7,symbol='x',color='#1B5E20')))
            fig_c.add_trace(go.Scatter(x=[0,cbd],y=[pct_,pct_],mode='lines',
                line=dict(color='red',width=2,dash='dash'),name=f'P{pct_:.0f}%'))
            fig_c.add_trace(go.Scatter(x=[cbd,cbd],y=[0,pct_],mode='lines',
                line=dict(color='red',width=2,dash='dash'),name=f'CBR={cbd:.2f}%'))
            fig_c.add_annotation(x=cbd,y=0,text=f"<b>{cbd:.2f}%</b>",
                showarrow=True,arrowhead=2,arrowcolor='red',font=dict(size=13,color='red'),ay=40)
            fig_c.update_layout(
                xaxis_title="CBR (%)",yaxis_title="Percentile (%)",
                plot_bgcolor='white',height=360,
                xaxis=dict(range=[0,max(u)*1.1],gridcolor='#E8F5E9'),
                yaxis=dict(range=[0,100],gridcolor='#E8F5E9'),
                margin=dict(l=50,r=20,t=20,b=50),showlegend=False)
            st.plotly_chart(fig_c,use_container_width=True)
            st.markdown('</div>',unsafe_allow_html=True)
            # Stats
            st.markdown('<div class="card"><h4>📋 สถิติ</h4>',unsafe_allow_html=True)
            s1,s2,s3,s4=st.columns(4)
            with s1: st.metric("n",n)
            with s2: st.metric("Min",f"{np.min(ss.cbr_values):.2f}%")
            with s3: st.metric("Max",f"{np.max(ss.cbr_values):.2f}%")
            with s4: st.metric("Mean",f"{np.mean(ss.cbr_values):.2f}%")
            st.markdown('</div>',unsafe_allow_html=True)

# ══════════════════════════════════════════════
#  TAB 3: FLEXIBLE DESIGN
# ══════════════════════════════════════════════
with tab3:
    st.markdown("### 🔧 Flexible Pavement Design — AASHTO 1993")
    fl,fr=st.columns([1,1])
    with fl:
        # Auto-fill params
        st.markdown('<div class="card"><h4>📥 พารามิเตอร์ (ดึงอัตโนมัติ)</h4>',unsafe_allow_html=True)
        st.markdown(badge('esal_flex','ESAL Flex'),unsafe_allow_html=True)
        st.markdown(badge('mr_psi','Mr (CBR)'),unsafe_allow_html=True)
        st.markdown(f"**R0 = {ss.r0}%** (ZR = {ZR_MAP[ss.r0]}) | **Pt = {ss.pt}**")
        # ESAL selector
        if ss.esal_flex:
            sn_keys=list(ss.esal_flex.keys())
            sel=st.selectbox("เลือก SN",range(len(sn_keys)),
                format_func=lambda i:f"SN {sn_keys[i]} → ESAL={ss.esal_flex[sn_keys[i]]:,.0f}",
                key="f3_sn")
            d_esal=ss.esal_flex[sn_keys[sel]]
            st.markdown(f'<div class="result-info">Design ESAL = <b>{d_esal:,.0f}</b></div>',
                       unsafe_allow_html=True)
        else:
            st.warning("⚠️ คำนวณ ESAL ใน Tab 1 ก่อน")
            d_esal=st.number_input("Design ESAL",value=0,step=100000,key="f3_esal_m")
        # Mr
        mr_f=st.number_input("Mr (psi)",value=float(ss.mr_psi) if ss.mr_psi else 4500.0,
                              step=500.0,min_value=500.0,key="f3_mr")
        st.markdown('</div>',unsafe_allow_html=True)
        # Design params
        st.markdown('<div class="card"><h4>⚙️ พารามิเตอร์ Flexible</h4>',unsafe_allow_html=True)
        fp1,fp2=st.columns(2)
        with fp1:
            pi_f=st.number_input("Pi (Flexible)",value=4.2,step=0.1,key="f3_pi")
        with fp2:
            so_f=st.number_input("So (Flexible)",value=0.45,step=0.01,
                                  min_value=0.3,max_value=0.6,key="f3_so")
        st.markdown('</div>',unsafe_allow_html=True)

    with fr:
        st.markdown('<div class="card"><h4>🔩 Layer Design</h4>',unsafe_allow_html=True)
        # Header
        hc=st.columns([3.5,1.2,0.8,0.8,0.9,1.1])
        for txt,col in zip(["**วัสดุ**","**หนา(cm)**","**ai**","**mi**","**SNi**","**ΣSNi**"],hc):
            col.markdown(txt)

        mat_opts=list(FLEX_MATERIALS.keys())
        layers_f=[]; cum_sn=0.0

        for li in range(6):
            lc0,lc1=st.columns([3.5,1.2])
            with lc0:
                mat=st.selectbox(f"L{li+1}",mat_opts,key=f"f3m_{li}",label_visibility="collapsed")
            with lc1:
                h=st.number_input("cm",value=0,step=1,min_value=0,
                                   key=f"f3h_{li}",label_visibility="collapsed")
            if mat!="None" and h>0:
                ai,mi=FLEX_MATERIALS[mat]
                h_in=h/2.54; sni=ai*h_in*mi; cum_sn+=sni
                layers_f.append({'layer':li+1,'mat':mat,'h':h,'ai':ai,'mi':mi,
                                 'sni':round(sni,3),'cum_sn':round(cum_sn,3)})
                # แถวผล — อยู่ในแถวเดียวกัน
                rc=st.columns([3.5,1.2,0.8,0.8,0.9,1.1])
                rc[0].markdown(f'<div class="layer-row">{h} cm</div>',unsafe_allow_html=True)
                rc[1].markdown(f'<div class="layer-row">{h} cm</div>',unsafe_allow_html=True)
                rc[2].markdown(f'<div class="layer-row">{ai:.2f}</div>',unsafe_allow_html=True)
                rc[3].markdown(f'<div class="layer-row">{mi:.1f}</div>',unsafe_allow_html=True)
                rc[4].markdown(f'<div class="layer-row">{sni:.3f}</div>',unsafe_allow_html=True)
                rc[5].markdown(f'<div class="layer-row"><b>{cum_sn:.3f}</b></div>',unsafe_allow_html=True)

        st.markdown(f'<div class="result-info">ΣSN Provided = <b>{cum_sn:.3f}</b></div>',
                   unsafe_allow_html=True)
        st.markdown('</div>',unsafe_allow_html=True)

        if st.button("✅ Design Check",type="primary",key="f3_check"):
            if d_esal<=0:
                st.warning("⚠️ ใส่ Design ESAL ก่อน")
            else:
                zr_f=ZR_MAP[ss.r0]
                sn_req=sn_required(d_esal,zr_f,so_f,pi_f,ss.pt,mr_f)
                if sn_req:
                    passed=cum_sn>=sn_req; margin=cum_sn-sn_req
                    css="result-pass" if passed else "result-fail"
                    chk="✅ PASS" if passed else "❌ FAIL"
                    c1,c2,c3=st.columns(3)
                    with c1: st.markdown(f"""<div class="metric-box">
                        <div class="val">{cum_sn:.3f}</div><div class="lbl">SN Provided</div></div>""",unsafe_allow_html=True)
                    with c2: st.markdown(f"""<div class="metric-box">
                        <div class="val">{sn_req:.3f}</div><div class="lbl">SN Required</div></div>""",unsafe_allow_html=True)
                    with c3:
                        clr='#1B5E20' if passed else '#B71C1C'
                        st.markdown(f"""<div class="metric-box">
                        <div class="val" style="color:{clr}">{margin:+.3f}</div>
                        <div class="lbl">Safety Margin</div></div>""",unsafe_allow_html=True)
                    st.markdown(f'<div class="{css}" style="margin-top:0.7rem;">'
                                f'<b>{chk}</b> — SN Req={sn_req:.3f} | SN Prov={cum_sn:.3f}</div>',
                               unsafe_allow_html=True)
                    ss.flex_results={'esal':d_esal,'sn_req':sn_req,'sn_prov':cum_sn,
                                     'pass':passed,'layers':layers_f,'mr_psi':mr_f}
                else:
                    st.error("ไม่สามารถคำนวณ SN Required ได้")

# ══════════════════════════════════════════════
#  TAB 4: K-VALUE NOMOGRAPH
# ══════════════════════════════════════════════
with tab4:
    st.markdown("### 📐 K-Value Nomograph — AASHTO 1993")
    sub_k,sub_ls=st.tabs(["📊 Composite k∞ (Fig.3.3)","📉 Loss of Support (Fig.3.4)"])

    LS_PRESETS={0.0:(138,715,753,84),0.5:(129,728,908,0),1.0:(150,718,903,84),
                1.5:(153,721,928,138),2.0:(164,718,929,220),3.0:(212,719,929,328)}

    with sub_k:
        up1=st.file_uploader("📂 Upload Figure 3.3",type=['png','jpg','jpeg'],key='up_k')
        if up1: raw1=up1.read(); ss['img1_original']=raw1
        elif ss.get('img1_original'): raw1=ss['img1_original']
        else: raw1=None

        if raw1:
            img1=PILImage.open(io.BytesIO(raw1)).convert("RGB")
            w1,h1=img1.size; d1=PILDraw.Draw(img1.copy())
            img1_draw=img1.copy(); draw1=PILDraw.Draw(img1_draw)
            kc1,ki1=st.columns([1,2])
            with kc1:
                st.markdown('<div class="card"><h4>⚙️ ปรับเส้นอ่านค่า</h4>',unsafe_allow_html=True)
                with st.expander("1. Turning Line (เขียว)",expanded=True):
                    gx1=st.slider("X เริ่ม",0,w1,ss.get('gx1',int(w1*0.40)),key="gx1")
                    gy1=st.slider("Y เริ่ม",0,h1,ss.get('gy1',int(h1*0.45)),key="gy1")
                    gx2=st.slider("X จบ",0,w1,ss.get('gx2',int(w1*0.45)),key="gx2")
                    gy2=st.slider("Y จบ",0,h1,ss.get('gy2',int(h1*0.52)),key="gy2")
                    draw1.line([(gx1,gy1),(gx2,gy2)],fill="green",width=5)
                    slope_g=(gy2-gy1)/(gx2-gx1) if gx2!=gx1 else 0
                with st.expander("2. พารามิเตอร์ (ส้ม/แดง/น้ำเงิน)",expanded=True):
                    sx =st.slider("ตำแหน่งแกน D_SB",0,w1,ss.get('s1_sx',int(w1*0.15)),key="s1_sx")
                    sey=st.slider("ระดับ ESB (บน)",0,h1,ss.get('s1_sy_esb',int(h1*0.10)),key="s1_sy_esb")
                    smy=st.slider("ระดับ MR (ล่าง)",0,h1,ss.get('s1_sy_mr',int(h1*0.55)),key="s1_sy_mr")
                    cx=int(gx1+(smy-gy1)/slope_g) if slope_g!=0 else gx1
                draw_arrow(draw1,(sx,sey),(cx,sey),"orange")
                draw_arrow(draw1,(sx,sey),(sx,smy),"red")
                draw_arrow(draw1,(sx,smy),(cx,smy),"darkblue")
                draw_arrow(draw1,(cx,smy),(cx,sey),"blue")
                r_=8; draw1.ellipse([(cx-r_,smy-r_),(cx+r_,smy+r_)],fill="black",outline="white")

                st.markdown('</div>',unsafe_allow_html=True)
                st.markdown('<div class="card"><h4>📝 บันทึกค่า</h4>',unsafe_allow_html=True)
                # Auto-fill จาก session
                mr_k =st.number_input("MR (psi)",value=int(ss.mr_psi) if ss.mr_psi else 7000,step=500,key="nomo_mr")
                esb_k=st.number_input("ESB (psi)",value=int(ss.nomo_esb),step=1000,key="nomo_esb")
                dsb_k=st.number_input("DSB (in)",value=float(ss.nomo_dsb),step=0.5,key="nomo_dsb")
                k_inf_read=st.number_input("k∞ ที่อ่านได้ (pci)",value=int(ss.get('nomo_k_inf',400)),step=10,key="nomo_k_inf")
                if ss.mr_psi: st.markdown(f'<div class="badge-ready">📊 MR จาก CBR={ss.mr_psi:,.0f} psi</div>',unsafe_allow_html=True)
                if ss.nomo_esb: st.markdown(f'<div class="badge-ready">📐 ESB จาก Rigid={ss.nomo_esb:,} psi</div>',unsafe_allow_html=True)
                if ss.nomo_dsb: st.markdown(f'<div class="badge-ready">📏 DSB จาก Rigid={ss.nomo_dsb:.1f} in</div>',unsafe_allow_html=True)
                if st.button("✅ บันทึก k∞",type="primary",key="save_k"):
                    ss.k_inf=float(k_inf_read)
                    buf1=io.BytesIO(); img1_draw.save(buf1,format='PNG')
                    ss['nomograph_img_k']=buf1.getvalue()
                    ss['img1_original']=raw1
                    st.success(f"k∞ = {k_inf_read} pci → Tab Loss of Support & Rigid")
                st.markdown('</div>',unsafe_allow_html=True)
            with ki1:
                if ss.k_inf>0: st.markdown(f'<div class="result-pass">k∞ = <b>{ss.k_inf:.0f} pci</b></div>',unsafe_allow_html=True)
                buf1b=io.BytesIO(); img1_draw.save(buf1b,format='PNG')
                ss['nomograph_img_k']=buf1b.getvalue()
                st.image(img1_draw,caption="Composite k∞ Nomograph (AASHTO 1993 Fig.3.3)",use_container_width=True)
        else:
            st.markdown('<div class="result-warn">👆 Upload รูป <b>Figure 3.3</b> หรือกรอก k∞ โดยตรง</div>',unsafe_allow_html=True)
            st.markdown('<div class="card"><h4>📝 กรอก k∞ โดยตรง</h4>',unsafe_allow_html=True)
            k_m=st.number_input("k∞ (pci)",value=float(ss.k_inf) if ss.k_inf else 200.0,step=10.0,key="k_manual")
            if st.button("✅ ใช้ค่านี้",key="use_k_m"):
                ss.k_inf=k_m; st.success(f"k∞={k_m:.0f} pci")
            st.markdown('</div>',unsafe_allow_html=True)

    with sub_ls:
        st.markdown(f'<div class="result-info">k∞ = <b>{ss.k_inf:.0f} pci</b></div>',unsafe_allow_html=True)
        up2=st.file_uploader("📂 Upload Figure 3.4",type=['png','jpg','jpeg'],key='up_ls')
        if up2: raw2=up2.read(); ss['img2_original']=raw2
        elif ss.get('img2_original'): raw2=ss['img2_original']
        else: raw2=None

        if raw2:
            img2=PILImage.open(io.BytesIO(raw2)).convert("RGB")
            w2,h2=img2.size; img2_draw=img2.copy(); draw2=PILDraw.Draw(img2_draw)
            lc2,li2=st.columns([1,2])
            with lc2:
                st.markdown('<div class="card"><h4>⚙️ กำหนดเส้น LS</h4>',unsafe_allow_html=True)
                ls_opts=[0.0,0.5,1.0,1.5,2.0,3.0]
                cur_ls=ss.get('ls_select_box',1.0)
                ls_sel=st.selectbox("ค่า LS",ls_opts,
                                    index=ls_opts.index(cur_ls) if cur_ls in ls_opts else 2,
                                    key="ls_select_box")
                if ss.get('_last_ls')!=ls_sel:
                    ss['_last_ls']=ls_sel
                    c_=LS_PRESETS.get(ls_sel,(150,718,903,84))
                    ss['_ls_x1'],ss['_ls_y1']=c_[0],c_[1]
                    ss['_ls_x2'],ss['_ls_y2']=c_[2],c_[3]
                with st.expander("ปรับละเอียด",expanded=False):
                    lx1=st.slider("เริ่ม X",-100,w2+100,ss.get('_ls_x1',150),key="_ls_x1")
                    ly1=st.slider("เริ่ม Y",-100,h2+100,ss.get('_ls_y1',718),key="_ls_y1")
                    lx2=st.slider("จบ X",-100,w2+100,ss.get('_ls_x2',903),key="_ls_x2")
                    ly2=st.slider("จบ Y",-100,h2+100,ss.get('_ls_y2',84),key="_ls_y2")
                draw2.line([(lx1,ly1),(lx2,ly2)],fill="red",width=6)
                m_r=(ly2-ly1)/(lx2-lx1) if lx2!=lx1 else None
                c_r=ly1-m_r*lx1 if m_r else 0
                with st.expander("ตำแหน่งแกน",expanded=True):
                    ax_l=st.number_input("แกน Y ซ้าย",value=ss.get('axis_left',100),step=5,key="axis_left")
                    ax_b=st.number_input("แกน X ล่าง",value=ss.get('axis_bottom',h2-50),step=5,key="axis_bottom")
                st.caption(f"k∞ = {ss.k_inf:.0f} pci")
                kpx=st.slider("ตำแหน่ง k บนแกน X",0,w2,ss.get('k_pos_x',w2//2),key="k_pos_x")
                iy=int(m_r*kpx+c_r) if m_r else h2//2
                draw2.line([(kpx,ax_b),(kpx,iy)],fill="blue",width=5)
                draw_arrow(draw2,(kpx,iy),(int(ax_l),iy),"blue")
                draw2.ellipse([(kpx-8,iy-8),(kpx+8,iy+8)],fill="black",outline="white",width=2)
                st.markdown('</div>',unsafe_allow_html=True)
                st.markdown('<div class="card"><h4>📝 บันทึก k_eff</h4>',unsafe_allow_html=True)
                kc_val=st.number_input("k_eff Corrected (pci)",
                    value=int(ss.get('k_corrected',max(10,ss.k_inf-100)) if ss.k_inf else 200),
                    step=10,min_value=10,key="k_corr_input")
                if kc_val>ss.k_inf>0:
                    st.warning(f"⚠️ k_eff >{ss.k_inf:.0f} pci (k∞)")
                if st.button("✅ บันทึก k_eff",type="primary",key="save_keff"):
                    ss.k_corrected=float(kc_val); ss.ls_value=ls_sel
                    buf2=io.BytesIO(); img2_draw.save(buf2,format='PNG')
                    ss['nomograph_img_ls']=buf2.getvalue()
                    ss['img2_original']=raw2
                    st.success(f"k_eff={kc_val} pci → Tab Rigid Design")
                st.markdown('</div>',unsafe_allow_html=True)
            with li2:
                if ss.k_corrected>0: st.markdown(f'<div class="result-pass">k_eff = <b>{ss.k_corrected:.0f} pci</b> (LS={ls_sel})</div>',unsafe_allow_html=True)
                buf2b=io.BytesIO(); img2_draw.save(buf2b,format='PNG')
                ss['nomograph_img_ls']=buf2b.getvalue()
                st.image(img2_draw,caption=f"Loss of Support (LS={ls_sel})",use_container_width=True)
        else:
            st.markdown('<div class="result-warn">👆 Upload รูป <b>Figure 3.4</b> หรือกรอก k_eff โดยตรง</div>',unsafe_allow_html=True)
            st.markdown('<div class="card"><h4>📝 กรอก k_eff โดยตรง</h4>',unsafe_allow_html=True)
            ls_m=st.select_slider("LS",options=[0.0,0.5,1.0,1.5,2.0,3.0],
                                   value=ss.ls_value if ss.ls_value in [0.0,0.5,1.0,1.5,2.0,3.0] else 1.0)
            ke_m=st.number_input("k_eff (pci)",value=float(ss.k_corrected) if ss.k_corrected else 200.0,
                                  step=10.0,min_value=10.0,key="ke_manual")
            if st.button("✅ ใช้ค่านี้",key="use_ke_m"):
                ss.k_corrected=ke_m; ss.ls_value=ls_m
                st.success(f"k_eff={ke_m:.0f} pci")
            st.markdown('</div>',unsafe_allow_html=True)

# ══════════════════════════════════════════════
#  TAB 5: RIGID DESIGN
# ══════════════════════════════════════════════
with tab5:
    st.markdown("### 🏗️ Rigid Pavement Design — AASHTO 1993")

    # Status
    sc1,sc2,sc3=st.columns(3)
    with sc1: st.markdown(badge('esal_rigid','ESAL Rigid'),unsafe_allow_html=True)
    with sc2: st.markdown(badge('k_corrected','k_eff'),unsafe_allow_html=True)
    with sc3: st.markdown(badge('mr_psi','CBR/Mr'),unsafe_allow_html=True)
    st.markdown("---")

    col5l,col5r=st.columns([1.1,1])

    # ════ คอลัมน์ซ้าย — Layer Editor ════
    with col5l:
        st.markdown('<div class="card"><h4>🔩 ชั้นโครงสร้าง (JPCP/JRCP/CRCP ใช้ร่วมกัน)</h4>',unsafe_allow_html=True)

        # Header
        rh=st.columns([3,1.2,1.5,0.4])
        rh[0].markdown("**วัสดุ**"); rh[1].markdown("**หนา(cm)**")
        rh[2].markdown("**E (MPa)**"); rh[3].markdown("")

        r_layers=[]; total_h=0.0; e_eq_psi=0.0; dsb_in=0.0

        for li in range(6):
            rc0,rc1,rc2,rc3=st.columns([3,1.2,1.5,0.4])
            with rc0:
                mat=st.selectbox(f"R{li+1}",RIGID_MAT_NAMES,key=f"r5m_{li}",label_visibility="collapsed")
            with rc1:
                h_cm=st.number_input("cm",value=0,step=1,min_value=0,
                                     key=f"r5h_{li}",label_visibility="collapsed")
            with rc2:
                e_def=RIGID_MATERIALS[mat]["e_mpa"] if mat!="None" and mat in RIGID_MATERIALS else 0
                # Track material change → reset E
                pk=f"_r5pm_{li}"
                if ss.get(pk)!=mat:
                    ss[pk]=mat; ss[f"r5e_{li}"]=e_def
                e_mpa=st.number_input("MPa",value=ss.get(f"r5e_{li}",e_def),
                                       step=50,min_value=0,key=f"r5e_{li}",
                                       label_visibility="collapsed",
                                       disabled=(mat=="None" or h_cm==0))
            with rc3:
                if mat!="None" and h_cm>0: st.markdown("✅")

            if mat!="None" and h_cm>0 and e_mpa>0:
                r_layers.append({"name":mat,"h_cm":h_cm,"e_mpa":e_mpa})
                total_h+=h_cm
                # แถวผล — แถวเดียวกัน inline
                ri=st.columns([3,1.2,1.5,0.4])
                ri[0].markdown(f'<div class="layer-row">{mat[:30]}</div>',unsafe_allow_html=True)
                ri[1].markdown(f'<div class="layer-row">{h_cm} cm</div>',unsafe_allow_html=True)
                ri[2].markdown(f'<div class="layer-row">{e_mpa:,} MPa</div>',unsafe_allow_html=True)
                ri[3].markdown("")

        # คำนวณ E_eq
        if r_layers:
            sum_hE=sum(l["h_cm"]*(l["e_mpa"]**(1/3)) for l in r_layers)
            e_eq_mpa=(sum_hE/total_h)**3
            e_eq_psi=e_eq_mpa*145.038; dsb_in=total_h/2.54
            ss['nomo_esb']=int(e_eq_psi); ss['nomo_dsb']=round(dsb_in,2)
            m1,m2,m3,m4=st.columns(4)
            for col_,val_,lbl_ in [(m1,f"{total_h:.0f}","รวม (cm)"),
                                    (m2,f"{dsb_in:.1f}","DSB (in)"),
                                    (m3,f"{e_eq_psi:,.0f}","E_eq (psi)"),
                                    (m4,f"{e_eq_mpa:.1f}","E_eq (MPa)")]:
                col_.markdown(f'<div class="metric-box"><div class="val" style="font-size:1.1rem;">{val_}</div><div class="lbl">{lbl_}</div></div>',unsafe_allow_html=True)
            st.markdown(f'<div class="result-info" style="font-size:0.82rem;">→ Tab K-Value: ESB=<b>{e_eq_psi:,.0f} psi</b> | DSB=<b>{dsb_in:.1f} in</b></div>',unsafe_allow_html=True)
        st.markdown('</div>',unsafe_allow_html=True)

    # ════ คอลัมน์ขวา — Parameters & Calc ════
    with col5r:
        st.markdown('<div class="card"><h4>⚙️ พารามิเตอร์ออกแบบ</h4>',unsafe_allow_html=True)
        rp1,rp2=st.columns(2)
        with rp1:
            fc=st.number_input("f'c (ksc)",value=350,step=10,min_value=200,key="r5_fc")
            fc_c=0.8*fc; fc_p=fc_c*14.223; ec_p=57000*math.sqrt(fc_p)
            sc_a=min(600,10.0*math.sqrt(fc_p))
            sc=st.number_input("Sc (psi)",value=int(sc_a),step=10,min_value=100,max_value=700,key="r5_sc")
        with rp2:
            pi_r=st.number_input("Pi (Rigid)",value=4.5,step=0.1,key="r5_pi")
            so_r=st.number_input("So (Rigid)",value=0.35,step=0.01,min_value=0.2,max_value=0.5,key="r5_so")
        rp3,rp4=st.columns(2)
        with rp3:
            cd=st.number_input("Cd (Drainage)",value=1.0,step=0.05,min_value=0.5,max_value=1.25,key="r5_cd")
            j_jr=st.number_input("J – JPCP/JRCP",value=2.8,step=0.1,min_value=1.0,max_value=5.0,key="r5_j_jr")
        with rp4:
            j_cr=st.number_input("J – CRCP",value=2.6,step=0.1,min_value=1.0,max_value=5.0,key="r5_j_crcp")
            st.markdown(f"<small>Ec = {ec_p:,.0f} psi</small>",unsafe_allow_html=True)
            st.markdown(f"<small>ZR = {ZR_MAP[ss.r0]} (R0={ss.r0}%)</small>",unsafe_allow_html=True)
        st.markdown('</div>',unsafe_allow_html=True)

        # ESAL & k_eff
        st.markdown('<div class="card"><h4>📊 Design ESAL & k_eff</h4>',unsafe_allow_html=True)
        # ESAL: แต่ละ Slab ใช้ค่าของตัวเอง — ใช้ slab 30cm เป็น representative display
        _esal_30=int(ss.esal_rigid.get(30,ss.esal_rigid.get(30.0,0)))
        if _esal_30>0:
            st.markdown(f'<div class="badge-ready">📊 ESAL (Slab 30cm ref) = {_esal_30:,}</div>',unsafe_allow_html=True)
        else:
            st.markdown('<div class="badge-wait">⚠️ คำนวณ ESAL ใน Tab 1 ก่อน</div>',unsafe_allow_html=True)
        st.caption("หมายเหตุ: แต่ละ Slab จะใช้ ESAL ของตัวเองในการคำนวณ")

        _k=float(ss.k_corrected) if ss.k_corrected else 0.0
        if _k>0:
            st.markdown(f'<div class="badge-ready">📐 k_eff = {_k:.0f} pci</div>',unsafe_allow_html=True)
            # update widget
            if ss.get('_r5k_prev')!=_k: ss['_r5k_prev']=_k; ss['r5_keff']=_k
        else:
            st.markdown('<div class="badge-wait">⚠️ คำนวณ k_eff ใน Tab K-Value ก่อน</div>',unsafe_allow_html=True)
        k_eff=st.number_input("k_eff (pci) — แก้ไขได้",
                               value=ss.get('r5_keff',_k if _k>0 else 200.0),
                               step=10.0,min_value=10.0,key="r5_keff")
        st.markdown('</div>',unsafe_allow_html=True)

        # ── คำนวณปุ่มเดียว ──
        if st.button("✅ คำนวณ JPCP/JRCP + CRCP × ทุก Slab",
                     type="primary",use_container_width=True,key="r5_calc"):
            any_esal=any(v>0 for v in ss.esal_rigid.values())
            if not any_esal:
                st.warning("⚠️ ใส่ ESAL ก่อน (Tab 1)")
            elif k_eff<=0:
                st.warning("⚠️ ใส่ k_eff ก่อน")
            else:
                zr_r=ZR_MAP[ss.r0]
                results={}
                for D in SLAB_THICKNESSES:
                    # ESAL ตรงตาม Slab
                    esal_D=int(ss.esal_rigid.get(D,ss.esal_rigid.get(int(D),
                               ss.esal_rigid.get(float(D),_esal_30))))
                    results[D]={}
                    for gkey,j_val in [("JPCP_JRCP",j_jr),("CRCP",j_cr)]:
                        wc=w18_rigid(D,pi_r,ss.pt,zr_r,so_r,sc,cd,j_val,ec_p,k_eff)
                        results[D][gkey]={
                            "w18_cap":wc or 0,
                            "pass":(wc is not None and wc>=esal_D),
                            "esal":esal_D,"j":j_val
                        }
                # min slab per group
                min_slab={}
                for gk in ["JPCP_JRCP","CRCP"]:
                    passed=[D for D in SLAB_THICKNESSES if results[D][gk]["pass"]]
                    min_slab[gk]=min(passed) if passed else None
                # best type
                valid_m={k:v for k,v in min_slab.items() if v is not None}
                best_g=min(valid_m,key=lambda k:valid_m[k]) if valid_m else None

                ss.rigid_results={
                    "results":results,"min_slab":min_slab,
                    "best_group":best_g,
                    "design_esal":_esal_30,"k_eff":k_eff,
                    "fc":fc,"sc":sc,"j_jr":j_jr,"j_crcp":j_cr,"cd":cd,
                    "ec_psi":ec_p,"layers":r_layers,
                    "e_eq_psi":e_eq_psi,"dsb_in":dsb_in,
                }
                st.rerun()

    # ════ ผลลัพธ์ตารางเปรียบเทียบ ════
    rr=ss.get("rigid_results",{})
    if rr and "results" in rr:
        results=rr["results"]; min_slab=rr["min_slab"]; best_g=rr.get("best_group")
        de=rr.get("design_esal",0); kv=rr.get("k_eff",0)

        st.markdown("---")
        st.markdown("#### 📊 ผลการออกแบบ — เปรียบเทียบทุก Slab")
        st.markdown(f'<div class="result-info">W18 Required (ref Slab 30cm) = <b>{de:,.0f}</b> &nbsp;|&nbsp; k_eff = <b>{kv:.0f} pci</b></div>',unsafe_allow_html=True)

        # HTML Table
        html_r=""
        for D,lbl in zip(SLAB_THICKNESSES,SLAB_LABELS):
            html_r+='<tr>'
            html_r+=f'<td style="padding:7px 10px;font-weight:600;border:1px solid #E0E0E0;">{lbl}</td>'
            for gk,gcol in [("JPCP_JRCP","#1565C0"),("CRCP","#B71C1C")]:
                rv=results[D][gk]
                wc=rv["w18_cap"]; passed=rv["pass"]; esal_D=rv["esal"]
                is_min=min_slab.get(gk)==D
                is_best=best_g==gk and is_min
                bg="#C8E6C9" if (passed and is_min) else ("#F1F8E9" if passed else "#FFEBEE")
                bdr=f"3px solid #FFD600" if is_best else "1px solid #E0E0E0"
                clr="#1B5E20" if passed else "#B71C1C"
                fw="700" if is_min else "400"
                icon="★ " if is_min else ""
                trph=" 🏆" if is_best else ""
                st_="PASS" if passed else "FAIL"
                ws=f"{wc/1e6:.2f}M" if wc>0 else "—"
                mg=f"+{(wc/esal_D-1)*100:.1f}%" if (passed and esal_D>0 and wc>0) else ""
                html_r+=(f'<td style="background:{bg};border:{bdr};padding:7px 10px;'
                         f'text-align:center;color:{clr};font-weight:{fw};">'
                         f'{icon}{st_}{trph}<br>'
                         f'<small style="color:#555;">W18={ws} {mg}</small><br>'
                         f'<small style="color:#888;">ESAL={esal_D/1e6:.2f}M</small>'
                         f'</td>')
            html_r+='</tr>'

        st.markdown(f"""<table style="width:100%;border-collapse:collapse;
                font-family:'Sarabun',sans-serif;font-size:0.9rem;margin:0.5rem 0;">
            <thead><tr style="background:#1B5E20;color:white;">
                <th style="padding:8px 10px;text-align:left;">Slab</th>
                <th style="padding:8px 10px;text-align:center;color:#90CAF9;">
                    🟦 JPCP / JRCP<br><small>J={rr.get('j_jr',2.8)}</small></th>
                <th style="padding:8px 10px;text-align:center;color:#EF9A9A;">
                    🟥 CRCP<br><small>J={rr.get('j_crcp',2.6)}</small></th>
            </tr></thead>
            <tbody>{html_r}</tbody></table>
            <div style="font-size:0.75rem;color:#666;margin-top:0.3rem;">
            ★ = ความหนาน้อยสุดที่ PASS | 🏆 = ประหยัดที่สุด |
            W18=Capacity | ESAL=Required (ตาม Slab นั้น)</div>""",
            unsafe_allow_html=True)

        # Summary
        st.markdown("#### 🎯 ความหนาออกแบบที่เลือก")
        sc1,sc2=st.columns(2)
        for col_,gk,icon_,lbl_ in [(sc1,"JPCP_JRCP","🟦","JPCP / JRCP"),(sc2,"CRCP","🟥","CRCP")]:
            with col_:
                dm=min_slab.get(gk)
                if dm is not None:
                    lbl_m=SLAB_LABELS[SLAB_THICKNESSES.index(dm)]
                    rv=results[dm][gk]
                    mg=(rv["w18_cap"]/rv["esal"]-1)*100 if rv["esal"]>0 else 0
                    trph=" 🏆" if best_g==gk else ""
                    bclr="#FFD600" if best_g==gk else "#C8E6C9"
                    st.markdown(f"""<div class="metric-box" style="border-color:{bclr};">
                        <div style="font-size:1rem;">{icon_} {lbl_}{trph}</div>
                        <div class="val" style="font-size:1.3rem;">{lbl_m}</div>
                        <div class="lbl">W18 Cap = {rv['w18_cap']:,.0f}</div>
                        <div class="lbl" style="color:#2E7D32;">Safety = {mg:+.1f}%</div>
                    </div>""",unsafe_allow_html=True)
                else:
                    st.markdown(f"""<div class="metric-box" style="border-color:#EF9A9A;">
                        <div style="font-size:1rem;">{icon_} {lbl_}</div>
                        <div class="val" style="color:#B71C1C;font-size:0.95rem;">ไม่มี Slab ที่ผ่าน</div>
                        <div class="lbl">พิจารณาเพิ่ม k_eff หรือ f'c</div>
                    </div>""",unsafe_allow_html=True)

        if rr.get("layers"):
            with st.expander("📋 ชั้นโครงสร้าง"):
                st.dataframe(pd.DataFrame(rr["layers"]),use_container_width=True,hide_index=True)

# ══════════════════════════════════════════════
#  TAB 6: REPORT & SAVE
# ══════════════════════════════════════════════
with tab6:
    st.markdown("### 📄 Report & Save")

    # Status
    sc=st.columns(5)
    for i,(k,lbl) in enumerate([('esal_rigid','ESAL'),('cbr_values','CBR'),
                                  ('flex_results','Flexible'),('k_corrected','K-Value'),
                                  ('rigid_results','Rigid')]):
        sc[i].markdown(badge(k,lbl),unsafe_allow_html=True)
    st.markdown("---")

    st.markdown("#### 📥 Download Word Report")
    st.markdown("**แยกส่วน:**")
    c1,c2,c3,c4,c5=st.columns(5)
    for col_,key_,fn_,fname_ in [
        (c1,'esal_rigid',report_esal,"ESAL_Report.docx"),
        (c2,'cbr_values',report_cbr,"CBR_Report.docx"),
        (c3,'flex_results',report_flex,"Flexible_Report.docx"),
        (c4,'k_corrected',report_kvalue,"KValue_Report.docx"),
        (c5,'rigid_results',report_rigid,"Rigid_Report.docx"),
    ]:
        with col_:
            if ss.get(key_):
                b=fn_(dict(ss))
                if b:
                    st.download_button(f"📥 {fname_.split('_')[0]}",b,fname_,
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True)
            else:
                st.button(f"📥 {fname_.split('_')[0]}",disabled=True,use_container_width=True)

    st.markdown("**รวมทุกส่วน:**")
    if st.button("🗂️ สร้างรายงานรวม",type="primary",use_container_width=True):
        b_full=report_full(dict(ss))
        if b_full:
            st.download_button("📥 Full Report",b_full,
                f"ITM_Pave_Report_{datetime.now().strftime('%Y%m%d')}.docx",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True)

    st.markdown("---")
    st.markdown("""<div style='text-align:center;color:#558B2F;font-size:0.82rem;padding:0.5rem;'>
        🛣️ <b>ITM Pave Pro v2.0</b> — AASHTO 1993 Pavement Design System<br>
        พัฒนาโดย รศ.ดร.อิทธิพล มีผล | ภาควิชาครุศาสตร์โยธา | คณะครุศาสตร์อุตสาหกรรม | มจพ.
    </div>""",unsafe_allow_html=True)
