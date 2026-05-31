# ╔══════════════════════════════════════════════════════════════════╗
# ║  ui/core.py — ITM Pave Pro                                      ║
# ║  Session State Init | CSS | Status Badge | Save/Load JSON       ║
# ╚══════════════════════════════════════════════════════════════════╝

import json
import streamlit as st
import pandas as pd
from datetime import datetime


# ─────────────────────────────────────────────
#  Session State Init
# ─────────────────────────────────────────────

def ss_init():
    defaults = {
        # Project
        'project_name':    '',
        # Traffic & ESAL
        'traffic_df':       None,
        'esal_rigid':       {},
        'esal_flex':        {},
        'ldf':              0.9,
        'ddf':              0.5,
        'pt_global':        2.5,
        'pt_rigid':         2.5,
        'pt_flex':          2.5,
        '_pt_sync':         2.5,
        'sn_list':          [6.5, 7.1, 7.5, 8.0],
        # CBR
        'cbr_values':       [],
        'cbr_percentile':   90.0,
        'cbr_design':       3.0,
        'mr_subgrade_psi':  4500.0,
        'k_subgrade_pci':   231.9,
        # Flexible
        'flex_results':     {},
        'cbr_fl_val':       3.0,
        'mr_fl_val':        4500.0,
        # K-Value / Nomograph
        'k_inf':            0.0,
        'k_corrected':      0.0,
        'ls_value':         1.0,
        'nomograph_img_k':  None,
        'nomograph_img_ls': None,
        'layer_esb_psi':    50000,
        'layer_dsb_in':     6.0,
        # Rigid
        'rigid_results':    {},
        # Navigation
        'current_page':     'ESAL Calculator',
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


# ─────────────────────────────────────────────
#  CSS
# ─────────────────────────────────────────────

def inject_css():
    st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;600;700&family=IBM+Plex+Mono:wght@400;600&display=swap');

html, body, [class*="css"] {
    font-family: 'Sarabun', sans-serif;
}

/* ── Header ── */
.main-header {
    background: linear-gradient(135deg, #0B1F3A 0%, #1565C0 60%, #1976D2 100%);
    color: white;
    padding: 1.2rem 2rem;
    border-radius: 12px;
    margin-bottom: 1.2rem;
    box-shadow: 0 4px 16px rgba(11,31,58,0.35);
    border-left: 6px solid #00B8D4;
}
.main-header h1 { margin:0; font-size:1.6rem; font-weight:700; letter-spacing:-0.5px; }
.main-header p  { margin:0.3rem 0 0; font-size:0.88rem; opacity:0.88; }

/* ── Cards ── */
.card {
    background: #fff;
    border: 1px solid #CBD5E1;
    border-left: 4px solid #1565C0;
    border-radius: 10px;
    padding: 1rem 1.3rem;
    margin-bottom: 1rem;
    box-shadow: 0 2px 8px rgba(21,101,192,0.07);
}
.card h4 { color:#0B1F3A; margin:0 0 0.8rem; font-size:1rem; font-weight:700; }

/* ── Status Badges ── */
.badge-ready { background:#E8F5E9; color:#2E7D32; border:1px solid #A5D6A7;
               border-radius:20px; padding:0.2rem 0.75rem;
               font-size:0.82rem; font-weight:600; display:inline-block; }
.badge-wait  { background:#FFF8E1; color:#E65100; border:1px solid #FFE082;
               border-radius:20px; padding:0.2rem 0.75rem;
               font-size:0.82rem; font-weight:600; display:inline-block; }
.badge-na    { background:#F5F5F5; color:#757575; border:1px solid #E0E0E0;
               border-radius:20px; padding:0.2rem 0.75rem;
               font-size:0.82rem; font-weight:600; display:inline-block; }

/* ── Result Boxes ── */
.result-pass { background:#E8F5E9; border:1px solid #A5D6A7; border-radius:8px;
               padding:0.7rem 1rem; color:#1B5E20; font-weight:600; margin:0.3rem 0; }
.result-fail { background:#FFEBEE; border:1px solid #EF9A9A; border-radius:8px;
               padding:0.7rem 1rem; color:#B71C1C; font-weight:600; margin:0.3rem 0; }
.result-info { background:#E3F2FD; border:1px solid #90CAF9; border-radius:8px;
               padding:0.7rem 1rem; color:#0D47A1; font-weight:600; margin:0.3rem 0; }
.result-warn { background:#FFF8E1; border:1px solid #FFE082; border-radius:8px;
               padding:0.7rem 1rem; color:#E65100; font-weight:600; margin:0.3rem 0; }

/* ── Metric Box ── */
.metric-box {
    background:#fff; border:1px solid #CBD5E1; border-radius:12px;
    padding:0.9rem; text-align:center;
    box-shadow:0 2px 8px rgba(21,101,192,0.08);
}
.metric-box .val { font-size:1.4rem; font-weight:700; color:#0B1F3A;
                   font-family:'IBM Plex Mono', monospace; }
.metric-box .lbl { font-size:0.78rem; color:#4A5568; margin-top:0.2rem; }

/* ── Number inputs ── */
.stNumberInput > div > div > input {
    font-family:'IBM Plex Mono', monospace; font-weight:600;
}

/* ── DataFrames ── */
.stDataFrame { border-radius:8px; overflow:hidden; }

/* ── Sidebar ── */
[data-testid="stSidebar"] { background:#0B1F3A; }
[data-testid="stSidebar"] * { color:#E8EEF4 !important; }
[data-testid="stSidebar"] hr { border-color:#1B3A5C; }

/* ── Buttons ── */
button[kind="primary"] {
    background:#1565C0 !important;
    border-radius:8px !important;
    font-weight:700 !important;
}

/* ── Flow Arrow ── */
.flow-arrow {
    text-align:center; font-size:1.5rem; color:#1565C0;
    margin:0.3rem 0; line-height:1;
}

/* ── Workflow Steps ── */
.wf-bar {
    display:flex; align-items:center; gap:4px;
    padding:6px 0; margin-bottom:12px;
    flex-wrap:wrap;
}
.wf-step-done {
    background:#E8F5E9; color:#2E7D32; border:1px solid #A5D6A7;
    padding:3px 12px; border-radius:12px;
    font-size:11px; font-weight:600;
}
.wf-step-active {
    background:#E3F2FD; color:#1565C0; border:1px solid #90CAF9;
    padding:3px 12px; border-radius:12px;
    font-size:11px; font-weight:700;
}
.wf-step-pending {
    background:#F5F7FA; color:#8A9BB0; border:1px solid #CBD5E1;
    padding:3px 12px; border-radius:12px;
    font-size:11px; font-weight:500;
}
.wf-arrow { color:#8A9BB0; font-size:10px; }

/* ── Tabs — Main & Sub ── */

/* Tab bar background */
[data-baseweb="tab-list"] {
    background: #F0F4F8;
    border-radius: 10px;
    padding: 4px;
    gap: 4px;
    border-bottom: none !important;
}

/* Tab ปกติ */
[data-baseweb="tab"] {
    background: transparent !important;
    border-radius: 8px !important;
    padding: 6px 18px !important;
    font-weight: 600 !important;
    font-size: 0.9rem !important;
    color: #4A5568 !important;
    border: none !important;
    transition: all 0.18s ease !important;
}

/* Tab hover */
[data-baseweb="tab"]:hover {
    background: #DBEAFE !important;
    color: #1565C0 !important;
}

/* Tab active */
[data-baseweb="tab"][aria-selected="true"] {
    background: #0B1F3A !important;
    color: #FFFFFF !important;
    border-radius: 8px !important;
    box-shadow: 0 2px 8px rgba(11,31,58,0.30) !important;
}

/* ซ่อน underline เส้นใต้ tab */
[data-baseweb="tab-highlight"],
[data-baseweb="tab-border"] {
    display: none !important;
}

/* Tab panel — ลด padding บน */
[data-baseweb="tab-panel"] {
    padding-top: 1rem !important;
}

</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
#  Status Badge Helper
# ─────────────────────────────────────────────

def status_badge(key, label=None):
    ss  = st.session_state
    val = ss.get(key)
    has = (val is not None and val != {} and val != [] and val != 0.0)
    cls  = "badge-ready" if has else "badge-wait"
    icon = "✅" if has else "⚠️"
    lbl  = label or key
    return f'<span class="{cls}">{icon} {lbl}</span>'


def workflow_bar(current_page: str):
    """แสดง workflow step indicator บนสุดของแต่ละหน้า"""
    pages = [
        ("🚛", "ESAL Calculator"),
        ("🌱", "CBR Analysis"),
        ("🔧", "Flexible Design"),
        ("🏗️", "Rigid Design"),
        ("📄", "Report & Save"),
    ]
    ss    = st.session_state
    ready = {
        "ESAL Calculator":  bool(ss.get('esal_rigid') or ss.get('esal_flex')),
        "CBR Analysis":     bool(ss.get('cbr_values')),
        "Flexible Design":  bool(ss.get('flex_results')),
        "Rigid Design":     bool(ss.get('rigid_results')),
        "Report & Save":    False,
    }

    parts = []
    for i, (icon, name) in enumerate(pages):
        if i > 0:
            parts.append('<span class="wf-arrow">›</span>')
        if name == current_page:
            parts.append(f'<span class="wf-step-active">{icon} {name}</span>')
        elif ready.get(name):
            parts.append(f'<span class="wf-step-done">{icon} {name} ✓</span>')
        else:
            parts.append(f'<span class="wf-step-pending">{icon} {name}</span>')

    st.markdown(
        f'<div class="wf-bar">{"".join(parts)}</div>',
        unsafe_allow_html=True
    )


# ─────────────────────────────────────────────
#  Sidebar
# ─────────────────────────────────────────────

def render_sidebar():
    ss = st.session_state

    with st.sidebar:
        # Logo
        st.markdown("""
        <div style='text-align:center;padding:1rem 0 0.8rem;'>
            <div style='font-size:2rem;'>🛣️</div>
            <div style='font-weight:800;font-size:1.15rem;
                        color:#00B8D4;letter-spacing:0.5px;'>ITM Pave Pro</div>
            <div style='font-size:0.75rem;color:#7EB3D8;
                        letter-spacing:2px;text-transform:uppercase;'>AASHTO 1993</div>
        </div>
        """, unsafe_allow_html=True)

        # Project name — ใช้ key= เพียงอย่างเดียว Streamlit จัดการ state เอง
        st.text_input("📁 ชื่อโครงการ",
                      key="project_name",
                      placeholder="กรอกชื่อโครงการ...")

        st.divider()

        # Navigation
        st.markdown("<div style='font-size:0.72rem;letter-spacing:2px;"
                    "color:#4A7FA5;text-transform:uppercase;"
                    "font-weight:600;margin-bottom:6px;'>การคำนวณ</div>",
                    unsafe_allow_html=True)

        pages = [
            ("🚛", "ESAL Calculator"),
            ("🌱", "CBR Analysis"),
            ("🔧", "Flexible Design"),
            ("🏗️", "Rigid Design"),
        ]
        for icon, name in pages:
            active = ss.get('current_page') == name
            if st.button(f"{icon}  {name}",
                         key=f"nav_{name}",
                         use_container_width=True,
                         type="primary" if active else "secondary"):
                ss['current_page'] = name
                st.rerun()

        st.markdown("<div style='font-size:0.72rem;letter-spacing:2px;"
                    "color:#4A7FA5;text-transform:uppercase;"
                    "font-weight:600;margin:12px 0 6px;'>ผลลัพธ์</div>",
                    unsafe_allow_html=True)

        if st.button("📄  Report & Save",
                     key="nav_report",
                     use_container_width=True,
                     type="primary" if ss.get('current_page') == 'Report & Save' else "secondary"):
            ss['current_page'] = 'Report & Save'
            st.rerun()

        st.divider()

        # Status panel
        st.markdown("**📊 สถานะข้อมูล**")
        for key, label in [
            ('esal_rigid',   '🚛 ESAL'),
            ('cbr_values',   '🌱 CBR'),
            ('flex_results', '🔧 Flexible'),
            ('k_corrected',  '📐 K-Value'),
            ('rigid_results','🏗️ Rigid'),
        ]:
            st.markdown(status_badge(key, label), unsafe_allow_html=True)

        st.divider()

        # Save / Load
        st.markdown("**💾 Save / Load Project**")
        if st.button("💾 Save JSON", use_container_width=True):
            _save_json(ss)

        uploaded_json = st.file_uploader("📂 Load JSON", type=['json'],
                                          key="load_json")
        if uploaded_json:
            _load_json(ss, uploaded_json)

        st.divider()

        # Footer
        st.markdown("""
        <div style='font-size:0.70rem;color:#4A7FA5;
                    text-align:center;line-height:1.9;'>
            รศ.ดร.อิทธิพล มีผล<br>
            ภาควิชาครุศาสตร์โยธา<br>
            คณะครุศาสตร์อุตสาหกรรม มจพ.<br>
            <b style='color:#00B8D4;'>ITM Pave Pro v3.0</b>
        </div>
        """, unsafe_allow_html=True)


# ─────────────────────────────────────────────
#  Save / Load JSON (private)
# ─────────────────────────────────────────────

def _save_json(ss):
    save_data = {
        'project_name':   ss.get('project_name', ''),
        'esal_rigid':     ss.esal_rigid,
        'esal_flex':      {str(k): v for k, v in ss.esal_flex.items()},
        'ldf':            ss.ldf,
        'ddf':            ss.ddf,
        'pt_global':      ss.pt_global,
        'pt_rigid':       ss.pt_rigid,
        'pt_flex':        ss.pt_flex,
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
        'traffic_df':     (ss.traffic_df.to_dict('records')
                           if ss.traffic_df is not None else None),
    }
    json_bytes = json.dumps(save_data, ensure_ascii=False, indent=2).encode('utf-8')
    st.download_button(
        "📥 Download JSON", json_bytes,
        file_name=f"itm_pave_{datetime.now().strftime('%Y%m%d_%H%M')}.json",
        mime="application/json",
        use_container_width=True,
    )


def _load_json(ss, uploaded_json):
    try:
        data = json.loads(uploaded_json.read().decode('utf-8'))
        keys = [
            'project_name', 'esal_rigid', 'ldf', 'ddf',
            'pt_global', 'pt_rigid', 'pt_flex', 'sn_list',
            'cbr_values', 'cbr_percentile', 'cbr_design',
            'mr_subgrade_psi', 'k_subgrade_pci', 'flex_results',
            'k_inf', 'k_corrected', 'ls_value', 'rigid_results',
        ]
        for k in keys:
            if k in data:
                ss[k] = data[k]
        if 'esal_flex' in data:
            ss.esal_flex = {float(k): v for k, v in data['esal_flex'].items()}
        if data.get('traffic_df'):
            ss.traffic_df = pd.DataFrame(data['traffic_df'])
        st.success("✅ โหลดข้อมูลสำเร็จ!")
        st.rerun()
    except Exception as e:
        st.error(f"❌ โหลดไม่สำเร็จ: {e}")
