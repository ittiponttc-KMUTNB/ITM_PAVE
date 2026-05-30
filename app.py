# ╔══════════════════════════════════════════════════════════════════╗
# ║  app.py — ITM Pave Pro v3.0                                     ║
# ║  Main Entry Point | Sidebar Navigation                          ║
# ║  พัฒนาโดย รศ.ดร.อิทธิพล มีผล | ภาควิชาครุศาสตร์โยธา มจพ.    ║
# ╚══════════════════════════════════════════════════════════════════╝

import streamlit as st

# ── Page Config (ต้องเป็น st call แรกสุด) ──
st.set_page_config(
    page_title="ITM Pave Pro – AASHTO 1993",
    page_icon="🛣️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Core UI ──
from ui.core import ss_init, inject_css, render_sidebar, workflow_bar

ss_init()
inject_css()
render_sidebar()

ss = st.session_state
page = ss.get('current_page', 'ESAL Calculator')

# ── Header ──
st.markdown(f"""
<div class="main-header">
    <h1>🛣️ ITM Pave Pro — ระบบออกแบบโครงสร้างชั้นทาง AASHTO 1993</h1>
    <p>ESAL Calculator · CBR Analysis · Flexible Design · Rigid Design · Report</p>
</div>
""", unsafe_allow_html=True)

# ── Workflow bar ──
workflow_bar(page)

# ── Route to page ──
if page == 'ESAL Calculator':
    from ui.tab1_esal import render
    render()

elif page == 'CBR Analysis':
    from ui.tab2_cbr import render
    render()

elif page == 'Flexible Design':
    from ui.tab3_flexible import render
    render()

elif page == 'Rigid Design':
    from ui.tab4_rigid import render
    render()

elif page == 'Report & Save':
    from ui.tab5_report import render
    render()
