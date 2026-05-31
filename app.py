# ╔══════════════════════════════════════════════════════════════════╗
# ║  app.py — ITM Pave Pro v3.0                                     ║
# ║  Main Entry Point | Top Tabs Navigation                         ║
# ║  พัฒนาโดย รศ.ดร.อิทธิพล มีผล | ภาควิชาครุศาสตร์โยธา มจพ.    ║
# ╚══════════════════════════════════════════════════════════════════╝

import streamlit as st

st.set_page_config(
    page_title="ITM Pave Pro – AASHTO 1993",
    page_icon="🛣️",
    layout="wide",
    initial_sidebar_state="collapsed",
)

from ui.core import ss_init, inject_css

ss_init()
inject_css()

ss = st.session_state

# ── Header ──
st.markdown("""
<div class="main-header">
    <h1>🛣️ ITM Pave Pro — ระบบออกแบบโครงสร้างชั้นทาง AASHTO 1993</h1>
    <p>พัฒนาโดย รศ.ดร.อิทธิพล มีผล | ภาควิชาครุศาสตร์โยธา | มจพ.</p>
</div>
""", unsafe_allow_html=True)

# ── Project name ──
with st.container():
    col_proj, col_info = st.columns([2, 3])
    with col_proj:
        st.text_input("📁 ชื่อโครงการ", key="project_name",
                      placeholder="กรอกชื่อโครงการ...")
    with col_info:
        # Status badges
        badges = []
        if ss.get('esal_rigid') or ss.get('esal_flex'):
            badges.append('✅ ESAL')
        if ss.get('cbr_design'):
            badges.append('✅ CBR')
        if ss.get('flex_results'):
            badges.append('✅ Flexible')
        if ss.get('rigid_results'):
            badges.append('✅ Rigid')
        if badges:
            st.markdown(
                '<div style="padding-top:1.8rem">' +
                ' &nbsp;|&nbsp; '.join(
                    f'<span style="background:#E8F5E9;color:#1B5E20;'
                    f'padding:3px 10px;border-radius:12px;font-size:0.85rem;'
                    f'font-weight:600">{b}</span>' for b in badges
                ) + '</div>',
                unsafe_allow_html=True)

st.markdown("---")

# ── Tabs ──
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "🚛 ESAL Calculator",
    "📊 CBR Analysis",
    "🔧 Flexible Design",
    "🏗️ Rigid Design",
    "💾 Project Save/Load",
])

with tab1:
    from ui.tab1_esal import render
    render()

with tab2:
    from ui.tab2_cbr import render
    render()

with tab3:
    from ui.tab3_flexible import render
    render()

with tab4:
    from ui.tab4_rigid import render
    render()

with tab5:
    from ui.tab5_report import render
    render()
