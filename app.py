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
    col_proj, col_steps = st.columns([2, 3])
    with col_proj:
        st.text_input("📁 ชื่อโครงการ", key="project_name",
                      placeholder="กรอกชื่อโครงการ...")
    with col_steps:
        steps = [
            # ESAL: มีค่าจริงเมื่อคำนวณแล้ว (dict ไม่ว่าง)
            (bool(ss.get('esal_rigid') or ss.get('esal_flex')),
             '🚛', 'ESAL Calculator'),
            # CBR: มีค่าจริงเมื่อ Upload/กรอก + กด "ใช้ค่านี้"
            (bool(ss.get('cbr_values') and ss.get('mr_subgrade_psi') != 4500.0),
             '📊', 'CBR Analysis'),
            # Flexible: มีค่าจริงเมื่อกด Design Check
            (bool(ss.get('flex_results') and ss.get('flex_results') != {}),
             '🔧', 'Flexible Design'),
            # Rigid: มีค่าจริงเมื่อกด Design Check
            (bool(ss.get('rigid_results') and ss.get('rigid_results') != {}),
             '🏗️', 'Rigid Design'),
        ]
        done  = [s for s in steps if s[0]]
        total = len(steps)
        n_done = len(done)

        if n_done == 0:
            intro = '<span style="color:#C62828;font-weight:700;font-size:0.88rem">⚠️ ยังไม่มีการดำเนินการ — เริ่มที่ ESAL Calculator</span>'
        else:
            intro = f'<span style="color:#1B5E20;font-weight:700;font-size:0.88rem">✅ ขั้นตอนที่ดำเนินการแล้ว ({n_done}/{total})</span>'

        badges_html = ''
        for done_flag, icon, label in steps:
            if done_flag:
                badges_html += (
                    f'<span style="background:#E8F5E9;color:#1B5E20;border:1px solid #A5D6A7;'
                    f'padding:3px 10px;border-radius:12px;font-size:0.82rem;font-weight:600;'
                    f'margin-right:4px">✅ {icon} {label}</span>')
            else:
                badges_html += (
                    f'<span style="background:#F5F5F5;color:#9E9E9E;border:1px solid #E0E0E0;'
                    f'padding:3px 10px;border-radius:12px;font-size:0.82rem;font-weight:500;'
                    f'margin-right:4px">⏳ {icon} {label}</span>')

        st.markdown(
            f'<div style="padding-top:0.3rem">{intro}</div>'
            f'<div style="padding-top:0.4rem">{badges_html}</div>',
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
