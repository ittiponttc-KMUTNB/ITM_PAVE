# ╔══════════════════════════════════════════════════════════════════╗
# ║  ui/tab5_report.py — ITM Pave Pro                               ║
# ║  Report & Save — Word / JSON                                    ║
# ╚══════════════════════════════════════════════════════════════════╝

import streamlit as st
from datetime import datetime

from ui.core import status_badge
from engine.report import (
    build_report_esal,
    build_report_cbr,
    build_report_flexible,
    build_report_kvalue,
    build_report_rigid,
    build_report_full,
)

_MIME_DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"


def render():
    ss = st.session_state
    st.markdown("### 📄 Report & Save — Word / JSON")

    # ── สถานะข้อมูล ──
    st.markdown('<div class="card"><h4>📋 สถานะข้อมูลแต่ละส่วน</h4>',
                unsafe_allow_html=True)
    c1, c2, c3, c4, c5 = st.columns(5)
    with c1: st.markdown(status_badge('esal_rigid',   'ESAL'),            unsafe_allow_html=True)
    with c2: st.markdown(status_badge('cbr_values',   'CBR Analysis'),    unsafe_allow_html=True)
    with c3: st.markdown(status_badge('flex_results', 'Flexible Design'), unsafe_allow_html=True)
    with c4: st.markdown(status_badge('k_corrected',  'K-Value'),         unsafe_allow_html=True)
    with c5: st.markdown(status_badge('rigid_results','Rigid Design'),    unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("#### 📥 Download รายงาน Word")

    # ── เลือกส่วนที่ต้องการ ──
    st.markdown("**เลือกส่วนที่ต้องการ Report:**")
    r1, r2, r3, r4, r5 = st.columns(5)
    with r1: chk_esal  = st.checkbox("🚛 ESAL",            value=True, key="chk_esal")
    with r2: chk_cbr   = st.checkbox("📊 CBR Analysis",    value=True, key="chk_cbr")
    with r3: chk_flex  = st.checkbox("🔧 Flexible Design", value=True, key="chk_flex")
    with r4: chk_kval  = st.checkbox("📐 K-Value",         value=True, key="chk_kval")
    with r5: chk_rigid = st.checkbox("🏗️ Rigid Design",    value=True, key="chk_rigid")

    st.markdown("---")
    col_l, col_r = st.columns(2)

    # ── คอลัมน์ซ้าย: แยกส่วน ──
    with col_l:
        st.markdown("**📑 Download แยกส่วน:**")
        ss_d = _get_ss_dict(ss)

        if chk_esal and ss.get('esal_rigid'):
            b = build_report_esal(ss_d)
            if b:
                st.download_button("📥 ESAL Report", b,
                                   "ESAL_Report.docx", mime=_MIME_DOCX,
                                   use_container_width=True, key="dl_esal")

        if chk_cbr and ss.get('cbr_values'):
            b = build_report_cbr(ss_d)
            if b:
                st.download_button("📥 CBR Report", b,
                                   "CBR_Report.docx", mime=_MIME_DOCX,
                                   use_container_width=True, key="dl_cbr")

        if chk_flex and ss.get('flex_results'):
            b = build_report_flexible(ss_d)
            if b:
                st.download_button("📥 Flexible Report", b,
                                   "Flexible_Report.docx", mime=_MIME_DOCX,
                                   use_container_width=True, key="dl_flex")

        if chk_kval and ss.get('k_corrected'):
            b = build_report_kvalue(ss_d)
            if b:
                st.download_button("📥 K-Value Report", b,
                                   "KValue_Report.docx", mime=_MIME_DOCX,
                                   use_container_width=True, key="dl_kval")

        if chk_rigid and ss.get('rigid_results'):
            b = build_report_rigid(ss_d)
            if b:
                st.download_button("📥 Rigid Report", b,
                                   "Rigid_Report.docx", mime=_MIME_DOCX,
                                   use_container_width=True, key="dl_rigid")

    # ── คอลัมน์ขวา: รวมทุกส่วน ──
    with col_r:
        st.markdown("**🗂️ Download รวมทุกส่วน:**")

        ready_parts = []
        if ss.get('esal_rigid') or ss.get('esal_flex'): ready_parts.append("ESAL")
        if ss.get('cbr_values'):    ready_parts.append("CBR")
        if ss.get('flex_results'):  ready_parts.append("Flexible")
        if ss.get('k_corrected'):   ready_parts.append("K-Value")
        if ss.get('rigid_results'): ready_parts.append("Rigid")

        if ready_parts:
            st.markdown(
                f'<div class="result-info">✅ พร้อม Export: '
                f'<b>{" | ".join(ready_parts)}</b></div>',
                unsafe_allow_html=True,
            )
        else:
            st.markdown(
                '<div class="result-warn">⚠️ ยังไม่มีข้อมูล — คำนวณให้ครบก่อนครับ</div>',
                unsafe_allow_html=True,
            )

        if st.button("🗂️ สร้างรายงานรวม", type="primary",
                     use_container_width=True, key="btn_full_report"):
            ss_d   = _get_ss_dict(ss)
            b_full = build_report_full(ss_d)
            if b_full:
                fname = f"ITM_Pave_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
                st.download_button(
                    "📥 Download Full Report", b_full,
                    fname, mime=_MIME_DOCX,
                    use_container_width=True, key="dl_full",
                )
            else:
                st.warning("ไม่มีข้อมูลสำหรับสร้างรายงาน หรือ python-docx ไม่พร้อม")

    # ── JSON ──
    st.markdown("---")
    st.markdown("#### 💾 JSON Save / Load")
    st.info("💡 ใช้ปุ่ม Save/Load JSON ใน Sidebar ด้านซ้ายเพื่อบันทึกและโหลดโปรเจกต์ทั้งหมด")

    st.divider()
    st.markdown("""
    <div style='text-align:center;color:#4A5568;font-size:0.85rem;padding:0.5rem;'>
        🛣️ <b>ITM Pave Pro v3.0</b> — AASHTO 1993 Pavement Design System<br>
        พัฒนาโดย รศ.ดร.อิทธิพล มีผล | ภาควิชาครุศาสตร์โยธา | คณะครุศาสตร์อุตสาหกรรม | มจพ.
    </div>
    """, unsafe_allow_html=True)


def _get_ss_dict(ss) -> dict:
    keys = [
        'project_name',
        'esal_rigid', 'esal_flex', 'ldf', 'ddf',
        'pt_global', 'pt_rigid', 'pt_flex',
        'r0_flex', 'so_flex', 'pi_flex',
        'r0_rig',  'so_rig',  'pi_rig', 'zr_rig',
        'cbr_values', 'cbr_percentile', 'cbr_design',
        'mr_subgrade_psi', 'k_subgrade_pci',
        'flex_results',
        'k_inf', 'k_corrected', 'ls_value',
        'nomograph_img_k', 'nomograph_img_ls',
        'rigid_results',
    ]
    return {k: ss.get(k) for k in keys}
