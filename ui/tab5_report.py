# ╔══════════════════════════════════════════════════════════════════╗
# ║  ui/tab5_report.py — ITM Pave Pro                               ║
# ║  Project Save / Load — JSON                                     ║
# ║  พัฒนาโดย รศ.ดร.อิทธิพล มีผล | ภาควิชาครุศาสตร์โยธา มจพ.    ║
# ╚══════════════════════════════════════════════════════════════════╝

import json
import io
from datetime import datetime
import streamlit as st
import pandas as pd
import numpy as np


# ─────────────────────────────────────────────
#  Keys ที่ต้องการ Save/Load
# ─────────────────────────────────────────────

# ── เก็บเฉพาะข้อมูลผลคำนวณ ไม่เก็บ widget keys ──
# widget keys (fmat_*, fh_*, jpcp_name_*, ฯลฯ) Streamlit จัดการเอง
# ห้าม set หลัง widget render แล้ว → แยกออกจาก SAVE_KEYS
SAVE_KEYS = [
    # Traffic & ESAL
    'ldf', 'ddf', 'pt_global', 'pt_rigid', 'pt_flex',
    'esal_rigid', 'esal_flex', 'sn_list',
    # CBR
    'cbr_values', 'cbr_percentile', 'cbr_design',
    'mr_subgrade_psi', 'k_subgrade_pci',
    'odemark_result',
    'improve_soil_check',
    # Flexible
    'flex_results', 'r0_flex', 'so_flex', 'pi_flex',
    # Rigid
    'rigid_results', 'r0_rig', 'so_rig', 'zr_rig',
    'k_inf', 'k_corrected', 'ls_value',
    'jpcp_rec_d_cm', 'crcp_rec_d_cm',
    'jpcp_design_params', 'crcp_design_params',
    'jpcp_design_rows', 'crcp_design_rows',
    'jpcp_k_eff', 'crcp_k_eff',
    'jpcp_k_inf', 'crcp_k_inf',
    'jpcp_ls_val', 'crcp_ls_val',
    'jpcp_dsb', 'crcp_dsb',
    'jpcp_esb', 'crcp_esb',
    'jpcp_layers', 'crcp_layers',
    'flex_structure_img',
    # Layer editor Flexible — save เพื่อให้กลับมาแสดงได้
    *[f'fmat_{i}'  for i in range(6)],
    *[f'fh_{i}'    for i in range(6)],
    *[f'fmi_{i}'   for i in range(6)],
    *[f'fwear_{i}' for i in range(6)],
    *[f'fbind_{i}' for i in range(6)],
    *[f'fbase_{i}' for i in range(6)],
    # Layer editor Rigid
    *[f'jpcp_name_{i}'  for i in range(6)],
    *[f'jpcp_thick_{i}' for i in range(6)],
    *[f'crcp_name_{i}'  for i in range(6)],
    *[f'crcp_thick_{i}' for i in range(6)],
]

# Widget keys ที่ห้าม set โดยตรง (เฉพาะที่ render ตลอดเวลา)
# fmat_*, fh_*, jpcp_name_*, jpcp_thick_* ฯลฯ สามารถ set ได้
# เพราะ set แล้ว st.rerun() ทันที ก่อน widget render
_WIDGET_KEYS = {
    'project_name',
    'improve_soil_check',
}


def _make_serializable(obj):
    """แปลง object ให้ JSON serializable"""
    if isinstance(obj, dict):
        return {k: _make_serializable(v) for k, v in obj.items()}
    elif isinstance(obj, list):
        return [_make_serializable(v) for v in obj]
    elif isinstance(obj, (np.integer,)):
        return int(obj)
    elif isinstance(obj, (np.floating,)):
        return float(obj)
    elif isinstance(obj, np.ndarray):
        return obj.tolist()
    elif isinstance(obj, pd.DataFrame):
        return obj.to_dict(orient='records')
    elif isinstance(obj, bool):
        return obj
    elif obj is None or isinstance(obj, (int, float, str)):
        return obj
    else:
        try:
            return str(obj)
        except Exception:
            return None


def render():
    ss = st.session_state
    st.markdown("### 💾 Project Save / Load")
    st.markdown("บันทึกและโหลดข้อมูลโปรเจกต์ทั้งหมดเป็นไฟล์ JSON")

    st.markdown("---")

    col_save, col_load = st.columns(2)

    # ════════════════════════════════
    #  SAVE
    # ════════════════════════════════
    with col_save:
        st.markdown("#### 📥 บันทึกโปรเจกต์")

        # แสดงสรุปข้อมูลที่มี
        with st.container(border=True):
            st.markdown("**ข้อมูลที่จะบันทึก:**")
            items = [
                ('esal_rigid',    '🚛 ESAL Rigid'),
                ('esal_flex',     '🚛 ESAL Flexible'),
                ('cbr_values',    '📊 CBR Analysis'),
                ('flex_results',  '🔧 Flexible Design'),
                ('rigid_results', '🏗️ Rigid Design'),
            ]
            for key, label in items:
                val = ss.get(key)
                has = val is not None and val != {} and val != []
                icon = '✅' if has else '—'
                color = '#2E7D32' if has else '#9E9E9E'
                st.markdown(
                    f'<div style="font-size:0.9rem;color:{color};padding:2px 0">'
                    f'{icon} {label}</div>',
                    unsafe_allow_html=True)

        if st.button("📥 สร้างไฟล์ Save", type="primary",
                      use_container_width=True, key="btn_save"):
            save_data = {}
            for key in SAVE_KEYS:
                val = ss.get(key)
                if val is not None:
                    save_data[key] = _make_serializable(val)

            # เพิ่ม traffic_df แยก
            if ss.get('traffic_df') is not None:
                try:
                    save_data['traffic_df'] = ss['traffic_df'].to_dict(orient='records')
                except Exception:
                    pass

            save_data['_saved_at']      = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            save_data['_version']       = 'ITM_Pave_Pro_v3'
            save_data['_project_name']  = ss.get('project_name', '')

            json_str = json.dumps(save_data, ensure_ascii=False, indent=2)
            proj     = ss.get('project_name', 'Project') or 'Project'
            fname    = f"{proj}_{datetime.now().strftime('%Y%m%d_%H%M')}.json"

            st.download_button(
                "💾 Download JSON",
                data=json_str.encode('utf-8'),
                file_name=fname,
                mime="application/json",
                use_container_width=True,
                key="dl_json")

            st.success(f"✅ พร้อม Download — {len(save_data)} รายการ")

    # ════════════════════════════════
    #  LOAD
    # ════════════════════════════════
    with col_load:
        st.markdown("#### 📤 โหลดโปรเจกต์")

        uploaded_json = st.file_uploader(
            "เลือกไฟล์ JSON", type=['json'], key="json_uploader")

        if uploaded_json:
            try:
                raw      = uploaded_json.read().decode('utf-8')
                data     = json.loads(raw)
                saved_at = data.get('_saved_at', 'ไม่ทราบ')
                proj_name = data.get('_project_name', '')
                version  = data.get('_version', '')

                with st.container(border=True):
                    st.markdown(f"**โปรเจกต์:** {proj_name or '—'}")
                    st.markdown(f"**บันทึกเมื่อ:** {saved_at}")
                    st.markdown(f"**Version:** {version}")

                    # แสดงสรุปข้อมูลในไฟล์
                    has_items = []
                    if data.get('esal_rigid') or data.get('esal_flex'):
                        has_items.append('🚛 ESAL')
                    if data.get('cbr_values'):
                        has_items.append('📊 CBR')
                    if data.get('flex_results'):
                        has_items.append('🔧 Flexible')
                    if data.get('rigid_results'):
                        has_items.append('🏗️ Rigid')
                    if has_items:
                        st.markdown("**มีข้อมูล:** " + " | ".join(has_items))

                if st.button("📤 โหลดข้อมูลนี้", type="primary",
                              use_container_width=True, key="btn_load"):
                    loaded = 0
                    # widget keys ที่ Streamlit จัดการเอง — ห้าม set โดยตรง
                    # ใช้ _WIDGET_KEYS ที่ define ไว้ด้านบน
                    for key in SAVE_KEYS:
                        if key in data and key not in _WIDGET_KEYS:
                            ss[key] = data[key]
                            loaded += 1
                    # project_name — ใช้ st.session_state ผ่าน internal key
                    if 'project_name' in data:
                        # set ผ่าน key ที่ไม่ผูกกับ widget
                        ss['_loaded_project_name'] = data['project_name']

                    # โหลด traffic_df กลับเป็น DataFrame
                    if 'traffic_df' in data and data['traffic_df']:
                        try:
                            ss['traffic_df'] = pd.DataFrame(data['traffic_df'])
                        except Exception:
                            pass

                    st.success(f"✅ โหลดสำเร็จ — {loaded} รายการ")
                    st.info("💡 กด Refresh หรือสลับ Tab เพื่อดูข้อมูลที่โหลดมา")
                    st.rerun()

            except Exception as e:
                st.error(f"❌ ไม่สามารถอ่านไฟล์ได้: {e}")

    # ════════════════════════════════
    #  Reset
    # ════════════════════════════════
    st.markdown("---")
    with st.expander("🗑️ ล้างข้อมูลทั้งหมด", expanded=False):
        st.warning("⚠️ จะลบข้อมูลทั้งหมดในเซสชันนี้ — ไม่สามารถย้อนกลับได้")
        if st.button("🗑️ ล้างข้อมูลทั้งหมด", type="primary", key="btn_reset"):
            # ใช้ _WIDGET_KEYS ที่ define ไว้ด้านบน
            for key in SAVE_KEYS + ['traffic_df']:
                if key in ss and key not in _WIDGET_KEYS:
                    del ss[key]
            from ui.core import ss_init
            ss_init()
            st.success("✅ ล้างข้อมูลแล้ว")
            st.rerun()

    # Footer
    st.markdown("---")
    st.markdown("""
    <div style='text-align:center;color:#4A5568;font-size:0.85rem;padding:0.5rem;'>
        🛣️ <b>ITM Pave Pro v3.0</b> — AASHTO 1993 Pavement Design System<br>
        พัฒนาโดย รศ.ดร.อิทธิพล มีผล | ภาควิชาครุศาสตร์โยธา | คณะครุศาสตร์อุตสาหกรรม | มจพ.
    </div>
    """, unsafe_allow_html=True)
