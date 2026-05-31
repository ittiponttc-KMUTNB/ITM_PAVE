# ╔══════════════════════════════════════════════════════════════════╗
# ║  ui/tab1_esal.py — ITM Pave Pro                                 ║
# ║  ESAL Calculator — AASHTO 1993                                  ║
# ╚══════════════════════════════════════════════════════════════════╝

import pandas as pd
import streamlit as st

from constants import (
    VEHICLE_COLS, VEHICLE_LABELS,
    SLAB_THICKNESSES, SLAB_LABELS,
)
from engine.esal import (
    truck_factor_rigid, truck_factor_flex,
    compute_esal_from_df, grow_traffic,
)

try:
    import openpyxl
    OPENPYXL_OK = True
except ImportError:
    OPENPYXL_OK = False


def _badge(label, value, unit='', bg='#EEF2F7', color='#546E7A'):
    return (
        f'<div style="display:inline-block;background:{bg};border-radius:6px;'
        f'padding:4px 12px;margin-right:6px;margin-bottom:4px;text-align:center">'
        f'<div style="font-size:0.72rem;color:#90A4AE;margin-bottom:1px">{label}</div>'
        f'<div style="font-family:IBM Plex Mono,monospace;font-size:0.88rem;'
        f'font-weight:600;color:{color}">{value}'
        f'<span style="font-size:0.7rem;font-weight:400;margin-left:3px;color:#90A4AE">{unit}</span>'
        f'</div></div>'
    )


def render():
    ss = st.session_state
    st.markdown("### 🚛 ESAL Calculator — AASHTO 1993")

    # ── Pt Global ──
    with st.container(border=True):
        st.markdown('<div class="rp-card-title">🎯 Terminal Serviceability (Pt) — ค่าร่วม</div>',
                    unsafe_allow_html=True)
        c_pt1, c_pt2 = st.columns([1, 3])
        with c_pt1:
            pt_global = st.number_input(
                "Pt (Global Default)",
                value=float(ss.pt_global), step=0.1,
                min_value=2.0, max_value=3.0,
                key="pt_global_input",
                help="ค่านี้จะเป็น default ใน Flexible และ Rigid Design — แก้ได้อิสระในแต่ละหน้า",
            )
            if pt_global != ss.pt_global:
                ss.pt_global  = pt_global
                ss['_pt_sync'] = pt_global
                st.rerun()
        with c_pt2:
            dpsi_r = round(4.5 - pt_global, 2)
            dpsi_f = round(4.2 - pt_global, 2)
            st.markdown('<div style="margin-top:0.5rem"></div>', unsafe_allow_html=True)
            st.markdown(
                _badge('Pt (Global)', f'{pt_global}', '', bg='#E8F5E9', color='#1B5E20') +
                _badge('ΔPSI Rigid (4.5−Pt)', f'{dpsi_r:.2f}', '', bg='#EEF2F7', color='#546E7A') +
                _badge('ΔPSI Flexible (4.2−Pt)', f'{dpsi_f:.2f}', '', bg='#EEF2F7', color='#546E7A'),
                unsafe_allow_html=True
            )
            st.markdown(
                '<div style="font-size:0.78rem;color:#90A4AE;margin-top:4px">'
                '💡 ค่านี้เป็น default ใน Flexible และ Rigid Design — แก้ได้อิสระในแต่ละหน้า</div>',
                unsafe_allow_html=True
            )

    # ── Sub-tabs ──
    sub_rigid, sub_flex = st.tabs(["🔴 Rigid Pavement", "🟢 Flexible Pavement"])

    # ── Traffic Input (shared expander) ──
    with st.expander("📋 ข้อมูลปริมาณจราจร (ใช้ร่วมกันทั้ง Rigid & Flexible)",
                     expanded=True):
        _render_traffic_input(ss)

    # ── Rigid ESAL ──
    with sub_rigid:
        _render_esal_rigid(ss)

    # ── Flexible ESAL ──
    with sub_flex:
        _render_esal_flex(ss)

    # ── Export ──
    render_export()


# ─────────────────────────────────────────────
#  Traffic Input
# ─────────────────────────────────────────────

def _render_traffic_input(ss):
    col_inp1, col_inp2 = st.columns([1, 1])

    with col_inp1:
        with st.container(border=True):
            st.markdown('<div class="rp-card-title">📁 Upload Excel / กรอกมือ</div>',
                        unsafe_allow_html=True)
            input_mode = st.radio(
                "วิธีกรอกข้อมูล",
                ["📁 Upload Excel", "✏️ กรอกมือ + Growth Rate"],
                horizontal=True, key="traffic_input_mode",
            )

            if input_mode == "📁 Upload Excel":
                uploaded_xl = st.file_uploader(
                    "เลือกไฟล์ Excel (.xlsx)", type=['xlsx'], key="traffic_xl"
                )
                st.caption("รูปแบบ: คอลัมน์ Year, MB, HB, MT, HT, TR, STR")
                if uploaded_xl:
                    if not OPENPYXL_OK:
                        st.error("❌ openpyxl ไม่ได้ติดตั้ง — กรุณาใช้วิธี 'กรอกมือ' แทน")
                    else:
                        try:
                            df_up = pd.read_excel(uploaded_xl, engine='openpyxl')
                            df_up.columns = [c.strip() for c in df_up.columns]
                            col_map = {}
                            for c in df_up.columns:
                                for vc in ['Year'] + VEHICLE_COLS:
                                    if c.upper() == vc.upper():
                                        col_map[c] = vc
                            df_up = df_up.rename(columns=col_map)
                            for vc in VEHICLE_COLS:
                                if vc not in df_up.columns:
                                    df_up[vc] = 0
                            ss.traffic_df = df_up[['Year'] + VEHICLE_COLS].fillna(0)
                            st.success(f"✅ อ่านข้อมูล {len(df_up)} ปีสำเร็จ")
                        except Exception as e:
                            st.error(f"❌ {e}")
            else:
                st.markdown("**ปริมาณจราจรปีแรก (คัน/วัน)**")
                base_cols    = st.columns(6)
                base_row     = {}
                defaults_base = {"MB":120,"HB":60,"MT":250,"HT":180,"TR":100,"STR":120}
                for i, vc in enumerate(VEHICLE_COLS):
                    with base_cols[i]:
                        base_row[vc] = st.number_input(
                            vc, value=defaults_base[vc],
                            min_value=0, step=10, key=f"base_{vc}",
                        )
                gc1, gc2 = st.columns(2)
                with gc1:
                    growth_rate  = st.number_input(
                        "Growth Rate (%/ปี)", value=4.5,
                        step=0.5, min_value=0.0, max_value=20.0, key="growth_rate",
                    )
                with gc2:
                    design_years = st.number_input(
                        "Design Period (ปี)", value=20,
                        min_value=1, max_value=40, step=1, key="design_years",
                    )
                if st.button("🔄 สร้างตารางจราจร", type="primary", key="gen_traffic"):
                    ss.traffic_df = grow_traffic(base_row, growth_rate, int(design_years))
                    st.success(f"✅ สร้างตาราง {int(design_years)} ปีสำเร็จ")

    with col_inp2:
        if ss.traffic_df is not None:
            with st.container(border=True):
                st.markdown('<div class="rp-card-title">📊 ตารางปริมาณจราจร</div>',
                            unsafe_allow_html=True)
                st.dataframe(
                    ss.traffic_df.style.format({c: "{:,.0f}" for c in VEHICLE_COLS}),
                    use_container_width=True, height=280,
                )
                total_row = {vc: ss.traffic_df[vc].sum() for vc in VEHICLE_COLS}
                st.markdown(
                    '<div class="result-info">📊 รวมตลอดอายุออกแบบ: '
                    + " | ".join(f"<b>{vc}</b>: {total_row[vc]:,.0f}" for vc in VEHICLE_COLS)
                    + '</div>', unsafe_allow_html=True,
                )
        else:
            st.info("⬅️ กรอกหรือ Upload ข้อมูลจราจรก่อน")


# ─────────────────────────────────────────────
#  ESAL Rigid
# ─────────────────────────────────────────────

def _render_esal_rigid(ss):
    with st.container(border=True):
        st.markdown('<div class="rp-card-title">⚙️ พารามิเตอร์ – Rigid Pavement</div>',
                    unsafe_allow_html=True)
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            ldf_r = st.number_input("Lane Distribution Factor",
                                    value=0.9, step=0.05, min_value=0.1, max_value=1.0,
                                    key="ldf_r")
        with c2:
            ddf_r = st.number_input("Directional Dist. Factor",
                                    value=0.5, step=0.05, min_value=0.1, max_value=1.0,
                                    key="ddf_r")
        with c3:
            pt_r  = st.number_input("Terminal Serviceability Pt",
                                    value=float(ss.get('pt_rigid', 2.5)),
                                    step=0.1, min_value=1.5, max_value=3.5, key="pt_r")
        with c4:
            st.markdown('<div style="margin-top:0.4rem"></div>', unsafe_allow_html=True)
            st.markdown(
                _badge('Pi (Rigid)', '4.5', '', bg='#EEF2F7', color='#546E7A') +
                _badge('ΔPSI', f'{4.5 - pt_r:.2f}', '', bg='#EEF2F7', color='#546E7A'),
                unsafe_allow_html=True
            )

    with st.container(border=True):
        st.markdown('<div class="rp-card-title">📋 Truck Factor (EALF/คัน) ตาม Slab Thickness</div>',
                    unsafe_allow_html=True)
        tf_rows = []
        for vt in VEHICLE_COLS:
            row = {"ประเภทรถ": f"{VEHICLE_LABELS[vt]} ({vt})"}
            for D, lbl in zip(SLAB_THICKNESSES, SLAB_LABELS):
                row[lbl] = f"{truck_factor_rigid(vt, D, pt_r):.3f}"
            tf_rows.append(row)
        st.dataframe(pd.DataFrame(tf_rows), use_container_width=True, hide_index=True)

    if st.button("🔄 คำนวณ ESAL Rigid", type="primary", key="calc_r"):
        if ss.traffic_df is None:
            st.warning("⚠️ กรุณากรอกข้อมูลจราจรก่อน")
        else:
            esal_r        = compute_esal_from_df(ss.traffic_df, ldf_r, ddf_r, pt_r,
                                                  mode="rigid")
            ss.esal_rigid = esal_r
            ss.ldf        = ldf_r
            ss.ddf        = ddf_r
            ss.pt_rigid   = pt_r

            st.markdown("---")
            st.markdown("#### 📊 ผลการคำนวณ ESAL – Rigid Pavement")
            cols_m = st.columns(len(SLAB_THICKNESSES))
            for i, (D, lbl) in enumerate(zip(SLAB_THICKNESSES, SLAB_LABELS)):
                with cols_m[i]:
                    st.markdown(f"""<div class="metric-box">
                        <div class="val">{esal_r[D]:,.0f}</div>
                        <div class="lbl">ESAL – {lbl}</div>
                    </div>""", unsafe_allow_html=True)
            st.markdown(
                '<div class="result-info">✅ บันทึกแล้ว → ใช้ได้ใน K-Value และ Rigid Design</div>',
                unsafe_allow_html=True,
            )

    if ss.esal_rigid:
        st.markdown("**ค่า ESAL Rigid ปัจจุบัน:**")
        df_er = pd.DataFrame({
            "Slab": SLAB_LABELS,
            "ESAL": [f"{ss.esal_rigid.get(D, 0):,.0f}" for D in SLAB_THICKNESSES],
        })
        st.dataframe(df_er, use_container_width=True, hide_index=True)


# ─────────────────────────────────────────────
#  ESAL Flexible
# ─────────────────────────────────────────────

def _render_esal_flex(ss):
    with st.container(border=True):
        st.markdown('<div class="rp-card-title">⚙️ พารามิเตอร์ – Flexible Pavement</div>',
                    unsafe_allow_html=True)
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            ldf_f = st.number_input("Lane Distribution Factor",
                                    value=0.9, step=0.05, min_value=0.1, max_value=1.0,
                                    key="ldf_f")
        with c2:
            ddf_f = st.number_input("Directional Dist. Factor",
                                    value=0.5, step=0.05, min_value=0.1, max_value=1.0,
                                    key="ddf_f")
        with c3:
            pt_f  = st.number_input("Terminal Serviceability Pt",
                                    value=float(ss.get('pt_flex', 2.5)),
                                    step=0.1, min_value=1.5, max_value=3.5, key="pt_f")
        with c4:
            st.markdown('<div style="margin-top:0.4rem"></div>', unsafe_allow_html=True)
            st.markdown(
                _badge('Pi (Flexible)', '4.2', '', bg='#EEF2F7', color='#546E7A') +
                _badge('ΔPSI', f'{4.2 - pt_f:.2f}', '', bg='#EEF2F7', color='#546E7A'),
                unsafe_allow_html=True
            )

    with st.container(border=True):
        st.markdown('<div class="rp-card-title">📐 กำหนด Structure Number (SN)</div>',
                    unsafe_allow_html=True)
        sn_cols = st.columns(4)
        sn_defs = [6.5, 7.1, 7.5, 8.0]
        user_sn = []
        for i, col in enumerate(sn_cols):
            with col:
                user_sn.append(round(
                    st.number_input(f"SN {i+1}", value=sn_defs[i],
                                    min_value=1.0, max_value=20.0,
                                    step=0.1, key=f"sn_{i}", format="%.1f"),
                    2,
                ))

    with st.container(border=True):
        st.markdown('<div class="rp-card-title">📋 Truck Factor (EALF/คัน) ตาม SN</div>',
                    unsafe_allow_html=True)
        tf_rows_f = []
        for vt in VEHICLE_COLS:
            row = {"ประเภทรถ": f"{VEHICLE_LABELS[vt]} ({vt})"}
            for sn in user_sn:
                row[f"SN={sn}"] = f"{truck_factor_flex(vt, sn, pt_f):.3f}"
            tf_rows_f.append(row)
        st.dataframe(pd.DataFrame(tf_rows_f), use_container_width=True, hide_index=True)

    if st.button("🔄 คำนวณ ESAL Flexible", type="primary", key="calc_f"):
        if ss.traffic_df is None:
            st.warning("⚠️ กรุณากรอกข้อมูลจราจรก่อน")
        else:
            esal_fv      = compute_esal_from_df(ss.traffic_df, ldf_f, ddf_f, pt_f,
                                                 mode="flex", sn_list=user_sn)
            ss.esal_flex = esal_fv
            ss.sn_list   = user_sn
            ss.pt_flex   = pt_f

            st.markdown("---")
            st.markdown("#### 📊 ผลการคำนวณ ESAL – Flexible Pavement")
            cols_m2 = st.columns(len(user_sn))
            for i, sn in enumerate(user_sn):
                with cols_m2[i]:
                    st.markdown(f"""<div class="metric-box">
                        <div class="val">{esal_fv[sn]:,.0f}</div>
                        <div class="lbl">ESAL – SN {sn}</div>
                    </div>""", unsafe_allow_html=True)
            st.markdown(
                '<div class="result-info">✅ บันทึกแล้ว → ใช้ได้ใน Flexible Design</div>',
                unsafe_allow_html=True,
            )

    if ss.esal_flex:
        st.markdown("**ค่า ESAL Flexible ปัจจุบัน:**")
        df_ef = pd.DataFrame({
            "SN":   [f"SN {k}" for k in ss.esal_flex.keys()],
            "ESAL": [f"{v:,.0f}" for v in ss.esal_flex.values()],
        })
        st.dataframe(df_ef, use_container_width=True, hide_index=True)


def render_export():
    """ปุ่ม Export ESAL Report — แสดงเสมอหลัง sub-tabs"""
    import streamlit as st
    ss = st.session_state

    st.markdown("---")
    st.markdown("#### 📄 Export ESAL Report")

    has_data = bool(ss.get('esal_flex') or ss.get('esal_rigid'))

    if not has_data:
        st.markdown(
            '<div class="result-warn">⚠️ คำนวณ ESAL ก่อนแล้วจึง Export ได้ครับ</div>',
            unsafe_allow_html=True)
        return

    st.markdown(
        '<div class="result-info">' +
        '📋 Report: หัวข้อ · บทเกริ่นนำ · สมการ · Truck Factor · ปริมาณจราจรรายปี · ESAL + ACC.ESAL'
        '</div>',
        unsafe_allow_html=True)

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        flex_sec = st.text_input("หัวข้อ Flexible", value="4.2.2", key="esal_flex_sec")
    with c2:
        flex_tbl = st.text_input("ตารางเริ่มต้น Flexible", value="4-1", key="esal_flex_tbl")
    with c3:
        rigid_sec = st.text_input("หัวข้อ Rigid", value="4.2.3", key="esal_rigid_sec")
    with c4:
        rigid_tbl = st.text_input("ตารางเริ่มต้น Rigid", value="4-4", key="esal_rigid_tbl")

    if st.button("📄 สร้าง ESAL Report", type="primary",
                  use_container_width=True, key="btn_esal_report"):
        try:
            from engine.report_esal import build_esal_report
            ss_dict = dict(ss)
            ss_dict['report_settings'] = {
                'flex_section_number':  flex_sec,
                'flex_table_start':     flex_tbl,
                'rigid_section_number': rigid_sec,
                'rigid_table_start':    rigid_tbl,
            }
            b = build_esal_report(ss_dict)
            if b:
                st.session_state['_esal_report_bytes'] = b
                st.success("✅ สร้าง Report สำเร็จ — กด Download ด้านล่าง")
            else:
                st.error("❌ ไม่สามารถสร้าง Report ได้ — ตรวจสอบข้อมูลจราจรและ ESAL")
        except Exception as e:
            st.error(f"❌ {e}")

    if ss.get('_esal_report_bytes'):
        proj = ss.get('project_name', '') or 'Report'
        st.download_button(
            "📥 Download ESAL Report (.docx)",
            ss['_esal_report_bytes'],
            file_name=f"ESAL_Report_{proj}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True, key="dl_esal_report")
