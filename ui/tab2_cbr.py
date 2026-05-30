# ╔══════════════════════════════════════════════════════════════════╗
# ║  ui/tab2_cbr.py — ITM Pave Pro                                  ║
# ║  CBR Analysis — Percentile Method                               ║
# ╚══════════════════════════════════════════════════════════════════╝

import re
import numpy as np
import pandas as pd
import streamlit as st
import plotly.graph_objects as go

from constants import SAMPLE_CBR
from engine.design import calc_percentile_cbr, cbr_to_mr, mr_to_k


def render():
    ss = st.session_state
    st.markdown("### 📊 CBR Analysis — Percentile Method")

    col_l, col_r = st.columns([1, 1])

    # ════════════════════════════════
    #  คอลัมน์ซ้าย — Input + Design CBR
    # ════════════════════════════════
    with col_l:

        # ── Input mode ──
        st.markdown('<div class="card"><h4>📁 ข้อมูล CBR</h4>', unsafe_allow_html=True)
        cbr_mode = st.radio(
            "แหล่งข้อมูล",
            ["📁 Upload Excel", "✏️ กรอกค่า", "📌 ใช้ข้อมูลตัวอย่าง"],
            horizontal=True,
            key="cbr_mode",
        )

        cbr_vals_input = None

        if cbr_mode == "📁 Upload Excel":
            cbr_xl = st.file_uploader(
                "ไฟล์ Excel (คอลัมน์ CBR)", type=['xlsx'], key="cbr_xl"
            )
            if cbr_xl:
                try:
                    df_cbr  = pd.read_excel(cbr_xl, engine='openpyxl')
                    col_cbr = next(
                        (c for c in df_cbr.columns if 'cbr' in c.lower()),
                        df_cbr.columns[0]
                    )
                    cbr_vals_input = (
                        pd.to_numeric(df_cbr[col_cbr], errors='coerce')
                        .dropna().tolist()
                    )
                    st.success(f"✅ {len(cbr_vals_input)} ตัวอย่าง")
                except Exception as e:
                    st.error(str(e))

        elif cbr_mode == "✏️ กรอกค่า":
            cbr_txt = st.text_area(
                "กรอกค่า CBR (%) คั่นด้วย , หรือ Enter",
                placeholder="6.5, 7.2, 8.1, 5.3, ...",
                height=120,
                key="cbr_txt",
            )
            if cbr_txt.strip():
                parts = re.split(r'[,\n\r\s]+', cbr_txt.strip())
                try:
                    cbr_vals_input = [float(x) for x in parts if x]
                    st.success(f"✅ {len(cbr_vals_input)} ค่า")
                except Exception:
                    st.error("กรุณากรอกตัวเลขเท่านั้น")

        else:
            cbr_vals_input = SAMPLE_CBR
            st.info(f"📌 ใช้ข้อมูลตัวอย่าง {len(SAMPLE_CBR)} ค่า")

        if cbr_vals_input:
            ss.cbr_values = cbr_vals_input

        target_pct        = st.slider(
            "Percentile ที่ต้องการ (%)", 50, 99,
            int(ss.cbr_percentile), step=1, key="pct_slider"
        )
        ss.cbr_percentile = float(target_pct)
        st.markdown('</div>', unsafe_allow_html=True)

        # ── Design CBR ──
        if ss.cbr_values:
            arr, n, u_cbr, u_pct = calc_percentile_cbr(ss.cbr_values)
            cbr_at_pct = float(np.interp(target_pct, u_pct[::-1], u_cbr[::-1]))
            mr_auto    = cbr_to_mr(cbr_at_pct)
            k_auto     = mr_to_k(mr_auto)

            st.markdown('<div class="card"><h4>🎯 ค่า CBR ที่ใช้ออกแบบ</h4>',
                        unsafe_allow_html=True)

            c1, c2, c3 = st.columns(3)
            with c1:
                st.markdown(f"""<div class="metric-box">
                    <div class="val">{cbr_at_pct:.2f}</div>
                    <div class="lbl">CBR @ P{target_pct:.0f} (%)</div>
                </div>""", unsafe_allow_html=True)
            with c2:
                st.markdown(f"""<div class="metric-box">
                    <div class="val">{mr_auto:,.0f}</div>
                    <div class="lbl">Mr (psi) = 1500×CBR</div>
                </div>""", unsafe_allow_html=True)
            with c3:
                st.markdown(f"""<div class="metric-box">
                    <div class="val">{k_auto:.1f}</div>
                    <div class="lbl">k subgrade (pci)</div>
                </div>""", unsafe_allow_html=True)

            design_cbr = st.number_input(
                "CBR ที่ใช้ออกแบบจริง (ปรับได้)",
                value=float(round(cbr_at_pct, 1)),
                min_value=0.5, max_value=100.0, step=0.5,
                key="design_cbr_input",
            )
            mr_design = cbr_to_mr(design_cbr)
            k_design  = mr_to_k(mr_design)

            st.markdown(f"""
            <div class="result-info">
                CBR ออกแบบ = <b>{design_cbr:.1f}%</b> →
                Mr = <b>{mr_design:,.0f} psi</b> →
                k_subgrade = <b>{k_design:.1f} pci</b>
            </div>""", unsafe_allow_html=True)

            if st.button("✅ ใช้ค่านี้", type="primary", key="use_cbr"):
                ss.cbr_design      = design_cbr
                ss.mr_subgrade_psi = mr_design
                ss.k_subgrade_pci  = k_design
                st.success(
                    "✅ บันทึกค่า CBR/Mr/k แล้ว "
                    "→ ใช้ได้ใน Flexible Design, K-Value, Rigid Design"
                )

            st.markdown('</div>', unsafe_allow_html=True)

    # ════════════════════════════════
    #  คอลัมน์ขวา — กราฟ + สถิติ
    # ════════════════════════════════
    with col_r:
        if ss.cbr_values:
            arr, n, u_cbr, u_pct = calc_percentile_cbr(ss.cbr_values)
            cbr_at_pct = float(np.interp(target_pct, u_pct[::-1], u_cbr[::-1]))

            # ── กราฟ Plotly ──
            st.markdown('<div class="card"><h4>📈 กราฟ Percentile vs CBR</h4>',
                        unsafe_allow_html=True)

            fig = go.Figure()
            fig.add_trace(go.Scatter(
                x=u_cbr, y=u_pct,
                mode='lines+markers',
                name='CBR Distribution',
                line=dict(color='#1565C0', width=2.5),
                marker=dict(size=7, symbol='x', color='#0B1F3A'),
            ))
            fig.add_trace(go.Scatter(
                x=[0, cbr_at_pct], y=[target_pct, target_pct],
                mode='lines', name=f'P{target_pct:.0f}%',
                line=dict(color='red', width=2, dash='dash'),
            ))
            fig.add_trace(go.Scatter(
                x=[cbr_at_pct, cbr_at_pct], y=[0, target_pct],
                mode='lines', name=f'CBR={cbr_at_pct:.2f}%',
                line=dict(color='red', width=2, dash='dash'),
            ))
            fig.add_annotation(
                x=cbr_at_pct, y=0,
                text=f"<b>{cbr_at_pct:.2f}%</b>",
                showarrow=True, arrowhead=2, arrowcolor='red',
                font=dict(size=14, color='red'), ay=40,
            )
            fig.update_layout(
                xaxis_title="CBR (%)",
                yaxis_title="Percentile (%)",
                plot_bgcolor='white',
                height=360,
                xaxis=dict(range=[0, max(u_cbr) * 1.1], gridcolor='#E3F2FD'),
                yaxis=dict(range=[0, 100], gridcolor='#E3F2FD'),
                legend=dict(bgcolor='rgba(255,255,255,0.8)', bordercolor='#CBD5E1'),
                margin=dict(l=50, r=30, t=30, b=50),
            )
            st.plotly_chart(fig, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)

            # ── สถิติ ──
            st.markdown('<div class="card"><h4>📋 สถิติ CBR</h4>',
                        unsafe_allow_html=True)
            s1, s2, s3, s4 = st.columns(4)
            with s1: st.metric("n",    n)
            with s2: st.metric("Min",  f"{np.min(ss.cbr_values):.2f}%")
            with s3: st.metric("Max",  f"{np.max(ss.cbr_values):.2f}%")
            with s4: st.metric("Mean", f"{np.mean(ss.cbr_values):.2f}%")
            st.markdown('</div>', unsafe_allow_html=True)
