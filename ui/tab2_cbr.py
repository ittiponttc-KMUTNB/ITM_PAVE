# ╔══════════════════════════════════════════════════════════════════╗
# ║  ui/tab2_cbr.py — ITM Pave Pro                                  ║
# ║  CBR Analysis — Percentile Method                               ║
# ║  Adapted from CBR Calculator V4                                 ║
# ╚══════════════════════════════════════════════════════════════════╝

import re
import math
import numpy as np
import pandas as pd
import streamlit as st
import plotly.graph_objects as go

from constants import SAMPLE_CBR
from engine.design import cbr_to_mr, mr_to_k
from engine.report_cbr import calc_max_rank_percentile, interp_cbr

MPA_PER_CBR = 1500 * 0.006895

IMPROVE_MATERIALS = {
    'หินคลุก CBR 80%':                   350.0,
    'รองพื้นทางวัสดุมวลรวม (CBR 25%)':  150.0,
    'วัสดุคัดเลือก ก':                   100.0,
}


# ─────────────────────────────────────────────
#  Main render
# ─────────────────────────────────────────────

def render():
    ss = st.session_state
    st.markdown('### 📊 CBR Analysis — Percentile Method')

    # ════════════════════════════════
    #  Row 1: Input + กราฟ
    # ════════════════════════════════
    col_l, col_r = st.columns([1, 1.2])

    with col_l:
        _render_input(ss)

    cbr_values = ss.get('cbr_values')

    with col_r:
        if cbr_values:
            _render_chart(ss, cbr_values)

    # ════════════════════════════════
    #  Row 2: ตาราง CBR data
    # ════════════════════════════════
    if cbr_values:
        _render_table(ss, cbr_values)

    # ════════════════════════════════
    #  Row 3: Odemark
    # ════════════════════════════════
    if ss.get('cbr_design') is not None:
        _render_odemark(ss)

    # ════════════════════════════════
    #  Row 4: Export
    # ════════════════════════════════
    _render_export(ss)


# ─────────────────────────────────────────────
#  Input Section
# ─────────────────────────────────────────────

def _render_input(ss):
    st.markdown('<div class="card"><h4>📁 ข้อมูล CBR</h4>', unsafe_allow_html=True)

    cbr_mode = st.radio('แหล่งข้อมูล',
                        ['📁 Upload Excel', '✏️ กรอกค่า', '📌 ใช้ข้อมูลตัวอย่าง'],
                        horizontal=True, key='cbr_mode')

    cbr_vals_input = None
    if cbr_mode == '📁 Upload Excel':
        xl = st.file_uploader('ไฟล์ Excel (คอลัมน์ CBR)', type=['xlsx'], key='cbr_xl')
        if xl:
            try:
                df = pd.read_excel(xl, engine='openpyxl')
                col = next((c for c in df.columns if 'cbr' in c.lower()), df.columns[0])
                cbr_vals_input = pd.to_numeric(df[col], errors='coerce').dropna().tolist()
                st.success(f'✅ {len(cbr_vals_input)} ตัวอย่าง')
            except Exception as e:
                st.error(str(e))

    elif cbr_mode == '✏️ กรอกค่า':
        txt = st.text_area('กรอกค่า CBR (%) คั่นด้วย , หรือ Enter',
                           placeholder='6.5, 7.2, 8.1, 5.3, ...',
                           height=120, key='cbr_txt')
        if txt.strip():
            parts = re.split(r'[,\n\r\s]+', txt.strip())
            try:
                cbr_vals_input = [float(x) for x in parts if x]
                st.success(f'✅ {len(cbr_vals_input)} ค่า')
            except Exception:
                st.error('กรุณากรอกตัวเลขเท่านั้น')
    else:
        cbr_vals_input = list(SAMPLE_CBR)
        st.info(f'📌 ข้อมูลตัวอย่าง {len(SAMPLE_CBR)} ค่า')

    if cbr_vals_input:
        ss.cbr_values = cbr_vals_input

    target_pct = st.slider('Percentile ที่ต้องการ (%)', 50, 99,
                           int(ss.cbr_percentile), step=1, key='pct_slider')
    ss.cbr_percentile = float(target_pct)
    st.markdown('</div>', unsafe_allow_html=True)

    # ── CBR Reference Panel (read-only) ──
    if ss.cbr_values:
        _, n, u_cbr, u_pct, _ = calc_max_rank_percentile(ss.cbr_values)
        cbr_at_pct = interp_cbr(target_pct, u_pct, u_cbr)
        ss['cbr_p90'] = cbr_at_pct  # บันทึกให้ TAB 3/4 ดึงไปใช้

        st.markdown('<div class="card"><h4>📌 ค่าอ้างอิง CBR</h4>',
                    unsafe_allow_html=True)

        # ① ดินเดิม P90
        st.markdown(
            f'<div style="background:#E3F2FD;border:1px solid #90CAF9;'
            f'border-radius:8px;padding:8px 14px;margin-bottom:8px">'
            f'<b>① ดินเดิม</b> — CBR @ P{target_pct:.0f} = '
            f'<b style="color:#0D47A1">{cbr_at_pct:.2f}%</b>'
            f'<span style="color:#546E7A;font-size:0.85rem;margin-left:12px">'
            f'Mr = {cbr_to_mr(cbr_at_pct):,.0f} psi</span></div>',
            unsafe_allow_html=True)

        # ② ดินถมในพื้นที่ — กรอกได้
        c_fill, c_fill_mr = st.columns([2, 2])
        with c_fill:
            cbr_fill = st.number_input(
                '② CBR ดินถมในพื้นที่ (%)',
                value=float(ss.get('cbr_fill_input') or 10.0),
                min_value=0.5, max_value=100.0, step=0.5,
                key='cbr_fill_input')
            ss['cbr_fill'] = cbr_fill
        with c_fill_mr:
            st.markdown(
                f'<div style="padding-top:1.7rem;color:#546E7A;font-size:0.88rem">'
                f'Mr = <b>{cbr_to_mr(cbr_fill):,.0f} psi</b></div>',
                unsafe_allow_html=True)

        # ③ หลังปรับปรุง (Odemark) — แสดงเมื่อมีผลคำนวณ
        ode = ss.get('odemark_result')
        if ss.get('improve_soil_check') and ode:
            cbr_imp = ode.get('cbr_eq_design', int(ode.get('cbr_eq', 0)))
            st.markdown(
                f'<div style="background:#E8F5E9;border:1px solid #A5D6A7;'
                f'border-radius:8px;padding:8px 14px;margin-top:6px">'
                f'<b>③ หลังปรับปรุงดินคันทาง (Odemark)</b> — '
                f'CBR_eq = <b style="color:#1B5E20">{cbr_imp}%</b>'
                f'<span style="color:#546E7A;font-size:0.85rem;margin-left:12px">'
                f'Mr = {cbr_to_mr(cbr_imp):,.0f} psi</span></div>',
                unsafe_allow_html=True)

        st.markdown(
            '<div style="color:#78909C;font-size:0.82rem;margin-top:8px">'
            '💡 ไปที่ TAB Flexible Design หรือ Rigid Design เพื่อเลือกค่าและคำนวณ</div>',
            unsafe_allow_html=True)

        # สถิติ
        with st.expander('📋 สถิติ CBR', expanded=False):
            s1, s2, s3, s4, s5 = st.columns(5)
            s1.metric('n',    n)
            s2.metric('Min',  f'{np.min(ss.cbr_values):.2f}%')
            s3.metric('Max',  f'{np.max(ss.cbr_values):.2f}%')
            s4.metric('Mean', f'{np.mean(ss.cbr_values):.2f}%')
            s5.metric('Std',  f'{np.std(ss.cbr_values):.2f}%')

        st.markdown('</div>', unsafe_allow_html=True)


# ─────────────────────────────────────────────
#  กราฟ
# ─────────────────────────────────────────────

def _render_chart(ss, cbr_values):
    _, n, u_cbr, u_pct, _ = calc_max_rank_percentile(cbr_values)
    target_pct = float(ss.cbr_percentile)
    cbr_at_pct = interp_cbr(target_pct, u_pct, u_cbr)

    st.markdown('<div class="card"><h4>📈 กราฟ Percentile vs CBR</h4>',
                unsafe_allow_html=True)

    x_max = max(u_cbr) * 1.1

    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=u_cbr, y=u_pct, mode='lines+markers',
        name='CBR Distribution',
        line=dict(color='blue', width=2),
        marker=dict(size=6, symbol='x', color='black')))
    fig.add_trace(go.Scatter(
        x=[0, cbr_at_pct], y=[target_pct, target_pct],
        mode='lines', name=f'Percentile {target_pct:.0f}%',
        line=dict(color='red', width=2, dash='dash')))
    fig.add_trace(go.Scatter(
        x=[cbr_at_pct, cbr_at_pct], y=[0, target_pct],
        mode='lines', name=f'CBR={cbr_at_pct:.2f}%',
        line=dict(color='red', width=2, dash='dash')))
    fig.add_annotation(
        x=cbr_at_pct, y=0,
        text=f'<b>{cbr_at_pct:.2f}</b>',
        showarrow=True, arrowhead=2, arrowcolor='red',
        font=dict(size=14, color='red'), ay=40)
    fig.update_layout(
        xaxis_title='CBR (%)', yaxis_title='Percentile (%)',
        plot_bgcolor='white', height=420,
        xaxis=dict(range=[0, x_max], gridcolor='lightgray',
                   showline=False, zeroline=False),
        yaxis=dict(range=[0, 100], gridcolor='lightgray',
                   showline=False, zeroline=False),
        legend=dict(bgcolor='rgba(255,255,255,0.8)', bordercolor='black', borderwidth=1),
        title=dict(text=f'CBR ที่ Percentile {target_pct:.0f}%', x=0.5, xanchor='center'),
        margin=dict(l=60, r=40, t=60, b=60),
    )
    fig.add_shape(type='rect', x0=0, y0=0, x1=x_max, y1=100,
                  line=dict(color='black', width=2), xref='x', yref='y')

    st.plotly_chart(fig, use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)


# ─────────────────────────────────────────────
#  ตาราง CBR data
# ─────────────────────────────────────────────

def _render_table(ss, cbr_values):
    _, n, _, _, full_table = calc_max_rank_percentile(cbr_values)
    st.markdown('---')
    st.markdown('### 📋 ตารางข้อมูล (Max Rank Method)')

    df_display = pd.DataFrame({
        'ลำดับ':         [ft['order'] for ft in full_table],
        'CBR (%)':       [ft['cbr']   for ft in full_table],
        'จำนวนที่≥':     [ft['count_gte'] if ft['show_pct'] else None for ft in full_table],
        'Percentile (%)': [round(ft['pct_gte'], 1) if ft['show_pct'] else None
                           for ft in full_table],
    })
    half = len(df_display) // 2 + 1
    c1, c2 = st.columns(2)
    with c1:
        st.dataframe(df_display.iloc[:half], use_container_width=True, hide_index=True)
    with c2:
        st.dataframe(df_display.iloc[half:], use_container_width=True, hide_index=True)


# ─────────────────────────────────────────────
#  Odemark
# ─────────────────────────────────────────────

def _render_odemark(ss):
    st.markdown('---')
    with st.expander('🔧 การปรับปรุงดินคันทาง (Odemark)', expanded=False):
        improve = st.checkbox('ต้องการปรับปรุงดินคันทาง (CBR ดินเดิมต่ำ)',
                               key='improve_soil_check')
        if not improve:
            return

        st.markdown('#### กำหนดชั้นดิน 2 ชั้น')

        # init defaults
        if 'imp_mat1' not in ss:
            first_mat = list(IMPROVE_MATERIALS.keys())[0]
            ss['imp_mat1'] = first_mat
            ss['imp_mr1']  = IMPROVE_MATERIALS[first_mat]

        # ชั้นที่ 1
        st.markdown('**ชั้นที่ 1 — วัสดุปรับปรุง**')
        cm1, ch1, cmr1 = st.columns(3)
        with cm1:
            mat1 = st.selectbox('ชนิดวัสดุ', list(IMPROVE_MATERIALS.keys()),
                                 key='imp_mat1')
        with ch1:
            h1_cm = st.number_input('ความหนา (ซม.)', 1.0, 150.0, 30.0, 5.0, key='imp_h1')
        with cmr1:
            mr1_def = IMPROVE_MATERIALS.get(mat1, 100.0)
            mr1_mpa = st.number_input('MR (MPa)', 10.0, 1000.0,
                                       float(ss.get('imp_mr1') or mr1_def),
                                       10.0, key='imp_mr1')

        # ชั้นที่ 2
        st.markdown('**ชั้นที่ 2 — ดินถมคันทางใหม่**')
        ch2, cc2 = st.columns(2)
        with ch2:
            h2_cm = st.number_input('ความหนา (ซม.)', 1.0, 300.0, 50.0, 5.0, key='imp_h2')
        with cc2:
            cbr2 = st.number_input('CBR ดินถมคันทางใหม่ (%)', 0.1, 100.0, 10.0, 1.0,
                                    key='imp_cbr2')
        mr2_mpa = cbr2 * MPA_PER_CBR
        st.caption(f'MR ชั้นที่ 2 = {cbr2:.1f} × {MPA_PER_CBR:.4f} = **{mr2_mpa:.2f} MPa**')

        if st.button('คำนวณ CBR_equivalent (Odemark)', type='primary', key='btn_odemark'):
            sum_h    = h1_cm + h2_cm
            sum_hE13 = h1_cm * (mr1_mpa**(1/3)) + h2_cm * (mr2_mpa**(1/3))
            mr_eq    = (sum_hE13 / sum_h)**3
            cbr_eq   = mr_eq / MPA_PER_CBR
            ss['odemark_result'] = {
                'mat1': mat1, 'h1_cm': h1_cm, 'mr1_mpa': mr1_mpa,
                'h2_cm': h2_cm, 'cbr2': cbr2, 'mr2_mpa': mr2_mpa,
                'sum_h': sum_h, 'sum_hE13': sum_hE13,
                'mr_eq_mpa': mr_eq, 'cbr_eq': cbr_eq,
                'cbr_eq_design': math.floor(cbr_eq),
            }
            st.rerun()

        if ss.get('odemark_result'):
            res = ss['odemark_result']
            st.markdown('---')
            st.markdown('#### ผลการคำนวณ')
            r1, r2, r3 = st.columns(3)
            r1.metric('MR equivalent',              f"{res['mr_eq_mpa']:.2f} MPa")
            r2.metric('CBR equivalent (คำนวณ)',     f"{res['cbr_eq']:.2f} %")
            r3.metric('CBR equivalent (ใช้ออกแบบ)', f"{res.get('cbr_eq_design', math.floor(res['cbr_eq']))} %")
            st.info(
                f"**สรุป:** ใช้ CBR_eq = **{res.get('cbr_eq_design', math.floor(res['cbr_eq']))} %** "
                f"แทนค่า CBR ดินเดิม ({ss.get('cbr_design', 0):.2f} %)")

            with st.expander('แสดงวิธีการคำนวณ'):
                st.latex(r'MR_{eq} = \left(\frac{\sum h_i \cdot MR_i^{1/3}}{\sum h_i}\right)^3')
                st.write(f"- ชั้นที่ 1 ({res['mat1']}): h={res['h1_cm']:.1f} cm, MR={res['mr1_mpa']:.1f} MPa, MR¹ᐟ³={res['mr1_mpa']**(1/3):.4f}")
                st.write(f"- ชั้นที่ 2: h={res['h2_cm']:.1f} cm, MR={res['mr2_mpa']:.2f} MPa, MR¹ᐟ³={res['mr2_mpa']**(1/3):.4f}")
                st.write(f"- Σh = {res['sum_h']:.1f} cm, Σ(h·MR¹ᐟ³) = {res['sum_hE13']:.4f}")
                st.write(f"- MR_eq = {res['mr_eq_mpa']:.2f} MPa, CBR_eq = **{res['cbr_eq']:.2f} %**")


# ─────────────────────────────────────────────
#  Export
# ─────────────────────────────────────────────

def _render_export(ss):
    st.markdown('---')
    st.markdown('#### 📄 Export CBR Report')

    if not ss.get('cbr_values'):
        st.markdown('<div class="result-warn">⚠️ กรอกข้อมูล CBR ก่อนครับ</div>',
                    unsafe_allow_html=True)
        return

    st.markdown('<div class="result-info">'
                '📋 Report: หัวข้อ · บทเกริ่นนำ · ตาราง CBR 6 คอลัมน์ · สถิติ · กราฟ'
                + (' · Odemark' if ss.get('odemark_result') else '')
                + '</div>', unsafe_allow_html=True)

    # Report settings
    c1, c2, c3 = st.columns(3)
    with c1:
        sec_num = st.text_input('เลขหัวข้อ',  value='4.3',  key='cbr_sec_num')
    with c2:
        tbl_num = st.text_input('เลขตาราง',   value='4-7',  key='cbr_tbl_num')
    with c3:
        fig_num = st.text_input('เลขรูป',      value='4-7',  key='cbr_fig_num')

    sec_title = st.text_input(
        'ชื่อหัวข้อ',
        value='ข้อมูลความแข็งแรงของดินฐานรากบริเวณพื้นที่โครงการ',
        key='cbr_sec_title')
    c4, c5 = st.columns(2)
    with c4:
        tbl_cap = st.text_input(
            'คำบรรยายตาราง',
            value='ค่าเปอร์เซ็นต์ไทล์ และค่า CBR ของตัวอย่างดินฐานรากตามแนวสายทาง',
            key='cbr_tbl_cap')
    with c5:
        fig_cap = st.text_input(
            'คำบรรยายรูป',
            value='กราฟแสดงความสัมพันธ์ระหว่าง Percentile และ CBR ของดินฐานรากตามแนวสายทาง',
            key='cbr_fig_cap')

    design_cbr_rpt = st.number_input(
        'CBR ที่ใช้ออกแบบในรายงาน (%)',
        value=float(ss.get('cbr_design') or 4.0),
        min_value=0.5, max_value=100.0, step=0.5,
        key='cbr_design_rpt')

    # Preview intro
    cbr_vals = ss.cbr_values
    _, n, u_cbr, u_pct, _ = calc_max_rank_percentile(cbr_vals)
    cbr_at_pct = interp_cbr(float(ss.cbr_percentile), u_pct, u_cbr)
    res   = ss.get('odemark_result')
    cbr_rpt = res.get('cbr_eq_design', math.floor(res['cbr_eq'])) if (
        ss.get('improve_soil_check') and res) else int(math.floor(design_cbr_rpt))

    PURPLE = 'background-color:#D8B4FE;padding:1px 4px;border-radius:3px;font-weight:bold'
    YELLOW = 'background-color:#FDE68A;padding:1px 4px;border-radius:3px;font-weight:bold'
    preview_html = (
        f'<div style="font-family:TH SarabunPSK,Tahoma,sans-serif;font-size:15px;'
        f'line-height:1.8;background:#f9f9f9;padding:15px;border-radius:8px;'
        f'border:1px solid #ddd;">'
        f'<p><b>{sec_num} &nbsp;&nbsp; {sec_title}</b></p>'
        f'<p style="text-indent:40px;text-align:justify;">'
        f'ความแข็งแรงของดินฐานรากบริเวณโดยรอบพื้นที่โครงการ ... '
        f'ผลการทดสอบค่า CBR ของดินฐานรากตามแนวสายทาง จำนวน '
        f'<span style="{PURPLE}">{n}</span> ตัวอย่าง '
        f'พบว่าที่เปอร์เซ็นต์ไทล์ ร้อยละ '
        f'<span style="{PURPLE}">{ss.cbr_percentile:.0f}</span> '
        f'CBR = <span style="{PURPLE}">{cbr_at_pct:.1f}</span> % '
        f'ที่ปรึกษาเลือกค่า CBR = '
        f'<span style="{YELLOW}">{cbr_rpt}</span> % '
        f'ดังแสดงใน<span style="{YELLOW}">ตารางที่ {tbl_num}</span> '
        f'และ<span style="{YELLOW}">รูปที่ {fig_num}</span>'
        f'</p></div>'
    )
    with st.expander('👁️ Preview บทเกริ่นนำ', expanded=False):
        st.markdown(preview_html, unsafe_allow_html=True)
        st.caption('🟣 สีม่วง = คำนวณอัตโนมัติ | 🟡 สีเหลือง = ผู้ใช้กรอกเอง')

    if st.button('📄 สร้าง CBR Report', type='primary',
                  use_container_width=True, key='btn_cbr_report'):
        try:
            from engine.report_cbr import build_cbr_report
            ss_dict = dict(ss)
            ss_dict.update({
                'section_number': sec_num,
                'table_number':   tbl_num,
                'figure_number':  fig_num,
                'section_title':  sec_title,
                'table_caption':  tbl_cap,
                'figure_caption': fig_cap,
                'cbr_design':     design_cbr_rpt,
            })
            b = build_cbr_report(ss_dict)
            if b:
                ss['_cbr_report_bytes'] = b
                st.success('✅ สร้าง Report สำเร็จ — กด Download ด้านล่าง')
            else:
                st.error('❌ ไม่สามารถสร้าง Report ได้')
        except Exception as e:
            st.error(f'❌ {e}')

    if ss.get('_cbr_report_bytes'):
        proj = ss.get('project_name', '') or 'Report'
        st.download_button(
            '📥 Download CBR Report (.docx)',
            ss['_cbr_report_bytes'],
            file_name=f'CBR_Report_{proj}.docx',
            mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            use_container_width=True, key='dl_cbr_report')
