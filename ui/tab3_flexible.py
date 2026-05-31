# ╔══════════════════════════════════════════════════════════════════╗
# ║  ui/tab3_flexible.py — ITM Pave Pro                             ║
# ║  Flexible Pavement Design — AASHTO 1993                         ║
# ╚══════════════════════════════════════════════════════════════════╝

import streamlit as st

from constants import (
    FLEX_LAYER_MATERIALS, AC_SURFACE_MATERIALS,
    AC_MATERIALS_LOCK_MI, ZR_MAP,
)
from engine.design import aashto_sn_required, cbr_to_mr, mr_to_k


# ─────────────────────────────────────────────
#  Callbacks CBR ↔ Mr
# ─────────────────────────────────────────────

def _on_cbr_fl_change():
    cbr = max(0.5, float(st.session_state.get('cbr_fl_input', 3.0)))
    st.session_state['cbr_fl_val'] = cbr
    st.session_state['mr_fl_val']  = max(500.0, cbr_to_mr(cbr))

def _on_mr_fl_change():
    mr = max(500.0, float(st.session_state.get('mr_fl_input', 4500.0)))
    st.session_state['mr_fl_val']  = mr
    st.session_state['cbr_fl_val'] = max(0.5, mr / 1500.0)


# ─────────────────────────────────────────────
#  AASHTO D_min Badge
# ─────────────────────────────────────────────

def _aashto_badge(esal, zr, so, pi, pt, mr_psi_layer,
                  cum_sn_before, ai, mi, h_cm):
    if not ai or ai <= 0 or not mr_psi_layer or mr_psi_layer <= 0:
        return ""
    if not esal or esal <= 0:
        return ""
    try:
        sn_req_i = aashto_sn_required(esal, zr, so, pi, pt, mr_psi_layer) or 0.0
    except Exception:
        return ""
    if sn_req_i <= 0:
        return ""
    sn_remaining = sn_req_i - cum_sn_before
    if sn_remaining <= 0:
        return ('<div style="font-size:0.78rem;font-weight:600;'
                'background:#E8F5E9;border-radius:5px;padding:0.2rem 0.5rem;'
                'margin-top:0.25rem;border:1px solid #A5D6A7;color:#1B5E20;">'
                '✅ ผ่าน — ชั้นบนรับ SN ครบแล้ว</div>')
    d_min_in = sn_remaining / (ai * mi)
    d_min_cm = d_min_in * 2.54
    passed   = h_cm >= d_min_cm - 0.05
    if passed:
        return (f'<div style="font-size:0.78rem;font-weight:600;'
                f'background:#E8F5E9;border-radius:5px;padding:0.2rem 0.5rem;'
                f'margin-top:0.25rem;border:1px solid #A5D6A7;color:#1B5E20;">'
                f'✅ ผ่าน &nbsp;(D<sub>min</sub> = {d_min_cm:.1f} ซม. | {d_min_in:.2f} in)</div>')
    else:
        return (f'<div style="font-size:0.78rem;font-weight:600;'
                f'background:#FFF8E1;border-radius:5px;padding:0.2rem 0.5rem;'
                f'margin-top:0.25rem;border:1px solid #FFE082;color:#E65100;">'
                f'💡 ต้องการ D<sub>min</sub> = {d_min_cm:.1f} ซม. ({d_min_in:.2f} in)'
                f' &nbsp;(กรอกอยู่ {h_cm} ซม.)</div>')


# ─────────────────────────────────────────────
#  Main render
# ─────────────────────────────────────────────

def render():
    ss = st.session_state
    st.markdown("### 🔧 Flexible Pavement Design — AASHTO 1993")

    col_fl, col_fr = st.columns([1, 1])

    # ════════════════════════════════
    #  คอลัมน์ซ้าย — ESAL + Subgrade + Parameters
    # ════════════════════════════════
    with col_fl:

        # ── Design ESAL ──
        st.markdown('<div class="card"><h4>📥 Design ESAL</h4>', unsafe_allow_html=True)
        if ss.esal_flex:
            sn_keys   = list(ss.esal_flex.keys())
            sel_idx   = st.selectbox(
                "เลือก SN", range(len(sn_keys)),
                format_func=lambda i: f"SN {sn_keys[i]}  →  ESAL = {ss.esal_flex[sn_keys[i]]:,.0f}",
                key="flex_sn_sel",
            )
            design_esal_f = ss.esal_flex[sn_keys[sel_idx]]
            st.markdown(
                f'<div class="result-info">📊 Design ESAL = <b>{design_esal_f:,.0f}</b></div>',
                unsafe_allow_html=True,
            )
        else:
            st.warning("⚠️ ยังไม่มีค่า ESAL — คำนวณใน ESAL Calculator ก่อน หรือกรอกเอง")
            design_esal_f = st.number_input(
                "Design ESAL (กรอกเอง)", value=0, step=100000, key="flex_esal_manual"
            )
        st.markdown('</div>', unsafe_allow_html=True)

        # ── Subgrade ──
        st.markdown('<div class="card"><h4>🌍 Subgrade</h4>', unsafe_allow_html=True)

        # ── Reference panel จาก TAB 2 ──
        _cbr1 = float(ss.get('cbr_p90') or ss.cbr_design or 0) or None
        _cbr2 = float(ss.get('cbr_fill') or 0)
        _ode  = ss.get('odemark_result')
        _cbr3 = float(_ode.get('cbr_eq_design', _ode.get('cbr_eq', 0))) if (
                    ss.get('improve_soil_check') and _ode) else None

        # ── Reference panel — แสดงข้อมูลอย่างเดียว ผู้ใช้กรอกเองด้านล่าง ──
        _ref_parts = []
        if _cbr1:
            _ref_parts.append(
                f'<span style="background:#E3F2FD;color:#0D47A1;border-radius:6px;'
                f'padding:3px 10px;font-size:0.82rem;margin-right:8px;font-weight:600">'
                f'① ดินเดิม P90 = {_cbr1:.2f}%  (Mr={cbr_to_mr(_cbr1):,.0f} psi)</span>')
        if _cbr2 and _cbr2 > 0:
            _ref_parts.append(
                f'<span style="background:#FFF8E1;color:#E65100;border-radius:6px;'
                f'padding:3px 10px;font-size:0.82rem;margin-right:8px;font-weight:600">'
                f'② ดินถม = {_cbr2:.1f}%  (Mr={cbr_to_mr(_cbr2):,.0f} psi)</span>')
        if _cbr3:
            _ref_parts.append(
                f'<span style="background:#E8F5E9;color:#1B5E20;border-radius:6px;'
                f'padding:3px 10px;font-size:0.82rem;margin-right:8px;font-weight:600">'
                f'③ หลังปรับปรุง = {_cbr3:.0f}%  (Mr={cbr_to_mr(_cbr3):,.0f} psi)</span>')
        if _ref_parts:
            st.markdown(
                '<div style="background:#F8F9FA;border:1px solid #E0E0E0;'
                'border-radius:8px;padding:8px 12px;margin-bottom:10px">'
                '<div style="font-size:0.8rem;color:#78909C;margin-bottom:5px">'
                '📌 ค่าอ้างอิงจาก TAB CBR Analysis:</div>'
                + ' '.join(_ref_parts) + '</div>',
                unsafe_allow_html=True)

        ss['cbr_fl_val'] = max(0.5,   float(ss.get('cbr_fl_val') or 3.0))
        ss['mr_fl_val']  = max(500.0, float(ss.get('mr_fl_val') or 4500.0))

        c1, c2 = st.columns(2)
        with c1:
            st.number_input(
                "CBR (%)", value=ss['cbr_fl_val'],
                step=0.5, min_value=0.5,
                key="cbr_fl_input", on_change=_on_cbr_fl_change,
            )
            cbr_fl = ss['cbr_fl_val']
        with c2:
            st.number_input(
                "Mr (psi)",
                value=ss['mr_fl_val'],
                step=500.0, min_value=500.0,
                key="mr_fl_input", on_change=_on_mr_fl_change,
            )
            mr_fl = ss['mr_fl_val']

        # Warning ถ้ากรอกสูงกว่า P90
        if _cbr1 and cbr_fl > _cbr1 * 1.2:
            st.markdown(
                f'<div class="result-warn">⚠️ CBR ({cbr_fl:.1f}%) สูงกว่า '
                f'ดินเดิม P90 ({_cbr1:.2f}%) มากกว่า 20%</div>',
                unsafe_allow_html=True)

        st.markdown(f"Mr = **{mr_fl:,.0f} psi**  ({mr_fl/145.038:.1f} MPa)")
        st.markdown('</div>', unsafe_allow_html=True)

        # ── Design Parameters ──
        st.markdown('<div class="card"><h4>⚙️ Design Parameters</h4>', unsafe_allow_html=True)
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            r0_fl = st.selectbox("Reliability R0 (%)", list(ZR_MAP.keys()),
                                  index=6, key="r0_fl")
            st.caption(f"ZR = {ZR_MAP[r0_fl]}")
        with c2:
            so_fl = st.number_input("So", value=0.45, step=0.01,
                                    min_value=0.3, max_value=0.6, key="so_fl")
        with c3:
            pi_fl = st.number_input("Pi", value=4.2, step=0.1, key="pi_fl")
        with c4:
            use_pt_global = st.checkbox("ใช้ Pt Global",
                                         value=ss.get('use_pt_global_fl', True),
                                         key="use_pt_global_fl")
            if use_pt_global:
                pt_fl2 = float(ss.get('pt_global', 2.5))
                st.caption(f"Pt = {pt_fl2} (Global)")
            else:
                pt_fl2 = st.number_input(
                    "Pt (Override)",
                    value=float(ss.get('pt_fl2_override') or ss.get('pt_global') or 2.5),
                    step=0.1, min_value=2.0, max_value=3.0, key="pt_fl2_override",
                )
        st.markdown('</div>', unsafe_allow_html=True)

    # ════════════════════════════════
    #  คอลัมน์ขวา — Layer Editor
    # ════════════════════════════════
    with col_fr:
        st.markdown('<div class="card"><h4>🔩 Layer Design</h4>', unsafe_allow_html=True)

        hc0, hc1, hc2, hc3 = st.columns([3, 1, 1, 4])
        hc0.markdown("**วัสดุ**")
        hc1.markdown("**หนา (cm)**")
        hc2.markdown("**mi**")
        hc3.markdown("**ผลคำนวณ**")

        mat_options  = list(FLEX_LAYER_MATERIALS.keys())
        layer_results = []
        cum_sn        = 0.0
        cum_sn_provided = 0.0

        # ── พารามิเตอร์สำหรับ D_min badge ──
        _zr_fl  = ZR_MAP.get(ss.get('r0_fl', 90), -1.282)
        _so_fl  = float(ss.get('so_fl') or 0.45)
        _pi_fl  = float(ss.get('pi_fl') or 4.2)
        _pt_fl2 = (float(ss.get('pt_fl2_override') or ss.get('pt_global') or 2.5)
                   if not ss.get('use_pt_global_fl', True)
                   else float(ss.get('pt_global') or 2.5))
        _esal_f     = ss.get('esal_flex', {})
        _esal_f_val = (list(_esal_f.values())[ss.get('flex_sn_sel', 0)]
                       if _esal_f else float(ss.get('flex_esal_manual') or 0))
        _mr_sub = float(ss.get('mr_fl_val') or 4500.0)

        try:
            _sn_req = aashto_sn_required(_esal_f_val, _zr_fl, _so_fl,
                                          _pi_fl, _pt_fl2, _mr_sub) or 0.0
        except Exception:
            _sn_req = 0.0

        def _get_mr_of_layer(layer_idx):
            mat = ss.get(f"fmat_{layer_idx}", "ไม่เลือก")
            if mat == "ไม่เลือก" or mat not in FLEX_LAYER_MATERIALS:
                return None
            _, _, mr = FLEX_LAYER_MATERIALS[mat]
            if mat == "ดินถมคันทาง CBR กรอกเอง":
                cbr = ss.get(f"fcbr_sub_{layer_idx}", 10.0)
                mr  = cbr * 10.0 * 145.038
            return mr if mr and mr > 0 else None

        # ── Layer rows ──
        for li in range(6):
            lc0, lc1, lc2, lc3 = st.columns([3, 1, 1, 4])

            _is_ac_mat  = ss.get(f"fmat_{li}", mat_options[0]) in AC_SURFACE_MATERIALS
            _do_sub_now = ss.get(f"fsub_{li}", False) and _is_ac_mat

            with lc0:
                mat_f = st.selectbox(f"L{li+1}", mat_options,
                                     key=f"fmat_{li}",
                                     label_visibility="collapsed")
            with lc1:
                if _do_sub_now:
                    _h_total_now = (ss.get(f"fwear_{li}", 5)
                                   + ss.get(f"fbind_{li}", 5)
                                   + ss.get(f"fbase_{li}", 7))
                    st.markdown(
                        f'<div style="padding:0.45rem 0.3rem;font-size:0.88rem;'
                        f'font-weight:700;text-align:center;color:#1B5E20;'
                        f'background:#E8F5E9;border-radius:6px;border:1px solid #A5D6A7;">'
                        f'{_h_total_now} 🔒</div>',
                        unsafe_allow_html=True,
                    )
                    h_f = _h_total_now
                else:
                    h_f = st.number_input("cm",
                                          value=int(ss.get(f"fh_{li}", 0)),
                                          step=1, min_value=0,
                                          key=f"fh_{li}",
                                          label_visibility="collapsed")
            with lc2:
                is_ac = mat_f in AC_MATERIALS_LOCK_MI
                if mat_f != "ไม่เลือก":
                    if is_ac:
                        st.markdown(
                            '<div style="padding:0.4rem 0.3rem;font-size:0.82rem;'
                            'text-align:center;color:#666;">1.0 🔒</div>',
                            unsafe_allow_html=True,
                        )
                        mi_f = 1.0
                    else:
                        mi_f = st.number_input(
                            "mi",
                            value=float(ss.get(f"fmi_{li}", 1.0)),
                            step=0.1, min_value=0.6, max_value=1.4,
                            key=f"fmi_{li}",
                            label_visibility="collapsed",
                            format="%.1f",
                        )
                else:
                    mi_f = 1.0
                    st.markdown("")

            with lc3:
                if mat_f != "ไม่เลือก" and h_f > 0:
                    ai, _, mr_psi_layer = FLEX_LAYER_MATERIALS[mat_f]

                    # ดินถมคันทาง — กรอก CBR
                    if mat_f == "ดินถมคันทาง CBR กรอกเอง":
                        cbr_sub = st.number_input(
                            "CBR ดินถม (%)", value=float(ss.get(f"fcbr_sub_{li}", 10.0)), step=1.0,
                            min_value=2.0, max_value=30.0,
                            key=f"fcbr_sub_{li}",
                        )
                        mr_psi_layer = cbr_sub * 10.0 * 145.038
                        st.caption(f"Mr = {mr_psi_layer:,.0f} psi ({cbr_sub*10:.0f} MPa)")

                    # AC sub-layers
                    if mat_f in AC_SURFACE_MATERIALS:
                        do_sub = st.checkbox("แบ่งชั้นย่อย", key=f"fsub_{li}",
                                             help="แบ่งเป็น Wearing / Binder / Base Course")
                        if do_sub:
                            sc1, sc2, sc3 = st.columns(3)
                            with sc1:
                                h_wear = st.number_input("Wearing (cm)", value=int(ss.get(f"fwear_{li}", 5)), step=1, min_value=0, key=f"fwear_{li}")
                            with sc2:
                                h_bind = st.number_input("Binder (cm)", value=int(ss.get(f"fbind_{li}", 5)), step=1, min_value=0, key=f"fbind_{li}")
                            with sc3:
                                h_base = st.number_input("Base (cm)", value=int(ss.get(f"fbase_{li}", 7)), step=1, min_value=0, key=f"fbase_{li}")
                            warn_msgs = []
                            if h_wear > 0 and not (4 <= h_wear <= 7):
                                warn_msgs.append(f"⚠️ Wearing {h_wear} cm เกินช่วงมาตรฐาน (4–7 cm)")
                            if h_bind > 0 and not (4 <= h_bind <= 8):
                                warn_msgs.append(f"⚠️ Binder {h_bind} cm เกินช่วงมาตรฐาน (4–8 cm)")
                            if h_base > 0 and not (7 <= h_base <= 10):
                                warn_msgs.append(f"⚠️ Base {h_base} cm เกินช่วงมาตรฐาน (7–10 cm)")
                            if warn_msgs:
                                st.markdown(
                                    '<div class="result-warn" style="font-size:0.82rem;">'
                                    + "<br>".join(warn_msgs) + '</div>',
                                    unsafe_allow_html=True,
                                )
                            h_total = h_wear + h_bind + h_base
                            h_in    = h_total / 2.54
                            sn_i    = ai * h_in * mi_f
                            _cum_sn_before   = cum_sn_provided
                            cum_sn          += sn_i
                            cum_sn_provided += sn_i
                            try:
                                _sn_req_i = aashto_sn_required(
                                    _esal_f_val, _zr_fl, _so_fl, _pi_fl, _pt_fl2,
                                    (_get_mr_of_layer(li+1) or _mr_sub)) or 0.0
                            except Exception:
                                _sn_req_i = 0.0
                            layer_results.append({
                                'layer': li+1, 'material': mat_f,
                                'h_cm': h_total, 'ai': ai, 'mi': mi_f,
                                'sni': round(sn_i, 3), 'cum_sn': round(cum_sn, 3),
                                'sub': {'wear': h_wear, 'bind': h_bind, 'base': h_base},
                            })
                            _sn_req_str = f' | SN_req={_sn_req_i:.3f}' if _sn_req_i > 0 else ''
                            st.markdown(
                                f'<div style="padding:0.35rem 0.5rem;font-size:0.80rem;'
                                f'font-family:monospace;background:#F0F4F8;border-radius:6px;">'
                                f'W={h_wear}+B={h_bind}+Base={h_base}=<b>{h_total} cm</b>'
                                f' | ai={ai:.2f} | SNi={sn_i:.3f}'
                                f' | <b>ΣSNi={cum_sn:.3f}</b>{_sn_req_str}</div>',
                                unsafe_allow_html=True,
                            )
                            _badge = _aashto_badge(
                                _esal_f_val, _zr_fl, _so_fl, _pi_fl, _pt_fl2,
                                (_get_mr_of_layer(li+1) or _mr_sub),
                                _cum_sn_before, ai, mi_f, h_total,
                            )
                            if _badge:
                                st.markdown(_badge, unsafe_allow_html=True)
                        else:
                            _render_layer_result(
                                st, li, mat_f, h_f, ai, mi_f, mr_psi_layer,
                                cum_sn, cum_sn_provided,
                                _esal_f_val, _zr_fl, _so_fl, _pi_fl, _pt_fl2,
                                _get_mr_of_layer, _mr_sub, layer_results,
                            )
                            cum_sn          = layer_results[-1]['cum_sn'] if layer_results else cum_sn
                            cum_sn_provided = cum_sn
                    else:
                        _render_layer_result(
                            st, li, mat_f, h_f, ai, mi_f, mr_psi_layer,
                            cum_sn, cum_sn_provided,
                            _esal_f_val, _zr_fl, _so_fl, _pi_fl, _pt_fl2,
                            _get_mr_of_layer, _mr_sub, layer_results,
                        )
                        cum_sn          = layer_results[-1]['cum_sn'] if layer_results else cum_sn
                        cum_sn_provided = cum_sn
                else:
                    st.markdown("")

        st.markdown(f"""<div class="result-info" style="margin-top:0.5rem;">
            ΣSN Provided = <b>{cum_sn:.3f}</b>
        </div>""", unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

        # ── Design Check ──
        if st.button("✅ Design Check", type="primary", key="flex_check"):
            _design_check(ss, design_esal_f, r0_fl, so_fl, pi_fl, pt_fl2,
                          mr_fl, cbr_fl, cum_sn, layer_results)


# ─────────────────────────────────────────────
#  Helpers
# ─────────────────────────────────────────────

    render_export()


def _render_layer_result(st_obj, li, mat_f, h_f, ai, mi_f, mr_psi_layer,
                          cum_sn, cum_sn_provided,
                          esal_f_val, zr, so, pi, pt,
                          get_mr_fn, mr_sub, layer_results):
    """Render ผล SNi สำหรับ layer ที่ไม่มี sub-layer"""
    h_in  = h_f / 2.54
    sn_i  = ai * h_in * mi_f
    _cum_sn_before = cum_sn_provided
    new_cum = cum_sn + sn_i
    try:
        _sn_req_i = aashto_sn_required(
            esal_f_val, zr, so, pi, pt,
            (get_mr_fn(li+1) or mr_sub)) or 0.0
    except Exception:
        _sn_req_i = 0.0
    layer_results.append({
        'layer': li+1, 'material': mat_f,
        'h_cm': h_f, 'ai': ai, 'mi': mi_f,
        'sni': round(sn_i, 3), 'cum_sn': round(new_cum, 3),
    })
    _sn_req_str = f' | SN_req={_sn_req_i:.3f}' if _sn_req_i > 0 else ''
    st_obj.markdown(
        f'<div style="padding:0.35rem 0.5rem;font-size:0.82rem;'
        f'font-family:monospace;background:#F0F4F8;border-radius:6px;">'
        f'<b>{h_f} cm</b> | ai={ai:.2f} | mi={mi_f:.1f} | '
        f'SNi={sn_i:.3f} | <b>ΣSNi={new_cum:.3f}</b>{_sn_req_str}</div>',
        unsafe_allow_html=True,
    )
    _badge = _aashto_badge(
        esal_f_val, zr, so, pi, pt,
        (get_mr_fn(li+1) or mr_sub),
        _cum_sn_before, ai, mi_f, h_f,
    )
    if _badge:
        st_obj.markdown(_badge, unsafe_allow_html=True)


def _design_check(ss, design_esal_f, r0_fl, so_fl, pi_fl, pt_fl2,
                  mr_fl, cbr_fl, cum_sn, layer_results):
    """คำนวณและแสดงผล Design Check"""
    import pandas as pd

    if design_esal_f <= 0:
        st.warning("⚠️ กรุณาใส่ Design ESAL")
        return

    sn_req = aashto_sn_required(design_esal_f, ZR_MAP[r0_fl],
                                  so_fl, pi_fl, pt_fl2, mr_fl)
    if not sn_req:
        st.error("ไม่สามารถคำนวณ SN Required ได้ — ตรวจสอบ ESAL และ Mr")
        return

    passed   = cum_sn >= sn_req
    margin   = cum_sn - sn_req
    sn_ratio = cum_sn / sn_req if sn_req > 0 else 0.0
    css      = "result-pass" if passed else "result-fail"
    chk      = "✅ PASS" if passed else "❌ FAIL"

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(f"""<div class="metric-box">
            <div class="val">{cum_sn:.3f}</div>
            <div class="lbl">SN Provided</div></div>""", unsafe_allow_html=True)
    with c2:
        st.markdown(f"""<div class="metric-box">
            <div class="val">{sn_req:.3f}</div>
            <div class="lbl">SN Required</div></div>""", unsafe_allow_html=True)
    with c3:
        color = '#1B5E20' if passed else '#B71C1C'
        st.markdown(f"""<div class="metric-box">
            <div class="val" style="color:{color}">{margin:+.3f}</div>
            <div class="lbl">Safety Margin</div></div>""", unsafe_allow_html=True)
    with c4:
        color = '#1B5E20' if sn_ratio >= 1.0 else '#B71C1C'
        st.markdown(f"""<div class="metric-box">
            <div class="val" style="color:{color}">{sn_ratio:.3f}</div>
            <div class="lbl">SN Ratio (≥1.0 = ผ่าน)</div></div>""",
                    unsafe_allow_html=True)

    st.markdown(
        f'<div class="{css}" style="margin-top:0.8rem;font-size:1.05rem;">'
        f'{chk} — SN Required = {sn_req:.3f} | SN Provided = {cum_sn:.3f}</div>',
        unsafe_allow_html=True,
    )

    ss.flex_results = {
        'esal':    design_esal_f,
        'sn_req':  sn_req,
        'sn_prov': cum_sn,
        'pass':    passed,
        'layers':  layer_results,
        'mr_psi':  mr_fl,
        'cbr':     cbr_fl,
    }
    ss['r0_flex']        = r0_fl
    ss['so_flex']        = so_fl
    ss['pi_flex']        = pi_fl
    # ── ส่งค่าต่อไป TAB 4 ──
    ss['cbr_design']      = cbr_fl
    ss['mr_subgrade_psi'] = mr_fl
    ss['k_subgrade_pci']  = mr_fl / 19.4

    if layer_results:
        st.dataframe(pd.DataFrame(layer_results),
                     use_container_width=True, hide_index=True)

    # ── รูปโครงสร้าง ──
    if layer_results:
        try:
            from engine.figures import draw_pavement_structure, fig_to_bytes
            import matplotlib.pyplot as plt

            fig_layers = [
                {'name':         l.get('material', ''),
                 'thickness_cm': l.get('h_cm', 0),
                 'ai':           l.get('ai', None),
                 'sni':          l.get('sni', None)}
                for l in layer_results
            ]
            fig = draw_pavement_structure(fig_layers, mode="flex",
                                           cbr_subgrade=cbr_fl)
            if fig:
                st.markdown("#### 🖼️ รูปโครงสร้างชั้นทางลาดยาง")
                st.pyplot(fig, use_container_width=True)
                ss['flex_structure_img'] = fig_to_bytes(fig)
                plt.close(fig)
        except Exception as e:
            st.warning(f"⚠️ ไม่สามารถสร้างรูปโครงสร้างได้: {e}")


def render_export():
    """ปุ่ม Export Flexible Report — เรียกจาก render() ท้าย tab"""
    import streamlit as st
    ss = st.session_state

    st.markdown('---')
    st.markdown('#### 📄 Export Flexible Pavement Report')

    if not ss.get('flex_results'):
        st.markdown('<div class="result-warn">⚠️ กด Design Check ก่อนแล้วจึง Export ได้ครับ</div>',
                    unsafe_allow_html=True)
        return

    fr = ss['flex_results']
    st.markdown(
        f'<div class="result-info">'
        f'✅ SN Required = <b>{fr.get("sn_req",0):.3f}</b> | '
        f'SN Provided = <b>{fr.get("sn_prov",0):.3f}</b> | '
        f'{"✅ PASS" if fr.get("pass") else "❌ FAIL"}'
        f'</div>', unsafe_allow_html=True)

    # Report settings
    c1, c2, c3, c4 = st.columns(4)
    with c1: sec_no  = st.text_input('เลขหัวข้อ',        value='4.4',  key='flex_sec_no')
    with c2: tbl_inp = st.text_input('ตาราง Inputs',      value='4-8',  key='flex_tbl_inp')
    with c3: tbl_mat = st.text_input('ตาราง Materials',   value='4-9',  key='flex_tbl_mat')
    with c4: tbl_sn  = st.text_input('ตาราง SN สรุป',    value='4-10', key='flex_tbl_sn')

    c5, c6 = st.columns(2)
    with c5: num_lanes = st.number_input('จำนวนช่องจราจร', value=2, min_value=1, max_value=8, key='flex_num_lanes')
    with c6: direction = st.text_input('ทิศทาง', value='2 ทิศทาง (ไป-กลับ)', key='flex_direction')

    if st.button('📄 สร้าง Flexible Report', type='primary',
                  use_container_width=True, key='btn_flex_report'):
        try:
            from engine.report_flexible import build_flexible_report
            ss_dict = dict(ss)
            ss_dict['report_settings'] = {
                'section_number':     sec_no,
                'table_inputs':       tbl_inp,
                'table_materials':    tbl_mat,
                'table_sn':           tbl_sn,
                'figure_number':      st.session_state.get('flex_fig_no', '4-8'),
                'num_lanes':          num_lanes,
                'direction':          direction,
            }
            b = build_flexible_report(ss_dict)
            if b:
                ss['_flex_report_bytes'] = b
                st.success('✅ สร้าง Report สำเร็จ — กด Download ด้านล่าง')
            else:
                st.error('❌ ไม่สามารถสร้าง Report ได้')
        except Exception as e:
            st.error(f'❌ {e}')

    if ss.get('_flex_report_bytes'):
        proj = ss.get('project_name', '') or 'Report'
        st.download_button(
            '📥 Download Flexible Report (.docx)',
            ss['_flex_report_bytes'],
            file_name=f'Flexible_Report_{proj}.docx',
            mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            use_container_width=True, key='dl_flex_report')
