# ╔══════════════════════════════════════════════════════════════════╗
# ║  ui/tab4_rigid.py — ITM Pave Pro                                ║
# ║  Rigid Pavement Design — AASHTO 1993                            ║
# ║  Layout: 2 คอลัมน์ JPCP | CRCP side-by-side                    ║
# ╚══════════════════════════════════════════════════════════════════╝

import math
import io
import streamlit as st
import pandas as pd

from constants import (
    RIGID_LAYER_MATERIALS, RIGID_LAYER_E_DEFAULT,
    SLAB_THICKNESSES, SLAB_LABELS, J_VALUES, ZR_MAP,
)
from engine.design import aashto_rigid_w18, cbr_to_mr, mr_to_k

# ── สีหลัก ──
_JPCP_BD  = '#1565C0'
_JPCP_BG  = '#E3F2FD'
_JPCP_LT  = '#90CAF9'
_CRCP_BD  = '#2E7D32'
_CRCP_BG  = '#E8F5E9'
_CRCP_LT  = '#A5D6A7'


# ─────────────────────────────────────────────
#  HTML helpers
# ─────────────────────────────────────────────

def _card_header(text, color):
    st.markdown(
        f'<div style="background:{color};border-radius:6px 6px 0 0;'
        f'padding:7px 14px;font-size:13px;font-weight:700;color:#fff;'
        f'margin-bottom:0">{text}</div>',
        unsafe_allow_html=True)


def _mbox(label, value, unit='', vc='#1565C0', bg='#E3F2FD'):
    st.markdown(
        f'<div style="background:{bg};border-radius:8px;padding:8px 10px;'
        f'text-align:center;margin-bottom:6px">'
        f'<div style="font-size:11px;color:#78909C;margin-bottom:2px">{label}</div>'
        f'<div style="font-family:IBM Plex Mono,monospace;font-size:18px;'
        f'font-weight:700;color:{vc}">{value}'
        f'<span style="font-size:11px;margin-left:4px;color:#90A4AE">{unit}</span>'
        f'</div></div>',
        unsafe_allow_html=True)


def _row(label, value, hi=False, color='#1A237E'):
    c = color if hi else '#546E7A'
    st.markdown(
        f'<div style="display:flex;justify-content:space-between;'
        f'padding:3px 0;border-bottom:1px solid rgba(0,0,0,0.06);font-size:12px">'
        f'<span style="color:#78909C">{label}</span>'
        f'<span style="font-family:IBM Plex Mono,monospace;font-weight:600;color:{c}">'
        f'{value}</span></div>',
        unsafe_allow_html=True)


# ─────────────────────────────────────────────
#  Main render
# ─────────────────────────────────────────────

def render():
    ss = st.session_state
    st.markdown("### 🏗️ Rigid Pavement Design — AASHTO 1993")

    # ── Status bar ──
    s1, s2, s3 = st.columns(3)
    with s1:
        v = ss.get('esal_rigid')
        ok = bool(v)
        cls = 'badge-ready' if ok else 'badge-wait'
        st.markdown(f'<span class="{cls}">{"✅" if ok else "⚠️"} ESAL Rigid</span>',
                    unsafe_allow_html=True)
    with s2:
        v = ss.get('k_subgrade_pci')
        ok = bool(v)
        cls = 'badge-ready' if ok else 'badge-wait'
        st.markdown(f'<span class="{cls}">{"✅" if ok else "⚠️"} k subgrade</span>',
                    unsafe_allow_html=True)
    with s3:
        v = ss.get('cbr_design')
        ok = bool(v)
        cls = 'badge-ready' if ok else 'badge-wait'
        st.markdown(f'<span class="{cls}">{"✅" if ok else "⚠️"} CBR/Mr</span>',
                    unsafe_allow_html=True)

    st.markdown("---")

    # ════════════════════════════════════════
    #  SECTION A — K-Value Nomograph
    # ════════════════════════════════════════
    with st.expander("📐 SECTION A — K-Value Nomograph (Fig.3.3 & Fig.3.4)", expanded=True):
        _render_nomograph(ss)

    st.markdown("---")

    # ════════════════════════════════════════
    #  SECTION B — Shared Design Parameters
    # ════════════════════════════════════════
    st.markdown("### ⚙️ SECTION B — พารามิเตอร์ร่วม")
    with st.container(border=True):
        rp1, rp2, rp3, rp4, rp5, rp6 = st.columns(6)
        with rp1:
            fc_cube = st.number_input("f'c (ksc)", value=350, step=10,
                                       min_value=200, key="fc_cube")
        with rp2:
            fc_cyl  = 0.8 * fc_cube
            fc_psi  = fc_cyl * 14.223
            ec_psi  = 57000 * math.sqrt(fc_psi)
            sc_auto = min(650, 10.0 * math.sqrt(fc_psi))
            sc_inp  = st.number_input("Sc (psi)", value=int(sc_auto), step=10,
                                       min_value=100, max_value=750, key="sc_inp")
        with rp3:
            r0_rig = st.selectbox("Reliability R0 (%)", list(ZR_MAP.keys()),
                                   index=6, key="r0_rig")
            zr_rig = ZR_MAP[r0_rig]
        with rp4:
            so_rig = st.number_input("So", value=0.35, step=0.01,
                                      min_value=0.2, max_value=0.5, key="so_rig")
        with rp5:
            pi_rig = st.number_input("Pi", value=4.5, step=0.1, key="pi_rig")
        with rp6:
            use_pt_global = st.checkbox("ใช้ Pt Global",
                                         value=ss.get('use_pt_global_rig', True),
                                         key="use_pt_global_rig")
            if use_pt_global:
                pt_rig = float(ss.get('pt_global', 2.5))
                st.caption(f"Pt = {pt_rig} (Global)")
            else:
                pt_rig = st.number_input(
                    "Pt (Override)",
                    value=float(ss.get('pt_rig_override', ss.get('pt_global', 2.5))),
                    step=0.1, min_value=2.0, max_value=3.0, key="pt_rig_override")

        st.markdown(
            f"Ec = **{ec_psi:,.0f} psi** &nbsp;|&nbsp; "
            f"f'c cylinder = **{fc_cyl:.0f} ksc** &nbsp;|&nbsp; "
            f"ZR = **{zr_rig}** &nbsp;|&nbsp; "
            f"Sc = **{sc_inp} psi**"
        )

    # หมายเหตุ: r0_rig, so_rig, pi_rig ถูก manage โดย Streamlit ผ่าน key= แล้ว
    # zr_rig เป็น derived value ไม่ใช่ widget จึง save ได้
    ss['zr_rig'] = zr_rig

    st.markdown("---")

    # ════════════════════════════════════════
    #  SECTION C — JPCP | CRCP side-by-side
    # ════════════════════════════════════════
    st.markdown("### 🏗️ SECTION C — Design JPCP / JRCP vs CRCP")
    col_j, col_c = st.columns(2)

    with col_j:
        _card_header('🔲  JPCP / JRCP — Design', _JPCP_BD)
        with st.container(border=True):
            _design_panel(ss, 'jpcp', 'JPCP',
                           fc_cyl, ec_psi, sc_inp,
                           zr_rig, so_rig, pi_rig, pt_rig,
                           _JPCP_BD, _JPCP_BG, _JPCP_LT)

    with col_c:
        _card_header('〰️  CRCP — Design', _CRCP_BD)
        with st.container(border=True):
            _design_panel(ss, 'crcp', 'CRCP',
                           fc_cyl, ec_psi, sc_inp,
                           zr_rig, so_rig, pi_rig, pt_rig,
                           _CRCP_BD, _CRCP_BG, _CRCP_LT)

    # ════════════════════════════════════════
    #  Comparison Summary
    # ════════════════════════════════════════
    _render_comparison(ss)


# ─────────────────────────────────────────────
#  Nomograph Section
# ─────────────────────────────────────────────

def _render_nomograph(ss):
    """
    Nomograph Section — ใช้ engine.figures plot อัตโนมัติ
    ไม่ต้อง upload รูป — คำนวณและ plot จาก ESB, DSB, k_subgrade
    """
    sub_kinf, sub_ls = st.tabs(["📊 Composite k∞ (Fig.3.3)", "📉 Loss of Support (Fig.3.4)"])

    with sub_kinf:
        st.markdown('<div class="result-info">'
                    '💡 ค่า ESB และ DSB คำนวณอัตโนมัติจาก Layer Editor (Section C) '
                    '— ปรับแก้ได้ด้านล่าง</div>',
                    unsafe_allow_html=True)

        col_ctrl, col_fig = st.columns([1, 2])
        with col_ctrl:
            with st.container(border=True):
                # auto-fill จาก Layer Editor
                esb_auto = int(ss.get('layer_esb_psi', 50000))
                dsb_auto = float(ss.get('layer_dsb_in', 6.0))
                mr_auto  = int(ss.mr_subgrade_psi) if ss.mr_subgrade_psi else 7000
                k_sub    = float(ss.k_subgrade_pci) if ss.k_subgrade_pci else mr_auto / 19.4

                if esb_auto > 0:
                    st.markdown(
                        f'<div class="badge-ready" style="font-size:11px;margin-bottom:4px">ESB จาก Layer Editor = {esb_auto:,} psi</div>',
                        unsafe_allow_html=True)
                if dsb_auto > 0:
                    st.markdown(
                        f'<div class="badge-ready" style="font-size:11px;margin-bottom:4px">DSB จาก Layer Editor = {dsb_auto:.1f} in</div>',
                        unsafe_allow_html=True)

                esb_val = st.number_input("ESB (psi)", value=esb_auto,
                                           step=1000, min_value=1000, key="nomo_esb")
                dsb_val = st.number_input("DSB (inches)", value=dsb_auto,
                                           step=0.5, min_value=0.0, key="nomo_dsb")
                k_sub_val = st.number_input("k subgrade (pci)",
                                             value=float(round(k_sub)),
                                             step=10.0, min_value=10.0, key="nomo_ksub")

                if st.button("📊 Plot Nomograph k∞", type="primary",
                              key="plot_kinf", use_container_width=True):
                    ss['nomo_esb_val']  = esb_val
                    ss['nomo_dsb_val']  = dsb_val
                    ss['nomo_ksub_val'] = k_sub_val
                    ss['nomo_plot_k']   = True

        with col_fig:
            if ss.get('nomo_plot_k') or ss.get('k_inf'):
                try:
                    from engine.figures import draw_k_infinity_nomograph, fig_to_bytes
                    import matplotlib.pyplot as plt

                    esb_use  = ss.get('nomo_esb_val',  esb_auto)
                    dsb_use  = ss.get('nomo_dsb_val',  dsb_auto)
                    ksub_use = ss.get('nomo_ksub_val', k_sub)

                    fig, k_inf_calc = draw_k_infinity_nomograph(esb_use, dsb_use, ksub_use)
                    st.pyplot(fig, use_container_width=True)

                    # auto-save k∞
                    ss.k_inf               = k_inf_calc
                    ss['nomograph_img_k'] = fig_to_bytes(fig)
                    plt.close(fig)

                    st.markdown(
                        f'<div class="result-pass">✅ k∞ = <b>{k_inf_calc:.0f} pci</b> บันทึกแล้ว</div>',
                        unsafe_allow_html=True)
                except Exception as e:
                    st.error(f"❌ ไม่สามารถ plot ได้: {e}")
            else:
                st.info("⬅️ กรอกค่าแล้วกด Plot Nomograph k∞")

        # manual override
        with st.container(border=True):
            k_inf_manual = st.number_input(
                "หรือกรอก k∞ โดยตรง (pci)",
                value=float(ss.k_inf or 200.0),
                step=10.0, min_value=10.0, key="k_inf_manual")
            if st.button("✅ ใช้ค่านี้", key="use_kinf_manual"):
                ss.k_inf = k_inf_manual
                st.success(f"✅ k∞ = {k_inf_manual:.0f} pci")

    with sub_ls:
        col_ctrl2, col_fig2 = st.columns([1, 2])
        with col_ctrl2:
            with st.container(border=True):
                ls_opts = [0.0, 0.5, 1.0, 1.5, 2.0, 3.0]
                ls_sel  = st.select_slider(
                    "Loss of Support (LS)", ls_opts,
                    value=ss.ls_value if ss.ls_value in ls_opts else 1.0,
                    key="ls_sel")

                k_inf_now = float(ss.k_inf or 200.0)
                st.markdown(
                    f'<div class="badge-ready" style="font-size:11px;margin-bottom:4px">k∞ ปัจจุบัน = {k_inf_now:.0f} pci</div>',
                    unsafe_allow_html=True)

                if st.button("📊 Plot Nomograph LS", type="primary",
                              key="plot_ls", use_container_width=True):
                    ss['nomo_ls_val']  = ls_sel
                    ss['nomo_plot_ls'] = True

        with col_fig2:
            if ss.get('nomo_plot_ls') or ss.get('k_corrected'):
                try:
                    from engine.figures import draw_loss_of_support_nomograph, fig_to_bytes
                    import matplotlib.pyplot as plt

                    ls_use = ss.get('nomo_ls_val', ls_sel)
                    fig, k_eff_calc = draw_loss_of_support_nomograph(k_inf_now, ls_use)
                    st.pyplot(fig, use_container_width=True)

                    ss.k_corrected          = k_eff_calc
                    ss.ls_value             = ls_use
                    ss['nomograph_img_ls'] = fig_to_bytes(fig)
                    plt.close(fig)

                    st.markdown(
                        f'<div class="result-pass">✅ k_eff = <b>{k_eff_calc:.0f} pci</b> (LS={ls_use}) บันทึกแล้ว</div>',
                        unsafe_allow_html=True)
                except Exception as e:
                    st.error(f"❌ ไม่สามารถ plot ได้: {e}")
            else:
                st.info("⬅️ กด Plot Nomograph LS")

        # manual override
        with st.container(border=True):
            k_eff_man = st.number_input(
                "หรือกรอก k_eff โดยตรง (pci)",
                value=float(ss.k_corrected or 200.0),
                step=10.0, min_value=10.0, key="k_eff_manual")
            if st.button("✅ ใช้ค่านี้", key="use_keff_manual"):
                ss.k_corrected = k_eff_man
                ss.ls_value    = ls_sel
                st.success(f"✅ k_eff = {k_eff_man:.0f} pci")


def _design_panel(ss, prefix, ptype, fc_cyl, ec_psi, sc_inp,
                   zr, so, pi, pt, bd, bg, bdlt):
    mat_opts = list(RIGID_LAYER_MATERIALS.keys())
    j_key    = "JPCP/JRCP" if prefix == 'jpcp' else "CRCP"
    j_def    = J_VALUES[j_key]

    # ── Layer Editor ──
    st.markdown(f'<div style="font-size:12px;font-weight:700;color:{bd};'
                f'margin-bottom:6px">🔩 ชั้นโครงสร้าง</div>', unsafe_allow_html=True)

    # CRCP: copy จาก JPCP
    if prefix == 'crcp':
        copy_j = st.checkbox("ใช้ชั้นเดียวกับ JPCP/JRCP",
                              value=ss.get('crcp_copy_layers', False),
                              key='crcp_copy_layers')
        if copy_j:
            st.markdown(
                f'<div style="font-size:11px;color:{_CRCP_BD};background:{_CRCP_BG};'
                f'border-radius:6px;padding:4px 8px;margin-bottom:4px">'
                f'✅ ใช้ชั้นวัสดุเดียวกับ JPCP/JRCP</div>',
                unsafe_allow_html=True)

    h0, h1, h2, h3 = st.columns([3, 1.2, 1.5, 0.5])
    h0.markdown("**วัสดุ**")
    h1.markdown("**หนา (cm)**")
    h2.markdown("**E (MPa)**")
    h3.markdown("")

    layer_r    = []
    total_h_cm = 0.0
    e_eq_psi   = 0.0

    for li in range(5):
        lca, lcb, lcc, lcd = st.columns([3, 1.2, 1.5, 0.5])

        # ถ้า CRCP copy layers จาก JPCP
        if prefix == 'crcp' and ss.get('crcp_copy_layers'):
            mat_r = ss.get(f"rmat_jpcp_jrcp_{li}", mat_opts[0])
            h_r   = ss.get(f"rh_jpcp_jrcp_{li}", 0)
            e_def = RIGID_LAYER_E_DEFAULT.get(mat_r, 0)
            e_mpa = ss.get(f"re_jpcp_jrcp_{li}", e_def)
            with lca:
                st.markdown(
                    f'<div style="font-size:11px;padding:6px 4px;color:#546E7A">{mat_r}</div>',
                    unsafe_allow_html=True)
            with lcb:
                st.markdown(
                    f'<div style="font-size:11px;padding:6px 4px;text-align:center">{h_r}</div>',
                    unsafe_allow_html=True)
            with lcc:
                st.markdown(
                    f'<div style="font-size:11px;padding:6px 4px;text-align:center">{e_mpa}</div>',
                    unsafe_allow_html=True)
        else:
            with lca:
                mat_r = st.selectbox(f"ชั้น {li+1}", mat_opts,
                                      key=f"rmat_{prefix}_{li}",
                                      label_visibility="collapsed")
            with lcb:
                h_r = st.number_input("cm", value=0, step=1, min_value=0,
                                       key=f"rh_{prefix}_{li}",
                                       label_visibility="collapsed")
            with lcc:
                prev_key = f"_prev_mat_{prefix}_{li}"
                e_key    = f"re_{prefix}_{li}"
                e_def    = RIGID_LAYER_E_DEFAULT.get(mat_r, 0) if mat_r != "ไม่เลือก" else 0
                if ss.get(prev_key) != mat_r:
                    ss[prev_key] = mat_r
                    ss[e_key]    = e_def
                e_mpa = st.number_input(
                    "MPa", value=ss.get(e_key, e_def),
                    step=50, min_value=0,
                    key=e_key,
                    label_visibility="collapsed",
                    disabled=(mat_r == "ไม่เลือก" or h_r == 0))
            with lcd:
                if mat_r != "ไม่เลือก" and h_r > 0:
                    st.markdown("✅")

        if mat_r != "ไม่เลือก" and h_r > 0 and e_mpa > 0:
            layer_r.append({"name": mat_r, "thickness_cm": h_r, "E_MPa": e_mpa})
            total_h_cm += h_r

    # ── E_equivalent ──
    if layer_r:
        total_valid = sum(l["thickness_cm"] for l in layer_r)
        sum_h_e     = sum(l["thickness_cm"] * (l["E_MPa"] ** (1/3)) for l in layer_r)
        e_eq_mpa    = (sum_h_e / total_valid) ** 3 if total_valid > 0 else 0
        e_eq_psi    = e_eq_mpa * 145.038
        dsb_in      = total_valid / 2.54
        ss['layer_esb_psi'] = int(e_eq_psi)
        ss['layer_dsb_in']  = round(dsb_in, 2)
        st.markdown(
            f'<div class="result-info" style="font-size:0.82rem;margin-top:4px">'
            f'รวม = <b>{total_valid:.0f} cm</b> ({dsb_in:.1f} in) &nbsp;|&nbsp; '
            f'E_eq = <b>{e_eq_psi:,.0f} psi</b></div>',
            unsafe_allow_html=True)

    st.markdown('<div style="height:8px"></div>', unsafe_allow_html=True)

    # ── Parameters ──
    st.markdown(f'<div style="font-size:12px;font-weight:700;color:{bd};'
                f'margin-bottom:4px">⚙️ พารามิเตอร์</div>', unsafe_allow_html=True)

    pc1, pc2 = st.columns(2)
    with pc1:
        j_val = st.number_input(
            f"J ({ptype})", value=j_def, step=0.1,
            min_value=1.0, max_value=5.0, key=f"j_{prefix}")
        cd_val = st.number_input(
            "Cd (Drainage)", value=1.0, step=0.05,
            min_value=0.5, max_value=1.25, key=f"cd_{prefix}")
    with pc2:
        d_sel = st.selectbox(
            "Slab Thickness",
            SLAB_THICKNESSES,
            index=2,
            format_func=lambda d: SLAB_LABELS[SLAB_THICKNESSES.index(d)],
            key=f"d_{prefix}")

        # ESAL auto-fill
        _esal_r    = ss.get('esal_rigid') or {}
        esal_auto  = int(_esal_r.get(d_sel, _esal_r.get(int(d_sel), 0)))
        _w18_key   = f"w18_{prefix}"
        if esal_auto > 0 and ss.get(f"_prev_esal_{prefix}_{d_sel}") != esal_auto:
            ss[f"_prev_esal_{prefix}_{d_sel}"] = esal_auto
            ss[_w18_key] = esal_auto
        if esal_auto > 0:
            st.markdown(
                f'<div class="badge-ready" style="font-size:11px">ESAL = {esal_auto:,}</div>',
                unsafe_allow_html=True)

    w18_req = st.number_input(
        "W18 Design (ESAL)",
        value=ss.get(_w18_key, esal_auto),
        step=100000, min_value=0, key=_w18_key)

    # k_eff auto-fill
    k_eff_auto = float(ss.k_corrected) if ss.k_corrected else 0.0
    if k_eff_auto > 0:
        st.markdown(
            f'<div class="badge-ready" style="font-size:11px">k_eff = {k_eff_auto:.0f} pci</div>',
            unsafe_allow_html=True)
    k_eff_inp = st.number_input(
        "k_eff (pci)",
        value=k_eff_auto if k_eff_auto > 0 else 200.0,
        step=10.0, min_value=10.0, key=f"keff_{prefix}")

    # ── Design Check ──
    if st.button(f"✅ Design Check — {ptype}", type="primary",
                  key=f"dc_{prefix}", use_container_width=True):
        _run_design_check(ss, prefix, ptype, d_sel, w18_req, k_eff_inp,
                           pi, pt, zr, so, sc_inp, cd_val, j_val,
                           ec_psi, fc_cube_val=None, layer_r=layer_r,
                           e_eq_psi=e_eq_psi, bd=bd, bg=bg)


def _run_design_check(ss, prefix, ptype, d_sel, w18_req, k_eff_inp,
                       pi, pt, zr, so, sc_inp, cd_val, j_val,
                       ec_psi, fc_cube_val, layer_r, e_eq_psi, bd, bg):
    if w18_req <= 0:
        st.warning("⚠️ W18 Design = 0 — กรุณาใส่ ESAL")
        return

    w18_cap = aashto_rigid_w18(d_sel, pi, pt, zr, so,
                                sc_inp, cd_val, j_val, ec_psi, k_eff_inp)
    if w18_cap is None:
        st.error("ไม่สามารถคำนวณได้ — ตรวจสอบพารามิเตอร์")
        return

    passed = w18_cap >= w18_req
    margin = (w18_cap / w18_req - 1) * 100 if w18_req > 0 else 0
    ratio  = w18_cap / w18_req if w18_req > 0 else 0
    css    = "result-pass" if passed else "result-fail"
    chk    = "✅ PASS" if passed else "❌ FAIL"

    m1, m2, m3, m4 = st.columns(4)
    with m1:
        st.markdown(f"""<div class="metric-box">
            <div class="val">{w18_cap:,.0f}</div>
            <div class="lbl">W18 Capacity</div></div>""", unsafe_allow_html=True)
    with m2:
        st.markdown(f"""<div class="metric-box">
            <div class="val">{w18_req:,.0f}</div>
            <div class="lbl">W18 Required</div></div>""", unsafe_allow_html=True)
    with m3:
        color = '#1B5E20' if passed else '#B71C1C'
        st.markdown(f"""<div class="metric-box">
            <div class="val" style="color:{color}">{margin:+.1f}%</div>
            <div class="lbl">Safety Margin</div></div>""", unsafe_allow_html=True)
    with m4:
        color = '#1B5E20' if ratio >= 1.0 else '#B71C1C'
        st.markdown(f"""<div class="metric-box">
            <div class="val" style="color:{color}">{ratio:.3f}</div>
            <div class="lbl">W18 Ratio (≥1.0)</div></div>""", unsafe_allow_html=True)

    st.markdown(
        f'<div class="{css}" style="margin-top:6px">'
        f'<b>{chk}</b> — Slab {d_sel} cm | k_eff={k_eff_inp:.0f} pci | J={j_val} | Cd={cd_val}'
        f'</div>', unsafe_allow_html=True)

    # ── ตาราง W18 ทุก Slab ──
    rows_w18 = []
    for D, lbl in zip(SLAB_THICKNESSES, SLAB_LABELS):
        wc = aashto_rigid_w18(D, pi, pt, zr, so, sc_inp, cd_val, j_val, ec_psi, k_eff_inp)
        if wc is None:
            continue
        r   = wc / w18_req if w18_req > 0 else 0
        tag = " ← เลือก" if D == d_sel else ""
        rows_w18.append({
            "Slab": lbl,
            "W18 Capacity": f"{wc:,.0f}",
            "W18 Required": f"{w18_req:,.0f}",
            "Ratio": f"{r:.3f}",
            "สถานะ": ("✅ PASS" if r >= 1.0 else "❌ FAIL") + tag,
        })
    if rows_w18:
        st.markdown("**📊 W18 ทุก Slab Thickness:**")
        st.dataframe(pd.DataFrame(rows_w18),
                     use_container_width=True, hide_index=True)

    # ── บันทึกผล ──
    if not isinstance(ss.rigid_results, dict):
        ss.rigid_results = {}
    ss.rigid_results[ptype] = {
        'd_cm':    d_sel,   'k_eff':   k_eff_inp,
        'sc':      sc_inp,  'j':       j_val,
        'cd':      cd_val,  'w18_cap': w18_cap,
        'w18_req': w18_req, 'pass':    passed,
        'margin':  margin,  'layers':  layer_r,
        'e_eq_psi': e_eq_psi,
        'fc': ss.get('fc_cube', 350),
    }

    # ── รูปโครงสร้าง ──
    if layer_r:
        try:
            from engine.figures import draw_pavement_structure, fig_to_bytes
            import matplotlib.pyplot as plt
            fig_layers = [{"name": l["name"], "thickness_cm": l["thickness_cm"],
                           "E_MPa": l["E_MPa"]} for l in layer_r]
            fig = draw_pavement_structure(
                fig_layers, mode="rigid",
                cbr_subgrade=float(ss.get('cbr_design') or 3.0),
                d_concrete_cm=d_sel, ptype=ptype)
            if fig:
                st.markdown(f"#### 🖼️ โครงสร้างชั้นทาง ({ptype})")
                st.pyplot(fig, use_container_width=True)
                ss[f'rigid_structure_img_{ptype}'] = fig_to_bytes(fig)
                plt.close(fig)
        except Exception as e:
            st.warning(f"⚠️ ไม่สามารถสร้างรูป: {e}")


# ─────────────────────────────────────────────
#  Comparison Summary
# ─────────────────────────────────────────────

def _render_comparison(ss):
    rr = ss.get('rigid_results', {})
    rj = rr.get('JPCP') or rr.get('jpcp')
    rc = rr.get('CRCP') or rr.get('crcp')
    if not rj and not rc:
        return

    st.markdown("---")
    st.markdown("### 📊 สรุปเปรียบเทียบ JPCP vs CRCP")

    with st.container(border=True):
        rows = [
            ("Slab Thickness",
             f"{rj.get('d_cm','—')} cm" if rj else "—",
             f"{rc.get('d_cm','—')} cm" if rc else "—"),
            ("k_eff (pci)",
             f"{rj.get('k_eff',0):.0f}" if rj else "—",
             f"{rc.get('k_eff',0):.0f}" if rc else "—"),
            ("J factor",
             f"{rj.get('j',0):.1f}" if rj else "—",
             f"{rc.get('j',0):.1f}" if rc else "—"),
            ("W18 Capacity",
             f"{rj.get('w18_cap',0):,.0f}" if rj else "—",
             f"{rc.get('w18_cap',0):,.0f}" if rc else "—"),
            ("W18 Required",
             f"{rj.get('w18_req',0):,.0f}" if rj else "—",
             f"{rc.get('w18_req',0):,.0f}" if rc else "—"),
            ("Safety Margin",
             f"{rj.get('margin',0):+.1f}%" if rj else "—",
             f"{rc.get('margin',0):+.1f}%" if rc else "—"),
            ("ผลการตรวจสอบ",
             "✅ PASS" if (rj and rj.get('pass')) else "❌ FAIL",
             "✅ PASS" if (rc and rc.get('pass')) else "❌ FAIL"),
        ]

        hc_l, hc_j, hc_c = st.columns([2, 1, 1])
        hc_l.markdown("**รายการ**")
        hc_j.markdown(f'<div style="color:{_JPCP_BD};font-weight:700">🔲 JPCP/JRCP</div>',
                      unsafe_allow_html=True)
        hc_c.markdown(f'<div style="color:{_CRCP_BD};font-weight:700">〰️ CRCP</div>',
                      unsafe_allow_html=True)
        st.markdown("---")
        for label, vj, vc in rows:
            rc1, rc2, rc3 = st.columns([2, 1, 1])
            rc1.markdown(f'<div style="font-size:12px;color:#546E7A">{label}</div>',
                         unsafe_allow_html=True)
            rc2.markdown(
                f'<div style="font-family:IBM Plex Mono,monospace;font-size:12px;'
                f'font-weight:600;color:{_JPCP_BD}">{vj}</div>',
                unsafe_allow_html=True)
            rc3.markdown(
                f'<div style="font-family:IBM Plex Mono,monospace;font-size:12px;'
                f'font-weight:600;color:{_CRCP_BD}">{vc}</div>',
                unsafe_allow_html=True)
