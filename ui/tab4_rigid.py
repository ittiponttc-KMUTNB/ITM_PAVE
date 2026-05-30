# ╔══════════════════════════════════════════════════════════════════╗
# ║  ui/tab4_rigid.py — ITM Pave Pro                                ║
# ║  Rigid Pavement Design — AASHTO 1993                            ║
# ║  Adapted from Rigid Pavement Design V7                          ║
# ║  พัฒนาโดย รศ.ดร.อิทธิพล มีผล | ภาควิชาครุศาสตร์โยธา มจพ.    ║
# ╚══════════════════════════════════════════════════════════════════╝

import streamlit as st
import matplotlib.pyplot as plt

from engine.rigid_nomograph import (
    calc_composite_k, calc_odemark, apply_loss_of_support,
    plot_f33, plot_f34, plot_structure, fig_to_bytes,
    calc_w18, check_design, find_optimum_k, compare_d,
    convert_cube_to_cyl, calc_ec, calc_sc, get_zr,
    MATERIAL_MODULUS, D_PAIRS,
)

# ── สีหลัก ──
_JPCP_BD   = '#1565C0'
_JPCP_BDLT = '#90CAF9'
_JPCP_BG   = '#E3F2FD'
_CRCP_BD   = '#2E7D32'
_CRCP_BDLT = '#A5D6A7'
_CRCP_BG   = '#E8F5E9'

SC_FIXED  = 600.0   # psi — กรมทางหลวง กำหนด max
DSB_MIN   = 6
DSB_MAX   = 20

_DEF_JPCP = [
    {'name':'รองผิวทางคอนกรีตด้วย AC',                   'thick':5},
    {'name':'หินคลุกปรับปรุงคุณภาพด้วยปูนซีเมนต์ (CTB)', 'thick':20},
    {'name':'หินคลุก CBR 80%',                            'thick':15},
    {'name':'รองพื้นทางวัสดุมวลรวม CBR 25%',             'thick':25},
    {'name':'วัสดุคัดเลือก ก',                            'thick':30},
]
_DEF_CRCP = [
    {'name':'หินคลุกปรับปรุงคุณภาพด้วยปูนซีเมนต์ (CTB)', 'thick':10},
    {'name':'รองพื้นทางวัสดุมวลรวม CBR 25%',             'thick':15},
    {'name':'วัสดุคัดเลือก ก',                            'thick':20},
]


# ─────────────────────────────────────────────
#  UI Helpers (from V7)
# ─────────────────────────────────────────────

def _card_header(text, color):
    st.markdown(
        f'<div style="background:{color};border-radius:6px 6px 0 0;'
        f'padding:6px 12px;font-size:12px;font-weight:700;color:#fff;'
        f'margin-bottom:0">{text}</div>',
        unsafe_allow_html=True)

def _row(label, value, hi=False):
    c = _JPCP_BD if hi else '#1A237E'
    st.markdown(
        f'<div style="display:flex;justify-content:space-between;'
        f'padding:3px 0;border-bottom:1px solid rgba(0,0,0,0.06);font-size:12px">'
        f'<span style="color:#78909C">{label}</span>'
        f'<span style="font-family:IBM Plex Mono,monospace;font-weight:600;color:{c}">'
        f'{value}</span></div>', unsafe_allow_html=True)

def _mbox(label, value, unit='', vc='#1565C0', bg='#E3F2FD'):
    st.markdown(
        f'<div style="background:{bg};border:1px solid rgba(0,0,0,0.08);'
        f'border-radius:7px;padding:8px;text-align:center;margin-bottom:4px">'
        f'<div style="font-size:10px;color:#78909C;margin-bottom:2px">{label}</div>'
        f'<div style="font-family:IBM Plex Mono,monospace;font-size:20px;'
        f'font-weight:700;color:{vc}">{value}</div>'
        f'<div style="font-size:10px;color:#78909C">{unit}</div></div>',
        unsafe_allow_html=True)

def _verdict_bar(d_cm, d_in, w18_cap, w18_req, ratio, passed, bd_color):
    pct_cap   = min(ratio * 100, 100)
    bar_color = '#43A047' if passed else '#E53935'
    label     = f'✅ ผ่าน  (×{ratio:.2f})' if passed else f'❌ ไม่ผ่าน  (×{ratio:.2f})'
    ratio_txt = f'+{(ratio-1)*100:.0f}%' if passed else f'{ratio*100:.0f}%'
    st.markdown(
        f'<div style="background:#F5F5F5;border:1px solid {bd_color}33;'
        f'border-radius:8px;padding:8px 10px;margin-bottom:4px">'
        f'<div style="display:flex;justify-content:space-between;font-size:12px;margin-bottom:4px">'
        f'<span style="font-family:IBM Plex Mono,monospace;font-weight:700;color:{bd_color}">'
        f'D = {d_in} in ({d_cm} ซม.)</span>'
        f'<span style="font-weight:700;color:{bar_color}">{label}</span></div>'
        f'<div style="position:relative;background:#E0E0E0;border-radius:4px;height:10px">'
        f'<div style="background:{bar_color};width:{pct_cap:.1f}%;height:10px;border-radius:4px;'
        f'opacity:0.85"></div></div>'
        f'<div style="display:flex;justify-content:space-between;font-size:10px;'
        f'color:#90A4AE;margin-top:3px">'
        f'<span>W18_cap = {w18_cap:,.0f}</span>'
        f'<span style="color:{bar_color};font-weight:600">{ratio_txt} จาก W18_req</span>'
        f'<span>W18_req = {w18_req:,.0f}</span>'
        f'</div></div>',
        unsafe_allow_html=True)

def _kopt_box(prefix, rec_d_cm, k_opt, k_eff, bd):
    if k_opt is None:
        return
    delta  = k_eff - k_opt
    ok     = k_eff >= k_opt
    bg     = _CRCP_BG if ok else '#FFEBEE'
    bc     = _CRCP_BDLT if ok else '#EF9A9A'
    vc     = _CRCP_BD if ok else '#C62828'
    symbol = '✅' if ok else '⚠️'
    margin = f'{delta:+.0f} pci ({delta/k_opt*100:+.1f}%)'
    st.markdown(
        f'<div style="background:{bg};border:2px solid {bc};border-radius:8px;'
        f'padding:10px 12px;margin-top:6px">'
        f'<div style="font-size:12px;font-weight:700;color:{vc};margin-bottom:6px">'
        f'{symbol} k_opt vs k_eff  —  D = {rec_d_cm} ซม. ({round(rec_d_cm/2.54)} in)</div>'
        f'<div style="display:flex;gap:8px">'
        f'<div style="flex:1;background:white;border-radius:6px;padding:6px;text-align:center">'
        f'<div style="font-size:10px;color:#78909C">k_opt (min required)</div>'
        f'<div style="font-family:IBM Plex Mono,monospace;font-size:18px;font-weight:700;color:{bd}">'
        f'{k_opt:.0f} pci</div></div>'
        f'<div style="flex:1;background:white;border-radius:6px;padding:6px;text-align:center">'
        f'<div style="font-size:10px;color:#78909C">k_eff</div>'
        f'<div style="font-family:IBM Plex Mono,monospace;font-size:18px;font-weight:700;color:{vc}">'
        f'{k_eff:.0f} pci</div></div>'
        f'<div style="flex:1;background:white;border-radius:6px;padding:6px;text-align:center">'
        f'<div style="font-size:10px;color:#78909C">Δk = k_eff − k_opt</div>'
        f'<div style="font-family:IBM Plex Mono,monospace;font-size:14px;font-weight:700;color:{vc}">'
        f'{margin}</div></div>'
        f'</div></div>',
        unsafe_allow_html=True)

def _round_dsb(dsb_raw):
    dsb_rounded = round(dsb_raw)
    warn = None
    if dsb_raw < DSB_MIN:
        warn = f'⚠️ DSB จริง ({dsb_raw:.2f} in) น้อยกว่า {DSB_MIN} in — บังคับใช้ {DSB_MIN} in'
        dsb_rounded = DSB_MIN
    elif dsb_raw > DSB_MAX:
        warn = f'⚠️ DSB จริง ({dsb_raw:.2f} in) เกิน {DSB_MAX} in — บังคับใช้ {DSB_MAX} in'
        dsb_rounded = DSB_MAX
    return dsb_rounded, warn


# ─────────────────────────────────────────────
#  Layer Editor
# ─────────────────────────────────────────────

def _layers(prefix, n, defaults):
    mat    = list(MATERIAL_MODULUS.keys())
    result = []
    c0, c1, c2 = st.columns([3, 1, 1])
    with c0: st.markdown('<div style="font-size:10px;color:#90A4AE;font-weight:600">วัสดุ</div>', unsafe_allow_html=True)
    with c1: st.markdown('<div style="font-size:10px;color:#90A4AE;font-weight:600">ซม.</div>', unsafe_allow_html=True)
    with c2: st.markdown('<div style="font-size:10px;color:#90A4AE;font-weight:600">E (MPa)</div>', unsafe_allow_html=True)
    for i in range(n):
        dn = st.session_state.get(f'{prefix}_name_{i}',
             defaults[i]['name'] if i < len(defaults) else 'หินคลุก CBR 80%')
        dt = st.session_state.get(f'{prefix}_thick_{i}',
             defaults[i]['thick'] if i < len(defaults) else 20)
        if dn not in mat:
            dn = mat[-1]
        ca, cb, cc = st.columns([3, 1, 1])
        with ca:
            nm = st.selectbox(f'n{prefix}{i}', mat, index=mat.index(dn),
                               key=f'{prefix}_name_{i}', label_visibility='collapsed')
        with cb:
            th = st.number_input(f't{prefix}{i}', 0, 200, dt, step=5,
                                  key=f'{prefix}_thick_{i}', label_visibility='collapsed')
        de = st.session_state.get(f'{prefix}_E_{i}_{nm}', MATERIAL_MODULUS.get(nm, 100))
        with cc:
            ev = st.number_input(f'e{prefix}{i}', 10, 10000, de,
                                  key=f'{prefix}_E_{i}_{nm}', label_visibility='collapsed')
        result.append({'name': nm, 'thickness_cm': th, 'E_MPa': ev})
    return result


# ─────────────────────────────────────────────
#  k∞ Block
# ─────────────────────────────────────────────

def _kblock(prefix, layers, MR_psi):
    od = calc_odemark([(l['thickness_cm'], l['E_MPa']) for l in layers])
    if od is None:
        st.warning('⚠️ กรุณากรอกความหนาและ E ให้ครบ')
        return None

    DSB_raw, ESB_psi = od
    DSB_used, warn   = _round_dsb(DSB_raw)
    if warn:
        st.warning(warn)

    res   = calc_composite_k(MR_psi, ESB_psi, float(DSB_used))
    k_inf = res['k_inf_pci']

    ls_val = st.number_input(
        'Loss of Support (LS)', 0.0, 3.0,
        st.session_state.get(f'{prefix}_ls', 1.0), 0.5,
        key=f'{prefix}_ls', format='%.1f',
        help='LS=0: ไม่มี | LS=1: granular | LS=2-3: stabilized')

    k_eff = k_inf if ls_val <= 0 else apply_loss_of_support(k_inf, ls_val)

    _row('DSB (Odemark จริง)', f'{DSB_raw:.2f} in')
    _row('DSB (ใช้จริง)',      f'{DSB_used} in  ← nearest', hi=True)
    _row('ESB equivalent',     f'{ESB_psi:,.0f} psi')
    _row('MR (subgrade)',       f'{MR_psi:,.0f} psi')
    st.markdown('<div style="height:4px"></div>', unsafe_allow_html=True)

    if ls_val <= 0:
        _mbox('k∞ = k_eff (LS=0)', f'{k_inf:.0f}', 'pci', _JPCP_BD, _JPCP_BG)
    else:
        ca, cb = st.columns(2)
        with ca: _mbox('k∞ (Fig.3.3)', f'{k_inf:.0f}', 'pci', _JPCP_BD, _JPCP_BG)
        with cb: _mbox(f'k_eff (LS={ls_val:.1f})', f'{k_eff:.0f}', 'pci', _CRCP_BD, _CRCP_BG)

    # บันทึก session_state
    st.session_state[f'{prefix}_k_inf']   = k_inf
    st.session_state[f'{prefix}_k_eff']   = k_eff
    st.session_state[f'{prefix}_dsb_raw'] = DSB_raw
    st.session_state[f'{prefix}_dsb']     = DSB_used
    st.session_state[f'{prefix}_esb']     = ESB_psi
    st.session_state[f'{prefix}_res33']   = res
    st.session_state[f'{prefix}_ls_val']  = ls_val
    st.session_state[f'{prefix}_layers']  = layers

    # sync กับ ITM Pave session_state
    if prefix == 'jpcp':
        st.session_state['k_inf']       = k_inf
        st.session_state['k_corrected'] = k_eff
        st.session_state['ls_value']    = ls_val

    # ปุ่มกราฟ
    if ls_val <= 0:
        b1, b2 = st.columns(2)
        with b1:
            if st.button('📊 Fig.3.3', key=f'bf33_{prefix}', use_container_width=True):
                st.session_state[f'{prefix}_show_f33'] = not st.session_state.get(f'{prefix}_show_f33', False)
        with b2:
            if st.button('🏗️ โครงสร้าง', key=f'bstr_{prefix}', use_container_width=True):
                st.session_state[f'{prefix}_show_str'] = not st.session_state.get(f'{prefix}_show_str', False)
    else:
        b1, b2, b3 = st.columns(3)
        with b1:
            if st.button('📊 Fig.3.3', key=f'bf33_{prefix}', use_container_width=True):
                st.session_state[f'{prefix}_show_f33'] = not st.session_state.get(f'{prefix}_show_f33', False)
        with b2:
            if st.button('📉 Fig.3.4', key=f'bf34_{prefix}', use_container_width=True):
                st.session_state[f'{prefix}_show_f34'] = not st.session_state.get(f'{prefix}_show_f34', False)
        with b3:
            if st.button('🏗️ โครงสร้าง', key=f'bstr_{prefix}', use_container_width=True):
                st.session_state[f'{prefix}_show_str'] = not st.session_state.get(f'{prefix}_show_str', False)

    _graphs(prefix, MR_psi)
    return (k_inf, k_eff)


def _graphs(prefix, MR_psi):
    res    = st.session_state.get(f'{prefix}_res33')
    ls_val = st.session_state.get(f'{prefix}_ls_val', 0)
    k_inf  = st.session_state.get(f'{prefix}_k_inf')
    k_eff  = st.session_state.get(f'{prefix}_k_eff')
    DSB    = st.session_state.get(f'{prefix}_dsb')
    ESB    = st.session_state.get(f'{prefix}_esb')
    layers = st.session_state.get(f'{prefix}_layers', [])
    if res is None:
        return

    fig33 = plot_f33(MR_psi, ESB, DSB, res)
    st.session_state[f'{prefix}_fig33_bytes'] = fig_to_bytes(fig33)
    if prefix == 'jpcp':
        st.session_state['nomograph_img_k'] = fig_to_bytes(fig33)
    if st.session_state.get(f'{prefix}_show_f33'):
        st.pyplot(fig33, use_container_width=True)
        st.download_button('⬇️ Fig.3.3', st.session_state[f'{prefix}_fig33_bytes'],
                            f'fig33_{prefix}.png', 'image/png', key=f'dl33_{prefix}')
    plt.close(fig33)

    if ls_val > 0:
        fig34 = plot_f34(k_inf, ls_val, k_eff)
        st.session_state[f'{prefix}_fig34_bytes'] = fig_to_bytes(fig34)
        if prefix == 'jpcp':
            st.session_state['nomograph_img_ls'] = fig_to_bytes(fig34)
        if st.session_state.get(f'{prefix}_show_f34'):
            st.pyplot(fig34, use_container_width=True)
            st.download_button('⬇️ Fig.3.4', st.session_state[f'{prefix}_fig34_bytes'],
                                f'fig34_{prefix}.png', 'image/png', key=f'dl34_{prefix}')
        plt.close(fig34)

    if st.session_state.get(f'{prefix}_show_str') and layers:
        fig = plot_structure(layers)
        if fig:
            st.pyplot(fig, use_container_width=True)
            st.download_button('⬇️ โครงสร้าง', fig_to_bytes(fig),
                                f'str_{prefix}.png', 'image/png', key=f'dlstr_{prefix}')
            plt.close(fig)


# ─────────────────────────────────────────────
#  Design Block
# ─────────────────────────────────────────────

def _design_block(prefix, ptype, fc_cyl, ec_psi, cd, w18_req, pt, zr, so, bd, bdlt):
    dpsi  = 4.5 - pt
    k_eff = st.session_state.get(f'{prefix}_k_eff')

    if w18_req is None:
        st.markdown(
            '<div style="background:#FFF3E0;border:1px solid #FFB74D;border-radius:8px;'
            'padding:8px 12px;font-size:12px;color:#E65100">'
            '⚠️ ยังไม่มีข้อมูล W18 — กรุณากรอก W18 ด้านบน</div>',
            unsafe_allow_html=True)
        return None

    if k_eff is None:
        st.markdown(
            '<div style="background:#FFF3E0;border:1px solid #FFB74D;border-radius:8px;'
            'padding:8px 12px;font-size:12px;color:#E65100">'
            f'⚠️ ยังไม่มีค่า k_eff ({ptype}) — คำนวณใน Section A ก่อน</div>',
            unsafe_allow_html=True)
        return None

    j_opts  = [2.5, 2.6, 2.7, 2.8] if prefix == 'jpcp' else [2.3, 2.4, 2.5, 2.6]
    j_def   = st.session_state.get(f'{prefix}_j', j_opts[-1])
    j_label = f'J — Load Transfer Coefficient ({ptype})'
    if j_def not in j_opts:
        j_def = j_opts[-1]

    j_val = st.select_slider(j_label, options=j_opts, value=j_def,
                              key=f'{prefix}_j', format_func=lambda x: f'{x:.1f}')

    st.markdown('<div style="height:4px"></div>', unsafe_allow_html=True)
    _row(f'W18 (ref)', f'{w18_req:,.0f} ESALs')
    _row('k_eff',       f'{k_eff:.0f} pci')
    _row("f'c (cube)",  f"{st.session_state.get('fc_cube', 350):.0f} ksc")
    _row('Ec',           f'{ec_psi:,.0f} psi')
    _row('Sc (lock)',    f'{SC_FIXED:.0f} psi')
    _row('J',            f'{j_val:.1f}', hi=True)
    _row('Cd',           f'{cd:.1f}', hi=True)
    _row('Pt / ΔPSI',    f'{pt:.1f} / {dpsi:.1f}')
    _row('ZR / So',      f'{zr:.3f} / {so:.2f}')
    st.markdown('<div style="height:6px"></div>', unsafe_allow_html=True)

    # คำนวณ W18 ทุก D
    rows = []
    for d_in, d_cm in D_PAIRS:
        lw, wc = calc_w18(d_in, dpsi, pt, zr, so, SC_FIXED, cd, j_val, ec_psi, k_eff)
        passed = wc >= w18_req
        ratio  = round(wc / w18_req, 3) if w18_req > 0 else 0
        rows.append({
            'd_cm':    d_cm, 'd_inch': d_in,
            'log_w18': round(lw, 4),
            'w18_cap': round(wc, 0), 'w18_req': w18_req,
            'passed':  passed, 'ratio': ratio,
        })
    st.session_state[f'{prefix}_design_rows'] = rows

    for r in rows:
        _verdict_bar(r['d_cm'], r['d_inch'], r['w18_cap'], r['w18_req'],
                     r['ratio'], r['passed'], bd)

    passed_rows = [r for r in rows if r['passed']]
    if passed_rows:
        rec = min(passed_rows, key=lambda r: r['d_cm'])
        _mbox(f'✅ D แนะนำ ({ptype})',
              f"{rec['d_inch']} in ({rec['d_cm']} ซม.)",
              f"W18 capacity = {rec['w18_cap']:,.0f}",
              _CRCP_BD if prefix == 'crcp' else _JPCP_BD,
              _CRCP_BG if prefix == 'crcp' else _JPCP_BG)
        st.session_state[f'{prefix}_rec_d_cm'] = rec['d_cm']
    else:
        st.markdown(
            '<div style="background:#FFEBEE;border:1px solid #EF9A9A;'
            'border-radius:8px;padding:8px 12px;font-size:12px;color:#C62828">'
            '❌ ไม่มี D ที่ผ่านเกณฑ์ — พิจารณาเพิ่ม k_eff หรือลด J</div>',
            unsafe_allow_html=True)
        st.session_state[f'{prefix}_rec_d_cm'] = None

    sel_d_cm = st.session_state.get(f'{prefix}_rec_d_cm') or 30
    sel_d_in = round(sel_d_cm / 2.54)
    k_opt    = find_optimum_k(w18_req, sel_d_in, dpsi, pt, zr, so,
                               SC_FIXED, cd, j_val, ec_psi)
    _kopt_box(prefix, sel_d_cm, k_opt, k_eff, bd)

    # บันทึก rigid_results สำหรับ report
    layers = st.session_state.get(f'{prefix}_layers', [])
    if not isinstance(st.session_state.get('rigid_results'), dict):
        st.session_state['rigid_results'] = {}
    st.session_state['rigid_results'][ptype] = {
        'd_cm':    sel_d_cm, 'k_eff':   k_eff,
        'fc':      st.session_state.get('fc_cube', 350),
        'sc':      SC_FIXED, 'j':       j_val,
        'cd':      cd,       'w18_req': w18_req,
        'w18_cap': passed_rows[0]['w18_cap'] if passed_rows else 0,
        'pass':    bool(passed_rows), 'layers': layers,
    }

    # ปุ่มโครงสร้าง
    if layers:
        if st.button(f'🏗️ โครงสร้าง {ptype}', key=f'str_{prefix}_d', use_container_width=True):
            st.session_state[f'{prefix}_show_str3'] = not st.session_state.get(f'{prefix}_show_str3', False)
        if st.session_state.get(f'{prefix}_show_str3'):
            rec_cm = st.session_state.get(f'{prefix}_rec_d_cm')
            fig = plot_structure(layers, concrete_cm=rec_cm,
                                  title=f'{ptype}  D = {rec_cm} cm' if rec_cm else f'{ptype}')
            if fig:
                st.pyplot(fig, use_container_width=True)
                st.session_state[f'rigid_structure_img_{ptype}'] = fig_to_bytes(fig)
                st.download_button(f'⬇️ PNG {ptype}', fig_to_bytes(fig),
                                    f'struct_{prefix}.png', 'image/png', key=f'dl_str_{prefix}')
                plt.close(fig)

    st.session_state[f'{prefix}_design_params'] = {
        'w18': w18_req, 'pt': pt, 'so': so, 'k_eff': k_eff,
        'fc_cube': st.session_state.get('fc_cube', 350),
        'fc_cyl': fc_cyl, 'sc': SC_FIXED, 'ec': ec_psi,
        'j': j_val, 'cd': cd, 'dpsi': dpsi, 'k_opt': k_opt,
    }
    return {'rows': rows, 'j': j_val, 'k_eff': k_eff, 'k_opt': k_opt}


# ─────────────────────────────────────────────
#  Main Render
# ─────────────────────────────────────────────

def render():
    ss = st.session_state
    st.markdown("### 🏗️ Rigid Pavement Design — AASHTO 1993")

    # ── ดึง parameters จาก ITM Pave session_state ──
    pt = float(ss.get('pt_global', 2.5))
    so = float(ss.get('so_rig', 0.35))
    R  = int(ss.get('r0_rig', 90))
    zr = get_zr(R)

    # ── ดึง MR จาก CBR Analysis ──
    cbr_design = float(ss.get('cbr_design', 4.0))
    from engine.rigid_nomograph import mr_from_cbr
    MR_psi = float(ss.get('mr_subgrade_psi') or mr_from_cbr(cbr_design))

    # ── ดึง W18 จาก ESAL Calculator ──
    esal_rigid = ss.get('esal_rigid', {})
    w18_ref    = int(esal_rigid.get(30, esal_rigid.get(list(esal_rigid.keys())[0], 0)) if esal_rigid else 0)

    # ════════════════════════════════════════
    #  Card 1 — Status bar
    # ════════════════════════════════════════
    with st.container(border=True):
        st.markdown('<div class="rp-card-title">📋 สถานะข้อมูลจาก ESAL & CBR</div>',
                    unsafe_allow_html=True)
        s1, s2, s3 = st.columns(3)
        with s1:
            if w18_ref > 0:
                st.markdown(f'<div class="result-pass">✅ W18 = {w18_ref:,.0f} (D=30 cm)</div>',
                            unsafe_allow_html=True)
            else:
                st.markdown('<div class="result-warn">⚠️ ยังไม่มี ESAL — คำนวณใน ESAL Calculator ก่อน</div>',
                            unsafe_allow_html=True)
        with s2:
            kj = ss.get('jpcp_k_eff')
            if kj:
                st.markdown(f'<div class="result-pass">✅ k_eff JPCP = {kj:.0f} pci</div>',
                            unsafe_allow_html=True)
            else:
                st.markdown('<div class="result-warn">⚠️ ยังไม่มี k_eff JPCP (Section A)</div>',
                            unsafe_allow_html=True)
        with s3:
            kc = ss.get('crcp_k_eff')
            if kc:
                st.markdown(f'<div class="result-pass">✅ k_eff CRCP = {kc:.0f} pci</div>',
                            unsafe_allow_html=True)
            else:
                st.markdown('<div class="result-warn">⚠️ ยังไม่มี k_eff CRCP (Section A)</div>',
                            unsafe_allow_html=True)

    # ════════════════════════════════════════
    #  Card 2 — W18 Input
    # ════════════════════════════════════════
    with st.container(border=True):
        st.markdown('<div class="rp-card-title">🔢 W18 — ESAL ออกแบบ</div>',
                    unsafe_allow_html=True)
        if w18_ref > 0:
            use_manual = st.checkbox('กรอก W18 เองแทน', value=ss.get('w18_manual_mode', False),
                                      key='w18_manual_mode')
            if use_manual:
                w18_req = st.number_input('W18 (ESALs)', min_value=100_000, max_value=500_000_000,
                                           value=ss.get('w18_manual', w18_ref), step=100_000,
                                           key='w18_manual', format='%d')
            else:
                w18_req = w18_ref
                st.markdown(f'<div class="result-pass">✅ ใช้ W18 = <b>{w18_req:,.0f}</b> จาก ESAL Calculator (D=30 cm)</div>',
                            unsafe_allow_html=True)
        else:
            st.session_state['w18_manual_mode'] = True
            w18_req = st.number_input('W18 (ESALs) — กรอกเอง', min_value=100_000, max_value=500_000_000,
                                       value=ss.get('w18_manual', 5_000_000), step=100_000,
                                       key='w18_manual', format='%d')

    # ════════════════════════════════════════
    #  Card 3 — Shared Parameters
    # ════════════════════════════════════════
    with st.container(border=True):
        st.markdown('<div class="rp-card-title">⚙️ พารามิเตอร์ร่วม (JPCP & CRCP)</div>',
                    unsafe_allow_html=True)
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            fc_cube = st.number_input("f'c Cube (ksc)", 280, 600,
                                       ss.get('fc_cube', 350), step=10, key='fc_cube')
            if fc_cube < 350:
                st.warning('⚠️ ต่ำกว่า 350 ksc')
        fc_cyl = convert_cube_to_cyl(fc_cube)
        ec_psi = calc_ec(fc_cyl)
        with c2:
            _mbox("f'c,cyl / Ec", f'{ec_psi:,.0f}', 'psi', _JPCP_BD, _JPCP_BG)
        with c3:
            _mbox('Sc (ทล. lock)', f'{SC_FIXED:.0f}', 'psi', _CRCP_BD, _CRCP_BG)
        with c4:
            r0_val = st.selectbox('Reliability R0 (%)', [85, 90, 91, 92, 93, 94, 95, 96, 97, 98, 99],
                                   index=1, key='r0_rig')
            zr = get_zr(r0_val)
            st.caption(f'ZR = {zr:.3f}')

        c5, c6, c7 = st.columns(3)
        with c5:
            so = st.number_input('So', 0.20, 0.50, ss.get('so_rig', 0.35), 0.01, key='so_rig')
        with c6:
            pt = st.number_input('Pt', 1.5, 3.5, float(ss.get('pt_global', 2.5)), 0.1, key='pt_rig_v7')
        with c7:
            cd_str = st.radio('Cd', ['1.0 — ปกติ', '1.1 — ดี', '1.2 — ดีมาก'],
                               index=[1.0,1.1,1.2].index(ss.get('cd_rig', 1.0)),
                               key='cd_rig_radio', horizontal=True, label_visibility='collapsed')
            cd = float(cd_str.split(' — ')[0])
            st.session_state['cd_rig'] = cd

    # ════════════════════════════════════════
    #  Section A — Layers + k∞ (2 คอลัมน์)
    # ════════════════════════════════════════
    st.markdown('---')
    st.markdown('### 📐 Section A — Subbase Layers & k∞')

    # Row A: Layers
    col_j, col_c = st.columns(2)
    with col_j:
        _card_header('🔲  JPCP / JRCP — Subbase Layers', _JPCP_BD)
        with st.container(border=True):
            n_j = st.slider('จำนวนชั้น JPCP', 1, 6, ss.get('jpcp_n', 5), key='jpcp_n')
            layers_jpcp = _layers('jpcp', n_j, _DEF_JPCP)
            tot_j = sum(l['thickness_cm'] for l in layers_jpcp if l['thickness_cm'] > 0)
            st.caption(f'รวม = **{tot_j} ซม.**')

    with col_c:
        _card_header('〰️  CRCP — Subbase Layers', _CRCP_BD)
        with st.container(border=True):
            copy_jpcp = st.checkbox('ใช้ค่าเดียวกับ JPCP/JRCP',
                                     value=ss.get('crcp_copy', False), key='crcp_copy')
            if copy_jpcp:
                layers_crcp = layers_jpcp
                st.markdown(f'<div style="font-size:12px;color:{_CRCP_BD};background:{_CRCP_BG};'
                            f'border-radius:6px;padding:5px 10px;margin-bottom:6px">'
                            f'✅ ใช้ชั้นวัสดุเดียวกับ JPCP/JRCP</div>', unsafe_allow_html=True)
            else:
                n_c = st.slider('จำนวนชั้น CRCP', 1, 6, ss.get('crcp_n', 3), key='crcp_n')
                layers_crcp = _layers('crcp', n_c, _DEF_CRCP)
            tot_c = sum(l['thickness_cm'] for l in layers_crcp if l['thickness_cm'] > 0)
            st.caption(f'รวม = **{tot_c} ซม.**')

    # Row B: k∞ / LS / กราฟ
    col_j2, col_c2 = st.columns(2)
    with col_j2:
        _card_header('🔲  JPCP / JRCP — k∞ & Loss of Support', _JPCP_BD)
        with st.container(border=True):
            _kblock('jpcp', layers_jpcp, MR_psi)

    with col_c2:
        _card_header('〰️  CRCP — k∞ & Loss of Support', _CRCP_BD)
        with st.container(border=True):
            _kblock('crcp', layers_crcp, MR_psi)

    # Summary k_eff
    kj = ss.get('jpcp_k_eff')
    kc = ss.get('crcp_k_eff')
    if kj or kc:
        st.markdown('<div style="height:4px"></div>', unsafe_allow_html=True)
        with st.container(border=True):
            s_left, s_right = st.columns([2, 3])
            with s_left:
                st.markdown(f'<div style="font-size:13px;font-weight:700;color:{_CRCP_BD};'
                            f'padding:6px 0">✅ สรุป k_eff → ส่งต่อ Section B Design</div>',
                            unsafe_allow_html=True)
            with s_right:
                sc1, sc2 = st.columns(2)
                with sc1:
                    if kj: _mbox('k_eff — JPCP/JRCP', f'{kj:.0f}', 'pci', _JPCP_BD, _JPCP_BG)
                with sc2:
                    if kc: _mbox('k_eff — CRCP', f'{kc:.0f}', 'pci', _CRCP_BD, _CRCP_BG)

    # ════════════════════════════════════════
    #  Section B — Design (2 คอลัมน์)
    # ════════════════════════════════════════
    st.markdown('---')
    st.markdown('### 🏗️ Section B — Design JPCP / JRCP vs CRCP')

    col_j3, col_c3 = st.columns(2)
    with col_j3:
        _card_header('🔲  JPCP / JRCP — Design', _JPCP_BD)
        with st.container(border=True):
            res_j = _design_block('jpcp', 'JPCP/JRCP', fc_cyl, ec_psi, cd,
                                   w18_req, pt, zr, so, _JPCP_BD, _JPCP_BDLT)

    with col_c3:
        _card_header('〰️  CRCP — Design', _CRCP_BD)
        with st.container(border=True):
            res_c = _design_block('crcp', 'CRCP', fc_cyl, ec_psi, cd,
                                   w18_req, pt, zr, so, _CRCP_BD, _CRCP_BDLT)

    # ════════════════════════════════════════
    #  Comparison Summary
    # ════════════════════════════════════════
    if res_j or res_c:
        st.markdown('---')
        _render_comparison(res_j, res_c)


# ─────────────────────────────────────────────
#  Comparison Summary
# ─────────────────────────────────────────────

def _render_comparison(res_j, res_c):
    ss = st.session_state
    dj = ss.get('jpcp_rec_d_cm')
    dc = ss.get('crcp_rec_d_cm')
    kj = ss.get('jpcp_k_eff')
    kc = ss.get('crcp_k_eff')

    st.markdown('### 📊 สรุปเปรียบเทียบ JPCP vs CRCP')
    with st.container(border=True):
        items = [
            ('D แนะนำ',
             f'{round(dj/2.54)} in ({dj} ซม.)' if dj else '—',
             f'{round(dc/2.54)} in ({dc} ซม.)' if dc else '—'),
            ('k_eff (pci)',
             f'{kj:.0f}' if kj else '—',
             f'{kc:.0f}' if kc else '—'),
            ('J factor',
             f'{ss.get("jpcp_j", 2.8):.1f}',
             f'{ss.get("crcp_j", 2.6):.1f}'),
            ('W18 Capacity',
             f'{res_j["rows"][next((i for i,r in enumerate(res_j["rows"]) if r["d_cm"]==dj), 0)]["w18_cap"]:,.0f}' if res_j and dj else '—',
             f'{res_c["rows"][next((i for i,r in enumerate(res_c["rows"]) if r["d_cm"]==dc), 0)]["w18_cap"]:,.0f}' if res_c and dc else '—'),
            ('ผลการตรวจสอบ',
             '✅ PASS' if (dj and res_j) else '❌',
             '✅ PASS' if (dc and res_c) else '❌'),
        ]

        hc_l, hc_j, hc_c = st.columns([2, 1, 1])
        hc_l.markdown('**รายการ**')
        hc_j.markdown(f'<div style="color:{_JPCP_BD};font-weight:700">🔲 JPCP/JRCP</div>', unsafe_allow_html=True)
        hc_c.markdown(f'<div style="color:{_CRCP_BD};font-weight:700">〰️ CRCP</div>', unsafe_allow_html=True)
        st.markdown('---')
        for label, vj, vc in items:
            rc1, rc2, rc3 = st.columns([2, 1, 1])
            rc1.markdown(f'<div style="font-size:12px;color:#546E7A">{label}</div>', unsafe_allow_html=True)
            rc2.markdown(f'<div style="font-family:IBM Plex Mono,monospace;font-size:12px;font-weight:600;color:{_JPCP_BD}">{vj}</div>', unsafe_allow_html=True)
            rc3.markdown(f'<div style="font-family:IBM Plex Mono,monospace;font-size:12px;font-weight:600;color:{_CRCP_BD}">{vc}</div>', unsafe_allow_html=True)
