# ╔══════════════════════════════════════════════════════════════════╗
# ║  engine/figures.py — ITM Pave Pro                               ║
# ║  Matplotlib Figures: Pavement Structure + Nomographs            ║
# ║  ไม่มี st. ใดๆ ทั้งสิ้น — pure matplotlib functions            ║
# ║  พัฒนาโดย รศ.ดร.อิทธิพล มีผล | ภาควิชาครุศาสตร์โยธา มจพ.    ║
# ╚══════════════════════════════════════════════════════════════════╝

import io
import math

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as patches


# ─────────────────────────────────────────────
#  Layer Name / Color / Text Lookup Tables
# ─────────────────────────────────────────────

# แปลงชื่อวัสดุภาษาไทย → อังกฤษ สำหรับแสดงในรูป
_LAYER_NAME_EN = {
    "ผิวทางลาดยาง PMA":                                    "PMA Surface",
    "ผิวทางแอสฟัลต์คอนกรีต (AC)":                         "AC Surface",
    "(CTB) หินคลุกปรับปรุงด้วยปูนซีเมนต์ UCS 40 ksc ":   "Cement Treated Base",
    "หินคลุกผสมซีเมนต์ UCS 24.5 ksc":                     "MOD. Crushed Rock",
    "ดินซีเมนต์ UCS 17.5 ksc":                             "Soil Cement",
    "หินคลุก CBR 80%":                                     "Crushed Rock (CBR 80%)",
    "วัสดุหมุนเวียน (Recycling)":                          "Recycled Material",
    "วัสดุมวลรวม CBR 25%":                                 "Aggregate Subbase",
    "รองพื้นทางวัสดุมวลรวม CBR 25%":                      "Aggregate Subbase",
    "วัสดุคัดเลือก ก":                                     "Selected A",
    "AC รองใต้ผิวคอนกรีต":                                 "AC Interlayer",
    "ดินถมคันทาง CBR 10%":                                 "Embankment",
    "ดินถมคันทาง CBR กรอกเอง":                             "Embankment",
    "ดินถมคันทาง / ดินเดิม":                               "Embankment",
    "Concrete Slab":                                        "Concrete Slab",
}

# สีพื้นหลังแต่ละวัสดุ
_LAYER_COLORS = {
    "ผิวทางลาดยาง PMA":                                    "#1A252F",
    "ผิวทางแอสฟัลต์คอนกรีต (AC)":                         "#2C3E50",
    "Concrete Slab":                                        "#78909C",
    "(CTB) หินคลุกปรับปรุงด้วยปูนซีเมนต์ UCS 40 ksc ":   "#7F8C8D",
    "หินคลุกผสมซีเมนต์ UCS 24.5 ksc":                     "#95A5A6",
    "ดินซีเมนต์ UCS 17.5 ksc":                             "#AAB7B8",
    "หินคลุก CBR 80%":                                     "#BDC3C7",
    "วัสดุหมุนเวียน (Recycling)":                          "#85929E",
    "วัสดุมวลรวม CBR 25%":                                 "#FFCC99",
    "รองพื้นทางวัสดุมวลรวม CBR 25%":                      "#FFCC99",
    "วัสดุคัดเลือก ก":                                     "#E8DAEF",
    "AC รองใต้ผิวคอนกรีต":                                 "#34495E",
    "ดินถมคันทาง CBR 10%":                                 "#F5CBA7",
    "ดินถมคันทาง CBR กรอกเอง":                             "#F5CBA7",
    "ดินถมคันทาง / ดินเดิม":                               "#F5CBA7",
}

# วัสดุที่ใช้ text สีขาว (background เข้ม)
_DARK_LAYERS = {
    "ผิวทางลาดยาง PMA",
    "ผิวทางแอสฟัลต์คอนกรีต (AC)",
    "Concrete Slab",
    "AC รองใต้ผิวคอนกรีต",
    "(CTB) หินคลุกปรับปรุงด้วยปูนซีเมนต์ UCS 40 ksc ",
    "หินคลุกผสมซีเมนต์ UCS 24.5 ksc",
    "วัสดุหมุนเวียน (Recycling)",
}


# ─────────────────────────────────────────────
#  Helper
# ─────────────────────────────────────────────

def fig_to_bytes(fig) -> bytes:
    """แปลง matplotlib figure → PNG bytes สำหรับ Word report"""
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=150, bbox_inches='tight',
                facecolor=fig.get_facecolor())
    buf.seek(0)
    return buf.read()


# ─────────────────────────────────────────────
#  Pavement Structure Figure
# ─────────────────────────────────────────────

def draw_pavement_structure(
    layers: list,
    mode: str = "flex",
    cbr_subgrade: float = 3.0,
    d_concrete_cm: float = None,
    ptype: str = "JPCP",
):
    """
    วาดรูปโครงสร้างชั้นทาง (matplotlib) — ชื่อวัสดุภาษาอังกฤษ

    Parameters
    ----------
    layers        : list of dict — แต่ละ dict มี keys:
                    name, thickness_cm, ai (flex), sni (flex), E_MPa (rigid)
    mode          : "flex" | "rigid"
    cbr_subgrade  : CBR ดินเดิม (%) สำหรับ label Subgrade
    d_concrete_cm : ความหนา Slab cm (rigid เท่านั้น)
    ptype         : "JPCP" | "JRCP" | "CRCP"

    Returns
    -------
    matplotlib.figure.Figure | None
    """
    MIN_H   = 5       # ความสูงขั้นต่ำ (display units)
    W       = 3.0     # ความกว้าง block
    X_CTR   = 5.0     # กึ่งกลาง X
    X_START = X_CTR - W / 2

    # ── เตรียม layer list ──
    all_layers = []

    if mode == "rigid" and d_concrete_cm:
        all_layers.append({
            "name":         "Concrete Slab",
            "thickness_cm": d_concrete_cm,
            "label":        f"Concrete Slab ({ptype})",
            "side_info":    None,
        })

    valid = [l for l in layers if l.get("thickness_cm", 0) > 0]
    for l in valid:
        name  = l.get("name", "")
        h     = l.get("thickness_cm", 0)
        en    = _LAYER_NAME_EN.get(name, name)
        if mode == "flex":
            ai  = l.get("ai",  None)
            sni = l.get("sni", None)
            side = f"ai={ai:.2f} | SNi={sni:.3f}" if ai and sni else None
        else:
            e_mpa = l.get("E_MPa", None)
            side  = f"E={e_mpa:,} MPa" if e_mpa else None
        all_layers.append({
            "name":         name,
            "thickness_cm": h,
            "label":        en,
            "side_info":    side,
        })

    # Subgrade (แสดงเป็น infinite depth)
    all_layers.append({
        "name":         "ดินถมคันทาง / ดินเดิม",
        "thickness_cm": 0,
        "label":        f"Subgrade (CBR≥{cbr_subgrade:.0f}%)",
        "side_info":    None,
    })

    if len(all_layers) <= 1:
        return None

    # ── Display height — normalize ──
    n_layers  = len(all_layers)
    real_h    = [l["thickness_cm"] for l in all_layers]
    max_real  = max((h for h in real_h if h > 0), default=30)
    SCALE     = 40.0 / max_real   # layer ใหญ่สุด = 40 display units
    display_h = []
    for h in real_h:
        display_h.append(MIN_H if h == 0 else max(round(h * SCALE, 1), MIN_H))

    total_disp  = sum(display_h)
    total_thick = sum(h for h in real_h if h > 0)

    # ── Figure size ──
    fig_h  = max(4.0, min(n_layers * 1.4, 8.0))
    fig_w  = 9.0
    fs_lbl = max(10.0, 12.0 - n_layers * 0.3)
    fs_h   = max(10.0, 12.0 - n_layers * 0.3)

    fig, ax = plt.subplots(figsize=(fig_w, fig_h))
    ax.set_xlim(0, 11)
    ax.set_ylim(-4, total_disp + 5)
    ax.axis('off')
    fig.patch.set_facecolor('white')

    y = total_disp

    for i, layer in enumerate(all_layers):
        dh    = display_h[i]
        name  = layer["name"]
        label = layer["label"]
        h_cm  = layer["thickness_cm"]
        side  = layer["side_info"]

        color   = _LAYER_COLORS.get(name, "#D5D8DC")
        hatch   = '///' if "หมุนเวียน" in name else None
        is_dark = name in _DARK_LAYERS
        txt_col = 'white' if is_dark else '#1A1A1A'

        y_bot = y - dh
        y_ctr = y_bot + dh / 2

        # Rectangle
        rect = patches.Rectangle(
            (X_START, y_bot), W, dh,
            linewidth=1.5, edgecolor='#2C3E50',
            facecolor=color, hatch=hatch, zorder=2
        )
        ax.add_patch(rect)

        # ความหนาในกล่อง
        h_text = f"{h_cm} cm" if h_cm > 0 else "∞"
        ax.text(X_CTR, y_ctr, h_text,
                ha='center', va='center',
                fontsize=fs_h, fontweight='bold',
                color=txt_col, zorder=3)

        # ชื่อวัสดุซ้าย
        ax.text(X_START - 0.2, y_ctr, label,
                ha='right', va='center',
                fontsize=fs_lbl, fontweight='bold',
                color='#1B2631', zorder=3)

        # ข้อมูลขวา (ai/SNi หรือ E_MPa)
        if side:
            ax.text(X_START + W + 0.2, y_ctr, side,
                    ha='left', va='center',
                    fontsize=max(fs_lbl - 1, 7.0),
                    color='#154360', zorder=3)

        # เส้นคั่นระหว่าง layer
        if i > 0:
            ax.plot([X_START, X_START + W], [y, y],
                    color='#2C3E50', lw=1.0, zorder=3)

        y = y_bot

    # ── ลูกศร Total Thickness ──
    x_arr   = X_START + W + 2.8
    y_solid = sum(display_h[:-1])   # ไม่รวม subgrade
    ax.annotate('', xy=(x_arr, total_disp),
                xytext=(x_arr, total_disp - y_solid),
                arrowprops=dict(arrowstyle='<->', color='#C0392B', lw=1.8))
    ax.text(x_arr + 0.2, total_disp - y_solid / 2,
            f"Total\n{total_thick} cm",
            ha='left', va='center',
            fontsize=9, color='#C0392B', fontweight='bold')

    plt.tight_layout()
    return fig


# ─────────────────────────────────────────────
#  Composite k∞ Nomograph (AASHTO 1993 Fig. 3.3)
# ─────────────────────────────────────────────

def draw_k_infinity_nomograph(
    esb_psi: float,
    dsb_in: float,
    k_sub_pci: float,
) -> tuple:
    """
    Composite k∞ Nomograph (AASHTO 1993 Fig. 3.3 approximation)
    3 axes: Esb (left) | DSB (center) | k∞ (right)

    Returns
    -------
    (fig, k_inf_calc) : matplotlib Figure และค่า k∞ ที่คำนวณได้ (pci)
    """
    fig, ax = plt.subplots(figsize=(8, 9))
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 10)
    ax.axis('off')
    ax.set_facecolor('#F1F8E9')
    fig.patch.set_facecolor('#F1F8E9')

    x_esb, x_dsb, x_kinf = 1.5, 5.0, 8.5

    # ── Axis lines ──
    for x in [x_esb, x_dsb, x_kinf]:
        ax.plot([x, x], [0.5, 9.5], color='#1B5E20', lw=2.5)

    # ── Esb axis: 5,000–100,000 psi (log scale) ──
    esb_range   = [5000, 10000, 20000, 30000, 50000, 100000]
    esb_log_min = math.log10(5000)
    esb_log_max = math.log10(100000)

    def esb_to_y(v):
        return 0.5 + 9.0 * (math.log10(v) - esb_log_min) / (esb_log_max - esb_log_min)

    for v in esb_range:
        y = esb_to_y(v)
        ax.plot([x_esb - 0.15, x_esb + 0.15], [y, y], color='#1B5E20', lw=1.5)
        ax.text(x_esb - 0.25, y, f"{v:,}", ha='right', va='center',
                fontsize=8, color='#1B5E20')
    ax.text(x_esb, 9.8, "Esb (psi)", ha='center', va='bottom',
            fontsize=9, fontweight='bold', color='#1B5E20')

    # ── DSB axis: 0–36 in (linear scale) ──
    dsb_range = [0, 4, 8, 12, 16, 20, 24, 28, 32, 36]

    def dsb_to_y(v):
        return 0.5 + 9.0 * (v / 36.0)

    for v in dsb_range:
        y = dsb_to_y(v)
        ax.plot([x_dsb - 0.15, x_dsb + 0.15], [y, y], color='#2E7D32', lw=1.5)
        ax.text(x_dsb + 0.25, y, f"{v}", ha='left', va='center',
                fontsize=8, color='#2E7D32')
    ax.text(x_dsb, 9.8, "DSB (in)", ha='center', va='bottom',
            fontsize=9, fontweight='bold', color='#2E7D32')

    # ── k∞ axis: 50–1000 pci (log scale) ──
    kinf_range   = [50, 100, 150, 200, 300, 500, 700, 1000]
    kinf_log_min = math.log10(50)
    kinf_log_max = math.log10(1000)

    def kinf_to_y(v):
        return 0.5 + 9.0 * (math.log10(v) - kinf_log_min) / (kinf_log_max - kinf_log_min)

    for v in kinf_range:
        y = kinf_to_y(v)
        ax.plot([x_kinf - 0.15, x_kinf + 0.15], [y, y], color='#388E3C', lw=1.5)
        ax.text(x_kinf + 0.25, y, f"{v}", ha='left', va='center',
                fontsize=8, color='#388E3C')
    ax.text(x_kinf, 9.8, "k∞ (pci)", ha='center', va='bottom',
            fontsize=9, fontweight='bold', color='#388E3C')

    # ── คำนวณ k∞ จาก inputs (AASHTO Odemark approximation) ──
    if esb_psi > 0 and dsb_in >= 0 and k_sub_pci > 0:
        if dsb_in == 0:
            k_inf_calc = k_sub_pci
        else:
            h_eq       = dsb_in * (esb_psi / (k_sub_pci * 19.4)) ** (1 / 3)
            k_inf_calc = min(k_sub_pci * (1 + 0.55 * h_eq ** 0.45), 1000)
        k_inf_calc = max(50, min(1000, k_inf_calc))
    else:
        k_inf_calc = k_sub_pci

    # ── เส้น Reading ──
    y_esb  = esb_to_y(max(5000,  min(100000, esb_psi)))
    y_dsb  = dsb_to_y(max(0,     min(36,     dsb_in)))
    y_kinf = kinf_to_y(max(50,   min(1000,   k_inf_calc)))

    ax.annotate("", xy=(x_dsb, y_dsb), xytext=(x_esb, y_esb),
                arrowprops=dict(arrowstyle="-", color='red', lw=2, linestyle='dashed'))
    ax.annotate("", xy=(x_kinf, y_kinf), xytext=(x_dsb, y_dsb),
                arrowprops=dict(arrowstyle="->", color='red', lw=2, linestyle='dashed'))

    for xp, yp in [(x_esb, y_esb), (x_dsb, y_dsb), (x_kinf, y_kinf)]:
        ax.plot(xp, yp, 'o', color='red', markersize=8, zorder=5)

    ax.text(x_kinf + 1.0, y_kinf,
            f"k∞ = {k_inf_calc:.0f} pci",
            ha='left', va='center', fontsize=11, fontweight='bold', color='red',
            bbox=dict(boxstyle='round,pad=0.3', facecolor='white',
                      edgecolor='red', alpha=0.9))

    ax.set_title("Composite k∞ Nomograph (AASHTO 1993 Fig.3.3)",
                 fontsize=11, fontweight='bold', color='#1B5E20', pad=15)
    plt.tight_layout()
    return fig, k_inf_calc


# ─────────────────────────────────────────────
#  Loss of Support Nomograph (AASHTO 1993 Fig. 3.7)
# ─────────────────────────────────────────────

def draw_loss_of_support_nomograph(
    k_inf_pci: float,
    ls_value: float,
) -> tuple:
    """
    Loss of Support Nomograph (AASHTO 1993 Fig. 3.7)
    k_corrected = k_inf / 10^(LS × 0.5)

    Returns
    -------
    (fig, k_corr_calc) : matplotlib Figure และค่า k_eff ที่คำนวณได้ (pci)
    """
    fig, ax = plt.subplots(figsize=(7, 8))
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 10)
    ax.axis('off')
    ax.set_facecolor('#F1F8E9')
    fig.patch.set_facecolor('#F1F8E9')

    x_kinf, x_kcorr = 2.5, 7.5

    ls_colors = {
        0.0: '#1B5E20',
        0.5: '#2E7D32',
        1.0: '#43A047',
        1.5: '#66BB6A',
        2.0: '#EF6C00',
        3.0: '#B71C1C',
    }

    k_range   = [10, 20, 50, 100, 200, 300, 500, 700, 1000, 1500, 2000, 3000]
    k_log_min = math.log10(10)
    k_log_max = math.log10(3000)

    def k_to_y(v):
        return 0.5 + 9.0 * (math.log10(max(v, 10)) - k_log_min) / (k_log_max - k_log_min)

    # ── Axis lines ──
    for x in [x_kinf, x_kcorr]:
        ax.plot([x, x], [0.5, 9.5], color='#1B5E20', lw=2.5)

    # ── Scale marks ──
    for v in k_range:
        for x in [x_kinf, x_kcorr]:
            y = k_to_y(v)
            ax.plot([x - 0.15, x + 0.15], [y, y], color='#1B5E20', lw=1.5)
        ax.text(x_kinf - 0.25, k_to_y(v), f"{v}",
                ha='right', va='center', fontsize=8, color='#1B5E20')
        ax.text(x_kcorr + 0.25, k_to_y(v), f"{v}",
                ha='left', va='center', fontsize=8, color='#388E3C')

    ax.text(x_kinf,  9.8, "k∞ (pci)",   ha='center', va='bottom',
            fontsize=9, fontweight='bold', color='#1B5E20')
    ax.text(x_kcorr, 9.8, "k_eff (pci)", ha='center', va='bottom',
            fontsize=9, fontweight='bold', color='#388E3C')

    # ── LS family lines ──
    for ls, lc in ls_colors.items():
        for k_val in [20, 50, 100, 200, 500, 1000, 2000]:
            k_corr_ls = max(10, min(3000, k_val / (10 ** (ls * 0.5))))
            ax.plot([x_kinf, x_kcorr], [k_to_y(k_val), k_to_y(k_corr_ls)],
                    color=lc, lw=0.8, alpha=0.4)
        ax.text(5.0, k_to_y(50 / (10 ** (ls * 0.5))) + ls * 0.3,
                f"LS={ls}", ha='center', va='center',
                fontsize=7, color=lc, alpha=0.8)

    # ── คำนวณ k_eff ──
    k_corr_calc = max(10, min(3000, k_inf_pci / (10 ** (ls_value * 0.5))))

    # ── เส้น Reading ──
    y1 = k_to_y(max(10, min(3000, k_inf_pci)))
    y2 = k_to_y(k_corr_calc)
    ax.annotate("", xy=(x_kcorr, y2), xytext=(x_kinf, y1),
                arrowprops=dict(arrowstyle="->", color='red', lw=2.5))
    for xp, yp in [(x_kinf, y1), (x_kcorr, y2)]:
        ax.plot(xp, yp, 'o', color='red', markersize=9, zorder=5)

    ax.text(x_kcorr + 1.2, y2,
            f"k_eff =\n{k_corr_calc:.0f} pci",
            ha='left', va='center', fontsize=10, fontweight='bold', color='red',
            bbox=dict(boxstyle='round,pad=0.3', facecolor='white',
                      edgecolor='red', alpha=0.9))

    ax.set_title(f"Loss of Support Nomograph  (LS = {ls_value})",
                 fontsize=11, fontweight='bold', color='#1B5E20', pad=15)

    # ── Legend ──
    legend_x, legend_y = 3.5, 1.5
    for i, (ls, lc) in enumerate(ls_colors.items()):
        ax.plot(legend_x, legend_y - i * 0.3, 's', color=lc, markersize=7)
        ax.text(legend_x + 0.2, legend_y - i * 0.3,
                f"LS = {ls}", va='center', fontsize=7, color=lc)

    plt.tight_layout()
    return fig, k_corr_calc
