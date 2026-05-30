# ╔══════════════════════════════════════════════════════════════════╗
# ║  engine/design.py — ITM Pave Pro                                ║
# ║  AASHTO 1993 Design Equations                                   ║
# ║  Flexible (SN) + Rigid (W18) + CBR/Mr/k conversions            ║
# ╚══════════════════════════════════════════════════════════════════╝

import math
import numpy as np

try:
    from scipy.optimize import brentq as _brentq
except ImportError:
    def _brentq(f, a, b, xtol=1e-6, maxiter=500):
        """Fallback bisection method ถ้าไม่มี scipy"""
        fa, fb = f(a), f(b)
        if fa * fb > 0:
            raise ValueError("No sign change in interval")
        for _ in range(maxiter):
            mid = (a + b) / 2.0
            fm  = f(mid)
            if abs(fm) < xtol or (b - a) / 2.0 < xtol:
                return mid
            if fa * fm < 0:
                b, fb = mid, fm
            else:
                a, fa = mid, fm
        return (a + b) / 2.0


# ─────────────────────────────────────────────
#  Flexible Pavement
# ─────────────────────────────────────────────

def aashto_sn_required(
    esal: float,
    zr: float,
    so: float,
    pi: float,
    pt: float,
    mr_psi: float,
) -> float | None:
    """
    คำนวณ Structural Number (SN) ที่ต้องการ — AASHTO 1993 Flexible

    Parameters
    ----------
    esal   : Design ESAL
    zr     : Reliability factor (จาก ZR_MAP)
    so     : Overall standard deviation (0.40–0.50)
    pi     : Initial serviceability (4.0–4.5)
    pt     : Terminal serviceability (2.0–3.0)
    mr_psi : Resilient Modulus of subgrade (psi)

    Returns
    -------
    float : SN required | None ถ้าคำนวณไม่ได้
    """
    delta_psi = pi - pt
    logW18    = math.log10(max(esal, 1))

    def eq(SN):
        if SN <= 0:
            return -1e10
        t1 = zr * so
        t2 = 9.36 * math.log10(SN + 1) - 0.20
        t3 = math.log10(delta_psi / 2.7) / (0.40 + 1094 / (SN + 1) ** 5.19)
        t4 = 2.32 * math.log10(mr_psi) - 8.07
        return t1 + t2 + t3 + t4 - logW18

    try:
        return _brentq(eq, 0.1, 30, xtol=1e-4)
    except Exception:
        return None


# ─────────────────────────────────────────────
#  Rigid Pavement
# ─────────────────────────────────────────────

def aashto_rigid_w18(
    d_cm: float,
    pi: float,
    pt: float,
    zr: float,
    so: float,
    sc_psi: float,
    cd: float,
    j: float,
    ec_psi: float,
    k_pci: float,
) -> float | None:
    """
    คำนวณ Allowable W18 สำหรับ Rigid Pavement — AASHTO 1993

    Parameters
    ----------
    d_cm   : ความหนา slab (cm)
    pi     : Initial serviceability
    pt     : Terminal serviceability
    zr     : Reliability factor
    so     : Overall standard deviation
    sc_psi : Modulus of Rupture (psi)
    cd     : Drainage coefficient
    j      : Load transfer coefficient (JPCP=2.8, CRCP=2.6)
    ec_psi : Elastic Modulus of concrete (psi)
    k_pci  : Modulus of subgrade reaction (pci)

    Returns
    -------
    float : Allowable W18 | None ถ้าคำนวณไม่ได้
    """
    d_in      = round(d_cm / 2.54)   # AASHTO 1993 ใช้ความหนาเป็นจำนวนเต็มนิ้ว
    delta_psi = pi - pt

    t1 = zr * so
    t2 = 7.35 * math.log10(d_in + 1) - 0.06
    t3 = math.log10(delta_psi / 3.0) / (1 + 1.624e7 / (d_in + 1) ** 8.46)

    num4 = sc_psi * cd * (d_in ** 0.75 - 1.132)
    den4 = 215.63 * j * (d_in ** 0.75 - 18.42 / (ec_psi / k_pci) ** 0.25)

    if num4 <= 0 or den4 <= 0:
        return None

    inner = num4 / den4
    if inner <= 0:
        return None

    t4 = (4.22 - 0.32 * pt) * math.log10(inner)
    return 10 ** (t1 + t2 + t3 + t4)


# ─────────────────────────────────────────────
#  CBR / Mr / k Conversions
# ─────────────────────────────────────────────

def cbr_to_mr(cbr: float) -> float:
    """CBR (%) → Resilient Modulus (psi)  [AASHTO 1993]"""
    return 1500.0 * cbr


def mr_to_k(mr_psi: float) -> float:
    """Mr (psi) → k-value (pci)  [approximate]"""
    return mr_psi / 19.4


def calc_percentile_cbr(cbr_values: list) -> tuple:
    """
    คำนวณ CBR percentile สำหรับ Design CBR

    Returns
    -------
    arr        : sorted array ของค่า CBR ทั้งหมด
    n          : จำนวนตัวอย่าง
    unique_cbr : ค่า CBR ที่ไม่ซ้ำ (sorted)
    unique_pct : เปอร์เซ็นต์ที่มีค่า ≥ ค่านั้น
    """
    arr        = np.sort(np.array(cbr_values, dtype=float))
    n          = len(arr)
    unique_cbr = np.unique(arr)
    unique_pct = np.array([np.sum(arr >= v) / n * 100 for v in unique_cbr])
    return arr, n, unique_cbr, unique_pct
