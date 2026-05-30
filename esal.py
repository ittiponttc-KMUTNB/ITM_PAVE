# ╔══════════════════════════════════════════════════════════════════╗
# ║  engine/esal.py — ITM Pave Pro                                  ║
# ║  ESAL Calculation Functions (AASHTO 1993)                       ║
# ║  ไม่มี st. ใดๆ ทั้งสิ้น — pure Python functions               ║
# ╚══════════════════════════════════════════════════════════════════╝

import math
import pandas as pd
from constants import (
    TON_TO_KIP, VEHICLE_AXLES, VEHICLE_COLS,
    SLAB_THICKNESSES, SN_DEFAULTS,
)


def ealf_flex(L1_ton: float, L2: float, SN: float, Pt: float) -> float:
    """
    คำนวณ Equivalent Axle Load Factor สำหรับ Flexible Pavement
    L1_ton : น้ำหนักเพลา (ตัน)
    L2     : Axle configuration code
    SN     : Structural Number
    Pt     : Terminal Serviceability
    """
    L1  = L1_ton * TON_TO_KIP
    Gt  = math.log10((4.2 - Pt) / (4.2 - 1.5))
    Bx  = 0.40 + 0.081*(L1+L2)**3.23 / ((SN+1)**5.19 * L2**3.23)
    B18 = 0.40 + 0.081*(18+1)**3.23  / ((SN+1)**5.19 * 1.0**3.23)
    return 10**(
        4.79*math.log10(L1+L2) - 4.33*math.log10(L2)
        - 4.79*math.log10(19) + Gt*(1/B18 - 1/Bx)
    )


def ealf_rigid(L1_ton: float, L2: float, D_cm: float, Pt: float) -> float:
    """
    คำนวณ Equivalent Axle Load Factor สำหรับ Rigid Pavement
    D_cm : ความหนา slab (cm) — แปลงเป็นนิ้วก่อนคำนวณ
    """
    L1  = L1_ton * TON_TO_KIP
    D   = round(D_cm / 2.54)   # AASHTO 1993 ใช้ความหนาเป็นจำนวนเต็มนิ้ว
    Gt  = math.log10((4.5 - Pt) / (4.5 - 1.5))
    Bx  = 1.0 + 3.63*(L1+L2)**5.20 / ((D+1)**8.46 * L2**3.52)
    B18 = 1.0 + 3.63*(18+1)**5.20  / ((D+1)**8.46 * 1.0**3.52)
    return 10**(
        4.62*math.log10(L1+L2) - 3.28*math.log10(L2)
        - 4.62*math.log10(19) + Gt*(1/B18 - 1/Bx)
    )


def truck_factor_flex(vtype: str, SN: float, Pt: float) -> float:
    """Truck Factor รวมทุกเพลาสำหรับ Flexible Pavement"""
    return sum(
        ealf_flex(L1, L2, SN, Pt) * cnt
        for L1, L2, cnt in VEHICLE_AXLES[vtype]
    )


def truck_factor_rigid(vtype: str, D_cm: float, Pt: float) -> float:
    """Truck Factor รวมทุกเพลาสำหรับ Rigid Pavement"""
    return sum(
        ealf_rigid(L1, L2, D_cm, Pt) * cnt
        for L1, L2, cnt in VEHICLE_AXLES[vtype]
    )


def compute_esal_from_df(
    traffic_df: pd.DataFrame,
    ldf: float,
    ddf: float,
    Pt: float,
    mode: str = "rigid",
    sn_list: list = None,
) -> dict:
    """
    คำนวณ ESAL จาก traffic DataFrame (AASHTO 1993)

    สูตร: ESAL = Σ_ปี [ AADT × 365 × DDF × LDF × TF ]

    Parameters
    ----------
    traffic_df : DataFrame คอลัมน์ Year, MB, HB, MT, HT, TR, STR
                 ค่า = AADT 2 ทิศทาง (คัน/วัน) ต่อปี
    ldf        : Lane Distribution Factor
    ddf        : Directional Distribution Factor
    Pt         : Terminal Serviceability
    mode       : "rigid" หรือ "flex"
    sn_list    : list ของ SN สำหรับ flexible (ถ้า None ใช้ SN_DEFAULTS)

    Returns
    -------
    dict : {D_cm: esal} สำหรับ rigid | {SN: esal} สำหรับ flexible
    """
    DAYS_PER_YEAR = 365

    if mode == "rigid":
        keys    = SLAB_THICKNESSES
        results = {k: 0.0 for k in keys}
        for _, row in traffic_df.iterrows():
            for vtype in VEHICLE_COLS:
                cnt = float(row.get(vtype, 0) or 0)
                if cnt <= 0:
                    continue
                for D in keys:
                    tf_val = truck_factor_rigid(vtype, D, Pt)
                    results[D] += cnt * DAYS_PER_YEAR * ddf * ldf * tf_val
        return results
    else:
        keys    = sn_list or SN_DEFAULTS
        results = {k: 0.0 for k in keys}
        for _, row in traffic_df.iterrows():
            for vtype in VEHICLE_COLS:
                cnt = float(row.get(vtype, 0) or 0)
                if cnt <= 0:
                    continue
                for SN in keys:
                    tf_val = truck_factor_flex(vtype, SN, Pt)
                    results[SN] += cnt * DAYS_PER_YEAR * ddf * ldf * tf_val
        return results


def grow_traffic(base_row: dict, growth_rate_pct: float, years: int) -> pd.DataFrame:
    """
    สร้าง DataFrame ปริมาณจราจรรายปี จาก base year + อัตราการเติบโต

    Parameters
    ----------
    base_row        : dict {vtype: AADT} ปีฐาน
    growth_rate_pct : อัตราการเติบโต (%)
    years           : จำนวนปีออกแบบ

    Returns
    -------
    DataFrame คอลัมน์ Year + VEHICLE_COLS
    """
    r    = growth_rate_pct / 100.0
    rows = []
    for y in range(1, years + 1):
        factor = (1 + r) ** (y - 1)
        row    = {"Year": y}
        for v in VEHICLE_COLS:
            row[v] = int(round(base_row.get(v, 0) * factor))
        rows.append(row)
    return pd.DataFrame(rows)
