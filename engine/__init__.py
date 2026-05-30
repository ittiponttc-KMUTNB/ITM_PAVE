# engine/__init__.py
# expose ฟังก์ชันหลักทั้งหมดเพื่อ import สะดวก

from engine.esal import (
    ealf_flex,
    ealf_rigid,
    truck_factor_flex,
    truck_factor_rigid,
    compute_esal_from_df,
    grow_traffic,
)

from engine.design import (
    aashto_sn_required,
    aashto_rigid_w18,
    cbr_to_mr,
    mr_to_k,
    calc_percentile_cbr,
)
