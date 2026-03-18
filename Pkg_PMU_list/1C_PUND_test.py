# -*- coding: utf-8 -*-
"""PUND test script with shared PMU session management."""

from pathlib import Path

import matplotlib.pyplot as plt

from src.data_processing import analyze_pund_diff, read_both_channels, save_channels_separate_excel
from src.pmu_tests import hy_pund_segARB, power_off_outputs
from src.session import PMUSession

INST = "TCPIP0::129.125.87.80::1225::SOCKET"
CH1, CH2 = 1, 2
params = dict(
    rise_time=5e-5,
    rise_point=250,
    Vp=3,
    offset=0,
    area_cm2=1.2567e-5,
    Irange1=1e-4,
    Irange2=1e-4,
)

SAVE_DIR = Path(r"C:\Users\P317151\Documents\data\11-02-2026")
SAVE_DIR.mkdir(parents=True, exist_ok=True)
fname_base = SAVE_DIR / f"PUND_{int(params['rise_time'] * 1e6)}us_{params['Vp']}V"


with PMUSession(INST, channels=(CH1, CH2)) as session:
    Q = session.query
    print("Running PUND...")
    hy_pund_segARB(Q, CH1, CH2, params)
    df_ch1, df_ch2 = read_both_channels(Q, CH1, CH2)
    power_off_outputs(Q, (CH1, CH2))
    if df_ch1 is None or df_ch2 is None or df_ch1.empty or df_ch2.empty:
        raise ValueError("PUND returned empty channel data.")

    data = analyze_pund_diff(df_ch1, df_ch2, params)
    save_channels_separate_excel({1: df_ch1, 2: df_ch2}, f"{fname_base}_raw.xlsx")
    data["df_total"].to_excel(f"{fname_base}_total.xlsx", index=False)
    data["pund_diff"].to_excel(f"{fname_base}_diff.xlsx", index=False)

    fig, ax = plt.subplots(figsize=(7, 5))
    for seg in ["P", "U", "N", "D"]:
        sub = data["pund_diff"][data["pund_diff"]["Segment"] == seg]
        ax.plot(sub["Voltage"], sub["Polarization"], ".", label=seg, markersize=4)
    ax.set_xlabel("Voltage (V)")
    ax.set_ylabel("Polarization (uC/cm^2)")
    ax.set_title("PUND Differential Polarization")
    ax.legend()
    ax.grid(alpha=0.3)
    fig.tight_layout()
    fig.savefig(f"{fname_base}_loop.png", dpi=300)
    plt.close(fig)
    print("PUND complete.")
