# -*- coding: utf-8 -*-
"""Endurance test: cycle, PV2 readback, then PUND readback."""

from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd

from src.data_processing import analyze_pund_diff, calculate_polarization, read_both_channels, save_channels_separate_excel
from src.pmu_tests import execute_segARB_test, hy_pund_segARB, hy_pv2_segARB, power_off_outputs
from src.session import PMUSession

INST = "TCPIP0::129.125.87.80::1225::SOCKET"
CH1, CH2 = 1, 2

params_cycle = dict(
    rise_time_cycle=10e-6,
    Vc=3.0,
    offset_c=0,
    Irange1=1e-3,
    Irange2=1e-3,
)
params_pv2 = dict(
    rise_time=50e-6,
    Vp=2.0,
    offset=0,
    area_cm2=1.2567e-3,
    Irange1=1e-3,
    Irange2=1e-3,
)
params_pund = dict(
    rise_time=50e-6,
    Vp=1.5,
    offset=0,
    area_cm2=1.2567e-3,
    Irange1=1e-3,
    Irange2=1e-3,
)

cycle_counts = [1, 10, 100, 1000, 10000, 1e5, 1e6, 1e7]
SAVE_DIR = Path(r"C:\Users\P317151\Documents\data\FTJ\Refined\Endurance")
SAVE_DIR.mkdir(parents=True, exist_ok=True)


def run_cycle_block(query, n_cycles):
    """Run the non-measuring endurance cycle waveform."""
    current_ranges = {CH1: params_cycle["Irange1"], CH2: params_cycle["Irange2"]}
    rise_time = params_cycle["rise_time_cycle"]
    cycle_voltage = params_cycle["Vc"]
    offset = params_cycle.get("offset_c", 0)
    start_v = [offset, -cycle_voltage + offset, offset, cycle_voltage + offset]
    stop_v = [-cycle_voltage + offset, offset, cycle_voltage + offset, offset]
    time_v = [rise_time] * 4
    meas_types = [0, 0, 0, 0]
    ch1_config = (1, start_v, stop_v, time_v, meas_types, [0.0] * 4, [1] * 4)
    ch2_config = (1, [0.0] * 4, [0.0] * 4, time_v, meas_types, [0.0] * 4, [1] * 4)
    seq_configs = {CH1: [ch1_config], CH2: [ch2_config]}
    seq_list = {CH1: [(1, int(n_cycles))], CH2: [(1, int(n_cycles))]}
    execute_segARB_test(query, [CH1, CH2], seq_configs, seq_list=seq_list, current_ranges=current_ranges)
    power_off_outputs(query, (CH1, CH2))


with PMUSession(INST, channels=(CH1, CH2)) as session:
    query = session.query
    print(f"Starting endurance test with {len(cycle_counts)} cycle-count steps.")

    for index, n_cycles in enumerate(cycle_counts, start=1):
        print(f"Step {index}/{len(cycle_counts)}: {int(n_cycles)} cycles")
        run_cycle_block(query, n_cycles)

        hy_pv2_segARB(query, CH1, CH2, params_pv2)
        pv2_ch1, pv2_ch2 = read_both_channels(query, CH1, CH2)
        power_off_outputs(query, (CH1, CH2))
        if pv2_ch1 is None or pv2_ch1.empty:
            raise ValueError(f"PV2 readback after {int(n_cycles)} cycles returned no data.")

        polarization = calculate_polarization(
            pv2_ch1[f"Current {CH1}"].values,
            pv2_ch1[f"Timestamp {CH1}"].values,
            params_pv2.get("area_cm2", 1.0),
        )
        pv2_df = pd.DataFrame(
            {
                "Time": pv2_ch1[f"Timestamp {CH1}"].values,
                "Voltage": pv2_ch1[f"Voltage {CH1}"].values,
                "Current": pv2_ch1[f"Current {CH1}"].values,
                "Polarization": polarization,
            }
        )
        fname_pv2 = SAVE_DIR / f"Endurance_PV2_after_{int(n_cycles)}cycles"
        save_channels_separate_excel({1: pv2_ch1, 2: pv2_ch2}, f"{fname_pv2}_raw.xlsx")
        pv2_df.to_excel(f"{fname_pv2}_analysis.xlsx", index=False)

        fig, ax = plt.subplots(figsize=(6, 5))
        ax.plot(pv2_df["Voltage"], pv2_df["Polarization"], "b-")
        ax.set_xlabel("Voltage (V)")
        ax.set_ylabel("Polarization (uC/cm^2)")
        ax.set_title(f"PV2 after {int(n_cycles)} cycles")
        ax.grid(alpha=0.3)
        fig.tight_layout()
        fig.savefig(f"{fname_pv2}_loop.png", dpi=300)
        plt.close(fig)

        hy_pund_segARB(query, CH1, CH2, params_pund)
        df_ch1, df_ch2 = read_both_channels(query, CH1, CH2)
        power_off_outputs(query, (CH1, CH2))
        if df_ch1 is None or df_ch2 is None or df_ch1.empty or df_ch2.empty:
            raise ValueError(f"PUND readback after {int(n_cycles)} cycles returned no data.")

        pund_result = analyze_pund_diff(df_ch1, df_ch2, params_pund)
        fname_pund = SAVE_DIR / f"Endurance_PUND_after_{int(n_cycles)}cycles"
        save_channels_separate_excel({1: df_ch1, 2: df_ch2}, f"{fname_pund}_raw.xlsx")
        pund_result["df_total"].to_excel(f"{fname_pund}_total.xlsx", index=False)
        pund_result["pund_diff"].to_excel(f"{fname_pund}_diff.xlsx", index=False)

        fig, ax = plt.subplots(figsize=(7, 5))
        for seg in ["P", "U", "N", "D"]:
            sub = pund_result["pund_diff"][pund_result["pund_diff"]["Segment"] == seg]
            ax.plot(sub["Voltage"], sub["Polarization"], ".", label=seg, markersize=4)
        ax.set_xlabel("Voltage (V)")
        ax.set_ylabel("Polarization (uC/cm^2)")
        ax.set_title(f"PUND after {int(n_cycles)} cycles")
        ax.legend()
        ax.grid(alpha=0.3)
        fig.tight_layout()
        fig.savefig(f"{fname_pund}_loop.png", dpi=300)
        plt.close(fig)

        print(f"Completed step {index}.")
