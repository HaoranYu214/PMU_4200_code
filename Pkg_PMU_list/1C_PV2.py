# -*- coding: utf-8 -*-
"""PV2 segARB test with direct sequence definitions."""

from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd

from debug.waveform_preview import preview_sequence_configs
from src.data_processing import calculate_polarization, read_both_channels, save_channels_separate_excel
from src.pmu_tests import execute_segARB_test, power_off_outputs
from src.session import PMUSession

INST = "TCPIP0::129.125.87.80::1225::SOCKET"
CH1, CH2 = 1, 2
params = dict(
    rise_time=5e-5,
    rise_point=200,
    Vp=5,
    offset=0,
    area_cm2=7.0686e-6,
    Irange1=1e-4,
    Irange2=1e-4,
)
PREVIEW_ONLY = True
SAVE_DIR = Path(r"C:\Users\P317151\Documents\data\Jingtian\2025-12-14\BTO\Device2")
SAVE_DIR.mkdir(parents=True, exist_ok=True)
fname_base = SAVE_DIR / f"PV2_{int(params['rise_time'] * 1e6)}us_{params['Vp']}V"


def make_pv2_seq_configs():
    """Build PV2 seq_configs directly in this script."""
    rise_time = params["rise_time"]
    vp = params["Vp"]
    offset = params["offset"]

    start_voltages = [
        0,
        offset,
        offset,
        -vp + offset,
        offset,
        offset,
        vp + offset,
        -vp + offset,
        vp + offset,
        -vp + offset,
    ]
    stop_voltages = [
        offset,
        offset,
        -vp + offset,
        offset,
        offset,
        vp + offset,
        -vp + offset,
        vp + offset,
        -vp + offset,
        offset,
    ]
    time_values = [
        rise_time,
        rise_time,
        rise_time,
        rise_time,
        2 * rise_time,
        rise_time,
        2 * rise_time,
        2 * rise_time,
        2 * rise_time,
        rise_time,
    ]
    meas_types = [2] * len(time_values)

    ch1_config = (1, start_voltages, stop_voltages, time_values, meas_types)
    ch2_config = (1, [0.0] * len(time_values), [0.0] * len(time_values), time_values, meas_types)
    return {CH1: [ch1_config], CH2: [ch2_config]}


def preview_waveform(output_path=None):
    """Preview the PV2 waveform without connecting to the PMU."""
    return preview_sequence_configs(make_pv2_seq_configs()[CH1], output_path, title_prefix="PV2 CH1")


def main():
    """Run the PV2 measurement and save raw/analysis files."""
    with PMUSession(INST, channels=(CH1, CH2)) as session:
        Q = session.query
        print("Running PV2...")
        seq_configs = make_pv2_seq_configs()
        current_ranges = {CH1: params["Irange1"], CH2: params["Irange2"]}
        execute_segARB_test(Q, [CH1, CH2], seq_configs, current_ranges=current_ranges)
        df_ch1, df_ch2 = read_both_channels(Q, CH1, CH2)
        power_off_outputs(Q, (CH1, CH2))
        if df_ch1 is None or df_ch1.empty:
            raise ValueError("PV2 returned no channel 1 data.")

        polarization = calculate_polarization(
            df_ch1[f"Current {CH1}"].values,
            df_ch1[f"Timestamp {CH1}"].values,
            params.get("area_cm2", 1.0),
        )
        df_pv2 = pd.DataFrame(
            {
                "Time": df_ch1[f"Timestamp {CH1}"].values,
                "Voltage": df_ch1[f"Voltage {CH1}"].values,
                "Current": df_ch1[f"Current {CH1}"].values,
                "Polarization": polarization,
            }
        )

        save_channels_separate_excel({1: df_ch1, 2: df_ch2}, f"{fname_base}_raw.xlsx")
        df_pv2.to_excel(f"{fname_base}_pv2.xlsx", index=False)

        fig, ax = plt.subplots(figsize=(6, 5))
        ax.plot(df_pv2["Voltage"], df_pv2["Polarization"], "b-")
        ax.set_xlabel("Voltage (V)")
        ax.set_ylabel("Polarization (uC/cm^2)")
        ax.set_title("PV2 Loop")
        ax.grid(alpha=0.3)
        fig.tight_layout()
        fig.savefig(f"{fname_base}_loop.png", dpi=300)
        plt.close(fig)
        print("PV2 complete.")


if __name__ == "__main__":
    if PREVIEW_ONLY:
        preview_waveform()
    else:
        main()
