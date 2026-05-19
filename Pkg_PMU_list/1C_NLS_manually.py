# -*- coding: utf-8 -*-
"""NLS switch segARB test with direct sequence definitions."""

from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

from debug.waveform_preview import preview_sequence_configs
from src.data_processing import calculate_polarization, read_both_channels
from src.pmu_tests import execute_segARB_test, power_off_outputs
from src.session import PMUSession

INST = "TCPIP0::129.125.87.80::1225::SOCKET"
CH1, CH2 = 1, 2
params = dict(
    offset=0,
    Vp=5,
    Rt_p=5e-5,
    Delaytime=100e-6,
    Vsquare=1.5,
    Rt_s=1e-7,
    Dwell=1e-4,
    Irange1=1e-4,
    Irange2=1e-4,
    MeasureSquare=False,
    area_cm2=7.0686e-6,
)

SAVE_DIR = Path(r"C:\Users\P317151\Documents\data\Jingtian\2025-12-14\BTO\Device2")
fname_prefix = (
    f"NISswitch_Meas{params['Vp']}V_{int(params['Rt_p'] * 1e6)}us_"
    f"with{params['Vsquare']}V_{int(params['Rt_s'] * 1e6)}us"
)
PREVIEW_ONLY = False


def make_nls_seq_configs():
    """Build NLS switch seq_configs directly in this script."""
    offset = params["offset"]
    vp = params["Vp"]
    rt_p = params["Rt_p"]
    delay_time = params["Delaytime"]
    vsquare = params["Vsquare"]
    rt_s = params["Rt_s"]
    dwell = params["Dwell"]
    measure_square = params.get("MeasureSquare", True)

    start_voltages = [
        0,
        0,
        offset,
        -vp + offset,
        offset,
        offset,
        vsquare + offset,
        vsquare + offset,
        offset,
        offset,
        vp + offset,
        offset,
        offset,
        vp + offset,
    ]
    stop_voltages = [
        0,
        offset,
        -vp + offset,
        offset,
        offset,
        vsquare + offset,
        vsquare + offset,
        offset,
        offset,
        vp + offset,
        offset,
        offset,
        vp + offset,
        offset,
    ]
    time_values = [
        rt_p,
        rt_p,
        rt_p,
        rt_p,
        delay_time,
        rt_s,
        dwell,
        rt_s,
        delay_time,
        rt_p,
        rt_p,
        delay_time,
        rt_p,
        rt_p,
    ]
    measure_square_type = 2 if measure_square else 0
    meas_types = [0, 0, 0, 0, 0, measure_square_type, measure_square_type, measure_square_type, 0, 2, 2, 0, 2, 2]

    ch1_config = (1, start_voltages, stop_voltages, time_values, meas_types)
    ch2_config = (1, [0.0] * len(time_values), [0.0] * len(time_values), time_values, meas_types)
    return {CH1: [ch1_config], CH2: [ch2_config]}


def preview_waveform(output_path=None):
    """Preview the NLS waveform without connecting to the PMU."""
    return preview_sequence_configs(make_nls_seq_configs()[CH1], output_path, title_prefix="NLS CH1")


def process_nls_channel(df, channel):
    """Calculate differential polarization for one NLS channel."""
    voltage = df[f"Voltage {channel}"].values
    current = df[f"Current {channel}"].values
    total_points = len(voltage)
    seg_pts = total_points // 4

    current_first = current[: 2 * seg_pts]
    current_second = current[2 * seg_pts : 4 * seg_pts]
    voltage_first = voltage[: 2 * seg_pts]

    min_len = min(len(current_first), len(current_second))
    diff_current = current_first[:min_len] - current_second[:min_len]
    diff_voltage = voltage_first[:min_len]
    time_step = params["Rt_p"] * 4 / total_points
    diff_time = np.arange(1, min_len + 1) * time_step
    polarization = calculate_polarization(diff_current, diff_time, params.get("area_cm2", 1.0))

    return pd.DataFrame(
        {
            "Time": diff_time,
            "Voltage": diff_voltage,
            "DiffCurrent": diff_current,
            "Polarization": polarization,
        }
    )


def save_nls_results(df_ch1, df_ch2):
    """Save raw NLS data and processed plots."""
    SAVE_DIR.mkdir(parents=True, exist_ok=True)
    fname_base = SAVE_DIR / fname_prefix
    excel_path = f"{fname_base}.xlsx"

    if params["MeasureSquare"]:
        with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
            df_ch1.to_excel(writer, sheet_name="Raw_CH1", index=False)
            df_ch2.to_excel(writer, sheet_name="Raw_CH2", index=False)

        fig, axes = plt.subplots(2, 1, figsize=(12, 8), sharex=True)
        for axis, df, channel, title in (
            (axes[0], df_ch1, CH1, "CH1 Waveform"),
            (axes[1], df_ch2, CH2, "CH2 Waveform"),
        ):
            time = df[f"Timestamp {channel}"]
            voltage = df[f"Voltage {channel}"]
            current = df[f"Current {channel}"]
            axis_current = axis.twinx()
            axis.plot(time, voltage, "b-", linewidth=0.8)
            axis_current.plot(time, current * 1e6, "r-", linewidth=0.8)
            axis.set_ylabel("Voltage (V)", color="b")
            axis_current.set_ylabel("Current (uA)", color="r")
            axis.set_title(title)
            axis.grid(alpha=0.3)
        axes[1].set_xlabel("Time (s)")
        fig.suptitle("NLS Switch - Waveform Check", fontsize=14)
        fig.tight_layout()
        fig.savefig(f"{fname_base}_waveform.png", dpi=300)
        plt.close(fig)
        return

    df_vp_ch1 = process_nls_channel(df_ch1, CH1)
    df_vp_ch2 = process_nls_channel(df_ch2, CH2)

    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        df_ch1.to_excel(writer, sheet_name="Raw_CH1", index=False)
        df_ch2.to_excel(writer, sheet_name="Raw_CH2", index=False)
        df_vp_ch1.to_excel(writer, sheet_name="VP_CH1", index=False)
        df_vp_ch2.to_excel(writer, sheet_name="VP_CH2", index=False)

    fig, axes = plt.subplots(1, 2, figsize=(12, 5))
    axes[0].plot(df_vp_ch1["Voltage"], df_vp_ch1["Polarization"], "b-", linewidth=1)
    axes[0].set_xlabel("Voltage (V)")
    axes[0].set_ylabel("Polarization (uC/cm^2)")
    axes[0].set_title("CH1 V-P (Tri1 - Tri2)")
    axes[0].grid(alpha=0.3)

    axes[1].plot(df_vp_ch2["Voltage"], df_vp_ch2["Polarization"], "r-", linewidth=1)
    axes[1].set_xlabel("Voltage (V)")
    axes[1].set_ylabel("Polarization (uC/cm^2)")
    axes[1].set_title("CH2 V-P (Tri1 - Tri2)")
    axes[1].grid(alpha=0.3)

    fig.suptitle("NLS Switch - Differential Polarization", fontsize=14)
    fig.tight_layout()
    fig.savefig(f"{fname_base}_vp.png", dpi=300)
    plt.close(fig)


def main():
    """Run the NLS switch measurement."""
    with PMUSession(INST, channels=(CH1, CH2)) as session:
        query = session.query
        print(
            "Running NLS switch "
            f"(Vp={params['Vp']}V, Vsquare={params['Vsquare']:.2f}V, Dwell={params['Dwell']:.1e}s)..."
        )
        seq_configs = make_nls_seq_configs()
        current_ranges = {CH1: params["Irange1"], CH2: params["Irange2"]}
        execute_segARB_test(query, [CH1, CH2], seq_configs, current_ranges=current_ranges)
        df_ch1, df_ch2 = read_both_channels(query, CH1, CH2)
        power_off_outputs(query, (CH1, CH2))
        if df_ch1 is None or df_ch2 is None or df_ch1.empty or df_ch2.empty:
            raise ValueError("NLS switch returned empty channel data.")
        save_nls_results(df_ch1, df_ch2)


if __name__ == "__main__":
    if PREVIEW_ONLY:
        preview_waveform()
    else:
        main()

