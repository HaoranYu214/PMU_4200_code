# -*- coding: utf-8 -*-
"""NLS switch test helpers and standalone entrypoint."""

from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

from src.data_processing import calculate_polarization, read_both_channels
from src.pmu_tests import hy_NISswitch_segARB, power_off_outputs
from src.session import PMUSession


def run_nls_switch_test(Q, ch1, ch2, params, save_dir, fname_prefix=None):
    """Run one NLS switch measurement and save raw and processed outputs."""
    save_dir = Path(save_dir)
    save_dir.mkdir(parents=True, exist_ok=True)

    if fname_prefix is None:
        fname_prefix = (
            f"NLSswitch_Meas{params['Vp']}V_{int(params['Rt_p'] * 1e6)}us_"
            f"with{params['Vsquare']:.2f}V_Dwell{params['Dwell']:.1e}s"
        )
    fname_base = save_dir / fname_prefix

    print(
        "Running NLS switch "
        f"(Vp={params['Vp']}V, Vsquare={params['Vsquare']:.2f}V, Dwell={params['Dwell']:.1e}s)..."
    )
    hy_NISswitch_segARB(Q, ch1, ch2, params)
    df_ch1, df_ch2 = read_both_channels(Q, ch1, ch2)
    power_off_outputs(Q, (ch1, ch2))
    if df_ch1 is None or df_ch2 is None or df_ch1.empty or df_ch2.empty:
        raise ValueError("NLS switch returned empty channel data.")

    result = {"df_ch1": df_ch1, "df_ch2": df_ch2, "success": True}
    excel_path = f"{fname_base}.xlsx"

    if params["MeasureSquare"]:
        with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
            df_ch1.to_excel(writer, sheet_name="Raw_CH1", index=False)
            df_ch2.to_excel(writer, sheet_name="Raw_CH2", index=False)

        fig, axes = plt.subplots(2, 1, figsize=(12, 8), sharex=True)
        for axis, df, channel, title in (
            (axes[0], df_ch1, ch1, "CH1 Waveform"),
            (axes[1], df_ch2, ch2, "CH2 Waveform"),
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
        return result

    area_cm2 = params.get("area_cm2", 1.0)

    def process_channel(df, channel):
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
        polarization = calculate_polarization(diff_current, diff_time, area_cm2)
        return pd.DataFrame(
            {
                "Time": diff_time,
                "Voltage": diff_voltage,
                "DiffCurrent": diff_current,
                "Polarization": polarization,
            }
        )

    df_vp_ch1 = process_channel(df_ch1, ch1)
    df_vp_ch2 = process_channel(df_ch2, ch2)
    result["df_vp_ch1"] = df_vp_ch1
    result["df_vp_ch2"] = df_vp_ch2

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
    return result


if __name__ == "__main__":
    INST = "TCPIP0::129.125.87.80::1225::SOCKET"
    CH1, CH2 = 1, 2
    params = dict(
        offset=0,
        Vp=2,
        Rt_p=5e-5,
        Delaytime=100e-6,
        Vsquare=1,
        Rt_s=1e-7,
        Dwell=1e-6,
        Irange1=1e-4,
        Irange2=1e-4,
        MeasureSquare=False,
        area_cm2=7.0686e-6,
    )
    SAVE_DIR = Path(r"C:\Users\P317151\Documents\data\Jingtian\2025-12-14\BTO\Device3")

    with PMUSession(INST, channels=(CH1, CH2)) as session:
        run_nls_switch_test(session.query, CH1, CH2, params, SAVE_DIR)
