# -*- coding: utf-8 -*-
"""PUND segARB test with direct sequence definitions."""

from pathlib import Path
import sys
from datetime import datetime

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

PKG_ROOT = Path(__file__).resolve().parents[1]
if str(PKG_ROOT) not in sys.path:
    sys.path.insert(0, str(PKG_ROOT))

from debug.waveform_preview import preview_sequence_configs
from src.data_processing import calculate_polarization, read_both_channels, save_channels_separate_excel
from src.pmu_tests import execute_segARB_test, power_off_outputs
from src.session import PMUSession

INST = "TCPIP0::129.125.87.80::1225::SOCKET"
CH1, CH2 = 1, 2
params = dict(
    rise_time=1e-4,
    dwell_time=1e-5,
    delay_time=1e-5,
    Vp=5,
    offset=0,
    # area_cm2=1.2567e-5,
    area_cm2=7.854e-5,
    Irange1=1e-4,
    Irange2=1e-4,
)

SAVE_DIR = Path(r"C:\Users\P317151\Documents\data\19-05-2026\03A6\50um-2\PUND")
SAVE_DIR.mkdir(parents=True, exist_ok=True)
PREVIEW_ONLY = False


def build_fname_base():
    """Build a descriptive output stem for the current PUND parameters."""
    rise_us = int(round(params["rise_time"] * 1e6))
    dwell_us = int(round(params["dwell_time"] * 1e6))
    delay_us = int(round(params["delay_time"] * 1e6))
    vp_str = f"{params['Vp']:g}".replace("-", "m").replace(".", "p")
    offset_str = f"{params['offset']:g}".replace("-", "m").replace(".", "p")
    ir1_str = f"{params['Irange1']:.0e}".replace("-", "m")
    ir2_str = f"{params['Irange2']:.0e}".replace("-", "m")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    name = (
        f"PUND_Vp{vp_str}V_off{offset_str}V_"
        f"rise{rise_us}us_dwell{dwell_us}us_delay{delay_us}us_"
        f"I1{ir1_str}_I2{ir2_str}_{timestamp}"
    )
    return SAVE_DIR / name


def make_pund_seq_configs():
    """Build PUND seq_configs directly in this script."""
    rise_time = params["rise_time"]
    dwell_time = params["dwell_time"]
    delay_time = params["delay_time"]
    vp = params["Vp"]
    offset = params["offset"]

    start_voltages = [
        0,
        0,
        offset,
        -vp + offset,
        -vp + offset,
        offset,
        offset,
        vp + offset,
        vp + offset,
        offset,
        offset,
        vp + offset,
        vp + offset,
        offset,
        offset,
        -vp + offset,
        -vp + offset,
        offset,
        offset,
        -vp + offset,
        -vp + offset,
        offset,
    ]
    stop_voltages = [
        0,
        offset,
        -vp + offset,
        -vp + offset,
        offset,
        offset,
        vp + offset,
        vp + offset,
        offset,
        offset,
        vp + offset,
        vp + offset,
        offset,
        offset,
        -vp + offset,
        -vp + offset,
        offset,
        offset,
        -vp + offset,
        -vp + offset,
        offset,
        offset,
    ]
    time_values = [
        delay_time,
        delay_time,
        rise_time,
        dwell_time,
        rise_time,
        delay_time,
        rise_time,
        dwell_time,
        rise_time,
        delay_time,
        rise_time,
        dwell_time,
        rise_time,
        delay_time,
        rise_time,
        dwell_time,
        rise_time,
        delay_time,
        rise_time,
        dwell_time,
        rise_time,
        delay_time,
    ]
    # Measure only the pulse edges. Zero-delay and dwell segments stay in the
    # waveform but do not contribute sampled points to the PUND integration.
    meas_types = [0, 0, 2, 0, 2, 0, 2, 0, 2, 0, 2, 0, 2, 0, 2, 0, 2, 0, 2, 0, 2, 0]
    # meas_types = [2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2]

    ch1_config = (1, start_voltages, stop_voltages, time_values, meas_types)
    ch2_config = (1, [0.0] * len(time_values), [0.0] * len(time_values), time_values, meas_types)
    return {CH1: [ch1_config], CH2: [ch2_config]}


def preview_waveform(output_path=None):
    """Preview the PUND waveform without connecting to the PMU."""
    return preview_sequence_configs(make_pund_seq_configs()[CH1], output_path, title_prefix="PUND CH1")


def analyze_pund_edge_diff(df_ch1, df_ch2):
    """Analyze PUND data for arbitrary measured segment selections."""
    if df_ch1 is None or df_ch2 is None or df_ch1.empty or df_ch2.empty:
        raise ValueError("PUND returned empty channel data.")

    ch1_config = make_pund_seq_configs()[CH1][0]
    time_values = ch1_config[3]
    meas_types = ch1_config[4]
    meas_start = ch1_config[5] if len(ch1_config) > 5 else [0.0] * len(time_values)
    meas_stop = ch1_config[6] if len(ch1_config) > 6 else list(time_values)

    v_total = df_ch1[f"Voltage {CH1}"].values - df_ch2[f"Voltage {CH2}"].values
    i1_total = df_ch1[f"Current {CH1}"].values
    i_total = -df_ch2[f"Current {CH2}"].values
    t_total = df_ch1[f"Timestamp {CH1}"].values
    df_total = pd.DataFrame(
        {
            "Time": t_total,
            "Voltage": v_total,
            "CurrentI1": i1_total,
            "CurrentI2": i_total,
        }
    )

    measured_segments = [index for index, mode in enumerate(meas_types) if mode != 0]
    measured_durations = [max(0.0, meas_stop[index] - meas_start[index]) for index in measured_segments]
    total_duration = sum(measured_durations)
    if not measured_segments or total_duration <= 0:
        raise ValueError("PUND has no valid measured segments in the current seq config.")

    total_points = len(i_total)
    exact_counts = [total_points * duration / total_duration for duration in measured_durations]
    base_counts = [int(np.floor(value)) for value in exact_counts]
    remainder = total_points - sum(base_counts)
    order = np.argsort([value - base for value, base in zip(exact_counts, base_counts)])[::-1]
    for pick in order[:remainder]:
        base_counts[int(pick)] += 1

    segment_slices = {}
    cursor = 0
    for segment_index, point_count in zip(measured_segments, base_counts):
        next_cursor = cursor + point_count
        segment_slices[segment_index] = slice(cursor, next_cursor)
        cursor = next_cursor

    pulse_segments = {
        "Preset": [2, 3, 4],
        "P": [6, 7, 8],
        "U": [10, 11, 12],
        "N": [14, 15, 16],
        "D": [18, 19, 20],
    }

    pulses = {}
    for label, segments in pulse_segments.items():
        edge_chunks = []
        for segment_index in segments:
            segment_slice = segment_slices.get(segment_index)
            if segment_slice is None:
                continue
            edge_chunks.append(
                {
                    "voltage": v_total[segment_slice],
                    "current_i1": i1_total[segment_slice],
                    "current": i_total[segment_slice],
                }
            )
        if not edge_chunks:
            continue
        pulses[label] = {
            "voltage": np.concatenate([chunk["voltage"] for chunk in edge_chunks]),
            "current_i1": np.concatenate([chunk["current_i1"] for chunk in edge_chunks]),
            "current": np.concatenate([chunk["current"] for chunk in edge_chunks]),
        }

    positive_dt = np.diff(t_total)
    positive_dt = positive_dt[positive_dt > 0]
    sample_dt = float(np.min(positive_dt)) if len(positive_dt) else 0.0

    pair_defs = [("P", "U", "P-U"), ("N", "D", "N-D")]
    frames = []
    for first, second, label in pair_defs:
        min_len = min(len(pulses[first]["current"]), len(pulses[second]["current"]))
        diff_current = pulses[first]["current"][:min_len] - pulses[second]["current"][:min_len]
        diff_current_i1 = pulses[first]["current_i1"][:min_len] - pulses[second]["current_i1"][:min_len]
        voltage = pulses[first]["voltage"][:min_len]
        local_time = np.arange(min_len) * sample_dt
        polarization = calculate_polarization(diff_current, local_time, params.get("area_cm2", 1.0))
        polarization_i1 = calculate_polarization(diff_current_i1, local_time, params.get("area_cm2", 1.0))
        frames.append(
            pd.DataFrame(
                {
                    "Time": local_time,
                    "Voltage": voltage,
                    "DiffCurrent": diff_current,
                    "DiffCurrentI1": diff_current_i1,
                    "Polarization": polarization,
                    "PolarizationI1": polarization_i1,
                    "Segment": label,
                }
            )
        )

    if not frames:
        raise ValueError("PUND edge differential analysis failed.")

    return {
        "df_total": df_total,
        "pund_diff": pd.concat(frames, ignore_index=True),
        "meta": {
            "measured_segments": measured_segments,
            "segment_point_counts": {segment: base_counts[idx] for idx, segment in enumerate(measured_segments)},
            "measured_pulses": list(pulses.keys()),
            "pairs": pair_defs,
        },
    }


def main():
    """Run the PUND measurement and save raw/analysis files."""
    fname_base = build_fname_base()
    with PMUSession(INST, channels=(CH1, CH2)) as session:
        Q = session.query
        print("Running PUND...")
        seq_configs = make_pund_seq_configs()
        current_ranges = {CH1: params["Irange1"], CH2: params["Irange2"]}
        execute_segARB_test(Q, [CH1, CH2], seq_configs, current_ranges=current_ranges)
        df_ch1, df_ch2 = read_both_channels(Q, CH1, CH2)
        power_off_outputs(Q, (CH1, CH2))
        if df_ch1 is None or df_ch2 is None or df_ch1.empty or df_ch2.empty:
            raise ValueError("PUND returned empty channel data.")

        data = analyze_pund_edge_diff(df_ch1, df_ch2)
        save_channels_separate_excel({1: df_ch1, 2: df_ch2}, f"{fname_base}_raw.xlsx")
        data["df_total"].to_excel(f"{fname_base}_total.xlsx", index=False)
        data["pund_diff"].to_excel(f"{fname_base}_diff.xlsx", index=False)

        fig_i2, ax_i2 = plt.subplots(figsize=(7, 5))
        for seg in ["P-U", "N-D"]:
            sub = data["pund_diff"][data["pund_diff"]["Segment"] == seg]
            ax_i2.plot(sub["Voltage"], sub["Polarization"], ".", label=seg, markersize=4)
        ax_i2.set_xlabel("Voltage (V)")
        ax_i2.set_ylabel("Polarization (uC/cm^2)")
        ax_i2.set_title("PUND Polarization from I2 Difference")
        ax_i2.legend()
        ax_i2.grid(alpha=0.3)
        fig_i2.tight_layout()
        fig_i2.savefig(f"{fname_base}_loop_i2diff.png", dpi=300)
        plt.close(fig_i2)

        fig_i1, ax_i1 = plt.subplots(figsize=(7, 5))
        for seg in ["P-U", "N-D"]:
            sub = data["pund_diff"][data["pund_diff"]["Segment"] == seg]
            ax_i1.plot(sub["Voltage"], sub["PolarizationI1"], ".", label=seg, markersize=4)
        ax_i1.set_xlabel("Voltage (V)")
        ax_i1.set_ylabel("Polarization (uC/cm^2)")
        ax_i1.set_title("PUND Polarization from I1 Difference")
        ax_i1.legend()
        ax_i1.grid(alpha=0.3)
        fig_i1.tight_layout()
        fig_i1.savefig(f"{fname_base}_loop_i1.png", dpi=300)
        plt.close(fig_i1)

        fig_iv, ax_iv = plt.subplots(figsize=(7, 5))
        for seg in ["P-U", "N-D"]:
            sub = data["pund_diff"][data["pund_diff"]["Segment"] == seg]
            ax_iv.plot(sub["Voltage"], sub["DiffCurrent"], ".", label=f"{seg} I2", markersize=4)
            ax_iv.plot(sub["Voltage"], sub["DiffCurrentI1"], ".", label=f"{seg} I1", markersize=3, alpha=0.7)
        ax_iv.set_xlabel("Voltage (V)")
        ax_iv.set_ylabel("Differential Current (A)")
        ax_iv.set_title("PUND Differential I-V")
        ax_iv.legend()
        ax_iv.grid(alpha=0.3)
        fig_iv.tight_layout()
        fig_iv.savefig(f"{fname_base}_diff_iv.png", dpi=300)
        plt.close(fig_iv)
        print("PUND complete.")


if __name__ == "__main__":
    if PREVIEW_ONLY:
        preview_waveform()
    else:
        main()
