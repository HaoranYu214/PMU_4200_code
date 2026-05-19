# -*- coding: utf-8 -*-
"""PUND segARB test with direct sequence definitions."""

from pathlib import Path
import sys

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
    rise_time=5e-5,
    dwell_time=5e-5,
    delay_time=5e-5,
    Vp=1,
    offset=0,
    area_cm2=1.2567e-5,
    Irange1=1e-4,
    Irange2=1e-4,
)

SAVE_DIR = Path(r"C:\Users\P317151\Documents\data\19-05-2026\Test")
SAVE_DIR.mkdir(parents=True, exist_ok=True)
fname_base = SAVE_DIR / f"PUND_{int(params['rise_time'] * 1e6)}us_{params['Vp']}V"
PREVIEW_ONLY = False


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

    ch1_config = (1, start_voltages, stop_voltages, time_values, meas_types)
    ch2_config = (1, [0.0] * len(time_values), [0.0] * len(time_values), time_values, meas_types)
    return {CH1: [ch1_config], CH2: [ch2_config]}


def preview_waveform(output_path=None):
    """Preview the PUND waveform without connecting to the PMU."""
    return preview_sequence_configs(make_pund_seq_configs()[CH1], output_path, title_prefix="PUND CH1")


def analyze_pund_edge_diff(df_ch1, df_ch2):
    """Analyze PUND data when only pulse edge segments are measured."""
    if df_ch1 is None or df_ch2 is None or df_ch1.empty or df_ch2.empty:
        raise ValueError("PUND returned empty channel data.")

    v_total = df_ch1[f"Voltage {CH1}"].values - df_ch2[f"Voltage {CH2}"].values
    i_total = -df_ch2[f"Current {CH2}"].values
    t_total = df_ch1[f"Timestamp {CH1}"].values
    df_total = pd.DataFrame({"Time": t_total, "Voltage": v_total, "Current": i_total})

    pulse_labels = ["Preset", "P", "U", "N", "D"]
    edges_per_pulse = 2
    measured_edges = len(pulse_labels) * edges_per_pulse
    points_per_edge = len(i_total) // measured_edges
    if points_per_edge == 0:
        raise ValueError("PUND data has too few sampled points for edge-only analysis.")

    pulses = {}
    for pulse_index, label in enumerate(pulse_labels):
        edge_chunks = []
        for edge_index in range(edges_per_pulse):
            measured_index = pulse_index * edges_per_pulse + edge_index
            start = measured_index * points_per_edge
            stop = start + points_per_edge
            edge_chunks.append(
                {
                    "voltage": v_total[start:stop],
                    "current": i_total[start:stop],
                }
            )
        pulses[label] = {
            "voltage": np.concatenate([chunk["voltage"] for chunk in edge_chunks]),
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
        voltage = pulses[first]["voltage"][:min_len]
        local_time = np.arange(min_len) * sample_dt
        polarization = calculate_polarization(diff_current, local_time, params.get("area_cm2", 1.0))
        frames.append(
            pd.DataFrame(
                {
                    "Time": local_time,
                    "Voltage": voltage,
                    "DiffCurrent": diff_current,
                    "Polarization": polarization,
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
            "points_per_edge": points_per_edge,
            "edges_per_pulse": edges_per_pulse,
            "measured_pulses": pulse_labels,
            "pairs": pair_defs,
        },
    }


def main():
    """Run the PUND measurement and save raw/analysis files."""
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

        fig, ax = plt.subplots(figsize=(7, 5))
        for seg in ["P-U", "N-D"]:
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


if __name__ == "__main__":
    if PREVIEW_ONLY:
        preview_waveform()
    else:
        main()
