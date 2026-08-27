# -*- coding: utf-8 -*-
"""PV2 segARB test with direct sequence definitions."""

from pathlib import Path
import sys
from datetime import datetime

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

PKG_ROOT = Path(__file__).resolve().parents[1]
REPO_ROOT = PKG_ROOT.parent
for path in (PKG_ROOT, REPO_ROOT):
    if str(path) not in sys.path:
        sys.path.insert(0, str(path))

from debug.waveform_preview import preview_sequence_configs
from src.pmu.current_range import acquire_with_auto_current_range
from src.pmu.data_processing import read_both_channels
from src.pmu.pmu_tests import execute_segARB_test, power_off_outputs
from src.pmu.session import PMUSession

INST = "TCPIP0::129.125.87.80::1225::SOCKET"
CH1, CH2 = 1, 2
params = dict(
    rise_time=2.5e-5,
    delay_time=2.5e-5,
    Vp=4.5,
    offset=0,
    # area_cm2=(20*1e-4)**2*3.14,
    # area_cm2=4e-6,
    area_cm2=(20*1e-4)**2,
    Irange1=1e-4,
    Irange2=1e-4,
)
SEGARB_OPTIONS = {
    "ENABLE_CONNECTION_COMP": False,
    "ENABLE_LOAD_CONFIG": False,
    "LOAD_RESISTANCE": 1e3,
    "ENABLE_LLEC": False,
}
PREVIEW_ONLY = False
SAVE_DIR = Path(r"C:\Users\P317151\Documents\data\25-08-2026\03B4_hZO_2700_800\L20_4\after IV")
SAVE_DIR.mkdir(parents=True, exist_ok=True)

def build_fname_base():
    """Build a unique output path for the current PV2 parameters."""
    rise_us = int(round(params["rise_time"] * 1e6))
    delay_us = int(round(params["delay_time"] * 1e6))
    vp_str = f"{params['Vp']:g}".replace("-", "m").replace(".", "p")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    return SAVE_DIR / f"PV2_{rise_us}us_delay{delay_us}us_{vp_str}V_{timestamp}"


def make_pv2_seq_configs():
    """Build PV2 seq_configs directly in this script."""
    rise_time = params["rise_time"]
    delay_time = params["delay_time"]
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
        delay_time,
        rise_time,
        2 * rise_time,
        2 * rise_time,
        2 * rise_time,
        rise_time,
    ]
    # Segments 0-3 are the pre/post triangle and segment 4 is the offset hold.
    # Only the main PV2 sweep in segments 5-9 is measured and integrated.
    meas_types = [0, 0, 0, 0, 0, 2, 2, 2, 2, 2]

    ch1_config = (1, start_voltages, stop_voltages, time_values, meas_types)
    ch2_config = (1, [0.0] * len(time_values), [0.0] * len(time_values), time_values, meas_types)
    return {CH1: [ch1_config], CH2: [ch2_config]}


def preview_waveform(output_path=None):
    """Preview the PV2 waveform without connecting to the PMU."""
    return preview_sequence_configs(make_pv2_seq_configs()[CH1], output_path, title_prefix="PV2 CH1")


def build_params_table():
    """Return the PV2 run parameters as a two-column table."""
    rows = [{"name": name, "value": repr(value)} for name, value in params.items()]
    rows.extend(
        {"name": name, "value": repr(value)}
        for name, value in SEGARB_OPTIONS.items()
    )
    return pd.DataFrame(rows)


def acquire_with_auto_range(query):
    """Repeat PV2 acquisition until both channel ranges are suitable."""
    def acquire_once(ranges):
        current_ranges = {CH1: ranges["Irange1"], CH2: ranges["Irange2"]}
        execute_segARB_test(
            query,
            [CH1, CH2],
            make_pv2_seq_configs(),
            current_ranges=current_ranges,
            options=SEGARB_OPTIONS,
        )
        df_ch1, df_ch2 = read_both_channels(query, CH1, CH2)
        power_off_outputs(query, (CH1, CH2))
        if df_ch1 is None or df_ch2 is None or df_ch1.empty or df_ch2.empty:
            raise ValueError("PV2 returned empty channel data during range check.")
        return df_ch1, df_ch2

    result, final_ranges, _assessments = acquire_with_auto_current_range(
        acquire_once,
        {
            "Irange1": params["Irange1"],
            "Irange2": params["Irange2"],
        },
        {
            "Irange1": lambda data: data[0][f"Current {CH1}"].to_numpy(),
            "Irange2": lambda data: data[1][f"Current {CH2}"].to_numpy(),
        },
        labels={"Irange1": "I1", "Irange2": "I2"},
        test_name="PV2",
    )
    params.update(final_ranges)
    return result


def _integrate_loop_from_zero(time, voltage, current, area_cm2):
    """Integrate from Q=0, then shift polarization so the two Pr values are symmetric."""
    if len(current) < 3:
        raise ValueError("PV2 loop has too few points for integration.")
    if area_cm2 <= 0:
        raise ValueError("area_cm2 must be positive.")

    raw_time = np.asarray(time, dtype=float)
    voltage = np.asarray(voltage, dtype=float)
    current = np.asarray(current, dtype=float)
    loop_duration = 4.0 * params["rise_time"]
    # PMU timestamps are quantized to 100 ns, while waveform samples can be much
    # denser (20 ns in the current setup). Use the programmed loop duration so
    # every measured point receives its actual uniform integration interval.
    local_time = np.linspace(0.0, loop_duration, len(current))
    charge = np.zeros(len(current), dtype=float)
    dt = np.diff(local_time)
    charge[1:] = np.cumsum(0.5 * (current[:-1] + current[1:]) * dt)
    polarization = charge / area_cm2 * 1e6

    positive_peak = int(np.argmax(voltage))
    negative_peak = positive_peak + int(np.argmin(voltage[positive_peak:]))
    if negative_peak <= positive_peak:
        raise ValueError("PV2 loop does not contain positive and negative peaks.")

    offset = params["offset"]
    positive_pr_index = positive_peak + int(
        np.argmin(np.abs(voltage[positive_peak : negative_peak + 1] - offset))
    )
    negative_pr_index = negative_peak + int(
        np.argmin(np.abs(voltage[negative_peak:] - offset))
    )
    pr_center = 0.5 * (
        polarization[positive_pr_index] + polarization[negative_pr_index]
    )
    polarization_centered = polarization - pr_center

    return pd.DataFrame(
        {
            "Time": local_time,
            "RawTime": raw_time,
            "Voltage": voltage,
            "Current": current,
            "Polarization": polarization_centered,
        }
    )


def _split_complete_loops(time, voltage, current):
    """Split measured points equally into delayed and non-delayed loops."""
    point_count = len(voltage)
    split_index = point_count // 2
    if split_index < 3 or point_count - split_index < 3:
        raise ValueError("PV2 returned too few points to split into two loops.")

    return (
        (time[:split_index], voltage[:split_index], current[:split_index]),
        (time[split_index:], voltage[split_index:], current[split_index:]),
    )


def _build_loop_sheet(delay_loop, no_delay_loop):
    """Build one four-column loop table, padding unequal point counts with NaN."""
    return pd.DataFrame(
        {
            "Voltage_Delay": pd.Series(delay_loop["Voltage"].to_numpy()),
            "Polarization_Delay": pd.Series(delay_loop["Polarization"].to_numpy()),
            "Voltage_NoDelay": pd.Series(no_delay_loop["Voltage"].to_numpy()),
            "Polarization_NoDelay": pd.Series(no_delay_loop["Polarization"].to_numpy()),
        }
    )


def analyze_pv2(df_ch1, df_ch2):
    """Split points in half and integrate each PV2 loop independently from zero."""
    if df_ch1 is None or df_ch2 is None or df_ch1.empty or df_ch2.empty:
        raise ValueError("PV2 returned empty channel data.")

    point_count = min(len(df_ch1), len(df_ch2))
    time = df_ch1[f"Timestamp {CH1}"].values[:point_count]
    voltage = (
        df_ch1[f"Voltage {CH1}"].values[:point_count]
        - df_ch2[f"Voltage {CH2}"].values[:point_count]
    )
    current_i1 = df_ch1[f"Current {CH1}"].values[:point_count]
    current_i2 = -df_ch2[f"Current {CH2}"].values[:point_count]
    area_cm2 = params.get("area_cm2", 1.0)

    df_total = pd.DataFrame(
        {
            "Time": time,
            "Voltage": voltage,
            "CurrentI1": current_i1,
            "CurrentI2": current_i2,
        }
    )

    i1_parts = _split_complete_loops(time, voltage, current_i1)
    i2_parts = _split_complete_loops(time, voltage, current_i2)
    i1_delay = _integrate_loop_from_zero(*i1_parts[0], area_cm2)
    i1_no_delay = _integrate_loop_from_zero(*i1_parts[1], area_cm2)
    i2_delay = _integrate_loop_from_zero(*i2_parts[0], area_cm2)
    i2_no_delay = _integrate_loop_from_zero(*i2_parts[1], area_cm2)

    return {
        "df_total": df_total,
        "i1_delay": i1_delay,
        "i1_no_delay": i1_no_delay,
        "i2_delay": i2_delay,
        "i2_no_delay": i2_no_delay,
        "i1_loops": _build_loop_sheet(i1_delay, i1_no_delay),
        "i2_loops": _build_loop_sheet(i2_delay, i2_no_delay),
    }


def save_pv2_workbook(output_path, df_ch1, df_ch2, data):
    """Save raw channels, processed PV2 data, and parameters in one workbook."""
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df_ch1.to_excel(writer, sheet_name="Channel_1", index=False)
        df_ch2.to_excel(writer, sheet_name="Channel_2", index=False)
        data["df_total"].to_excel(writer, sheet_name="Total", index=False)
        data["i1_loops"].to_excel(writer, sheet_name="I1_Loops", index=False)
        data["i2_loops"].to_excel(writer, sheet_name="I2_Loops", index=False)
        build_params_table().to_excel(writer, sheet_name="Parameters", index=False)
    return output_path


def main():
    """Run PV2 and save one workbook plus the loop figure."""
    fname_base = build_fname_base()
    with PMUSession(INST, channels=(CH1, CH2)) as session:
        Q = session.query
        print("Running PV2...")
        df_ch1, df_ch2 = acquire_with_auto_range(Q)
        data = analyze_pv2(df_ch1, df_ch2)
        workbook_path = save_pv2_workbook(
            f"{fname_base}.xlsx",
            df_ch1,
            df_ch2,
            data,
        )

        fig_i2, ax_i2 = plt.subplots(figsize=(6, 5))
        ax_i2.plot(
            data["i2_delay"]["Voltage"],
            data["i2_delay"]["Polarization"],
            "b-",
            label=f"Delay {params['delay_time'] * 1e3:g} ms",
        )
        ax_i2.plot(
            data["i2_no_delay"]["Voltage"],
            data["i2_no_delay"]["Polarization"],
            "c-",
            label="No delay",
        )
        ax_i2.set_xlabel("Voltage (V)")
        ax_i2.set_ylabel("Polarization (uC/cm^2)")
        ax_i2.set_title("PV2 Loop from I2")
        ax_i2.legend()
        ax_i2.grid(alpha=0.3)
        fig_i2.tight_layout()
        i2_plot_path = Path(f"{fname_base}_loop_i2.png")
        fig_i2.savefig(i2_plot_path, dpi=300)
        plt.close(fig_i2)

        fig_i1, ax_i1 = plt.subplots(figsize=(6, 5))
        ax_i1.plot(
            data["i1_delay"]["Voltage"],
            data["i1_delay"]["Polarization"],
            "r-",
            label=f"Delay {params['delay_time'] * 1e3:g} ms",
        )
        ax_i1.plot(
            data["i1_no_delay"]["Voltage"],
            data["i1_no_delay"]["Polarization"],
            "m-",
            label="No delay",
        )
        ax_i1.set_xlabel("Voltage (V)")
        ax_i1.set_ylabel("Polarization (uC/cm^2)")
        ax_i1.set_title("PV2 Loop from I1")
        ax_i1.legend()
        ax_i1.grid(alpha=0.3)
        fig_i1.tight_layout()
        i1_plot_path = Path(f"{fname_base}_loop_i1.png")
        fig_i1.savefig(i1_plot_path, dpi=300)
        plt.close(fig_i1)

        print(f"Saved PV2 workbook: {workbook_path.resolve()}")
        print(f"Saved PV2 I2 loop: {i2_plot_path.resolve()}")
        print(f"Saved PV2 I1 loop: {i1_plot_path.resolve()}")
        print("PV2 complete.")


if __name__ == "__main__":
    if PREVIEW_ONLY:
        preview_waveform()
    else:
        main()
