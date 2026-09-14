# -*- coding: utf-8 -*-
"""Endurance test: cycle, PV2 readback, then PUND readback."""

from pathlib import Path
import sys

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

REPO_ROOT = next(
    parent for parent in Path(__file__).resolve().parents
    if (parent / "pyproject.toml").is_file()
)
SRC_ROOT = REPO_ROOT / "src"
for path in (SRC_ROOT, REPO_ROOT):
    if str(path) not in sys.path:
        sys.path.insert(0, str(path))

from keithley4200.tools.waveform_preview import preview_sequence_configs
from keithley4200.output import prepare_output_dir, reserve_output_stem, measurement_name, time_tag, saved_at
from keithley4200.pmu.data_processing import analyze_pund_diff, read_both_channels
from keithley4200.pmu.pmu_tests import execute_segARB_test, power_off_outputs
from keithley4200.pmu.session import PMUSession

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
    delay_time=100e-6,
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
SEGARB_OPTIONS = {
    "ENABLE_CONNECTION_COMP": False,
    "ENABLE_LOAD_CONFIG": False,
    "LOAD_RESISTANCE": 1e6,
    "ENABLE_LLEC": False,
}

# These are cumulative readback milestones, not per-step cycle increments.
cycle_counts = [1, 10, 100, 1000, 10000, 1e5, 1e6, 1e7]
SAVE_DIR = Path(r"C:\Users\P317151\Documents\data\FTJ\Refined\Endurance")


def build_params_table(readback_name, readback_params, completed_cycles):
    """Return cycle, readback, and common PMU settings for one saved result."""
    rows = [
        {"section": "endurance", "name": "completed_cycles", "value": repr(int(completed_cycles))}
    ]
    rows.extend(
        {"section": "cycle", "name": name, "value": repr(value)}
        for name, value in params_cycle.items()
    )
    rows.extend(
        {"section": readback_name, "name": name, "value": repr(value)}
        for name, value in readback_params.items()
    )
    rows.extend(
        {"section": "SEGARB_OPTIONS", "name": name, "value": repr(value)}
        for name, value in SEGARB_OPTIONS.items()
    )
    rows.append({"name": "saved_at", "value": saved_at()})
    return pd.DataFrame(rows)


def save_channels_with_params(dfs, path, readback_name, readback_params, completed_cycles):
    """Save both channels and the exact endurance/PMU configuration together."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for channel, frame in dfs.items():
            if frame is not None and not frame.empty:
                frame.to_excel(writer, sheet_name=f"Channel_{channel}", index=False)
        build_params_table(
            readback_name,
            readback_params,
            completed_cycles,
        ).to_excel(writer, sheet_name="Parameters", index=False)
    return path


def make_cycle_seq_configs():
    """Build the non-measuring endurance cycle seq_configs."""
    rise_time = params_cycle["rise_time_cycle"]
    cycle_voltage = params_cycle["Vc"]
    offset = params_cycle.get("offset_c", 0)
    start_v = [offset, -cycle_voltage + offset, offset, cycle_voltage + offset]
    stop_v = [-cycle_voltage + offset, offset, cycle_voltage + offset, offset]
    time_v = [rise_time] * 4
    meas_types = [0, 0, 0, 0]
    no_measure_window = [0.0] * 4
    ch1_config = (
        1,
        start_v,
        stop_v,
        time_v,
        meas_types,
        no_measure_window.copy(),
        no_measure_window.copy(),
    )
    ch2_config = (
        1,
        [0.0] * 4,
        [0.0] * 4,
        time_v,
        meas_types,
        no_measure_window.copy(),
        no_measure_window.copy(),
    )
    return {CH1: [ch1_config], CH2: [ch2_config]}


def make_pv2_seq_configs():
    """Build PV2 readback seq_configs directly in this script."""
    rise_time = params_pv2["rise_time"]
    delay_time = params_pv2["delay_time"]
    vp = params_pv2["Vp"]
    offset = params_pv2["offset"]
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
    # Match the standalone Fe_cap/PV2.py implementation: segments 0-3 are
    # preset/post-conditioning and segment 4 is the delay. Only segments 5-9
    # contain the two complete PV2 loops used by the integration.
    meas_types = [0, 0, 0, 0, 0, 2, 2, 2, 2, 2]
    ch1_config = (1, start_voltages, stop_voltages, time_values, meas_types)
    ch2_config = (1, [0.0] * len(time_values), [0.0] * len(time_values), time_values, meas_types)
    return {CH1: [ch1_config], CH2: [ch2_config]}


def build_cycle_schedule(target_counts):
    """Convert strictly increasing cumulative milestones into cycle increments."""
    schedule = []
    completed = 0
    for raw_target in target_counts:
        target = int(raw_target)
        if target != raw_target or target <= 0:
            raise ValueError(f"Cycle target must be a positive integer, got {raw_target!r}.")
        if target <= completed:
            raise ValueError("Cycle targets must be strictly increasing.")
        schedule.append((target, target - completed))
        completed = target
    return schedule


def _integrate_pv2_loop(time, voltage, current, area_cm2):
    """Integrate one complete PV2 loop and center its two remanent states."""
    if area_cm2 <= 0:
        raise ValueError("PV2 area_cm2 must be positive.")
    if len(current) < 3:
        raise ValueError("PV2 loop has too few points for integration.")

    raw_time = np.asarray(time, dtype=float)
    voltage = np.asarray(voltage, dtype=float)
    current = np.asarray(current, dtype=float)
    loop_time = np.linspace(0.0, 4.0 * params_pv2["rise_time"], len(current))
    charge = np.zeros(len(current), dtype=float)
    charge[1:] = np.cumsum(
        0.5 * (current[:-1] + current[1:]) * np.diff(loop_time)
    )
    polarization = charge / area_cm2 * 1e6

    positive_peak = int(np.argmax(voltage))
    negative_peak = positive_peak + int(np.argmin(voltage[positive_peak:]))
    if negative_peak <= positive_peak:
        raise ValueError("PV2 loop does not contain positive and negative peaks.")
    offset = params_pv2["offset"]
    positive_pr = positive_peak + int(
        np.argmin(np.abs(voltage[positive_peak : negative_peak + 1] - offset))
    )
    negative_pr = negative_peak + int(
        np.argmin(np.abs(voltage[negative_peak:] - offset))
    )
    polarization -= 0.5 * (polarization[positive_pr] + polarization[negative_pr])

    return pd.DataFrame(
        {
            "Time": loop_time,
            "RawTime": raw_time,
            "Voltage": voltage,
            "Current": current,
            "Polarization": polarization,
        }
    )


def _pv2_loop_sheet(delay_loop, no_delay_loop):
    """Combine the two PV2 loops in one padded table for Excel output."""
    return pd.DataFrame(
        {
            "Voltage_Delay": pd.Series(delay_loop["Voltage"].to_numpy()),
            "Polarization_Delay": pd.Series(delay_loop["Polarization"].to_numpy()),
            "Voltage_NoDelay": pd.Series(no_delay_loop["Voltage"].to_numpy()),
            "Polarization_NoDelay": pd.Series(no_delay_loop["Polarization"].to_numpy()),
        }
    )


def analyze_pv2_readback(df_ch1, df_ch2):
    """Analyze only the measured PV2 sweep and integrate its two loops separately."""
    if df_ch1 is None or df_ch2 is None or df_ch1.empty or df_ch2.empty:
        raise ValueError("PV2 returned empty channel data.")

    point_count = min(len(df_ch1), len(df_ch2))
    split_index = point_count // 2
    if split_index < 3 or point_count - split_index < 3:
        raise ValueError("PV2 returned too few points to split into two loops.")

    time = df_ch1[f"Timestamp {CH1}"].to_numpy()[:point_count]
    voltage = (
        df_ch1[f"Voltage {CH1}"].to_numpy()[:point_count]
        - df_ch2[f"Voltage {CH2}"].to_numpy()[:point_count]
    )
    currents = {
        "i1": df_ch1[f"Current {CH1}"].to_numpy()[:point_count],
        "i2": -df_ch2[f"Current {CH2}"].to_numpy()[:point_count],
    }
    area_cm2 = params_pv2["area_cm2"]
    result = {
        "df_total": pd.DataFrame(
            {
                "Time": time,
                "Voltage": voltage,
                "CurrentI1": currents["i1"],
                "CurrentI2": currents["i2"],
            }
        )
    }
    for label, current in currents.items():
        delay_loop = _integrate_pv2_loop(
            time[:split_index],
            voltage[:split_index],
            current[:split_index],
            area_cm2,
        )
        no_delay_loop = _integrate_pv2_loop(
            time[split_index:],
            voltage[split_index:],
            current[split_index:],
            area_cm2,
        )
        result[f"{label}_delay"] = delay_loop
        result[f"{label}_no_delay"] = no_delay_loop
        result[f"{label}_loops"] = _pv2_loop_sheet(delay_loop, no_delay_loop)
    return result


def make_pund_seq_configs():
    """Build PUND readback seq_configs directly in this script."""
    rise_time = params_pund["rise_time"]
    vp = params_pund["Vp"]
    offset = params_pund["offset"]
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
    time_values = [rise_time] * len(start_voltages)
    meas_types = [2] * len(time_values)
    ch1_config = (1, start_voltages, stop_voltages, time_values, meas_types)
    ch2_config = (1, [0.0] * len(time_values), [0.0] * len(time_values), time_values, meas_types)
    return {CH1: [ch1_config], CH2: [ch2_config]}


def run_cycle_block(query, n_cycles):
    """Run the non-measuring endurance cycle waveform."""
    n_cycles = int(n_cycles)
    if n_cycles <= 0:
        raise ValueError("Endurance cycle increment must be positive.")
    current_ranges = {CH1: params_cycle["Irange1"], CH2: params_cycle["Irange2"]}
    seq_configs = make_cycle_seq_configs()
    seq_list = {CH1: [(1, n_cycles)], CH2: [(1, n_cycles)]}
    try:
        execute_segARB_test(
            query,
            [CH1, CH2],
            seq_configs,
            seq_list=seq_list,
            current_ranges=current_ranges,
            options=SEGARB_OPTIONS,
        )
    finally:
        power_off_outputs(query, (CH1, CH2))


def acquire_readback(query, seq_configs, current_ranges):
    """Acquire both channels and always switch the PMU outputs off afterward."""
    try:
        execute_segARB_test(
            query,
            [CH1, CH2],
            seq_configs,
            current_ranges=current_ranges,
            options=SEGARB_OPTIONS,
        )
        return read_both_channels(query, CH1, CH2)
    finally:
        power_off_outputs(query, (CH1, CH2))


def save_pv2_analysis(path, data, completed_cycles):
    """Save processed PV2 loops and their exact endurance parameters."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        data["df_total"].to_excel(writer, sheet_name="Total", index=False)
        data["i1_loops"].to_excel(writer, sheet_name="I1_Loops", index=False)
        data["i2_loops"].to_excel(writer, sheet_name="I2_Loops", index=False)
        build_params_table("pv2", params_pv2, completed_cycles).to_excel(
            writer,
            sheet_name="Parameters",
            index=False,
        )
    return path


def preview_cycle_waveform(output_path=None):
    """Preview the endurance cycle waveform without connecting to the PMU."""
    return preview_sequence_configs(make_cycle_seq_configs()[CH1], output_path, title_prefix="Endurance cycle CH1")


def preview_pv2_waveform(output_path=None):
    """Preview the endurance PV2 readback waveform without connecting to the PMU."""
    return preview_sequence_configs(make_pv2_seq_configs()[CH1], output_path, title_prefix="Endurance PV2 CH1")


def preview_pund_waveform(output_path=None):
    """Preview the endurance PUND readback waveform without connecting to the PMU."""
    return preview_sequence_configs(make_pund_seq_configs()[CH1], output_path, title_prefix="Endurance PUND CH1")


def main():
    """Run the endurance cycle/readback sequence."""
    schedule = build_cycle_schedule(cycle_counts)
    output_dir = prepare_output_dir(SAVE_DIR)
    with PMUSession(INST, channels=(CH1, CH2)) as session:
        query = session.query
        print(f"Starting endurance test with {len(schedule)} cumulative cycle milestones.")

        for index, (completed_cycles, cycle_increment) in enumerate(schedule, start=1):
            print(
                f"Step {index}/{len(schedule)}: run {cycle_increment} cycles "
                f"to cumulative {completed_cycles}"
            )
            run_cycle_block(query, cycle_increment)

            pv2_ranges = {CH1: params_pv2["Irange1"], CH2: params_pv2["Irange2"]}
            pv2_ch1, pv2_ch2 = acquire_readback(
                query,
                make_pv2_seq_configs(),
                pv2_ranges,
            )
            pv2_data = analyze_pv2_readback(pv2_ch1, pv2_ch2)
            fname_pv2 = reserve_output_stem(output_dir, measurement_name(
                "PV2", params_pv2["Vp"], "tr" + time_tag(params_pv2["rise_time"]),
                "td" + time_tag(params_pv2["delay_time"]), f"n{completed_cycles:09d}",
            ))
            save_channels_with_params(
                {1: pv2_ch1, 2: pv2_ch2},
                f"{fname_pv2}_raw.xlsx",
                "pv2",
                params_pv2,
                completed_cycles,
            )
            save_pv2_analysis(
                f"{fname_pv2}_analysis.xlsx",
                pv2_data,
                completed_cycles,
            )

            fig, ax = plt.subplots(figsize=(6, 5))
            ax.plot(
                pv2_data["i2_delay"]["Voltage"],
                pv2_data["i2_delay"]["Polarization"],
                "b-",
                label=f"Delay {params_pv2['delay_time'] * 1e6:g} us",
            )
            ax.plot(
                pv2_data["i2_no_delay"]["Voltage"],
                pv2_data["i2_no_delay"]["Polarization"],
                "c-",
                label="No delay",
            )
            ax.set_xlabel("Voltage (V)")
            ax.set_ylabel("Polarization (uC/cm^2)")
            ax.set_title(f"PV2 from I2 after {completed_cycles} cycles")
            ax.legend()
            ax.grid(alpha=0.3)
            fig.tight_layout()
            fig.savefig(f"{fname_pv2}_loop.png", dpi=300)
            plt.close(fig)

            pund_ranges = {CH1: params_pund["Irange1"], CH2: params_pund["Irange2"]}
            df_ch1, df_ch2 = acquire_readback(
                query,
                make_pund_seq_configs(),
                pund_ranges,
            )
            if df_ch1 is None or df_ch2 is None or df_ch1.empty or df_ch2.empty:
                raise ValueError(f"PUND readback after {completed_cycles} cycles returned no data.")

            pund_result = analyze_pund_diff(df_ch1, df_ch2, params_pund)
            fname_pund = reserve_output_stem(output_dir, measurement_name(
                "PUND", params_pund["Vp"], "tr" + time_tag(params_pund["rise_time"]),
                "td" + time_tag(params_pund["rise_time"]),
                "tw" + time_tag(params_pund["rise_time"]), f"n{completed_cycles:09d}",
            ))
            save_channels_with_params(
                {1: df_ch1, 2: df_ch2},
                f"{fname_pund}_raw.xlsx",
                "pund",
                params_pund,
                completed_cycles,
            )
            pund_result["df_total"].to_excel(f"{fname_pund}_total.xlsx", index=False)
            pund_result["pund_diff"].to_excel(f"{fname_pund}_diff.xlsx", index=False)

            fig, ax = plt.subplots(figsize=(7, 5))
            for seg in ["P", "U", "N", "D"]:
                sub = pund_result["pund_diff"][pund_result["pund_diff"]["Segment"] == seg]
                ax.plot(sub["Voltage"], sub["Polarization"], ".", label=seg, markersize=4)
            ax.set_xlabel("Voltage (V)")
            ax.set_ylabel("Polarization (uC/cm^2)")
            ax.set_title(f"PUND after {completed_cycles} cycles")
            ax.legend()
            ax.grid(alpha=0.3)
            fig.tight_layout()
            fig.savefig(f"{fname_pund}_loop.png", dpi=300)
            plt.close(fig)

            print(f"Completed step {index}.")


if __name__ == "__main__":
    main()
