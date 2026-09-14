# -*- coding: utf-8 -*-
"""FTJ RV script with one prepost sequence and one full write+read scan sequence."""
"""从+VP开始测试, 到-Vp然后回到VP"""

from pathlib import Path
import sys

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
from keithley4200.output import measurement_name, reserve_output_stem, time_tag, saved_at
from keithley4200.pmu.data_processing import read_both_channels
from keithley4200.pmu.pmu_tests import (
    MAX_SEGMENTS_PER_SEQUENCE,
    execute_segARB_test,
    power_off_outputs,
    validate_segment_arb_configs,
)
from keithley4200.pmu.session import PMUSession

INST = "TCPIP0::129.125.87.80::1225::SOCKET"
CH1, CH2 = 1, 2
SAVE_DIR = Path(r"C:\Users\P317151\Documents\data\02-09-2026\03B4_2700_800_100\L30_1\FTJ\RV2")
FILE_STEM = "RV2"

CURRENT_RANGES = {CH1: 1e-4, CH2: 1e-4}
SEGARB_OPTIONS = {
    "ENABLE_CONNECTION_COMP": False,
    "ENABLE_LOAD_CONFIG": False,
    "LOAD_RESISTANCE": 1e6,
    "ENABLE_LLEC": False,
}

# PREVIEW_ONLY = True
PREVIEW_ONLY = False

params = {
    # Voltage held before/after every pulse. This is deliberately independent
    # from offset_v, which only shifts the commanded write-voltage window.
    "base_v": 0.0,
    "offset_v": -2,
    "vp": 6,
    "write_level_step": 0.2,
    "read_v": -1,
    "scan_cycles": 1,
    "prepost_dwell": 5e-5,
    "write_dwell": 5e-5,
    "read_dwell": 5e-5,
    "prepost_rise": 1e-5,
    "prepost_fall": 1e-5,
    "prepost_idle": 1e-3,
    "write_rise": 1e-5,
    "write_fall": 1e-5,
    "write_idle": 1e-3,
    "read_rise": 1e-5,
    "read_fall": 1e-5,
    "read_idle": 1e-3,
}


def _sync_parameter_aliases():
    """Expose readable legacy names while keeping ``params`` authoritative."""
    global BASE_V, OFFSET_V, VP, WRITE_LEVEL_STEP, READ, READ_LEVEL, PREPOST_LEVEL
    global SCAN_CYCLES, PREPOST_DWELL, WRITE_DWELL, READ_DWELL
    global PREPOST_RISE, PREPOST_FALL, PREPOST_IDLE_2
    global WRITE_RISE, WRITE_FALL, WRITE_IDLE_2
    global READ_RISE, READ_FALL, READ_IDLE_2

    BASE_V = float(params["base_v"])
    OFFSET_V = float(params["offset_v"])
    VP = float(params["vp"])
    WRITE_LEVEL_STEP = float(params["write_level_step"])
    READ = float(params["read_v"])
    READ_LEVEL = READ - OFFSET_V
    # This is a relative level; the physical prepost voltage is OFFSET_V - VP.
    PREPOST_LEVEL = -VP
    SCAN_CYCLES = int(params["scan_cycles"])
    PREPOST_DWELL = float(params["prepost_dwell"])
    WRITE_DWELL = float(params["write_dwell"])
    READ_DWELL = float(params["read_dwell"])
    PREPOST_RISE = float(params["prepost_rise"])
    PREPOST_FALL = float(params["prepost_fall"])
    PREPOST_IDLE_2 = float(params["prepost_idle"])
    WRITE_RISE = float(params["write_rise"])
    WRITE_FALL = float(params["write_fall"])
    WRITE_IDLE_2 = float(params["write_idle"])
    READ_RISE = float(params["read_rise"])
    READ_FALL = float(params["read_fall"])
    READ_IDLE_2 = float(params["read_idle"])


_sync_parameter_aliases()

PREPOST_SEQ_ID = 1
FIRST_SCAN_SEQ_ID = 2
SEGMENTS_PER_SCAN_POINT = 8
MAX_SEGMENTS_PER_SEQ = MAX_SEGMENTS_PER_SEQUENCE



time_values_prepost = [PREPOST_RISE, PREPOST_DWELL, PREPOST_FALL, PREPOST_IDLE_2]
time_values_write = [WRITE_RISE, WRITE_DWELL, WRITE_FALL, WRITE_IDLE_2]
time_values_read = [READ_RISE, READ_DWELL, READ_FALL, READ_IDLE_2]

meas_types_prepost = [0, 0, 0, 0]
meas_start_prepost = [0.0] * len(time_values_prepost)
meas_stop_prepost = [0.0] * len(time_values_prepost)

meas_types_write = [0, 0, 0, 0]
meas_start_write = [0.0, WRITE_DWELL* 0, 0.0, 0.0]
meas_stop_write = [0.0, 0.0, 0.0, 0.0]

meas_types_read = [0, 1, 0, 0]
meas_start_read = [0.0, READ_DWELL * 0.5, 0.0, 0.0]
meas_stop_read = [0.0, READ_DWELL * 0.9, 0.0, 0.0]


def build_pulse_block(level, time_values):
    """Return a base -> (write offset + relative level) -> base pulse."""
    target_v = OFFSET_V + level
    start_v = [BASE_V, target_v, target_v, BASE_V]
    stop_v = [target_v, target_v, BASE_V, BASE_V]
    return start_v, stop_v, list(time_values)


def _levels_between(start_level, stop_level, step):
    """Return inclusive levels from start_level to stop_level."""
    if start_level == stop_level:
        return [round(start_level, 10)]

    step = abs(step)
    direction = 1 if stop_level > start_level else -1
    step *= direction

    levels = []
    current = start_level
    while (direction > 0 and current < stop_level) or (direction < 0 and current > stop_level):
        levels.append(round(current, 10))
        current += step
    levels.append(round(stop_level, 10))
    return levels


def voltage_sweep_path(vp, step, cycles=1):
    """Return relative levels: +Vp -> -Vp -> +Vp, repeated for the requested cycles."""
    if vp == 0:
        return []

    path = []
    for cycle_index in range(cycles):
        down_leg = _levels_between(vp, -vp, step)
        up_leg = _levels_between(-vp, vp, step)
        if cycle_index == 0:
            path.extend(down_leg)
        else:
            path.extend(down_leg[1:])
        path.extend(up_leg[1:])
    return path


def make_prepost_sequence():
    """Build the prepost sequence: one pulse only."""
    start_v, stop_v, time_values = build_pulse_block(PREPOST_LEVEL, time_values_prepost)
    ch1_config = (PREPOST_SEQ_ID, start_v, stop_v, time_values, meas_types_prepost, meas_start_prepost, meas_stop_prepost)
    ch2_config = (
        PREPOST_SEQ_ID,
        [0.0] * len(time_values),
        [0.0] * len(time_values),
        time_values,
        meas_types_prepost,
        meas_start_prepost,
        meas_stop_prepost,
    )
    return ch1_config, ch2_config


def make_scan_sequence(seq_id, voltages):
    """Build one scan sequence: write pulse, then small read pulse, for a chunk of scan points."""
    ch1_start_v = []
    ch1_stop_v = []
    ch2_start_v = []
    ch2_stop_v = []
    time_values = []
    meas_types = []
    meas_start = []
    meas_stop = []

    for voltage in voltages:
        write_start_v, write_stop_v, write_times = build_pulse_block(voltage, time_values_write)
        ch1_start_v.extend(write_start_v)
        ch1_stop_v.extend(write_stop_v)
        ch2_start_v.extend([0.0] * len(write_times))
        ch2_stop_v.extend([0.0] * len(write_times))
        time_values.extend(write_times)
        meas_types.extend(meas_types_write)
        meas_start.extend(meas_start_write)
        meas_stop.extend(meas_stop_write)

        read_start_v, read_stop_v, read_times = build_pulse_block(READ_LEVEL, time_values_read)
        ch1_start_v.extend(read_start_v)
        ch1_stop_v.extend(read_stop_v)
        ch2_start_v.extend([0.0] * len(read_times))
        ch2_stop_v.extend([0.0] * len(read_times))
        time_values.extend(read_times)
        meas_types.extend(meas_types_read)
        meas_start.extend(meas_start_read)
        meas_stop.extend(meas_stop_read)

    if len(time_values) > MAX_SEGMENTS_PER_SEQ:
        raise ValueError(
            f"RV scan sequence has {len(time_values)} segments, "
            f"above MAX_SEGMENTS_PER_SEQ={MAX_SEGMENTS_PER_SEQ}."
        )

    ch1_config = (seq_id, ch1_start_v, ch1_stop_v, time_values, meas_types, meas_start, meas_stop)
    ch2_config = (seq_id, ch2_start_v, ch2_stop_v, time_values, meas_types, meas_start, meas_stop)
    return ch1_config, ch2_config


def chunk_scan_voltages(voltages):
    """Split scan voltages into multiple sequences to stay below the PMU segment limit."""
    max_points_per_seq = max(1, MAX_SEGMENTS_PER_SEQ // SEGMENTS_PER_SCAN_POINT)
    return [
        voltages[index : index + max_points_per_seq]
        for index in range(0, len(voltages), max_points_per_seq)
    ]


def _rebuild_runtime_config():
    """Rebuild every waveform-derived value after workflow configuration."""
    global time_values_prepost, time_values_write, time_values_read
    global meas_start_read, meas_stop_read
    global SCAN_LEVELS, SCAN_VOLTAGES, SCAN_LEVEL_CHUNKS
    global ch1_prepost_config, ch2_prepost_config, scan_seq_ids
    global scan_config_pairs, ch1_scan_configs, ch2_scan_configs
    global seq_configs, SEQ_PLAN, SEQ_LIST, CURRENT_RANGES

    _sync_parameter_aliases()
    time_values_prepost = [PREPOST_RISE, PREPOST_DWELL, PREPOST_FALL, PREPOST_IDLE_2]
    time_values_write = [WRITE_RISE, WRITE_DWELL, WRITE_FALL, WRITE_IDLE_2]
    time_values_read = [READ_RISE, READ_DWELL, READ_FALL, READ_IDLE_2]
    meas_start_read = [0.0, READ_DWELL * 0.5, 0.0, 0.0]
    meas_stop_read = [0.0, READ_DWELL * 0.9, 0.0, 0.0]

    SCAN_LEVELS = voltage_sweep_path(VP, WRITE_LEVEL_STEP, SCAN_CYCLES)
    SCAN_VOLTAGES = [OFFSET_V + level for level in SCAN_LEVELS]
    SCAN_LEVEL_CHUNKS = chunk_scan_voltages(SCAN_LEVELS)
    ch1_prepost_config, ch2_prepost_config = make_prepost_sequence()
    scan_seq_ids = [FIRST_SCAN_SEQ_ID + index for index in range(len(SCAN_LEVEL_CHUNKS))]
    scan_config_pairs = [
        make_scan_sequence(seq_id, chunk)
        for seq_id, chunk in zip(scan_seq_ids, SCAN_LEVEL_CHUNKS)
    ]
    ch1_scan_configs = [pair[0] for pair in scan_config_pairs]
    ch2_scan_configs = [pair[1] for pair in scan_config_pairs]
    seq_configs = {
        CH1: [ch1_prepost_config] + ch1_scan_configs,
        CH2: [ch2_prepost_config] + ch2_scan_configs,
    }
    SEQ_PLAN = [(PREPOST_SEQ_ID, 1)] + [(seq_id, 1) for seq_id in scan_seq_ids]
    SEQ_LIST = {CH1: list(SEQ_PLAN), CH2: list(SEQ_PLAN)}
    CURRENT_RANGES = {
        CH1: float(CURRENT_RANGES.get(CH1, next(iter(CURRENT_RANGES.values())))),
        CH2: float(CURRENT_RANGES.get(CH2, next(iter(CURRENT_RANGES.values())))),
    }
    validate_segment_arb_configs(seq_configs)


_rebuild_runtime_config()


def configure_measurement(
    *,
    params_override=None,
    inst=None,
    channels=None,
    current_ranges=None,
    segarb_options=None,
    save_dir=None,
    file_stem=None,
):
    """Apply one complete workflow configuration and rebuild the RV waveform."""
    global INST, CH1, CH2, SAVE_DIR, FILE_STEM, CURRENT_RANGES, SEGARB_OPTIONS

    if params_override is not None:
        params.clear()
        params.update(params_override)
    if inst is not None:
        INST = inst
    if channels is not None:
        CH1, CH2 = tuple(channels)
    if current_ranges is not None:
        CURRENT_RANGES = dict(current_ranges)
    if segarb_options is not None:
        SEGARB_OPTIONS = dict(segarb_options)
    if save_dir is not None:
        SAVE_DIR = Path(save_dir)
    if file_stem is not None:
        FILE_STEM = str(file_stem)
    _rebuild_runtime_config()


def preview_waveform(output_path=None, *, show=True, title_prefix="FTJ RV CH1"):
    """Preview the generated RV waveform on CH1."""
    return preview_sequence_configs(
        [ch1_prepost_config] + ch1_scan_configs,
        output_path,
        title_prefix=title_prefix,
        show=show,
    )


def _extend_trace_points(points, start_v, stop_v, time_values, start_time, *, add_gap=True):
    """Append t-V endpoint pairs for one waveform block."""
    cursor = start_time
    for segment_start_v, segment_stop_v, segment_time in zip(start_v, stop_v, time_values):
        next_cursor = cursor + segment_time
        points.append((cursor, segment_start_v))
        points.append((next_cursor, segment_stop_v))
        cursor = next_cursor
    if add_gap:
        points.append((None, None))
    return cursor


def build_waveform_trace_table():
    """Return one wide t-V table for plotting prepost, write, and read waveforms."""
    prepost_points = []
    write_points = []
    read_points = []
    cursor = 0.0

    cursor = _extend_trace_points(
        prepost_points,
        ch1_prepost_config[1],
        ch1_prepost_config[2],
        ch1_prepost_config[3],
        cursor,
        add_gap=False,
    )

    for chunk in SCAN_LEVEL_CHUNKS:
        for level in chunk:
            write_start_v, write_stop_v, write_times = build_pulse_block(level, time_values_write)
            cursor = _extend_trace_points(
                write_points,
                write_start_v,
                write_stop_v,
                write_times,
                cursor,
            )

            read_start_v, read_stop_v, read_times = build_pulse_block(READ_LEVEL, time_values_read)
            cursor = _extend_trace_points(
                read_points,
                read_start_v,
                read_stop_v,
                read_times,
                cursor,
            )

    trace_columns = {
        "Time_Prepost_s": [time for time, _voltage in prepost_points],
        "Voltage_Prepost_V": [voltage for _time, voltage in prepost_points],
        "Time_Write_s": [time for time, _voltage in write_points],
        "Voltage_Write_V": [voltage for _time, voltage in write_points],
        "Time_Read_s": [time for time, _voltage in read_points],
        "Voltage_Read_V": [voltage for _time, voltage in read_points],
    }
    return pd.DataFrame({name: pd.Series(values) for name, values in trace_columns.items()})


def build_readback_table(df_ch1, df_ch2):
    """Return only the readback points, one row per commanded write level."""
    if df_ch1 is None or df_ch2 is None or df_ch1.empty or df_ch2.empty:
        raise ValueError("RV run returned empty data.")

    expected_count = len(SCAN_VOLTAGES)
    actual_counts = (len(df_ch1), len(df_ch2))
    if actual_counts != (expected_count, expected_count):
        raise ValueError(
            "RV readback count mismatch: "
            f"expected {expected_count}, got CH{CH1}={actual_counts[0]} "
            f"and CH{CH2}={actual_counts[1]}."
        )
    count = expected_count
    rv_df = pd.DataFrame(
        {
            "CommandedWriteLevel": SCAN_LEVELS[:count],
            "CommandedWriteVoltage": SCAN_VOLTAGES[:count],
            "TimestampI1": df_ch1[f"Timestamp {CH1}"].values[:count],
            "TimestampI2": df_ch2[f"Timestamp {CH2}"].values[:count],
            "ReadVoltageI1": df_ch1[f"Voltage {CH1}"].values[:count],
            "ReadVoltageI2": df_ch2[f"Voltage {CH2}"].values[:count],
            "CurrentI1": df_ch1[f"Current {CH1}"].values[:count],
            "CurrentI2": df_ch2[f"Current {CH2}"].values[:count],
        }
    )
    rv_df["ReadVoltageDiff"] = rv_df["ReadVoltageI1"] - rv_df["ReadVoltageI2"]
    rv_df["ResistanceI1"] = rv_df["ReadVoltageI1"] / rv_df["CurrentI1"].replace(0, pd.NA)
    rv_df["ResistanceI2"] = rv_df["ReadVoltageDiff"] / (-rv_df["CurrentI2"]).replace(0, pd.NA)
    return rv_df


def run_ftj_test(*, save_results=True, save_dir=None, file_stem=None):
    """Run the FTJ RV test and optionally save readback-only results."""
    save_dir = SAVE_DIR if save_dir is None else Path(save_dir)
    file_stem = FILE_STEM if file_stem is None else str(file_stem)
    with PMUSession(INST, channels=(CH1, CH2)) as session:
        query = session.query
        execute_segARB_test(
            query,
            channels=[CH1, CH2],
            seq_configs=seq_configs,
            seq_list=SEQ_LIST,
            current_ranges=CURRENT_RANGES,
            options=SEGARB_OPTIONS,
        )

        df_ch1, df_ch2 = read_both_channels(query, CH1, CH2)
        power_off_outputs(query, (CH1, CH2))

    rv_df = build_readback_table(df_ch1, df_ch2)
    waveform_df = build_waveform_trace_table()

    output_path = None
    if save_results:
        save_dir.mkdir(parents=True, exist_ok=True)
        output_stem = reserve_output_stem(
            save_dir, measurement_name(file_stem, params["vp"], "tw" + time_tag(params["write_dwell"])),
        )
        output_path = Path(f"{output_stem}.xlsx")
        saved_params = {
            "saved_at": saved_at(),
            **params,
            "inst": INST,
            "channels": (CH1, CH2),
            "current_ranges": CURRENT_RANGES,
            "segarb_options": SEGARB_OPTIONS,
        }
        params_df = pd.DataFrame(
            {"name": saved_params.keys(), "value": map(repr, saved_params.values())}
        )

        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            rv_df.to_excel(writer, sheet_name="RV_ReadOnly", index=False)
            df_ch1.to_excel(writer, sheet_name="Channel_1_ReadOnly", index=False)
            df_ch2.to_excel(writer, sheet_name="Channel_2_ReadOnly", index=False)
            waveform_df.to_excel(writer, sheet_name="Waveform", index=False)
            params_df.to_excel(writer, sheet_name="Parameters", index=False)

    return {
        "df_ch1": df_ch1,
        "df_ch2": df_ch2,
        "rv_df": rv_df,
        "waveform_df": waveform_df,
        "output_path": output_path,
    }


def main():
    """Run the FTJ RV sequence and save readback data."""
    if PREVIEW_ONLY:
        preview_waveform()
        return
    run_ftj_test()


if __name__ == "__main__":
    main()
