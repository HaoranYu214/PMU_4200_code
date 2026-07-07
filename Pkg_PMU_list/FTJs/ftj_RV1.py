# -*- coding: utf-8 -*-
"""FTJ RV script with one prepost sequence and one full write+read scan sequence."""
"""从0开始测试, 逐步增加电压, 先正后负"""

from datetime import datetime
from pathlib import Path
import sys

import pandas as pd

PKG_ROOT = Path(__file__).resolve().parents[1]
if str(PKG_ROOT) not in sys.path:
    sys.path.insert(0, str(PKG_ROOT))

from debug.waveform_preview import preview_sequence_configs
from src.data_processing import read_both_channels
from src.pmu_tests import execute_segARB_test, power_off_outputs
from src.session import PMUSession

INST = "TCPIP0::129.125.87.80::1225::SOCKET"
CH1, CH2 = 1, 2
SAVE_DIR = Path(r"C:\Users\P317151\Documents\data\06-07-2026\03C6\L40um4\RV1")
SAVE_DIR.mkdir(parents=True, exist_ok=True)
FILE_STEM = "ftj_rv"

CURRENT_RANGES = {CH1: 1e-5, CH2: 1e-5}
SEGARB_OPTIONS = {
    "ENABLE_CONNECTION_COMP": False,
    "ENABLE_LOAD_CONFIG": False,
    "LOAD_RESISTANCE": 1e6,
    "ENABLE_LLEC": False,
}

OFFSET_V = -1
VP = 6
WRITE_LEVEL_STEP = 0.2
READ = -1
READ_LEVEL = READ-OFFSET_V
PREPOST_LEVEL =  - VP

PREPOST_DWELL = 5e-5
WRITE_DWELL = 5e-5
READ_DWELL = 5e-5

PREPOST_RISE = 1e-5
PREPOST_FALL = 1e-5
PREPOST_IDLE_2 = 0.5

WRITE_RISE = 1e-5
WRITE_FALL = 1e-5
WRITE_IDLE_2 = 0.5

READ_RISE = 1e-5
READ_FALL = 1e-5
READ_IDLE_2 = 0.5

PREPOST_SEQ_ID = 1
FIRST_SCAN_SEQ_ID = 2
SEGMENTS_PER_SCAN_POINT = 8
MAX_SEGMENTS_PER_SEQ = 1000

# PREVIEW_ONLY = True
PREVIEW_ONLY = False

time_values_prepost = [PREPOST_RISE, PREPOST_DWELL, PREPOST_FALL, PREPOST_IDLE_2]
time_values_write = [WRITE_RISE, WRITE_DWELL, WRITE_FALL, WRITE_IDLE_2]
time_values_read = [READ_RISE, READ_DWELL, READ_FALL, READ_IDLE_2]

meas_types_prepost = [0, 0, 0, 0]
meas_start_prepost = [0.0] * len(time_values_prepost)
meas_stop_prepost = [0.0] * len(time_values_prepost)

meas_types_write = [0, 0, 0, 0]
meas_start_write = [0.0] * len(time_values_write)
meas_stop_write = [0.0] * len(time_values_write)

meas_types_read = [0, 1, 0, 0]
meas_start_read = [0.0, READ_DWELL * 0.5, 0.0, 0.0]
meas_stop_read = [0.0, READ_DWELL * 0.9, 0.0, 0.0]


def build_pulse_block(level, time_values):
    """Return a 4-segment offset -> (offset + level) -> offset pulse block."""
    target_v = OFFSET_V + level
    start_v = [OFFSET_V, target_v, target_v, OFFSET_V]
    stop_v = [target_v, target_v, OFFSET_V, OFFSET_V]
    return start_v, stop_v, list(time_values)


def _leg_to_target(target_level, step):
    """Return 0 -> target_level as relative levels, excluding the initial 0."""
    if target_level == 0:
        return []

    step = abs(step)
    sign = 1 if target_level > 0 else -1
    magnitude = abs(target_level)
    n_steps = max(1, int(round(magnitude / step)))
    levels = [round(sign * step * i, 10) for i in range(1, n_steps + 1)]
    levels[-1] = round(target_level, 10)
    return levels


def voltage_sweep_path(vp, step):
    """Return relative levels: 0 -> +vp -> 0 -> -vp -> 0.

    Changing the sign of vp reverses which side is visited first.
    """
    first_leg = _leg_to_target(vp, step)
    second_leg = _leg_to_target(-vp, step)

    first_return = list(reversed(first_leg[:-1])) + [0.0] if first_leg else [0.0]
    second_return = list(reversed(second_leg[:-1])) + [0.0] if second_leg else [0.0]
    return first_leg + first_return + second_leg + second_return


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


SCAN_LEVELS = voltage_sweep_path(
    VP,
    WRITE_LEVEL_STEP,
)
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
SEQ_LIST = {
    CH1: SEQ_PLAN,
    CH2: SEQ_PLAN,
}


def preview_waveform(output_path=None):
    """Preview the generated RV waveform on CH1."""
    return preview_sequence_configs(
        [ch1_prepost_config] + ch1_scan_configs,
        output_path,
        title_prefix="FTJ RV CH1",
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

    count = min(len(df_ch1), len(df_ch2), len(SCAN_VOLTAGES))
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


def run_ftj_test(*, save_results=True, save_dir=SAVE_DIR, file_stem=FILE_STEM):
    """Run the FTJ RV test and optionally save readback-only results."""
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
        save_dir = Path(save_dir)
        save_dir.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = save_dir / f"{file_stem}_{timestamp}_rv.xlsx"
        params_df = pd.DataFrame(
            [
                {"name": name, "value": repr(value)}
                for name, value in globals().items()
                if name.isupper()
                or name.startswith(("time_values_", "meas_", "ch1_", "ch2_", "seq_configs", "SEQ_LIST"))
            ]
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
