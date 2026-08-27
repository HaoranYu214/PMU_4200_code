# -*- coding: utf-8 -*-
"""FTJ multilevel resistance distribution (MRD) measurement.

MRD means Multilevel Resistance Distribution.

Physical purpose:
    Evaluate FTJ reliability and cycle-to-cycle (C2C) variation. For each
    fixed write-voltage level, repeat the same reset/read/write/read cycle and
    measure how much the resulting resistance varies from trial to trial.

    The mean or median indicates the resistance level produced by a given
    programming voltage, while the standard deviation and full distribution
    quantify repeatability, state overlap, and programming reliability.

Important implementation detail:
    This script does not sweep write-pulse width. ``WRITE_DWELL`` is fixed.
    The programmed variable is ``WRITE_VOLTAGES``.

    It also does not wait for several accumulated write pulses before reading.
    Every cycle currently performs:

        reference/reset pulse
        -> reference-state read
        -> one write pulse at the selected voltage
        -> after-write read

    That four-pulse block is repeated ``CYCLES_PER_LEVEL`` times for every
    write voltage. Therefore the distribution comes from repeated, individually
    read trials at each voltage level, rather than from pulse-width modulation
    or sparse reading after every N programming pulses.
"""

from datetime import datetime
from pathlib import Path
import sys

import pandas as pd

PKG_ROOT = Path(__file__).resolve().parents[2]
REPO_ROOT = PKG_ROOT.parent
for path in (PKG_ROOT, REPO_ROOT):
    if str(path) not in sys.path:
        sys.path.insert(0, str(path))

from debug.waveform_preview import preview_sequence_configs
from src.pmu.data_processing import read_both_channels
from src.pmu.pmu_tests import execute_segARB_test, power_off_outputs
from src.pmu.session import PMUSession
from route_config import apply_route_config

INST = "TCPIP0::129.125.87.80::1225::SOCKET"
CH1, CH2 = 1, 2
SAVE_DIR = Path(r"C:\Users\P317151\Documents\data\06-07-2026\03C6\L40um\FTJ")
SAVE_DIR.mkdir(parents=True, exist_ok=True)
FILE_STEM = "ftj_mrd"

CURRENT_RANGES = {CH1: 1e-3, CH2: 1e-3}
SEGARB_OPTIONS = {
    "ENABLE_CONNECTION_COMP": False,
    "ENABLE_LOAD_CONFIG": False,
    "LOAD_RESISTANCE": 1,
    "ENABLE_LLEC": False,
}

OFFSET_V = 0.0
REFERENCE_V = -7
# WRITE_VOLTAGES = [-2, -2.5, -3, -3.5, -4, -4.5, -5, -5.5, -6, -6.5, -7]
# WRITE_VOLTAGES = [-0.5, -1, -1.5, -2, -2.5, -3, -3.5, -4, -4.5, -5]
WRITE_VOLTAGES = [0.1, 0.5, 1, 1.5, 2, 2.5, 3, 3.5, 4, 4.5, 5, 6, 7]
READ_V = -2
CYCLES_PER_LEVEL = 1

REFERENCE_RISE = 1e-6
REFERENCE_DWELL = 1e-3
REFERENCE_FALL = 1e-6
REFERENCE_IDLE_2 = 0.1

WRITE_RISE = 1e-6
WRITE_DWELL = 5e-5
WRITE_FALL = 1e-6
WRITE_IDLE_2 = 0.1

READ_RISE = 1e-6
READ_DWELL = 5e-5
READ_FALL = 1e-6
READ_IDLE_2 = 0.1
apply_route_config("mrd", globals())

BASE_SEQ_ID = 1
MAX_SEGMENTS_PER_SEQ = 1024

PREVIEW_ONLY = False

time_values_reference = [
    REFERENCE_RISE,
    REFERENCE_DWELL,
    REFERENCE_FALL,
    REFERENCE_IDLE_2,
]
time_values_write = [
    WRITE_RISE,
    WRITE_DWELL,
    WRITE_FALL,
    WRITE_IDLE_2,
]
time_values_read = [
    READ_RISE,
    READ_DWELL,
    READ_FALL,
    READ_IDLE_2,
]

meas_types_reference = [0] * len(time_values_reference)
meas_start_reference = [0.0] * len(time_values_reference)
meas_stop_reference = time_values_reference

meas_types_write = [0] * len(time_values_write)
meas_start_write = [0.0] * len(time_values_write)
meas_stop_write = time_values_write

meas_types_read = [0, 1, 0, 0]
meas_start_read = [0.0, READ_DWELL * 0.5, 0.0, 0.0]
meas_stop_read = [time_values_read[0], READ_DWELL * 0.9, time_values_read[2], time_values_read[3]]


def build_pulse_block(amplitude, time_values):
    """Return a 4-segment offset -> amplitude -> offset pulse block."""
    start_v = [OFFSET_V, amplitude, amplitude, OFFSET_V]
    stop_v = [amplitude, amplitude, OFFSET_V, OFFSET_V]
    return start_v, stop_v, list(time_values)


def build_mrd_sequence(seq_id, write_voltage):
    """Build one MRD sequence: reference pulse, read, write pulse, then read."""
    ch1_start_v = []
    ch1_stop_v = []
    ch2_start_v = []
    ch2_stop_v = []
    time_values = []
    meas_types = []
    meas_start = []
    meas_stop = []

    pulse_defs = [
        (REFERENCE_V, time_values_reference, meas_types_reference, meas_start_reference, meas_stop_reference),
        (READ_V, time_values_read, meas_types_read, meas_start_read, meas_stop_read),
        (write_voltage, time_values_write, meas_types_write, meas_start_write, meas_stop_write),
        (READ_V, time_values_read, meas_types_read, meas_start_read, meas_stop_read),
    ]

    for amplitude, pulse_times, pulse_meas_types, pulse_meas_start, pulse_meas_stop in pulse_defs:
        start_v, stop_v, local_times = build_pulse_block(amplitude, pulse_times)
        ch1_start_v.extend(start_v)
        ch1_stop_v.extend(stop_v)
        ch2_start_v.extend([0.0] * len(local_times))
        ch2_stop_v.extend([0.0] * len(local_times))
        time_values.extend(local_times)
        meas_types.extend(pulse_meas_types)
        meas_start.extend(pulse_meas_start)
        meas_stop.extend(pulse_meas_stop)

    if len(time_values) > MAX_SEGMENTS_PER_SEQ:
        raise ValueError(
            f"MRD sequence {seq_id} has {len(time_values)} segments, "
            f"above MAX_SEGMENTS_PER_SEQ={MAX_SEGMENTS_PER_SEQ}."
        )

    ch1_config = (seq_id, ch1_start_v, ch1_stop_v, time_values, meas_types, meas_start, meas_stop)
    ch2_config = (seq_id, ch2_start_v, ch2_stop_v, time_values, meas_types, meas_start, meas_stop)
    return ch1_config, ch2_config


def build_all_sequences(write_voltages):
    """Build one sequence per write voltage."""
    ch1_configs = []
    ch2_configs = []
    seq_plan = []
    seq_metadata = []

    for index, write_voltage in enumerate(write_voltages):
        seq_id = BASE_SEQ_ID + index
        ch1_config, ch2_config = build_mrd_sequence(seq_id, write_voltage)
        ch1_configs.append(ch1_config)
        ch2_configs.append(ch2_config)
        seq_plan.append((seq_id, CYCLES_PER_LEVEL))
        seq_metadata.append(
            {
                "seq_id": seq_id,
                "write_voltage": write_voltage,
                "cycles": CYCLES_PER_LEVEL,
            }
        )

    return ch1_configs, ch2_configs, seq_plan, seq_metadata


def expected_cycle_table(seq_metadata):
    """Return the expected read-point order from the sequence plan."""
    rows = []
    for item in seq_metadata:
        for cycle_index in range(1, item["cycles"] + 1):
            rows.extend(
                [
                    {
                        "SequenceID": item["seq_id"],
                        "WriteVoltage": item["write_voltage"],
                        "CycleIndex": cycle_index,
                        "ReadType": "RefStateRead",
                    },
                    {
                        "SequenceID": item["seq_id"],
                        "WriteVoltage": item["write_voltage"],
                        "CycleIndex": cycle_index,
                        "ReadType": "AfterWriteRead",
                    },
                ]
            )
    return pd.DataFrame(rows)


ch1_configs, ch2_configs, SEQ_PLAN, SEQ_METADATA = build_all_sequences(WRITE_VOLTAGES)

seq_configs = {
    CH1: ch1_configs,
    CH2: ch2_configs,
}

SEQ_LIST = {
    CH1: SEQ_PLAN,
    CH2: SEQ_PLAN,
}


def preview_waveform(output_path=None):
    """Preview all MRD sequences on CH1."""
    return preview_sequence_configs(
        ch1_configs,
        output_path,
        title_prefix="FTJ MRD CH1",
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
    """Return one wide t-V table for plotting reference, write, and read waveforms."""
    reference_points = []
    write_points = []
    read_points = []
    cursor = 0.0

    for item in SEQ_METADATA:
        for _cycle_index in range(item["cycles"]):
            reference_start_v, reference_stop_v, reference_times = build_pulse_block(REFERENCE_V, time_values_reference)
            cursor = _extend_trace_points(
                reference_points,
                reference_start_v,
                reference_stop_v,
                reference_times,
                cursor,
            )

            read_start_v, read_stop_v, read_times = build_pulse_block(READ_V, time_values_read)
            cursor = _extend_trace_points(
                read_points,
                read_start_v,
                read_stop_v,
                read_times,
                cursor,
            )

            write_start_v, write_stop_v, write_times = build_pulse_block(item["write_voltage"], time_values_write)
            cursor = _extend_trace_points(
                write_points,
                write_start_v,
                write_stop_v,
                write_times,
                cursor,
            )

            read_start_v, read_stop_v, read_times = build_pulse_block(READ_V, time_values_read)
            cursor = _extend_trace_points(
                read_points,
                read_start_v,
                read_stop_v,
                read_times,
                cursor,
            )

    trace_columns = {
        "Time_Reference_s": [time for time, _voltage in reference_points],
        "Voltage_Reference_V": [voltage for _time, voltage in reference_points],
        "Time_Write_s": [time for time, _voltage in write_points],
        "Voltage_Write_V": [voltage for _time, voltage in write_points],
        "Time_Read_s": [time for time, _voltage in read_points],
        "Voltage_Read_V": [voltage for _time, voltage in read_points],
    }
    return pd.DataFrame({name: pd.Series(values) for name, values in trace_columns.items()})


def build_readback_table(df_ch1, df_ch2):
    """Return one row per measured read point."""
    if df_ch1 is None or df_ch2 is None or df_ch1.empty or df_ch2.empty:
        raise ValueError("MRD run returned empty data.")

    expected_df = expected_cycle_table(SEQ_METADATA)
    count = min(len(df_ch1), len(df_ch2), len(expected_df))
    if count == 0:
        raise ValueError("MRD run returned zero read points.")

    expected_df = expected_df.iloc[:count].reset_index(drop=True)
    result_df = expected_df.copy()
    result_df["ReadTimestamp"] = df_ch1[f"Timestamp {CH1}"].values[:count]
    result_df["ReadVoltage"] = df_ch1[f"Voltage {CH1}"].values[:count]
    result_df["ReadCurrent"] = df_ch1[f"Current {CH1}"].values[:count]
    result_df["Resistance"] = result_df["ReadVoltage"] / result_df["ReadCurrent"].replace(0, pd.NA)
    return result_df


def build_distribution_summary(result_df):
    """Summarize resistance distribution for each write level."""
    if result_df.empty:
        return pd.DataFrame()

    summary_df = (
        result_df.groupby(["WriteVoltage", "ReadType"], dropna=False)["Resistance"]
        .agg(
            Count="count",
            Mean="mean",
            Std="std",
            Median="median",
            Min="min",
            Max="max",
        )
        .reset_index()
    )
    return summary_df


def run_ftj_test(*, save_results=True, save_dir=SAVE_DIR, file_stem=FILE_STEM):
    """Run the FTJ MRD test and optionally save the readback table."""
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

    result_df = build_readback_table(df_ch1, df_ch2)
    summary_df = build_distribution_summary(result_df)
    waveform_df = build_waveform_trace_table()

    output_path = None
    if save_results:
        save_dir = Path(save_dir)
        save_dir.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = save_dir / f"{file_stem}_{timestamp}_mrd.xlsx"
        params_df = pd.DataFrame(
            [
                {"name": name, "value": repr(value)}
                for name, value in globals().items()
                if name.isupper()
                or name.startswith(("time_values_", "meas_", "ch1_", "ch2_", "seq_configs", "SEQ_LIST", "SEQ_METADATA"))
            ]
        )

        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            result_df.to_excel(writer, sheet_name="MRD_ReadOnly", index=False)
            summary_df.to_excel(writer, sheet_name="MRD_Summary", index=False)
            df_ch1.to_excel(writer, sheet_name="Channel_1_ReadOnly", index=False)
            df_ch2.to_excel(writer, sheet_name="Channel_2_ReadOnly", index=False)
            waveform_df.to_excel(writer, sheet_name="Waveform", index=False)
            params_df.to_excel(writer, sheet_name="Parameters", index=False)

    return {
        "df_ch1": df_ch1,
        "df_ch2": df_ch2,
        "result_df": result_df,
        "summary_df": summary_df,
        "waveform_df": waveform_df,
        "output_path": output_path,
    }


def main():
    """Run the FTJ MRD sequence and save readback data."""
    if PREVIEW_ONLY:
        preview_waveform()
        return
    run_ftj_test()


if __name__ == "__main__":
    main()
