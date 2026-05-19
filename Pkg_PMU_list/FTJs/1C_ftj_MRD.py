## FTJ Multilevel Resistance Distribution Measurement
# -*- coding: utf-8 -*-
"""FTJ multilevel resistance distribution measurement."""

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
SAVE_DIR = Path(r"C:\Users\P317151\Documents\data\19-05-2026\Test")
SAVE_DIR.mkdir(parents=True, exist_ok=True)
FILE_STEM = "ftj_mrd"

CURRENT_RANGES = {CH1: 1e-5, CH2: 1e-5}
SEGARB_OPTIONS = {
    "ENABLE_CONNECTION_COMP": False,
    "ENABLE_LOAD_CONFIG": False,
    "LOAD_RESISTANCE": 1e7,
    "ENABLE_LLEC": True,
}

OFFSET_V = 0.0
REFERENCE_V = -1.0
WRITE_VOLTAGES = [0.5, 0.6, 0.7, 0.8]
READ_V = 0.1
CYCLES_PER_LEVEL = 10

REFERENCE_IDLE_1 = 1e-3
REFERENCE_RISE = 2e-7
REFERENCE_DWELL = 1e-6
REFERENCE_FALL = 2e-7
REFERENCE_IDLE_2 = 1e-3

WRITE_IDLE_1 = 1e-3
WRITE_RISE = 2e-7
WRITE_DWELL = 1e-6
WRITE_FALL = 2e-7
WRITE_IDLE_2 = 1e-3

READ_IDLE_1 = 1e-3
READ_RISE = 2e-7
READ_DWELL = 1e-5
READ_FALL = 2e-7
READ_IDLE_2 = 1e-3

BASE_SEQ_ID = 1
MAX_SEGMENTS_PER_SEQ = 10000

PREVIEW_ONLY = False

time_values_reference = [
    REFERENCE_IDLE_1,
    REFERENCE_RISE,
    REFERENCE_DWELL,
    REFERENCE_FALL,
    REFERENCE_IDLE_2,
]
time_values_write = [
    WRITE_IDLE_1,
    WRITE_RISE,
    WRITE_DWELL,
    WRITE_FALL,
    WRITE_IDLE_2,
]
time_values_read = [
    READ_IDLE_1,
    READ_RISE,
    READ_DWELL,
    READ_FALL,
    READ_IDLE_2,
]

meas_types_reference = [0, 0, 1, 0, 0]
meas_start_reference = [0.0] * len(time_values_reference)
meas_stop_reference = time_values_reference

meas_types_write = [0, 0, 1, 0, 0]
meas_start_write = [0.0] * len(time_values_write)
meas_stop_write = time_values_write

meas_types_read = [0, 0, 1, 0, 0]
meas_start_read = [0.0, 0.0, READ_DWELL * 0.5, 0.0, 0.0]
meas_stop_read = [time_values_read[0], time_values_read[1], READ_DWELL * 0.9, time_values_read[3], time_values_read[4]]


def build_pulse_block(amplitude, time_values):
    """Return a 5-segment offset -> amplitude -> offset pulse block."""
    start_v = [OFFSET_V, OFFSET_V, amplitude, amplitude, OFFSET_V]
    stop_v = [OFFSET_V, amplitude, amplitude, OFFSET_V, OFFSET_V]
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
            params_df.to_excel(writer, sheet_name="Parameters", index=False)

    return {
        "df_ch1": df_ch1,
        "df_ch2": df_ch2,
        "result_df": result_df,
        "summary_df": summary_df,
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
