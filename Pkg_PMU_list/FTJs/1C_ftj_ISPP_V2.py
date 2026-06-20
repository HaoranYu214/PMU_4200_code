# -*- coding: utf-8 -*-
"""FTJ ISPP V2 script."""

from datetime import datetime
from pathlib import Path

import pandas as pd
import sys
PKG_ROOT = Path(__file__).resolve().parents[1]
if str(PKG_ROOT) not in sys.path:
    sys.path.insert(0, str(PKG_ROOT))

from src.data_processing import read_both_channels
from src.pmu_tests import execute_segARB_test, power_off_outputs
from src.session import PMUSession
from debug.waveform_preview import preview_sequence_configs

INST = "TCPIP0::129.125.87.80::1225::SOCKET"
CH1, CH2 = 1, 2
SAVE_DIR = Path(r"C:\Users\P317151\Documents\data\19-05-2026\Test")
SAVE_DIR.mkdir(parents=True, exist_ok=True)
FILE_STEM = "ftj_ispp_v2"

CURRENT_RANGES = {CH1: 1e-5, CH2: 1e-5}
SEGARB_OPTIONS = {
    "ENABLE_CONNECTION_COMP": False,
    "ENABLE_LOAD_CONFIG": False,
    "LOAD_RESISTANCE": 1e6,
    "ENABLE_LLEC": False,
}


READ_V = 0.1
POS_V_START = 0.5
POS_V_STOP = 2.0
POS_STEPS = 10
NEG_V_START = -0.5
NEG_V_STOP = -2.0
NEG_STEPS = 10

WRITE_DWELL = 1e-6
READ_DWELL = 1e-5

WRITE_POSITIVE_SEQ_ID = 1
WRITE_NEGATIVE_SEQ_ID = 2
MAX_SEGMENTS_PER_SEQ = 10000
PREVIEW_ONLY = True


# Shared clock definition for each sequence.
time_values_write = [1e-3, 2e-7, WRITE_DWELL, 2e-7, 1e-3]
time_values_read = [1e-3, 2e-7, READ_DWELL, 2e-7, 1e-3]

# Measurement start/stop are times inside each segment window.
meas_types_write = [0, 0, 1, 0, 0]
meas_start_write = [0.0, 0.0, 0.0, 0.0, 0.0]
meas_stop_write = time_values_write

meas_types_read = [0, 0, 1, 0, 0]
meas_start_read = [x * 0.5 for x in time_values_read]
meas_stop_read = [x * 0.9 for x in time_values_read]


def voltage_steps(start, stop, steps):
    """Return inclusive voltage steps from start to stop."""
    if steps <= 1:
        return [stop]
    step = (stop - start) / (steps - 1)
    return [start + i * step for i in range(steps)]


def make_ispp_sequence(seq_id, voltages):
    """Build one sequence containing write+read blocks for a voltage ladder."""
    ch1_start_v = []
    ch1_stop_v = []
    ch2_start_v = []
    ch2_stop_v = []
    time_values = []
    meas_types = []
    meas_start = []
    meas_stop = []

    for voltage in voltages:
        ch1_start_v.extend([0.0, 0.0, voltage, voltage, 0])
        ch1_stop_v.extend([0.0, voltage, voltage, 0.0, 0])
        ch2_start_v.extend([0] * 5)
        ch2_stop_v.extend([0] * 5)
        time_values.extend(time_values_write)
        meas_types.extend(meas_types_write)
        meas_start.extend(meas_start_write)
        meas_stop.extend(meas_stop_write)

        ch1_start_v.extend([0.0, 0.0, READ_V, READ_V, 0])
        ch1_stop_v.extend([0.0, READ_V, READ_V, 0.0, 0])
        ch2_start_v.extend([0] * 5)
        ch2_stop_v.extend([0] * 5)
        time_values.extend(time_values_read)
        meas_types.extend(meas_types_read)
        meas_start.extend(meas_start_read)
        meas_stop.extend(meas_stop_read)

    if len(time_values) > MAX_SEGMENTS_PER_SEQ:
        raise ValueError(
            f"ISPP seq {seq_id} has {len(time_values)} segments, "
            f"above MAX_SEGMENTS_PER_SEQ={MAX_SEGMENTS_PER_SEQ}."
        )

    ch1_config = (seq_id, ch1_start_v, ch1_stop_v, time_values, meas_types, meas_start, meas_stop)
    ch2_config = (seq_id, ch2_start_v, ch2_stop_v, time_values, meas_types, meas_start, meas_stop)
    return ch1_config, ch2_config


POS_VOLTAGES = voltage_steps(POS_V_START, POS_V_STOP, POS_STEPS)
NEG_VOLTAGES = voltage_steps(NEG_V_START, NEG_V_STOP, NEG_STEPS)

ch1_pos_config, ch2_pos_config = make_ispp_sequence(WRITE_POSITIVE_SEQ_ID, POS_VOLTAGES)
ch1_neg_config, ch2_neg_config = make_ispp_sequence(WRITE_NEGATIVE_SEQ_ID, NEG_VOLTAGES)

seq_configs = {
    CH1: [ch1_pos_config, ch1_neg_config],
    CH2: [ch2_pos_config, ch2_neg_config],
}

# Equivalent to :PMU:SARB:WFM:SEQ:LIST.
# Each tuple is (seq_id, loop_count), so loop_count repeats that seq in hardware.
SEQ_PLAN = [(WRITE_POSITIVE_SEQ_ID, 1), (WRITE_NEGATIVE_SEQ_ID, 1)]

SEQ_LIST = {
    CH1: SEQ_PLAN,
    CH2: SEQ_PLAN,
}


def preview_waveform(output_path=None):
    """Preview the generated ISPP waveform on CH1."""
    return preview_sequence_configs(
        [ch1_pos_config, ch1_neg_config],
        output_path,
        title_prefix="FTJ ISPP V2 CH1",
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
    """Return one wide t-V table for plotting ISPP write and read waveforms."""
    write_points = []
    read_points = []
    cursor = 0.0

    for voltages in (POS_VOLTAGES, NEG_VOLTAGES):
        for voltage in voltages:
            write_start_v = [0.0, 0.0, voltage, voltage, 0]
            write_stop_v = [0.0, voltage, voltage, 0.0, 0]
            cursor = _extend_trace_points(
                write_points,
                write_start_v,
                write_stop_v,
                time_values_write,
                cursor,
            )

            read_start_v = [0.0, 0.0, READ_V, READ_V, 0]
            read_stop_v = [0.0, READ_V, READ_V, 0.0, 0]
            cursor = _extend_trace_points(
                read_points,
                read_start_v,
                read_stop_v,
                time_values_read,
                cursor,
            )

    trace_columns = {
        "Time_Write_s": [time for time, _voltage in write_points],
        "Voltage_Write_V": [voltage for _time, voltage in write_points],
        "Time_Read_s": [time for time, _voltage in read_points],
        "Voltage_Read_V": [voltage for _time, voltage in read_points],
    }
    return pd.DataFrame({name: pd.Series(values) for name, values in trace_columns.items()})


def run_ftj_test(*, save_results=True, save_dir=SAVE_DIR, file_stem=FILE_STEM):
    """Run the FTJ segARB sequence list and optionally save raw data."""
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

    if df_ch1 is None and df_ch2 is None:
        raise ValueError("No data returned from the FTJ segARB run.")

    waveform_df = build_waveform_trace_table()
    output_path = None
    if save_results:
        save_dir = Path(save_dir)
        save_dir.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = save_dir / f"{file_stem}_{timestamp}_raw.xlsx"
        params_df = pd.DataFrame(
            [
                {"name": name, "value": repr(value)}
                for name, value in globals().items()
                if name.isupper()
                or name.startswith(("WRITE_", "READ_", "POS_", "NEG_", "time_values_", "meas_", "ch1_", "ch2_", "seq_configs", "SEQ_LIST", "SEQ_PLAN"))
            ]
        )

        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            if df_ch1 is not None and not df_ch1.empty:
                df_ch1.to_excel(writer, sheet_name="Channel_1", index=False)
            if df_ch2 is not None and not df_ch2.empty:
                df_ch2.to_excel(writer, sheet_name="Channel_2", index=False)
            waveform_df.to_excel(writer, sheet_name="Waveform", index=False)
            params_df.to_excel(writer, sheet_name="Parameters", index=False)

    return {
        "df_ch1": df_ch1,
        "df_ch2": df_ch2,
        "waveform_df": waveform_df,
        "output_path": output_path,
    }


def main():
    """Run the FTJ segARB sequence list and save raw data."""
    if PREVIEW_ONLY:
        preview_waveform()
        return
    run_ftj_test()


if __name__ == "__main__":
    main()
