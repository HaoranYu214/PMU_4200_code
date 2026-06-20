# -*- coding: utf-8 -*-
"""FTJ identical-pulse script with fixed write/read levels."""

from datetime import datetime
from pathlib import Path

import pandas as pd

import sys
PKG_ROOT = Path(__file__).resolve().parents[1]
if str(PKG_ROOT) not in sys.path:
    sys.path.insert(0, str(PKG_ROOT))

from src.data_processing import read_both_channels
from debug.waveform_preview import preview_sequence_configs
from src.pmu_tests import execute_segARB_test, power_off_outputs
from src.session import PMUSession

INST = "TCPIP0::129.125.87.80::1225::SOCKET"
CH1, CH2 = 1, 2
SAVE_DIR = Path(r"C:\Users\P317151\Documents\data\11-06-2026\03D2\L40um1\FTJ\Identical")
SAVE_DIR.mkdir(parents=True, exist_ok=True)
FILE_STEM = "ftj_identical"

CURRENT_RANGES = {CH1: 1e-4, CH2: 1e-4}
SEGARB_OPTIONS = {
    "ENABLE_CONNECTION_COMP": False,
    "ENABLE_LOAD_CONFIG": False,
    "LOAD_RESISTANCE": 1,
    "ENABLE_LLEC": False,
}


WRITE_POSITIVE_V = 1.5
READ_V = -1
REVERSE_READ_V = -READ_V
WRITE_NEGATIVE_V = -8

WRITE_POSITIVE_DWELL = 1e-4
READ_DWELL = 5e-5
WRITE_NEGATIVE_DWELL = 5e-5

WRITE_POSITIVE_TRF = 1e-6
WRITE_NEGATIVE_TRF = 1e-6
READ_TRF = 1e-6

WRITE_POSITIVE_IDLE = 0.1
WRITE_NEGATIVE_IDLE = 0.1
READ_IDLE = 1e-3


PREVIEW_ONLY = False
SAVE_WAVEFORM_PREVIEW = False

WRITE_POSITIVE_SEQ_ID = 1
READ_SEQ_ID = 2
WRITE_NEGATIVE_SEQ_ID = 3
REVERSE_READ_SEQ_ID = 4

POSITIVE_REPEAT_COUNT = 40
NEGATIVE_REPEAT_COUNT = 40


time_values_write_positive = [WRITE_POSITIVE_TRF, WRITE_POSITIVE_DWELL, WRITE_POSITIVE_TRF, WRITE_POSITIVE_IDLE]
time_values_read = [READ_TRF, READ_DWELL, READ_TRF, READ_IDLE]
time_values_write_negative = [WRITE_NEGATIVE_TRF, WRITE_NEGATIVE_DWELL, WRITE_NEGATIVE_TRF, WRITE_NEGATIVE_IDLE]

# Measurement start/stop are absolute offsets in seconds within each segment.
meas_types_write_positive = [0, 0, 0, 0]
meas_start_write_positive = [ 0.0, 0.0, 0.0, 0.0]
meas_stop_write_positive = time_values_write_positive

meas_types_read = [0, 1, 0, 0]
meas_start_read = [x * 0.5 for x in time_values_read]
meas_stop_read = [x * 0.9 for x in time_values_read]

meas_types_write_negative = [0, 0, 0, 0]
meas_start_write_negative = [0.0, 0.0, 0.0, 0.0]
meas_stop_write_negative = time_values_write_negative

ch1_start_v_write_positive = [0.0, WRITE_POSITIVE_V, WRITE_POSITIVE_V, 0]
ch1_stop_v_write_positive = [WRITE_POSITIVE_V, WRITE_POSITIVE_V, 0.0, 0]
ch2_start_v_write_positive = [0] * 4
ch2_stop_v_write_positive = [0] * 4

ch1_start_v_read = [0.0, READ_V, READ_V, 0]
ch1_stop_v_read = [READ_V, READ_V, 0.0, 0]
ch1_start_v_reverse_read = [0.0, REVERSE_READ_V, REVERSE_READ_V, 0]
ch1_stop_v_reverse_read = [REVERSE_READ_V, REVERSE_READ_V, 0.0, 0]


ch2_start_v_read = [0] * 4
ch2_stop_v_read = [0] * 4

ch1_start_v_write_negative = [0.0, WRITE_NEGATIVE_V, WRITE_NEGATIVE_V, 0]
ch1_stop_v_write_negative = [WRITE_NEGATIVE_V, WRITE_NEGATIVE_V, 0.0, 0]
ch2_start_v_write_negative = [0] * 4
ch2_stop_v_write_negative = [0] * 4

ch1_write_positive_config = (WRITE_POSITIVE_SEQ_ID, ch1_start_v_write_positive, ch1_stop_v_write_positive, time_values_write_positive, meas_types_write_positive, meas_start_write_positive, meas_stop_write_positive)
ch2_write_positive_config = (WRITE_POSITIVE_SEQ_ID, ch2_start_v_write_positive, ch2_stop_v_write_positive, time_values_write_positive, meas_types_write_positive, meas_start_write_positive, meas_stop_write_positive)

ch1_read_config = (READ_SEQ_ID, ch1_start_v_read, ch1_stop_v_read, time_values_read, meas_types_read, meas_start_read, meas_stop_read)
ch2_read_config = (READ_SEQ_ID, ch2_start_v_read, ch2_stop_v_read, time_values_read, meas_types_read, meas_start_read, meas_stop_read)
ch1_reverse_read_config = (REVERSE_READ_SEQ_ID, ch1_start_v_reverse_read, ch1_stop_v_reverse_read, time_values_read, meas_types_read, meas_start_read, meas_stop_read)
ch2_reverse_read_config = (REVERSE_READ_SEQ_ID, ch2_start_v_read, ch2_stop_v_read, time_values_read, meas_types_read, meas_start_read, meas_stop_read)


ch1_write_negative_config = (WRITE_NEGATIVE_SEQ_ID, ch1_start_v_write_negative, ch1_stop_v_write_negative, time_values_write_negative, meas_types_write_negative, meas_start_write_negative, meas_stop_write_negative)
ch2_write_negative_config = (WRITE_NEGATIVE_SEQ_ID, ch2_start_v_write_negative, ch2_stop_v_write_negative, time_values_write_negative, meas_types_write_negative, meas_start_write_negative, meas_stop_write_negative)

seq_configs = {
    CH1: [ch1_write_positive_config, ch1_read_config, ch1_reverse_read_config, ch1_write_negative_config],
    CH2: [ch2_write_positive_config, ch2_read_config, ch2_reverse_read_config, ch2_write_negative_config],
}

# Equivalent to :PMU:SARB:WFM:SEQ:LIST.
# SEQ_LIST = {
#     CH1: (
#         [(3, 1), (2, 1)] * 100
#     ),
#     CH2: (
#         [(3, 1), (2, 1)] * 100
#     ),
# }

# SEQ_PLAN = (
#     [(WRITE_POSITIVE_SEQ_ID, 1), (READ_SEQ_ID, 1), (REVERSE_READ_SEQ_ID, 1)] * POSITIVE_REPEAT_COUNT +
#     [(WRITE_NEGATIVE_SEQ_ID, 1),  (REVERSE_READ_SEQ_ID, 1), (READ_SEQ_ID, 1)] * NEGATIVE_REPEAT_COUNT +
#     [(WRITE_POSITIVE_SEQ_ID, 1), (READ_SEQ_ID, 1), (REVERSE_READ_SEQ_ID, 1)] * POSITIVE_REPEAT_COUNT +
#     [(WRITE_NEGATIVE_SEQ_ID, 1), (REVERSE_READ_SEQ_ID, 1), (READ_SEQ_ID, 1)] * NEGATIVE_REPEAT_COUNT
# )


SEQ_PLAN = (
    [(WRITE_POSITIVE_SEQ_ID, 1), (READ_SEQ_ID, 1)] * POSITIVE_REPEAT_COUNT +
    [(WRITE_NEGATIVE_SEQ_ID, 1), (READ_SEQ_ID, 1)] * NEGATIVE_REPEAT_COUNT +
    [(WRITE_POSITIVE_SEQ_ID, 1), (READ_SEQ_ID, 1)] * POSITIVE_REPEAT_COUNT +
    [(WRITE_NEGATIVE_SEQ_ID, 1), (READ_SEQ_ID, 1)] * NEGATIVE_REPEAT_COUNT +
    [(WRITE_POSITIVE_SEQ_ID, 1), (READ_SEQ_ID, 1)] * POSITIVE_REPEAT_COUNT +
    [(WRITE_NEGATIVE_SEQ_ID, 1), (READ_SEQ_ID, 1)] * NEGATIVE_REPEAT_COUNT
)


SEQ_LIST = {
    CH1: SEQ_PLAN,
    CH2: SEQ_PLAN,
}


def validate_sequence_configs(configs_by_channel, seq_list_by_channel):
    """Catch common parameter edit mistakes before sending configs to the PMU."""
    for channel, configs in configs_by_channel.items():
        config_by_id = {config[0]: config for config in configs}
        for config in configs:
            seq_id, start_v, stop_v, times, meas_types, meas_start, meas_stop = config
            lengths = {
                len(start_v),
                len(stop_v),
                len(times),
                len(meas_types),
                len(meas_start),
                len(meas_stop),
            }
            if len(lengths) != 1:
                raise ValueError(f"CH{channel} seq {seq_id} has mismatched segment array lengths.")
            if any(time_value <= 0 for time_value in times):
                raise ValueError(f"CH{channel} seq {seq_id} has non-positive segment time.")

        missing_seq_ids = [
            seq_id
            for seq_id, _repeat_count in seq_list_by_channel[channel]
            if seq_id not in config_by_id
        ]
        if missing_seq_ids:
            raise ValueError(f"CH{channel} SEQ_LIST references missing seq IDs: {missing_seq_ids}")


validate_sequence_configs(seq_configs, SEQ_LIST)


def build_sequence_plan_preview_config(configs, seq_plan, *, seq_id=0):
    """Flatten the SEQ_LIST plan into one config for waveform preview."""
    config_by_id = {config[0]: config for config in configs}
    start_v = []
    stop_v = []
    time_values = []
    meas_types = []
    meas_start = []
    meas_stop = []

    for plan_seq_id, repeat_count in seq_plan:
        config = config_by_id[plan_seq_id]
        for _ in range(repeat_count):
            start_v.extend(config[1])
            stop_v.extend(config[2])
            time_values.extend(config[3])
            meas_types.extend(config[4])
            meas_start.extend(config[5])
            meas_stop.extend(config[6])

    return (seq_id, start_v, stop_v, time_values, meas_types, meas_start, meas_stop)


ch1_sequence_plan_preview_config = build_sequence_plan_preview_config(
    seq_configs[CH1],
    SEQ_PLAN,
    seq_id=0,
)


def preview_waveform(output_path=None):
    """Preview the generated identical-pulse waveform on CH1."""
    return preview_sequence_configs(
        [ch1_sequence_plan_preview_config],
        output_path,
        title_prefix="FTJ Identical CH1",
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
    """Return one wide t-V table for plotting write/read command waveforms."""
    config_by_id = {config[0]: config for config in seq_configs[CH1]}
    write_points = []
    read_points = []
    reverse_read_points = []
    cursor = 0.0

    for seq_id, repeat_count in SEQ_PLAN:
        config = config_by_id[seq_id]
        if seq_id in (WRITE_POSITIVE_SEQ_ID, WRITE_NEGATIVE_SEQ_ID):
            points = write_points
        elif seq_id == READ_SEQ_ID:
            points = read_points
        elif seq_id == REVERSE_READ_SEQ_ID:
            points = reverse_read_points
        else:
            points = []

        for _ in range(repeat_count):
            cursor = _extend_trace_points(
                points,
                config[1],
                config[2],
                config[3],
                cursor,
            )

    trace_columns = {
        "Time_Write_s": [time for time, _voltage in write_points],
        "Voltage_Write_V": [voltage for _time, voltage in write_points],
        "Time_Read_s": [time for time, _voltage in read_points],
        "Voltage_Read_V": [voltage for _time, voltage in read_points],
        "Time_ReverseRead_s": [time for time, _voltage in reverse_read_points],
        "Voltage_ReverseRead_V": [voltage for _time, voltage in reverse_read_points],
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
    preview_path = None
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
                or name.startswith(("WRITE_", "READ_", "time_values_", "meas_", "ch1_", "ch2_", "seq_configs", "SEQ_LIST", "SEQ_PLAN"))
            ]
        )

        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            if df_ch1 is not None and not df_ch1.empty:
                df_ch1.to_excel(writer, sheet_name="Channel_1", index=False)
            if df_ch2 is not None and not df_ch2.empty:
                df_ch2.to_excel(writer, sheet_name="Channel_2", index=False)
            waveform_df.to_excel(writer, sheet_name="Waveform", index=False)
            params_df.to_excel(writer, sheet_name="Parameters", index=False)

        if SAVE_WAVEFORM_PREVIEW:
            preview_path = save_dir / f"{file_stem}_{timestamp}_waveform.png"
            preview_waveform(preview_path)

    return {
        "df_ch1": df_ch1,
        "df_ch2": df_ch2,
        "waveform_df": waveform_df,
        "output_path": output_path,
        "preview_path": preview_path,
    }


def main():
    """Run the FTJ segARB sequence list and save raw data."""
    if PREVIEW_ONLY:
        preview_waveform()
        return
    run_ftj_test()


if __name__ == "__main__":
    main()
