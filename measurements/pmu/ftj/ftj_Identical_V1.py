# -*- coding: utf-8 -*-
"""FTJ identical-pulse test with fixed write and read levels.

Physical purpose:
    Apply the same write pulse repeatedly and read the FTJ state after every
    pulse. Unlike ISPP, the write amplitude does not increase. The changing
    variable is cumulative pulse count.

    This reveals pulse-to-pulse evolution at fixed programming conditions:
    gradual resistance change, switching probability, saturation, and the
    number of identical pulses required to reach a target state. Because the
    FTJ response is nonlinear, identical pulses may initially cause little
    change and then produce an abrupt jump or rapid saturation.

Difference from ISPP (Incremental Step Pulse Programming):
    ISPP increases write amplitude step by step to make resistance/conductance
    updates more controlled and approximately linear. Identical-pulse testing
    keeps amplitude fixed and exposes the device's nonlinear accumulation
    versus pulse number.
"""

from pathlib import Path

import pandas as pd

import sys
REPO_ROOT = next(
    parent for parent in Path(__file__).resolve().parents
    if (parent / "pyproject.toml").is_file()
)
SRC_ROOT = REPO_ROOT / "src"
for path in (SRC_ROOT, REPO_ROOT):
    if str(path) not in sys.path:
        sys.path.insert(0, str(path))

from keithley4200.output import measurement_name, reserve_output_stem, time_tag, saved_at
from keithley4200.pmu.data_processing import read_both_channels
from keithley4200.tools.waveform_preview import preview_sequence_configs
from keithley4200.pmu.pmu_tests import (
    MAX_SEGMENTS_PER_SEQUENCE,
    execute_segARB_test,
    power_off_outputs,
    validate_segment_arb_configs,
)
from keithley4200.pmu.session import PMUSession

INST = "TCPIP0::129.125.87.80::1225::SOCKET"
CH1, CH2 = 1, 2
SAVE_DIR = Path(r"C:\Users\P317151\Documents\data\02-09-2026\03B4_2700_800_100\L30_1\FTJ\Identical")
# SAVE_DIR = Path(r"D:\Code\data\20260620")
FILE_STEM = "Identical1"

CURRENT_RANGES = {CH1: 1e-5, CH2: 1e-5}
SEGARB_OPTIONS = {
    "ENABLE_CONNECTION_COMP": False,
    "ENABLE_LOAD_CONFIG": False,
    "LOAD_RESISTANCE": 1e6,
    "ENABLE_LLEC": False,
}

# PREVIEW_ONLY = True
PREVIEW_ONLY = False
SAVE_WAVEFORM_PREVIEW = False

params = {
    # Voltage held before/after every write and read pulse.
    "base_v": 0.0,
    "write_positive_v": 5,
    "write_negative_v": -5,
    "read_v": -1,
    "write_positive_dwell": 5e-5,
    "write_negative_dwell": 5e-5,
    "read_dwell": 5e-5,
    "write_positive_trf": 1e-6,
    "write_negative_trf": 1e-6,
    "read_trf": 1e-6,
    "write_positive_idle": 1e-2,
    "write_negative_idle": 1e-2,
    "read_idle": 1e-2,
    "positive_repeat_count": 50,
    "negative_repeat_count": 50,
    "sequence_cycle_count": 2,
}


def _sync_parameter_aliases():
    global BASE_V, WRITE_POSITIVE_V, WRITE_NEGATIVE_V, READ_V, REVERSE_READ_V
    global WRITE_POSITIVE_DWELL, WRITE_NEGATIVE_DWELL, READ_DWELL
    global WRITE_POSITIVE_TRF, WRITE_NEGATIVE_TRF, READ_TRF
    global WRITE_POSITIVE_IDLE, WRITE_NEGATIVE_IDLE, READ_IDLE
    global POSITIVE_REPEAT_COUNT, NEGATIVE_REPEAT_COUNT, SEQ_CYCLE_COUNT

    BASE_V = float(params["base_v"])
    WRITE_POSITIVE_V = float(params["write_positive_v"])
    WRITE_NEGATIVE_V = float(params["write_negative_v"])
    READ_V = float(params["read_v"])
    REVERSE_READ_V = -READ_V
    WRITE_POSITIVE_DWELL = float(params["write_positive_dwell"])
    WRITE_NEGATIVE_DWELL = float(params["write_negative_dwell"])
    READ_DWELL = float(params["read_dwell"])
    WRITE_POSITIVE_TRF = float(params["write_positive_trf"])
    WRITE_NEGATIVE_TRF = float(params["write_negative_trf"])
    READ_TRF = float(params["read_trf"])
    WRITE_POSITIVE_IDLE = float(params["write_positive_idle"])
    WRITE_NEGATIVE_IDLE = float(params["write_negative_idle"])
    READ_IDLE = float(params["read_idle"])
    POSITIVE_REPEAT_COUNT = int(params["positive_repeat_count"])
    NEGATIVE_REPEAT_COUNT = int(params["negative_repeat_count"])
    SEQ_CYCLE_COUNT = int(params["sequence_cycle_count"])


_sync_parameter_aliases()



WRITE_POSITIVE_SEQ_ID = 1
READ_SEQ_ID = 2
WRITE_NEGATIVE_SEQ_ID = 3
REVERSE_READ_SEQ_ID = 4

time_values_write_positive = [WRITE_POSITIVE_TRF, WRITE_POSITIVE_DWELL, WRITE_POSITIVE_TRF, WRITE_POSITIVE_IDLE]
time_values_read = [READ_TRF, READ_DWELL, READ_TRF, READ_IDLE]
time_values_write_negative = [WRITE_NEGATIVE_TRF, WRITE_NEGATIVE_DWELL, WRITE_NEGATIVE_TRF, WRITE_NEGATIVE_IDLE]

# Measurement start/stop are absolute offsets in seconds within each segment.
meas_types_write_positive = [0, 0, 0, 0]
meas_start_write_positive = [ 0.0, 0.0, 0.0, 0.0]
meas_stop_write_positive = [0.0, 0.0, 0.0, 0.0]

meas_types_read = [0, 1, 0, 0]
meas_start_read = [0.0, READ_DWELL * 0.5, 0.0, 0.0]
meas_stop_read = [0.0, READ_DWELL * 0.9, 0.0, 0.0]

meas_types_write_negative = [0, 0, 0, 0]
meas_start_write_negative = [0.0, 0.0, 0.0, 0.0]
meas_stop_write_negative = [0.0, 0.0, 0.0, 0.0]

ch1_start_v_write_positive = [BASE_V, WRITE_POSITIVE_V, WRITE_POSITIVE_V, BASE_V]
ch1_stop_v_write_positive = [WRITE_POSITIVE_V, WRITE_POSITIVE_V, BASE_V, BASE_V]
ch2_start_v_write_positive = [0] * 4
ch2_stop_v_write_positive = [0] * 4

ch1_start_v_read = [BASE_V, READ_V, READ_V, BASE_V]
ch1_stop_v_read = [READ_V, READ_V, BASE_V, BASE_V]
ch1_start_v_reverse_read = [BASE_V, REVERSE_READ_V, REVERSE_READ_V, BASE_V]
ch1_stop_v_reverse_read = [REVERSE_READ_V, REVERSE_READ_V, BASE_V, BASE_V]


ch2_start_v_read = [0] * 4
ch2_stop_v_read = [0] * 4

ch1_start_v_write_negative = [BASE_V, WRITE_NEGATIVE_V, WRITE_NEGATIVE_V, BASE_V]
ch1_stop_v_write_negative = [WRITE_NEGATIVE_V, WRITE_NEGATIVE_V, BASE_V, BASE_V]
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


FIRST_EXPANDED_SEQ_ID = 1

SINGLE_CYCLE_PLAN = (
    [(WRITE_POSITIVE_SEQ_ID, 1), (READ_SEQ_ID, 1)] * POSITIVE_REPEAT_COUNT +
    [(WRITE_NEGATIVE_SEQ_ID, 1), (READ_SEQ_ID, 1)] * NEGATIVE_REPEAT_COUNT
)
SEQ_PLAN = SINGLE_CYCLE_PLAN * SEQ_CYCLE_COUNT


def build_sequence_plan_config(configs, seq_plan, *, seq_id):
    """Flatten a sequence-list plan into one PMU sequence."""
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


base_seq_configs = seq_configs
expanded_seq_ids = [FIRST_EXPANDED_SEQ_ID + index for index in range(SEQ_CYCLE_COUNT)]
ch1_expanded_configs = [
    build_sequence_plan_config(base_seq_configs[CH1], SINGLE_CYCLE_PLAN, seq_id=seq_id)
    for seq_id in expanded_seq_ids
]
ch2_expanded_configs = [
    build_sequence_plan_config(base_seq_configs[CH2], SINGLE_CYCLE_PLAN, seq_id=seq_id)
    for seq_id in expanded_seq_ids
]

seq_configs = {
    CH1: ch1_expanded_configs,
    CH2: ch2_expanded_configs,
}


SEQ_LIST = {
    CH1: [(seq_id, 1) for seq_id in expanded_seq_ids],
    CH2: [(seq_id, 1) for seq_id in expanded_seq_ids],
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
            if len(times) > MAX_SEGMENTS_PER_SEQUENCE:
                raise ValueError(
                    f"CH{channel} seq {seq_id} has {len(times)} segments; "
                    f"limit is {MAX_SEGMENTS_PER_SEQUENCE}."
                )

        missing_seq_ids = [
            seq_id
            for seq_id, _repeat_count in seq_list_by_channel[channel]
            if seq_id not in config_by_id
        ]
        if missing_seq_ids:
            raise ValueError(f"CH{channel} SEQ_LIST references missing seq IDs: {missing_seq_ids}")


def _rebuild_runtime_config():
    """Rebuild all pulse arrays and sequence plans after parameter injection."""
    global time_values_write_positive, time_values_read, time_values_write_negative
    global meas_start_read, meas_stop_read
    global ch1_start_v_write_positive, ch1_stop_v_write_positive
    global ch1_start_v_read, ch1_stop_v_read
    global ch1_start_v_reverse_read, ch1_stop_v_reverse_read
    global ch1_start_v_write_negative, ch1_stop_v_write_negative
    global ch1_write_positive_config, ch2_write_positive_config
    global ch1_read_config, ch2_read_config
    global ch1_reverse_read_config, ch2_reverse_read_config
    global ch1_write_negative_config, ch2_write_negative_config
    global SINGLE_CYCLE_PLAN, SEQ_PLAN, base_seq_configs, expanded_seq_ids
    global ch1_expanded_configs, ch2_expanded_configs, seq_configs, SEQ_LIST
    global ch1_sequence_plan_preview_config, CURRENT_RANGES

    _sync_parameter_aliases()
    time_values_write_positive = [
        WRITE_POSITIVE_TRF,
        WRITE_POSITIVE_DWELL,
        WRITE_POSITIVE_TRF,
        WRITE_POSITIVE_IDLE,
    ]
    time_values_read = [READ_TRF, READ_DWELL, READ_TRF, READ_IDLE]
    time_values_write_negative = [
        WRITE_NEGATIVE_TRF,
        WRITE_NEGATIVE_DWELL,
        WRITE_NEGATIVE_TRF,
        WRITE_NEGATIVE_IDLE,
    ]
    meas_start_read = [0.0, READ_DWELL * 0.5, 0.0, 0.0]
    meas_stop_read = [0.0, READ_DWELL * 0.9, 0.0, 0.0]

    ch1_start_v_write_positive = [BASE_V, WRITE_POSITIVE_V, WRITE_POSITIVE_V, BASE_V]
    ch1_stop_v_write_positive = [WRITE_POSITIVE_V, WRITE_POSITIVE_V, BASE_V, BASE_V]
    ch1_start_v_read = [BASE_V, READ_V, READ_V, BASE_V]
    ch1_stop_v_read = [READ_V, READ_V, BASE_V, BASE_V]
    ch1_start_v_reverse_read = [BASE_V, REVERSE_READ_V, REVERSE_READ_V, BASE_V]
    ch1_stop_v_reverse_read = [REVERSE_READ_V, REVERSE_READ_V, BASE_V, BASE_V]
    ch1_start_v_write_negative = [BASE_V, WRITE_NEGATIVE_V, WRITE_NEGATIVE_V, BASE_V]
    ch1_stop_v_write_negative = [WRITE_NEGATIVE_V, WRITE_NEGATIVE_V, BASE_V, BASE_V]

    ch1_write_positive_config = (
        WRITE_POSITIVE_SEQ_ID, ch1_start_v_write_positive, ch1_stop_v_write_positive,
        time_values_write_positive, meas_types_write_positive,
        meas_start_write_positive, meas_stop_write_positive,
    )
    ch2_write_positive_config = (
        WRITE_POSITIVE_SEQ_ID, ch2_start_v_write_positive, ch2_stop_v_write_positive,
        time_values_write_positive, meas_types_write_positive,
        meas_start_write_positive, meas_stop_write_positive,
    )
    ch1_read_config = (
        READ_SEQ_ID, ch1_start_v_read, ch1_stop_v_read, time_values_read,
        meas_types_read, meas_start_read, meas_stop_read,
    )
    ch2_read_config = (
        READ_SEQ_ID, ch2_start_v_read, ch2_stop_v_read, time_values_read,
        meas_types_read, meas_start_read, meas_stop_read,
    )
    ch1_reverse_read_config = (
        REVERSE_READ_SEQ_ID, ch1_start_v_reverse_read, ch1_stop_v_reverse_read,
        time_values_read, meas_types_read, meas_start_read, meas_stop_read,
    )
    ch2_reverse_read_config = (
        REVERSE_READ_SEQ_ID, ch2_start_v_read, ch2_stop_v_read, time_values_read,
        meas_types_read, meas_start_read, meas_stop_read,
    )
    ch1_write_negative_config = (
        WRITE_NEGATIVE_SEQ_ID, ch1_start_v_write_negative, ch1_stop_v_write_negative,
        time_values_write_negative, meas_types_write_negative,
        meas_start_write_negative, meas_stop_write_negative,
    )
    ch2_write_negative_config = (
        WRITE_NEGATIVE_SEQ_ID, ch2_start_v_write_negative, ch2_stop_v_write_negative,
        time_values_write_negative, meas_types_write_negative,
        meas_start_write_negative, meas_stop_write_negative,
    )
    base_seq_configs = {
        CH1: [ch1_write_positive_config, ch1_read_config, ch1_reverse_read_config, ch1_write_negative_config],
        CH2: [ch2_write_positive_config, ch2_read_config, ch2_reverse_read_config, ch2_write_negative_config],
    }
    SINGLE_CYCLE_PLAN = (
        [(WRITE_POSITIVE_SEQ_ID, 1), (READ_SEQ_ID, 1)] * POSITIVE_REPEAT_COUNT
        + [(WRITE_NEGATIVE_SEQ_ID, 1), (READ_SEQ_ID, 1)] * NEGATIVE_REPEAT_COUNT
    )
    SEQ_PLAN = SINGLE_CYCLE_PLAN * SEQ_CYCLE_COUNT
    expanded_seq_ids = [FIRST_EXPANDED_SEQ_ID + index for index in range(SEQ_CYCLE_COUNT)]
    ch1_expanded_configs = [
        build_sequence_plan_config(base_seq_configs[CH1], SINGLE_CYCLE_PLAN, seq_id=seq_id)
        for seq_id in expanded_seq_ids
    ]
    ch2_expanded_configs = [
        build_sequence_plan_config(base_seq_configs[CH2], SINGLE_CYCLE_PLAN, seq_id=seq_id)
        for seq_id in expanded_seq_ids
    ]
    seq_configs = {CH1: ch1_expanded_configs, CH2: ch2_expanded_configs}
    SEQ_LIST = {
        CH1: [(seq_id, 1) for seq_id in expanded_seq_ids],
        CH2: [(seq_id, 1) for seq_id in expanded_seq_ids],
    }
    ch1_sequence_plan_preview_config = build_sequence_plan_config(
        base_seq_configs[CH1], SEQ_PLAN, seq_id=0
    )
    CURRENT_RANGES = {
        CH1: float(CURRENT_RANGES.get(CH1, next(iter(CURRENT_RANGES.values())))),
        CH2: float(CURRENT_RANGES.get(CH2, next(iter(CURRENT_RANGES.values())))),
    }
    validate_segment_arb_configs(seq_configs)


_rebuild_runtime_config()
validate_sequence_configs(seq_configs, SEQ_LIST)


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
    """Apply one complete workflow configuration and rebuild all sequences."""
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


def preview_waveform(output_path=None, *, show=True, title_prefix="FTJ Identical CH1"):
    """Preview the generated identical-pulse waveform on CH1."""
    return preview_sequence_configs(
        [ch1_sequence_plan_preview_config],
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
    """Return one wide t-V table for plotting write/read command waveforms."""
    config_by_id = {config[0]: config for config in base_seq_configs[CH1]}
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


def run_ftj_test(*, save_results=True, save_dir=None, file_stem=None):
    """Run the FTJ segARB sequence list and optionally save raw data."""
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

    if df_ch1 is None and df_ch2 is None:
        raise ValueError("No data returned from the FTJ segARB run.")

    waveform_df = build_waveform_trace_table()
    output_path = None
    preview_path = None
    if save_results:
        save_dir.mkdir(parents=True, exist_ok=True)
        output_stem = reserve_output_stem(
            save_dir, measurement_name(file_stem, params["write_positive_v"], "tw" + time_tag(params["write_positive_dwell"])),
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
            if df_ch1 is not None and not df_ch1.empty:
                df_ch1.to_excel(writer, sheet_name="Channel_1", index=False)
            if df_ch2 is not None and not df_ch2.empty:
                df_ch2.to_excel(writer, sheet_name="Channel_2", index=False)
            waveform_df.to_excel(writer, sheet_name="Waveform", index=False)
            params_df.to_excel(writer, sheet_name="Parameters", index=False)

        if SAVE_WAVEFORM_PREVIEW:
            preview_path = Path(f"{output_stem}_waveform.png")
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
