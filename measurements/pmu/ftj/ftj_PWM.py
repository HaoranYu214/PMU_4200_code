# -*- coding: utf-8 -*-
"""FTJ pulse-width-modulation (PWM) experiment using true width sweep.

The script keeps write voltage fixed, varies the dwell time of one write pulse,
and reads the resistance after each width:

    write width 1x   -> read
    write width 2x   -> read
    write width 5x   -> read

Positive write widths are swept first, followed by negative write widths. The
whole plan is flattened into one segARB sequence, so the PMU sequence list only
contains one sequence per channel.
"""

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
SAVE_DIR = Path(r"C:\Users\P317151\Documents\data\06-07-2026\03C6\L40um6\PWM")
FILE_STEM = "PWM"

CURRENT_RANGES = {CH1: 1e-4, CH2: 1e-4}
SEGARB_OPTIONS = {
    "ENABLE_CONNECTION_COMP": False,
    "ENABLE_LOAD_CONFIG": False,
    "LOAD_RESISTANCE": 1e6,
    # LLEC is not available for Segment Arb measurements.
    "ENABLE_LLEC": False,
}

params = {
    # Voltage held before/after every write and read pulse.
    "base_v": 0.0,
    "write_positive_v": 2,
    "write_negative_v": -6,
    "read_v": -1,
    "write_base_dwell": 1e-6,
    "width_multipliers": [1, 2, 5, 10, 20, 50, 100, 200, 500, 1000, 2000, 5000],
    "repeat_count": 5,
    "read_dwell": 5e-5,
    "write_positive_trf": 1e-6,
    "write_negative_trf": 1e-6,
    "read_trf": 1e-5,
    "write_positive_idle": 0.5,
    "write_negative_idle": 0.5,
    "read_idle": 1e-3,
}


def _sync_parameter_aliases():
    global BASE_V, WRITE_POSITIVE_V, WRITE_NEGATIVE_V, READ_V
    global WRITE_BASE_DWELL, WIDTH_MULTIPLIERS, WRITE_WIDTHS
    global PWM_REPEAT_COUNT, READ_DWELL
    global WRITE_POSITIVE_TRF, WRITE_NEGATIVE_TRF, READ_TRF
    global WRITE_POSITIVE_IDLE, WRITE_NEGATIVE_IDLE, READ_IDLE

    BASE_V = float(params["base_v"])
    WRITE_POSITIVE_V = float(params["write_positive_v"])
    WRITE_NEGATIVE_V = float(params["write_negative_v"])
    READ_V = float(params["read_v"])
    WRITE_BASE_DWELL = float(params["write_base_dwell"])
    WIDTH_MULTIPLIERS = list(params["width_multipliers"])
    WRITE_WIDTHS = [WRITE_BASE_DWELL * float(value) for value in WIDTH_MULTIPLIERS]
    PWM_REPEAT_COUNT = int(params["repeat_count"])
    READ_DWELL = float(params["read_dwell"])
    WRITE_POSITIVE_TRF = float(params["write_positive_trf"])
    WRITE_NEGATIVE_TRF = float(params["write_negative_trf"])
    READ_TRF = float(params["read_trf"])
    WRITE_POSITIVE_IDLE = float(params["write_positive_idle"])
    WRITE_NEGATIVE_IDLE = float(params["write_negative_idle"])
    READ_IDLE = float(params["read_idle"])


_sync_parameter_aliases()

FULL_PWM_SEQ_ID = 1
MAX_SEGMENTS_PER_SEQ = MAX_SEGMENTS_PER_SEQUENCE

PREVIEW_ONLY = False
SAVE_WAVEFORM_PREVIEW = False

meas_types_write = [0, 0, 0, 0]
meas_start_write = [0.0, 0.0, 0.0, 0.0]
meas_stop_write = [0.0, 0.0, 0.0, 0.0]

meas_types_read = [0, 1, 0, 0]
meas_start_read = [0.0, READ_DWELL * 0.5, 0.0, 0.0]
meas_stop_read = [0.0, READ_DWELL * 0.9, 0.0, 0.0]


def build_pulse_block(level, trf, dwell, idle):
    """Return a 4-segment base -> absolute target -> base pulse block."""
    start_v = [BASE_V, level, level, BASE_V]
    stop_v = [level, level, BASE_V, BASE_V]
    time_values = [trf, dwell, trf, idle]
    return start_v, stop_v, time_values


def append_block(target, start_v, stop_v, time_values, meas_types, meas_start, meas_stop):
    target["start_v"].extend(start_v)
    target["stop_v"].extend(stop_v)
    target["time_values"].extend(time_values)
    target["meas_types"].extend(meas_types)
    target["meas_start"].extend(meas_start)
    target["meas_stop"].extend(meas_stop)


def make_empty_sequence_accumulator():
    return {
        "start_v": [],
        "stop_v": [],
        "time_values": [],
        "meas_types": [],
        "meas_start": [],
        "meas_stop": [],
    }


def make_pwm_sequence(seq_id):
    """Build one sequence: positive width sweep, then negative width sweep."""
    ch1 = make_empty_sequence_accumulator()
    ch2 = make_empty_sequence_accumulator()
    steps = []

    for polarity, write_voltage, trf, idle in (
        ("positive", WRITE_POSITIVE_V, WRITE_POSITIVE_TRF, WRITE_POSITIVE_IDLE),
        ("negative", WRITE_NEGATIVE_V, WRITE_NEGATIVE_TRF, WRITE_NEGATIVE_IDLE),
    ):
        for width in WRITE_WIDTHS:
            write_start_v, write_stop_v, write_times = build_pulse_block(write_voltage, trf, width, idle)
            read_start_v, read_stop_v, read_times = build_pulse_block(READ_V, READ_TRF, READ_DWELL, READ_IDLE)

            append_block(ch1, write_start_v, write_stop_v, write_times, meas_types_write, meas_start_write, meas_stop_write)
            append_block(ch1, read_start_v, read_stop_v, read_times, meas_types_read, meas_start_read, meas_stop_read)

            zero_write = [0.0] * len(write_times)
            zero_read = [0.0] * len(read_times)
            append_block(ch2, zero_write, zero_write, write_times, meas_types_write, meas_start_write, meas_stop_write)
            append_block(ch2, zero_read, zero_read, read_times, meas_types_read, meas_start_read, meas_stop_read)

            steps.append(
                {
                    "Polarity": polarity,
                    "WriteVoltage": write_voltage,
                    "WriteWidth_s": width,
                    "WidthMultiplier": width / WRITE_BASE_DWELL,
                }
            )

    ch1_config = (
        seq_id,
        ch1["start_v"],
        ch1["stop_v"],
        ch1["time_values"],
        ch1["meas_types"],
        ch1["meas_start"],
        ch1["meas_stop"],
    )
    ch2_config = (
        seq_id,
        ch2["start_v"],
        ch2["stop_v"],
        ch2["time_values"],
        ch2["meas_types"],
        ch2["meas_start"],
        ch2["meas_stop"],
    )
    return ch1_config, ch2_config, steps


def _rebuild_runtime_config():
    global meas_start_read, meas_stop_read
    global ch1_full_pwm_config, ch2_full_pwm_config, PWM_SINGLE_RUN_STEPS
    global PWM_STEPS, seq_configs, SEQ_LIST, CURRENT_RANGES

    _sync_parameter_aliases()
    meas_start_read = [0.0, READ_DWELL * 0.5, 0.0, 0.0]
    meas_stop_read = [0.0, READ_DWELL * 0.9, 0.0, 0.0]
    ch1_full_pwm_config, ch2_full_pwm_config, PWM_SINGLE_RUN_STEPS = make_pwm_sequence(
        FULL_PWM_SEQ_ID
    )
    PWM_STEPS = [
        {"RepeatIndex": repeat_index, **step}
        for repeat_index in range(1, PWM_REPEAT_COUNT + 1)
        for step in PWM_SINGLE_RUN_STEPS
    ]
    seq_configs = {CH1: [ch1_full_pwm_config], CH2: [ch2_full_pwm_config]}
    SEQ_LIST = {
        CH1: [(FULL_PWM_SEQ_ID, PWM_REPEAT_COUNT)],
        CH2: [(FULL_PWM_SEQ_ID, PWM_REPEAT_COUNT)],
    }
    CURRENT_RANGES = {
        CH1: float(CURRENT_RANGES.get(CH1, next(iter(CURRENT_RANGES.values())))),
        CH2: float(CURRENT_RANGES.get(CH2, next(iter(CURRENT_RANGES.values())))),
    }
    validate_segment_arb_configs(seq_configs)


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
            if len(times) > MAX_SEGMENTS_PER_SEQ:
                raise ValueError(
                    f"CH{channel} seq {seq_id} has {len(times)} segments, "
                    f"above MAX_SEGMENTS_PER_SEQ={MAX_SEGMENTS_PER_SEQ}."
                )
            if any(time_value <= 0 for time_value in times):
                raise ValueError(f"CH{channel} seq {seq_id} has non-positive segment time.")
            for index, (meas_type, start, stop, segment_time) in enumerate(zip(meas_types, meas_start, meas_stop, times), start=1):
                if meas_type == 0 and (start != 0.0 or stop != 0.0):
                    raise ValueError(f"CH{channel} seq {seq_id} segment {index} has a window but no measurement.")
                if meas_type != 0 and not (0.0 <= start < stop <= segment_time):
                    raise ValueError(f"CH{channel} seq {seq_id} segment {index} has an invalid measurement window.")

        missing_seq_ids = [
            seq_id
            for seq_id, _repeat_count in seq_list_by_channel[channel]
            if seq_id not in config_by_id
        ]
        if missing_seq_ids:
            raise ValueError(f"CH{channel} SEQ_LIST references missing seq IDs: {missing_seq_ids}")


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
    """Apply one complete workflow configuration and rebuild the PWM waveform."""
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


def preview_waveform(output_path=None, *, show=True, title_prefix="FTJ PWM CH1"):
    """Preview the full generated PWM waveform on CH1."""
    preview_config = expand_config_for_preview(ch1_full_pwm_config, PWM_REPEAT_COUNT, seq_id=0)
    return preview_sequence_configs(
        [preview_config],
        output_path,
        title_prefix=title_prefix,
        show=show,
    )


def expand_config_for_preview(config, repeat_count, *, seq_id=0):
    """Return a repeated copy for preview/export-only waveform tracing."""
    start_v = []
    stop_v = []
    time_values = []
    meas_types = []
    meas_start = []
    meas_stop = []

    for _ in range(repeat_count):
        start_v.extend(config[1])
        stop_v.extend(config[2])
        time_values.extend(config[3])
        meas_types.extend(config[4])
        meas_start.extend(config[5])
        meas_stop.extend(config[6])

    return (seq_id, start_v, stop_v, time_values, meas_types, meas_start, meas_stop)


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
    write_positive_points = []
    write_negative_points = []
    read_points = []
    cursor = 0.0

    for step in PWM_STEPS:
        write_start_v, write_stop_v, write_times = build_pulse_block(
            step["WriteVoltage"],
            WRITE_POSITIVE_TRF if step["Polarity"] == "positive" else WRITE_NEGATIVE_TRF,
            step["WriteWidth_s"],
            WRITE_POSITIVE_IDLE if step["Polarity"] == "positive" else WRITE_NEGATIVE_IDLE,
        )
        write_points = write_positive_points if step["Polarity"] == "positive" else write_negative_points
        cursor = _extend_trace_points(write_points, write_start_v, write_stop_v, write_times, cursor)

        read_start_v, read_stop_v, read_times = build_pulse_block(READ_V, READ_TRF, READ_DWELL, READ_IDLE)
        cursor = _extend_trace_points(read_points, read_start_v, read_stop_v, read_times, cursor)

    trace_columns = {
        "Time_WritePositive_s": [time for time, _voltage in write_positive_points],
        "Voltage_WritePositive_V": [voltage for _time, voltage in write_positive_points],
        "Time_WriteNegative_s": [time for time, _voltage in write_negative_points],
        "Voltage_WriteNegative_V": [voltage for _time, voltage in write_negative_points],
        "Time_Read_s": [time for time, _voltage in read_points],
        "Voltage_Read_V": [voltage for _time, voltage in read_points],
    }
    return pd.DataFrame({name: pd.Series(values) for name, values in trace_columns.items()})


def build_readback_table(df_ch1, df_ch2):
    """Return one readback row per commanded PWM width."""
    if df_ch1 is None or df_ch2 is None or df_ch1.empty or df_ch2.empty:
        raise ValueError("PWM run returned empty data.")

    expected_count = len(PWM_STEPS)
    actual_counts = (len(df_ch1), len(df_ch2))
    if actual_counts != (expected_count, expected_count):
        raise ValueError(
            "PWM readback count mismatch: "
            f"expected {expected_count}, got CH{CH1}={actual_counts[0]} "
            f"and CH{CH2}={actual_counts[1]}."
        )
    count = expected_count
    step_df = pd.DataFrame(PWM_STEPS[:count])
    pwm_df = pd.DataFrame(
        {
            "TimestampI1": df_ch1[f"Timestamp {CH1}"].values[:count],
            "TimestampI2": df_ch2[f"Timestamp {CH2}"].values[:count],
            "ReadVoltageI1": df_ch1[f"Voltage {CH1}"].values[:count],
            "ReadVoltageI2": df_ch2[f"Voltage {CH2}"].values[:count],
            "CurrentI1": df_ch1[f"Current {CH1}"].values[:count],
            "CurrentI2": df_ch2[f"Current {CH2}"].values[:count],
        }
    )
    pwm_df = pd.concat([step_df, pwm_df], axis=1)
    pwm_df["ReadVoltageDiff"] = pwm_df["ReadVoltageI1"] - pwm_df["ReadVoltageI2"]
    pwm_df["ResistanceI1"] = pwm_df["ReadVoltageI1"] / pwm_df["CurrentI1"].replace(0, pd.NA)
    pwm_df["ResistanceI2"] = pwm_df["ReadVoltageDiff"] / (-pwm_df["CurrentI2"]).replace(0, pd.NA)
    return pwm_df


def run_ftj_test(*, save_results=True, save_dir=None, file_stem=None):
    """Run the FTJ PWM width sweep and optionally save data."""
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

    pwm_df = build_readback_table(df_ch1, df_ch2)
    waveform_df = build_waveform_trace_table()
    output_path = None
    preview_path = None
    if save_results:
        save_dir.mkdir(parents=True, exist_ok=True)
        output_stem = reserve_output_stem(
            save_dir, measurement_name(file_stem, params["write_positive_v"], "tw" + time_tag(params["write_base_dwell"])),
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
            pwm_df.to_excel(writer, sheet_name="PWM_ReadOnly", index=False)
            df_ch1.to_excel(writer, sheet_name="Channel_1_ReadOnly", index=False)
            df_ch2.to_excel(writer, sheet_name="Channel_2_ReadOnly", index=False)
            waveform_df.to_excel(writer, sheet_name="Waveform", index=False)
            params_df.to_excel(writer, sheet_name="Parameters", index=False)

        if SAVE_WAVEFORM_PREVIEW:
            preview_path = Path(f"{output_stem}_waveform.png")
            preview_waveform(preview_path)

    return {
        "df_ch1": df_ch1,
        "df_ch2": df_ch2,
        "pwm_df": pwm_df,
        "waveform_df": waveform_df,
        "output_path": output_path,
        "preview_path": preview_path,
    }


def main():
    """Run the FTJ PWM width sweep and save readback data."""
    if PREVIEW_ONLY:
        preview_waveform()
        return
    run_ftj_test()


if __name__ == "__main__":
    main()
