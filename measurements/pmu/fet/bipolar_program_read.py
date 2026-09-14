# -*- coding: utf-8 -*-
"""Dual-polarity FeFET program/read test using one synchronized Segment Arb run."""

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

from keithley4200.output import measurement_name, reserve_output_stem, voltage_tag, time_tag
from keithley4200.pmu.fet_three_terminal_common import (
    add_derived_fet_columns,
    execute_program_read_with_software_delay,
    report_last_error,
    save_fet_workbook,
    save_waveform_data_sheets,
    source_smu_off,
    source_smu_on,
)
from keithley4200.tools.waveform_preview import (
    preview_sequence_configs,
    save_ids_dual_axis_plot,
    sequence_configs_to_dataframe,
)
from keithley4200.pmu.data_processing import read_both_channels
from keithley4200.pmu.pmu_tests import (
    MAX_SEGMENTS_PER_SEQUENCE, execute_segARB_test, power_off_outputs,
)
from keithley4200.pmu.session import PMUSession


INST = "TCPIP0::129.125.87.80::1225::SOCKET"
GATE_CH, DRAIN_CH, SOURCE_SMU = 1, 2, 3
PREVIEW_ONLY = False
SAVE_WAVEFORM_PREVIEW = False
# Optional semantic filename supplied by wrapper scripts. None keeps the
# standalone parameter-based filename behavior.
OUTPUT_TAG = None

# FTJs-style experiment parameters.
CURRENT_RANGES = {GATE_CH: 1e-7, DRAIN_CH: 1e-6}
SEQ_CYCLE_COUNT = 3

# Positive and negative write trains are configured independently. Each train
# is followed by the same READ_DELAY + read pulse defined below.
POS_WRITE_VOLTAGE = 4
POS_PULSES_PER_TRAIN = 1
POS_PROGRAM_READ_REPEATS = 20
POS_WRITE_BASE = 0.0
POS_WRITE_RISE = 1e-7
POS_WRITE_DWELL = 1e-4
POS_WRITE_FALL = 1e-7
POS_WRITE_IDLE = 1e-5

NEG_WRITE_VOLTAGE = -4
NEG_PULSES_PER_TRAIN = 1
NEG_PROGRAM_READ_REPEATS = 20
NEG_WRITE_BASE = 0.0
NEG_WRITE_RISE = 1e-7
NEG_WRITE_DWELL = 1e-4
NEG_WRITE_FALL = 1e-7
NEG_WRITE_IDLE = 1e-5

READ_GATE_V = 0.5
READ_DRAIN_V = -2
READ_BASE = 0.0
READ_DELAY = 5
READ_RISE = 1e-4
READ_DWELL = 1e-3
READ_FALL = 1e-4
READ_IDLE = 1e-4
READ_MEAS_START = 0.2
READ_MEAS_STOP = 0.9

SOURCE_V = 0.0
SOURCE_COMPLIANCE = 1e-3
# False: connect Source to GNDU FORCE and leave SMU3 physically disconnected.
# True: connect Source to SMU3 only; do not connect it to GNDU at the same time.
USE_SOURCE_SMU = False

ENABLE_LLEC = False
LLEC_CHANNELS = (DRAIN_CH,)
ENABLE_LOAD_CONFIG = False
LOAD_RESISTANCES = {GATE_CH: 1e6, DRAIN_CH: 1e6}
ENABLE_CONNECTION_COMP = False
CONNECTION_COMP_CHANNELS = (GATE_CH, DRAIN_CH)

# Internal mapping retained so saving and shared builders use one snapshot.
PARAMS = {
    "cycles": SEQ_CYCLE_COUNT,
    "positive_write_voltage": POS_WRITE_VOLTAGE,
    "positive_train_count": POS_PULSES_PER_TRAIN,
    "positive_program_read_repeats": POS_PROGRAM_READ_REPEATS,
    "positive_write_base": POS_WRITE_BASE,
    "positive_write_rise": POS_WRITE_RISE,
    "positive_write_plateau": POS_WRITE_DWELL,
    "positive_write_fall": POS_WRITE_FALL,
    "positive_write_rest": POS_WRITE_IDLE,
    "negative_write_voltage": NEG_WRITE_VOLTAGE,
    "negative_train_count": NEG_PULSES_PER_TRAIN,
    "negative_program_read_repeats": NEG_PROGRAM_READ_REPEATS,
    "negative_write_base": NEG_WRITE_BASE,
    "negative_write_rise": NEG_WRITE_RISE,
    "negative_write_plateau": NEG_WRITE_DWELL,
    "negative_write_fall": NEG_WRITE_FALL,
    "negative_write_rest": NEG_WRITE_IDLE,
    "read_gate_voltage": READ_GATE_V,
    "read_drain_voltage": READ_DRAIN_V,
    "read_base": READ_BASE,
    "read_delay": READ_DELAY,
    "read_rise": READ_RISE,
    "read_plateau": READ_DWELL,
    "read_fall": READ_FALL,
    "read_rest": READ_IDLE,
    "measure_start_fraction": READ_MEAS_START,
    "measure_stop_fraction": READ_MEAS_STOP,
    "gate_current_range": CURRENT_RANGES[GATE_CH],
    "drain_current_range": CURRENT_RANGES[DRAIN_CH],
    "source_voltage": SOURCE_V,
    "source_compliance": SOURCE_COMPLIANCE,
    "max_segments_per_sequence": MAX_SEGMENTS_PER_SEQUENCE,
}

SAVE_DIR = Path(r"C:\Users\P317151\Documents\data\06-08-2026\03C5_FET\FeFET 2\D40-5um gap circular 2\Vg0.5 Vd-2")


def append_pulse(arrays, base, level, timing, measure=False):
    """Append seamless rise, plateau, fall, and rest segments."""
    rise, plateau, fall, rest = timing
    arrays["start"].extend([base, level, level, base])
    arrays["stop"].extend([level, level, base, base])
    arrays["time"].extend([rise, plateau, fall, rest])
    arrays["mode"].extend([0, 1 if measure else 0, 0, 0])
    if measure:
        m_start = plateau * PARAMS["measure_start_fraction"]
        m_stop = plateau * PARAMS["measure_stop_fraction"]
    else:
        m_start = m_stop = 0.0
    arrays["meas_start"].extend([0.0, m_start, 0.0, 0.0])
    arrays["meas_stop"].extend([0.0, m_stop, 0.0, 0.0])


def append_constant_delay(arrays, level, duration):
    """Append an unmeasured delay, split at the 1 s segment-time limit."""
    remaining = float(duration)
    while remaining > 0:
        segment_time = min(remaining, 1.0)
        arrays["start"].append(level)
        arrays["stop"].append(level)
        arrays["time"].append(segment_time)
        arrays["mode"].append(0)
        arrays["meas_start"].append(0.0)
        arrays["meas_stop"].append(0.0)
        remaining -= segment_time


def build_program_read_sequence(include_delay=True):
    """Build independent positive/negative trains, each followed by a read."""
    p = PARAMS
    plan = []
    configs = {GATE_CH: [], DRAIN_CH: []}
    execution_list = []
    read_timing = (p["read_rise"], p["read_plateau"], p["read_fall"], p["read_rest"])

    write_states = (
        {
            "name": "Positive",
            "voltage": p["positive_write_voltage"],
            "count": int(p["positive_train_count"]),
            "repeats": int(p["positive_program_read_repeats"]),
            "base": p["positive_write_base"],
            "timing": (
                p["positive_write_rise"], p["positive_write_plateau"],
                p["positive_write_fall"], p["positive_write_rest"],
            ),
        },
        {
            "name": "Negative",
            "voltage": p["negative_write_voltage"],
            "count": int(p["negative_train_count"]),
            "repeats": int(p["negative_program_read_repeats"]),
            "base": p["negative_write_base"],
            "timing": (
                p["negative_write_rise"], p["negative_write_plateau"],
                p["negative_write_fall"], p["negative_write_rest"],
            ),
        },
    )

    # Define one four-segment sequence for each polarity. CH2 gets a matching
    # zero-volt hold waveform so both PMU channels use the same time stamps.
    program_sequence_ids = []
    for sequence_id, state in enumerate(write_states, start=1):
        program_sequence_ids.append((sequence_id, state))
        for channel, base, level in (
            (GATE_CH, state["base"], state["voltage"]),
            (DRAIN_CH, p["read_base"], p["read_base"]),
        ):
            arrays = {
                key: [] for key in ("start", "stop", "time", "mode", "meas_start", "meas_stop")
            }
            append_pulse(arrays, base, level, state["timing"], measure=False)
            configs[channel].append((
                sequence_id,
                arrays["start"], arrays["stop"], arrays["time"], arrays["mode"],
                arrays["meas_start"], arrays["meas_stop"],
            ))

    # A shared four-segment read sequence follows each completed train.
    read_sequence_id = len(program_sequence_ids) + 1
    for channel, base, level in (
        (GATE_CH, p["read_base"], p["read_gate_voltage"]),
        (DRAIN_CH, p["read_base"], p["read_drain_voltage"]),
    ):
        arrays = {
            key: [] for key in ("start", "stop", "time", "mode", "meas_start", "meas_stop")
        }
        if include_delay:
            append_constant_delay(arrays, base, p["read_delay"])
        append_pulse(arrays, base, level, read_timing, measure=True)
        configs[channel].append((
            read_sequence_id,
            arrays["start"], arrays["stop"], arrays["time"], arrays["mode"],
            arrays["meas_start"], arrays["meas_stop"],
        ))

    # Both channels receive this exact same ordered list. Looping the program
    # sequence produces N pulses without defining 4*N segments.
    for cycle in range(1, int(p["cycles"]) + 1):
        for program_sequence_id, state in program_sequence_ids:
            for polarity_repeat in range(1, state["repeats"] + 1):
                execution_list.extend([
                    (program_sequence_id, state["count"]),
                    (read_sequence_id, 1),
                ])
                plan.append(
                    {
                        "Cycle": cycle,
                        "WritePolarity": state["name"],
                        "PolarityRepeat": polarity_repeat,
                        "ProgramVoltage": state["voltage"],
                        "TrainCount": state["count"],
                        "ProgramSequenceID": program_sequence_id,
                    }
                )

    for channel_configs in configs.values():
        for config in channel_configs:
            segment_count = len(config[3])
            if segment_count > int(p["max_segments_per_sequence"]):
                raise ValueError(
                    f"Sequence {config[0]} needs {segment_count} segments, above "
                    f"max_segments_per_sequence={p['max_segments_per_sequence']}."
                )
            if any(duration > 1.0 for duration in config[3]):
                raise ValueError(f"Sequence {config[0]} contains a segment longer than 1 s.")
    seq_lists = {
        GATE_CH: list(execution_list),
        DRAIN_CH: list(execution_list),
    }
    return pd.DataFrame(plan), configs, seq_lists


def validate_parameters():
    if int(PARAMS["cycles"]) < 1:
        raise ValueError("cycles must be a positive integer.")
    for name in (
        "positive_train_count", "negative_train_count",
        "positive_program_read_repeats", "negative_program_read_repeats",
    ):
        if int(PARAMS[name]) < 1:
            raise ValueError(f"{name} must be a positive integer.")
    if PARAMS["read_delay"] < 0:
        raise ValueError("READ_DELAY must be non-negative.")
    for name in (
        "positive_write_rise", "positive_write_plateau",
        "positive_write_fall", "positive_write_rest",
        "negative_write_rise", "negative_write_plateau",
        "negative_write_fall", "negative_write_rest",
        "read_rise", "read_plateau", "read_fall", "read_rest",
    ):
        if PARAMS[name] <= 0:
            raise ValueError(f"{name} must be positive.")
    if not 0 <= PARAMS["measure_start_fraction"] < PARAMS["measure_stop_fraction"] <= 1:
        raise ValueError("Measurement fractions must satisfy 0 <= start < stop <= 1.")


def build_sequence_plan_preview_config(configs, sequence_plan, sequence_id):
    """Expand hardware loops into one config solely for waveform preview."""
    config_by_id = {config[0]: config for config in configs}
    arrays = [[] for _ in range(6)]
    for seq_id, repeat_count in sequence_plan:
        config = config_by_id[seq_id]
        for _ in range(int(repeat_count)):
            for target, source in zip(arrays, config[1:7]):
                target.extend(source)
    return (sequence_id, *arrays)


def expanded_preview_configs():
    """Return fully expanded Gate and Drain configs for plotting/export."""
    _plan, seq_configs, seq_lists = build_program_read_sequence()
    gate_preview = build_sequence_plan_preview_config(
        seq_configs[GATE_CH], seq_lists[GATE_CH], GATE_CH
    )
    drain_preview = build_sequence_plan_preview_config(
        seq_configs[DRAIN_CH], seq_lists[DRAIN_CH], DRAIN_CH
    )
    return [gate_preview, drain_preview]


def preview_waveform(output_path=None, *, compress_delay=True):
    """Preview the expanded plan using compressed or real-time delay scaling."""
    return preview_sequence_configs(
        expanded_preview_configs(),
        output_path,
        title_prefix="FeFET program/read",
        channel_labels=("CH1 Gate", "CH2 Drain"),
        compress_constant_segments_above=0.1 if compress_delay else None,
    )


def waveform_data_frames():
    """Return real-time and compressed numeric waveform plotting tables."""
    configs = expanded_preview_configs()
    labels = ("CH1", "CH2")
    return (
        sequence_configs_to_dataframe(configs, channel_labels=labels),
        sequence_configs_to_dataframe(
            configs,
            channel_labels=labels,
            compress_constant_segments_above=0.1,
        ),
    )


def main():
    validate_parameters()
    split_delay = PARAMS["read_delay"] > 1.0
    plan, seq_configs, seq_lists = build_program_read_sequence(
        include_delay=not split_delay
    )
    defined_segments = sum(len(config[3]) for config in seq_configs[GATE_CH])
    print(
        f"Segment Arb FeFET plan: {len(plan)} program/read states, "
        f"{defined_segments} defined segments/channel; gate trains use sequence loops"
    )
    print(
        "Delay execution: "
        + ("separate write/read executions" if split_delay else "one synchronized execution")
    )
    print(plan.to_string(index=False))
    if PREVIEW_ONLY:
        preview_waveform()
        return

    with PMUSession(INST, channels=(GATE_CH, DRAIN_CH)) as session:
        query = session.query
        try:
            query(":ERROR:LAST:CLEAR")
            if USE_SOURCE_SMU:
                source_smu_on(
                    query,
                    SOURCE_SMU,
                    PARAMS["source_voltage"],
                    PARAMS["source_compliance"],
                )
            current_ranges = {
                GATE_CH: PARAMS["gate_current_range"],
                DRAIN_CH: PARAMS["drain_current_range"],
            }
            options = {
                "ENABLE_LLEC": ENABLE_LLEC,
                "LLEC_CHANNELS": LLEC_CHANNELS,
                "ENABLE_LOAD_CONFIG": ENABLE_LOAD_CONFIG,
                "LOAD_RESISTANCES": LOAD_RESISTANCES,
                "ENABLE_CONNECTION_COMP": ENABLE_CONNECTION_COMP,
                "CONNECTION_COMP_CHANNELS": CONNECTION_COMP_CHANNELS,
            }
            if split_delay:
                gate_df, drain_df = execute_program_read_with_software_delay(
                    query, plan, seq_configs, PARAMS["read_delay"],
                    GATE_CH, DRAIN_CH, current_ranges, options,
                )
            else:
                execute_segARB_test(
                    query, [GATE_CH, DRAIN_CH], seq_configs,
                    seq_list=seq_lists, current_ranges=current_ranges, options=options,
                )
                gate_df, drain_df = read_both_channels(query, GATE_CH, DRAIN_CH)
            if gate_df is None or drain_df is None or gate_df.empty or drain_df.empty:
                report_last_error(query)
                raise ValueError("Segment Arb FeFET read returned empty PMU data.")
            measured = pd.concat(
                [gate_df.reset_index(drop=True), drain_df.reset_index(drop=True)], axis=1
            )
            if len(measured) != len(plan):
                raise ValueError(
                    f"Expected {len(plan)} drain-read points, received {len(measured)}."
                )
            data = add_derived_fet_columns(
                pd.concat([plan, measured], axis=1),
                gate_channel=GATE_CH, drain_channel=DRAIN_CH,
            )
            name = measurement_name(
                "FETdual", PARAMS["positive_write_voltage"],
                "Vg" + voltage_tag(PARAMS["read_gate_voltage"]),
                "Vd" + voltage_tag(PARAMS["read_drain_voltage"]),
                "td" + time_tag(PARAMS["read_delay"]),
            )
            if OUTPUT_TAG:
                name += "_" + OUTPUT_TAG
            output_stem = reserve_output_stem(SAVE_DIR, name)
            path = Path(f"{output_stem}.xlsx")
            save_fet_workbook(
                path,
                data,
                {
                    "mode": "dual_polarity_segment_arb_program_then_read",
                    "delay_execution": "software_split" if split_delay else "single_execution",
                    "source_connection": "SMU3" if USE_SOURCE_SMU else "GNDU",
                    **PARAMS,
                    "CURRENT_RANGES": CURRENT_RANGES,
                    "ENABLE_LLEC": ENABLE_LLEC,
                    "LLEC_CHANNELS": LLEC_CHANNELS,
                    "ENABLE_LOAD_CONFIG": ENABLE_LOAD_CONFIG,
                    "LOAD_RESISTANCES": LOAD_RESISTANCES,
                    "ENABLE_CONNECTION_COMP": ENABLE_CONNECTION_COMP,
                    "CONNECTION_COMP_CHANNELS": CONNECTION_COMP_CHANNELS,
                },
                {GATE_CH: gate_df, DRAIN_CH: drain_df},
            )
            try:
                save_ids_dual_axis_plot(
                    expanded_preview_configs(),
                    data["Id"],
                    path.with_name(f"{path.stem}_Ids.png"),
                    read_gate_voltage=PARAMS["read_gate_voltage"],
                    read_drain_voltage=PARAMS["read_drain_voltage"],
                    read_delay=PARAMS["read_delay"],
                )
            except Exception as exc:
                print(f"Warning: Ids plot could not be saved; measurement data is safe: {exc}")
            if SAVE_WAVEFORM_PREVIEW:
                real_time_data, schematic_data = waveform_data_frames()
                save_waveform_data_sheets(path, real_time_data, schematic_data)
                real_time_path = path.with_name(f"{path.stem}_waveform_real_time.png")
                schematic_path = path.with_name(f"{path.stem}_waveform_compressed.png")
                preview_waveform(
                    real_time_path,
                    compress_delay=False,
                )
                preview_waveform(
                    schematic_path,
                    compress_delay=True,
                )
            print(f"Saved: {path.resolve()}")
            return path
        finally:
            power_off_outputs(query, (GATE_CH, DRAIN_CH))
            if USE_SOURCE_SMU:
                source_smu_off(query, SOURCE_SMU)


if __name__ == "__main__":
    main()
