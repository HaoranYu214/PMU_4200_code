# -*- coding: utf-8 -*-
"""FTJ ISPP V2: pack each complete voltage ladder into one sequence."""

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

from keithley4200.output import measurement_name, reserve_output_stem, time_tag, saved_at
from keithley4200.pmu.data_processing import read_both_channels
from keithley4200.pmu.pmu_tests import (
    MAX_SEGMENTS_PER_SEQUENCE,
    execute_segARB_test,
    power_off_outputs,
    validate_segment_arb_configs,
)
from keithley4200.pmu.session import PMUSession
from keithley4200.tools.waveform_preview import preview_sequence_configs


INST = "TCPIP0::129.125.87.80::1225::SOCKET"
CH1, CH2 = 1, 2
SAVE_DIR = Path(r"C:\Users\P317151\Documents\data\FTJ\ISPP_V2")
FILE_STEM = "ISPP2"

CURRENT_RANGES = {CH1: 1e-5, CH2: 1e-5}
SEGARB_OPTIONS = {
    "ENABLE_CONNECTION_COMP": False,
    "ENABLE_LOAD_CONFIG": False,
    "LOAD_RESISTANCE": 1e6,
    "ENABLE_LLEC": False,
}

params = {
    # Voltage held during pre-delay and after every write/read pulse.
    "base_v": 0.0,
    "read_v": 0.1,
    "positive_start_v": 0.5,
    "positive_stop_v": 2.0,
    "positive_steps": 10,
    "negative_start_v": -0.5,
    "negative_stop_v": -2.0,
    "negative_steps": 10,
    "write_dwell": 1e-6,
    "read_dwell": 1e-5,
    "pre_delay": 1e-3,
    "rise_time": 2e-7,
    "fall_time": 2e-7,
    "idle_time": 1e-3,
}

WRITE_POSITIVE_SEQ_ID = 1
WRITE_NEGATIVE_SEQ_ID = 2
MAX_SEGMENTS_PER_SEQ = MAX_SEGMENTS_PER_SEQUENCE
PREVIEW_ONLY = True


def voltage_steps(start, stop, steps):
    if int(steps) <= 0:
        raise ValueError("ISPP step count must be positive.")
    if int(steps) == 1:
        return [float(stop)]
    step = (float(stop) - float(start)) / (int(steps) - 1)
    return [float(start) + index * step for index in range(int(steps))]


def make_ispp_sequence(seq_id, voltages):
    arrays = {
        "ch1_start": [], "ch1_stop": [], "ch2_start": [], "ch2_stop": [],
        "times": [], "types": [], "starts": [], "stops": [],
    }
    for voltage in voltages:
        arrays["ch1_start"].extend([BASE_V, BASE_V, voltage, voltage, BASE_V])
        arrays["ch1_stop"].extend([BASE_V, voltage, voltage, BASE_V, BASE_V])
        arrays["ch2_start"].extend([0.0] * 5)
        arrays["ch2_stop"].extend([0.0] * 5)
        arrays["times"].extend(time_values_write)
        arrays["types"].extend(meas_types_write)
        arrays["starts"].extend(meas_start_write)
        arrays["stops"].extend(meas_stop_write)

        arrays["ch1_start"].extend([BASE_V, BASE_V, READ_V, READ_V, BASE_V])
        arrays["ch1_stop"].extend([BASE_V, READ_V, READ_V, BASE_V, BASE_V])
        arrays["ch2_start"].extend([0.0] * 5)
        arrays["ch2_stop"].extend([0.0] * 5)
        arrays["times"].extend(time_values_read)
        arrays["types"].extend(meas_types_read)
        arrays["starts"].extend(meas_start_read)
        arrays["stops"].extend(meas_stop_read)

    if len(arrays["times"]) > MAX_SEGMENTS_PER_SEQ:
        raise ValueError(
            f"ISPP seq {seq_id} has {len(arrays['times'])} segments; "
            f"limit is {MAX_SEGMENTS_PER_SEQ}."
        )
    ch1_config = (
        seq_id, arrays["ch1_start"], arrays["ch1_stop"], arrays["times"],
        arrays["types"], arrays["starts"], arrays["stops"],
    )
    ch2_config = (
        seq_id, arrays["ch2_start"], arrays["ch2_stop"], arrays["times"],
        arrays["types"], arrays["starts"], arrays["stops"],
    )
    return ch1_config, ch2_config


def _rebuild_runtime_config():
    global BASE_V, READ_V, POS_V_START, POS_V_STOP, POS_STEPS
    global NEG_V_START, NEG_V_STOP, NEG_STEPS
    global WRITE_DWELL, READ_DWELL, PRE_DELAY, RISE_TIME, FALL_TIME, IDLE_TIME
    global time_values_write, time_values_read
    global meas_types_write, meas_start_write, meas_stop_write
    global meas_types_read, meas_start_read, meas_stop_read
    global POS_VOLTAGES, NEG_VOLTAGES
    global ch1_pos_config, ch2_pos_config, ch1_neg_config, ch2_neg_config
    global seq_configs, SEQ_PLAN, SEQ_LIST, CURRENT_RANGES

    BASE_V = float(params["base_v"])
    READ_V = float(params["read_v"])
    POS_V_START = float(params["positive_start_v"])
    POS_V_STOP = float(params["positive_stop_v"])
    POS_STEPS = int(params["positive_steps"])
    NEG_V_START = float(params["negative_start_v"])
    NEG_V_STOP = float(params["negative_stop_v"])
    NEG_STEPS = int(params["negative_steps"])
    WRITE_DWELL = float(params["write_dwell"])
    READ_DWELL = float(params["read_dwell"])
    PRE_DELAY = float(params["pre_delay"])
    RISE_TIME = float(params["rise_time"])
    FALL_TIME = float(params["fall_time"])
    IDLE_TIME = float(params["idle_time"])

    time_values_write = [PRE_DELAY, RISE_TIME, WRITE_DWELL, FALL_TIME, IDLE_TIME]
    time_values_read = [PRE_DELAY, RISE_TIME, READ_DWELL, FALL_TIME, IDLE_TIME]
    # V2 intentionally measures both write and read dwell segments.
    meas_types_write = [0, 0, 1, 0, 0]
    meas_start_write = [0.0, 0.0, WRITE_DWELL * 0.5, 0.0, 0.0]
    meas_stop_write = [0.0, 0.0, WRITE_DWELL * 0.9, 0.0, 0.0]
    meas_types_read = [0, 0, 1, 0, 0]
    meas_start_read = [0.0, 0.0, READ_DWELL * 0.5, 0.0, 0.0]
    meas_stop_read = [0.0, 0.0, READ_DWELL * 0.9, 0.0, 0.0]

    POS_VOLTAGES = voltage_steps(POS_V_START, POS_V_STOP, POS_STEPS)
    NEG_VOLTAGES = voltage_steps(NEG_V_START, NEG_V_STOP, NEG_STEPS)
    ch1_pos_config, ch2_pos_config = make_ispp_sequence(
        WRITE_POSITIVE_SEQ_ID, POS_VOLTAGES
    )
    ch1_neg_config, ch2_neg_config = make_ispp_sequence(
        WRITE_NEGATIVE_SEQ_ID, NEG_VOLTAGES
    )
    seq_configs = {
        CH1: [ch1_pos_config, ch1_neg_config],
        CH2: [ch2_pos_config, ch2_neg_config],
    }
    SEQ_PLAN = [(WRITE_POSITIVE_SEQ_ID, 1), (WRITE_NEGATIVE_SEQ_ID, 1)]
    SEQ_LIST = {CH1: list(SEQ_PLAN), CH2: list(SEQ_PLAN)}
    CURRENT_RANGES = {
        CH1: float(CURRENT_RANGES.get(CH1, next(iter(CURRENT_RANGES.values())))),
        CH2: float(CURRENT_RANGES.get(CH2, next(iter(CURRENT_RANGES.values())))),
    }
    validate_segment_arb_configs(seq_configs)


_rebuild_runtime_config()


def configure_measurement(
    *, params_override=None, inst=None, channels=None, current_ranges=None,
    segarb_options=None, save_dir=None, file_stem=None,
):
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


def preview_waveform(output_path=None, *, show=True, title_prefix="FTJ ISPP V2 CH1"):
    return preview_sequence_configs(
        [ch1_pos_config, ch1_neg_config],
        output_path,
        title_prefix=title_prefix,
        show=show,
    )


def build_waveform_trace_table():
    rows = []
    elapsed = 0.0
    for polarity, voltages in (("positive", POS_VOLTAGES), ("negative", NEG_VOLTAGES)):
        for voltage in voltages:
            rows.append(
                {
                    "Polarity": polarity,
                    "WriteVoltage_V": voltage,
                    "ReadVoltage_V": READ_V,
                    "WriteDwell_s": WRITE_DWELL,
                    "ReadDwell_s": READ_DWELL,
                    "BlockStart_s": elapsed,
                }
            )
            elapsed += sum(time_values_write) + sum(time_values_read)
    return pd.DataFrame(rows)


def run_ftj_test(*, save_results=True, save_dir=None, file_stem=None):
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
        raise ValueError("No data returned from the FTJ ISPP V2 run.")

    waveform_df = build_waveform_trace_table()
    output_path = None
    if save_results:
        save_dir.mkdir(parents=True, exist_ok=True)
        output_stem = reserve_output_stem(
            save_dir, measurement_name(file_stem, params["positive_stop_v"], "tw" + time_tag(params["write_dwell"])),
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
    return {
        "df_ch1": df_ch1,
        "df_ch2": df_ch2,
        "waveform_df": waveform_df,
        "output_path": output_path,
    }


def main():
    if PREVIEW_ONLY:
        return preview_waveform()
    return run_ftj_test()


if __name__ == "__main__":
    main()
