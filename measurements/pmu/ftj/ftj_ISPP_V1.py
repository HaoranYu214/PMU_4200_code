# -*- coding: utf-8 -*-
"""FTJ ISPP V1: one write sequence per voltage, followed by a common read."""

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
    execute_segARB_test,
    power_off_outputs,
    validate_segment_arb_configs,
)
from keithley4200.pmu.session import PMUSession
from keithley4200.tools.waveform_preview import preview_sequence_configs


INST = "TCPIP0::129.125.87.80::1225::SOCKET"
CH1, CH2 = 1, 2
SAVE_DIR = Path(r"C:\Users\P317151\Documents\data\07-07-2026\FTJ")
FILE_STEM = "ISPP1"

CURRENT_RANGES = {CH1: 1e-5, CH2: 1e-5}
SEGARB_OPTIONS = {
    "ENABLE_CONNECTION_COMP": False,
    "ENABLE_LOAD_CONFIG": False,
    "LOAD_RESISTANCE": 1e6,
    # LLEC is not available for Segment Arb measurements.
    "ENABLE_LLEC": False,
}

params = {
    # Voltage held during pre-delay and after every write/read pulse.
    "base_v": 0.0,
    "read_v": 0.1,
    "positive_start_v": 0.5,
    "positive_stop_v": 2.0,
    "positive_steps": 15,
    "negative_start_v": -0.5,
    "negative_stop_v": -2.0,
    "negative_steps": 15,
    "write_dwell": 1e-6,
    "read_dwell": 1e-5,
    "pre_delay": 1e-3,
    "rise_time": 2e-7,
    "fall_time": 2e-7,
    "idle_time": 0.1,
}

READ_SEQ_ID = 1
WRITE_POSITIVE_SEQ_START_ID = 2
PREVIEW_ONLY = True


def voltage_steps(start, stop, steps):
    if int(steps) <= 0:
        raise ValueError("ISPP step count must be positive.")
    if int(steps) == 1:
        return [float(stop)]
    step = (float(stop) - float(start)) / (int(steps) - 1)
    return [float(start) + index * step for index in range(int(steps))]


def make_read_configs():
    ch1_start_v = [BASE_V, BASE_V, READ_V, READ_V, BASE_V]
    ch1_stop_v = [BASE_V, READ_V, READ_V, BASE_V, BASE_V]
    zeros = [0.0] * 5
    ch1_config = (
        READ_SEQ_ID, ch1_start_v, ch1_stop_v, time_values_read,
        meas_types_read, meas_start_read, meas_stop_read,
    )
    ch2_config = (
        READ_SEQ_ID, zeros.copy(), zeros.copy(), time_values_read,
        meas_types_read, meas_start_read, meas_stop_read,
    )
    return ch1_config, ch2_config


def make_write_configs(seq_start_id, voltages):
    ch1_configs = []
    ch2_configs = []
    for index, voltage in enumerate(voltages):
        seq_id = seq_start_id + index
        ch1_start_v = [BASE_V, BASE_V, voltage, voltage, BASE_V]
        ch1_stop_v = [BASE_V, voltage, voltage, BASE_V, BASE_V]
        zeros = [0.0] * 5
        ch1_configs.append(
            (
                seq_id, ch1_start_v, ch1_stop_v, time_values_write,
                meas_types_write, meas_start_write, meas_stop_write,
            )
        )
        ch2_configs.append(
            (
                seq_id, zeros.copy(), zeros.copy(), time_values_write,
                meas_types_write, meas_start_write, meas_stop_write,
            )
        )
    return ch1_configs, ch2_configs


def make_ispp_plan(seq_start_id, steps):
    return [
        item
        for index in range(int(steps))
        for item in ((seq_start_id + index, 1), (READ_SEQ_ID, 1))
    ]


def _rebuild_runtime_config():
    global BASE_V, READ_V, POS_V_START, POS_V_STOP, POS_STEPS
    global NEG_V_START, NEG_V_STOP, NEG_STEPS
    global WRITE_DWELL, READ_DWELL, PRE_DELAY, RISE_TIME, FALL_TIME, IDLE_TIME
    global WRITE_NEGATIVE_SEQ_START_ID, time_values_write, time_values_read
    global meas_types_write, meas_start_write, meas_stop_write
    global meas_types_read, meas_start_read, meas_stop_read
    global POS_VOLTAGES, NEG_VOLTAGES, ch1_read_config, ch2_read_config
    global ch1_pos_configs, ch2_pos_configs, ch1_neg_configs, ch2_neg_configs
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
    WRITE_NEGATIVE_SEQ_START_ID = WRITE_POSITIVE_SEQ_START_ID + POS_STEPS

    time_values_write = [PRE_DELAY, RISE_TIME, WRITE_DWELL, FALL_TIME, IDLE_TIME]
    time_values_read = [PRE_DELAY, RISE_TIME, READ_DWELL, FALL_TIME, IDLE_TIME]
    meas_types_write = [0, 0, 0, 0, 0]
    meas_start_write = [0.0] * 5
    meas_stop_write = [0.0] * 5
    meas_types_read = [0, 0, 1, 0, 0]
    meas_start_read = [0.0, 0.0, READ_DWELL * 0.5, 0.0, 0.0]
    meas_stop_read = [0.0, 0.0, READ_DWELL * 0.9, 0.0, 0.0]

    POS_VOLTAGES = voltage_steps(POS_V_START, POS_V_STOP, POS_STEPS)
    NEG_VOLTAGES = voltage_steps(NEG_V_START, NEG_V_STOP, NEG_STEPS)
    ch1_read_config, ch2_read_config = make_read_configs()
    ch1_pos_configs, ch2_pos_configs = make_write_configs(
        WRITE_POSITIVE_SEQ_START_ID, POS_VOLTAGES
    )
    ch1_neg_configs, ch2_neg_configs = make_write_configs(
        WRITE_NEGATIVE_SEQ_START_ID, NEG_VOLTAGES
    )
    seq_configs = {
        CH1: [ch1_read_config] + ch1_pos_configs + ch1_neg_configs,
        CH2: [ch2_read_config] + ch2_pos_configs + ch2_neg_configs,
    }
    SEQ_PLAN = (
        make_ispp_plan(WRITE_POSITIVE_SEQ_START_ID, POS_STEPS)
        + make_ispp_plan(WRITE_NEGATIVE_SEQ_START_ID, NEG_STEPS)
    )
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


def preview_waveform(output_path=None, *, show=True, title_prefix="FTJ ISPP V1 CH1"):
    return preview_sequence_configs(
        [ch1_read_config] + ch1_pos_configs + ch1_neg_configs,
        output_path,
        title_prefix=title_prefix,
        show=show,
    )


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
        raise ValueError("No data returned from the FTJ ISPP V1 run.")

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
            params_df.to_excel(writer, sheet_name="Parameters", index=False)

    return {"df_ch1": df_ch1, "df_ch2": df_ch2, "output_path": output_path}


def main():
    if PREVIEW_ONLY:
        return preview_waveform()
    return run_ftj_test()


if __name__ == "__main__":
    main()
