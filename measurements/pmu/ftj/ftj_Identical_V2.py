# -*- coding: utf-8 -*-
"""FTJ Identical V2: separate executions with real Python inter-pulse waits."""

from pathlib import Path
import sys
import time

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
from keithley4200.pmu.data_processing import merge_channels, read_both_channels
from keithley4200.pmu.pmu_tests import (
    execute_segARB_test,
    power_off_outputs,
    validate_segment_arb_configs,
)
from keithley4200.pmu.session import PMUSession
from keithley4200.tools.waveform_preview import preview_sequence_configs


INST = "TCPIP0::129.125.87.80::1225::SOCKET"
CH1, CH2 = 1, 2
SAVE_DIR = Path(r"C:\Users\P317151\Documents\data\FTJ\Identical_V2")
FILE_STEM = "Identical2"

CURRENT_RANGES = {CH1: 1e-4, CH2: 1e-4}
SEGARB_OPTIONS = {
    "ENABLE_CONNECTION_COMP": False,
    "ENABLE_LOAD_CONFIG": False,
    "LOAD_RESISTANCE": 1.0,
    "ENABLE_LLEC": False,
}

params = {
    # Voltage held before/after every write and read pulse.
    "base_v": 0.0,
    "write_positive_v": 1.5,
    "write_negative_v": -8,
    "read_v": -1,
    "write_positive_dwell": 1e-4,
    "write_negative_dwell": 5e-5,
    "read_dwell": 5e-5,
    "write_positive_trf": 1e-6,
    "write_negative_trf": 1e-6,
    "read_trf": 1e-6,
    "write_positive_idle": 0.1,
    "write_negative_idle": 0.1,
    "read_idle": 1e-3,
    "wait_after_positive_write_s": 1.0,
    "wait_after_read_s": 1.0,
    "wait_after_negative_write_s": 1.0,
    "positive_repeat_count": 40,
    "negative_repeat_count": 40,
    "plan_repeat_count": 3,
}

PREVIEW_ONLY = True
WRITE_POSITIVE_SEQ_ID = 1
READ_SEQ_ID = 2
WRITE_NEGATIVE_SEQ_ID = 3


def _pulse_config(seq_id, voltage, time_values, meas_types, meas_start, meas_stop):
    ch1_start = [BASE_V, voltage, voltage, BASE_V]
    ch1_stop = [voltage, voltage, BASE_V, BASE_V]
    zeros = [0.0] * 4
    ch1_config = (
        seq_id, ch1_start, ch1_stop, time_values,
        meas_types, meas_start, meas_stop,
    )
    ch2_config = (
        seq_id, zeros.copy(), zeros.copy(), time_values,
        meas_types, meas_start, meas_stop,
    )
    return ch1_config, ch2_config


def _rebuild_runtime_config():
    global BASE_V, WRITE_POSITIVE_V, WRITE_NEGATIVE_V, READ_V
    global WRITE_POSITIVE_DWELL, WRITE_NEGATIVE_DWELL, READ_DWELL
    global WRITE_POSITIVE_TRF, WRITE_NEGATIVE_TRF, READ_TRF
    global WRITE_POSITIVE_IDLE, WRITE_NEGATIVE_IDLE, READ_IDLE
    global WAIT_AFTER_WRITE_POSITIVE_S, WAIT_AFTER_READ_S, WAIT_AFTER_WRITE_NEGATIVE_S
    global POSITIVE_REPEAT_COUNT, NEGATIVE_REPEAT_COUNT, PLAN_REPEAT_COUNT
    global time_values_write_positive, time_values_write_negative, time_values_read
    global ch1_write_positive_config, ch2_write_positive_config
    global ch1_write_negative_config, ch2_write_negative_config
    global ch1_read_config, ch2_read_config
    global write_positive_entry, write_negative_entry, read_entry
    global ONE_PLAN, TEST_PLAN, seq_configs, CURRENT_RANGES

    BASE_V = float(params["base_v"])
    WRITE_POSITIVE_V = float(params["write_positive_v"])
    WRITE_NEGATIVE_V = float(params["write_negative_v"])
    READ_V = float(params["read_v"])
    WRITE_POSITIVE_DWELL = float(params["write_positive_dwell"])
    WRITE_NEGATIVE_DWELL = float(params["write_negative_dwell"])
    READ_DWELL = float(params["read_dwell"])
    WRITE_POSITIVE_TRF = float(params["write_positive_trf"])
    WRITE_NEGATIVE_TRF = float(params["write_negative_trf"])
    READ_TRF = float(params["read_trf"])
    WRITE_POSITIVE_IDLE = float(params["write_positive_idle"])
    WRITE_NEGATIVE_IDLE = float(params["write_negative_idle"])
    READ_IDLE = float(params["read_idle"])
    WAIT_AFTER_WRITE_POSITIVE_S = float(params["wait_after_positive_write_s"])
    WAIT_AFTER_READ_S = float(params["wait_after_read_s"])
    WAIT_AFTER_WRITE_NEGATIVE_S = float(params["wait_after_negative_write_s"])
    POSITIVE_REPEAT_COUNT = int(params["positive_repeat_count"])
    NEGATIVE_REPEAT_COUNT = int(params["negative_repeat_count"])
    PLAN_REPEAT_COUNT = int(params["plan_repeat_count"])
    if min(POSITIVE_REPEAT_COUNT, NEGATIVE_REPEAT_COUNT, PLAN_REPEAT_COUNT) < 0:
        raise ValueError("Identical V2 repeat counts cannot be negative.")

    time_values_write_positive = [
        WRITE_POSITIVE_TRF, WRITE_POSITIVE_DWELL,
        WRITE_POSITIVE_TRF, WRITE_POSITIVE_IDLE,
    ]
    time_values_write_negative = [
        WRITE_NEGATIVE_TRF, WRITE_NEGATIVE_DWELL,
        WRITE_NEGATIVE_TRF, WRITE_NEGATIVE_IDLE,
    ]
    time_values_read = [READ_TRF, READ_DWELL, READ_TRF, READ_IDLE]
    no_measure = [0, 0, 0, 0]
    zero_windows = [0.0] * 4
    read_types = [0, 1, 0, 0]
    read_start = [0.0, READ_DWELL * 0.5, 0.0, 0.0]
    read_stop = [0.0, READ_DWELL * 0.9, 0.0, 0.0]

    ch1_write_positive_config, ch2_write_positive_config = _pulse_config(
        WRITE_POSITIVE_SEQ_ID, WRITE_POSITIVE_V, time_values_write_positive,
        no_measure, zero_windows, zero_windows,
    )
    ch1_read_config, ch2_read_config = _pulse_config(
        READ_SEQ_ID, READ_V, time_values_read, read_types, read_start, read_stop,
    )
    ch1_write_negative_config, ch2_write_negative_config = _pulse_config(
        WRITE_NEGATIVE_SEQ_ID, WRITE_NEGATIVE_V, time_values_write_negative,
        no_measure, zero_windows, zero_windows,
    )
    write_positive_entry = {
        "name": "write_positive",
        "seq_configs": {
            CH1: [ch1_write_positive_config], CH2: [ch2_write_positive_config]
        },
        "wait_after_s": WAIT_AFTER_WRITE_POSITIVE_S,
    }
    read_entry = {
        "name": "read",
        "seq_configs": {CH1: [ch1_read_config], CH2: [ch2_read_config]},
        "wait_after_s": WAIT_AFTER_READ_S,
    }
    write_negative_entry = {
        "name": "write_negative",
        "seq_configs": {
            CH1: [ch1_write_negative_config], CH2: [ch2_write_negative_config]
        },
        "wait_after_s": WAIT_AFTER_WRITE_NEGATIVE_S,
    }
    ONE_PLAN = (
        [write_positive_entry, read_entry] * POSITIVE_REPEAT_COUNT
        + [write_negative_entry, read_entry] * NEGATIVE_REPEAT_COUNT
    )
    TEST_PLAN = ONE_PLAN * PLAN_REPEAT_COUNT
    # Exposed for common offline validators; execution still uses each entry.
    seq_configs = {
        CH1: [ch1_write_positive_config, ch1_read_config, ch1_write_negative_config],
        CH2: [ch2_write_positive_config, ch2_read_config, ch2_write_negative_config],
    }
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


def preview_waveform(
    output_path=None, *, show=True, title_prefix="FTJ Identical V2 CH1"
):
    return preview_sequence_configs(
        [ch1_write_positive_config, ch1_read_config, ch1_write_negative_config],
        output_path,
        title_prefix=title_prefix,
        show=show,
    )


preview_waveforms = preview_waveform


def run_single_test(query, test):
    execute_segARB_test(
        query,
        channels=[CH1, CH2],
        seq_configs=test["seq_configs"],
        current_ranges=CURRENT_RANGES,
        options=SEGARB_OPTIONS,
    )
    df_ch1, df_ch2 = read_both_channels(query, CH1, CH2)
    power_off_outputs(query, (CH1, CH2))
    return df_ch1, df_ch2


def run_ftj_test(*, save_results=True, save_dir=None, file_stem=None):
    save_dir = SAVE_DIR if save_dir is None else Path(save_dir)
    file_stem = FILE_STEM if file_stem is None else str(file_stem)
    summary_rows = []
    combined_frames = []

    with PMUSession(INST, channels=(CH1, CH2)) as session:
        query = session.query
        for index, test in enumerate(TEST_PLAN, start=1):
            print(f"Running {test['name']} ({index}/{len(TEST_PLAN)}) ...")
            df_ch1, df_ch2 = run_single_test(query, test)
            merged_df = merge_channels({CH1: df_ch1, CH2: df_ch2})
            if merged_df is not None and not merged_df.empty:
                merged_df.insert(0, "test_name", test["name"])
                merged_df.insert(1, "test_index", index)
                merged_df.insert(2, "wait_after_s", test["wait_after_s"])
                combined_frames.append(merged_df)
            summary_rows.append(
                {
                    "test_index": index,
                    "test_name": test["name"],
                    f"points_ch{CH1}": 0 if df_ch1 is None else len(df_ch1),
                    f"points_ch{CH2}": 0 if df_ch2 is None else len(df_ch2),
                    "wait_after_s": test["wait_after_s"],
                }
            )
            if test["wait_after_s"] > 0:
                time.sleep(test["wait_after_s"])

    summary_df = pd.DataFrame(summary_rows)
    combined_df = (
        pd.concat(combined_frames, ignore_index=True)
        if combined_frames else pd.DataFrame()
    )
    output_path = None
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
            summary_df.to_excel(writer, sheet_name="Summary", index=False)
            combined_df.to_excel(writer, sheet_name="RawCombined", index=False)
            params_df.to_excel(writer, sheet_name="Parameters", index=False)
    return {
        "summary_df": summary_df,
        "combined_df": combined_df,
        "output_path": output_path,
    }


def main():
    if PREVIEW_ONLY:
        return preview_waveform()
    return run_ftj_test()


if __name__ == "__main__":
    main()
