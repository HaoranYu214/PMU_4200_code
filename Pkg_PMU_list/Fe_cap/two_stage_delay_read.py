# -*- coding: utf-8 -*-
"""Run single segARB tests one by one with a delay between tests.

This is different from a seq_list waveform. Each test is executed separately:
test1 -> wait -> test2 -> wait -> test3 ...
"""

from pathlib import Path
import sys
import time

import pandas as pd

PKG_ROOT = Path(__file__).resolve().parents[1]
if str(PKG_ROOT) not in sys.path:
    sys.path.insert(0, str(PKG_ROOT))

from src.data_processing import merge_channels, read_both_channels
from debug.waveform_preview import preview_sequence_configs
from src.pmu_tests import execute_segARB_test, power_off_outputs
from src.session import PMUSession

INST = "TCPIP0::129.125.87.80::1225::SOCKET"
CH1, CH2 = 1, 2
SAVE_DIR = Path(r"C:\Users\P317151\Documents\data\PMU\StepByStepDelay")
SAVE_DIR.mkdir(parents=True, exist_ok=True)
FILE_STEM = "step_by_step_delay"

CURRENT_RANGES = {CH1: 1e-4, CH2: 1e-4}
SEGARB_OPTIONS = {
    "ENABLE_CONNECTION_COMP": False,
    "ENABLE_LOAD_CONFIG": False,
    "LOAD_RESISTANCE": 1e6,
    "ENABLE_LLEC": False,
}


Vibas_test1 = 4.0
Vibas_test2 = -2.0
Vibas_test3 = -4.0

Dwell_test1 = 1e-5
Dwell_test2 = 1e-3
Dwell_test3 = 1e-5

WAIT_AFTER_TEST1_S = 1.0
WAIT_AFTER_TEST2_S = 1.0
WAIT_AFTER_TEST3_S = 1.0

REPEAT_12 = 10
REPEAT_32 = 10


time_values_test1 = [1e-4, 2e-7, Dwell_test1, 2e-7, 0.5]
time_values_test2 = [1e-4, 2e-7, Dwell_test2, 2e-7, 0.5]
time_values_test3 = [1e-4, 2e-7, Dwell_test3, 2e-7, 0.5]

meas_types_test1 = [0, 0, 1, 0, 0]
meas_start_test1 = [0.0, 0.0, 0.5, 0.0, 0.0]
meas_stop_test1 = [1.0, 1.0, 0.9, 1.0, 1.0]

meas_types_test2 = [0, 0, 1, 0, 0]
meas_start_test2 = [0.0, 0.0, 0.5, 0.0, 0.0]
meas_stop_test2 = [1.0, 1.0, 0.9, 1.0, 1.0]

meas_types_test3 = [0, 0, 1, 0, 0]
meas_start_test3 = [0.0, 0.0, 0.5, 0.0, 0.0]
meas_stop_test3 = [1.0, 1.0, 0.9, 1.0, 1.0]

ch1_start_v_test1 = [0.0, 0.0, Vibas_test1, Vibas_test1, 0.0]
ch1_stop_v_test1 = [0.0, Vibas_test1, Vibas_test1, 0.0, 0.0]
ch2_start_v_test1 = [0.0] * 5
ch2_stop_v_test1 = [0.0] * 5

ch1_start_v_test2 = [0.0, 0.0, Vibas_test2, Vibas_test2, 0.0]
ch1_stop_v_test2 = [0.0, Vibas_test2, Vibas_test2, 0.0, 0.0]
ch2_start_v_test2 = [0.0] * 5
ch2_stop_v_test2 = [0.0] * 5

ch1_start_v_test3 = [0.0, 0.0, Vibas_test3, Vibas_test3, 0.0]
ch1_stop_v_test3 = [0.0, Vibas_test3, Vibas_test3, 0.0, 0.0]
ch2_start_v_test3 = [0.0] * 5
ch2_stop_v_test3 = [0.0] * 5

ch1_config_test1 = (1, ch1_start_v_test1, ch1_stop_v_test1, time_values_test1, meas_types_test1, meas_start_test1, meas_stop_test1)
ch2_config_test1 = (1, ch2_start_v_test1, ch2_stop_v_test1, time_values_test1, meas_types_test1, meas_start_test1, meas_stop_test1)

ch1_config_test2 = (2, ch1_start_v_test2, ch1_stop_v_test2, time_values_test2, meas_types_test2, meas_start_test2, meas_stop_test2)
ch2_config_test2 = (2, ch2_start_v_test2, ch2_stop_v_test2, time_values_test2, meas_types_test2, meas_start_test2, meas_stop_test2)

ch1_config_test3 = (3, ch1_start_v_test3, ch1_stop_v_test3, time_values_test3, meas_types_test3, meas_start_test3, meas_stop_test3)
ch2_config_test3 = (3, ch2_start_v_test3, ch2_stop_v_test3, time_values_test3, meas_types_test3, meas_start_test3, meas_stop_test3)

test1_entry = {
    "name": "test1",
    "seq_configs": {CH1: [ch1_config_test1], CH2: [ch2_config_test1]},
    "wait_after_s": WAIT_AFTER_TEST1_S,
}
test2_entry = {
    "name": "test2",
    "seq_configs": {CH1: [ch1_config_test2], CH2: [ch2_config_test2]},
    "wait_after_s": WAIT_AFTER_TEST2_S,
}
test3_entry = {
    "name": "test3",
    "seq_configs": {CH1: [ch1_config_test3], CH2: [ch2_config_test3]},
    "wait_after_s": WAIT_AFTER_TEST3_S,
}

TEST_PLAN = ([test1_entry, test2_entry] * REPEAT_12) + ([test3_entry, test2_entry] * REPEAT_32)


def preview_waveforms(output_path=None):
    """Preview the three standalone waveforms without connecting to the PMU."""
    configs = [ch1_config_test1, ch1_config_test2, ch1_config_test3]
    return preview_sequence_configs(configs, output_path, title_prefix="Two-stage delay CH1")


def run_single_test(query, test_name, seq_configs):
    """Run one standalone segARB test and return raw channel data."""
    execute_segARB_test(
        query,
        channels=[CH1, CH2],
        seq_configs=seq_configs,
        current_ranges=CURRENT_RANGES,
        options=SEGARB_OPTIONS,
    )
    df_ch1, df_ch2 = read_both_channels(query, CH1, CH2)
    power_off_outputs(query, (CH1, CH2))
    return df_ch1, df_ch2


def main():
    """Run each test separately and save one combined Excel workbook at the end."""
    summary_rows = []
    combined_frames = []

    with PMUSession(INST, channels=(CH1, CH2)) as session:
        query = session.query

        for index, test in enumerate(TEST_PLAN, start=1):
            print(f"Running {test['name']} ({index}/{len(TEST_PLAN)}) ...")
            df_ch1, df_ch2 = run_single_test(query, test["name"], test["seq_configs"])

            merged_df = merge_channels({1: df_ch1, 2: df_ch2})
            if merged_df is not None and not merged_df.empty:
                merged_df.insert(0, "test_name", test["name"])
                merged_df.insert(1, "test_index", index)
                merged_df.insert(2, "wait_after_s", test["wait_after_s"])
                combined_frames.append(merged_df)

            summary_rows.append(
                {
                    "test_name": test["name"],
                    "points_ch1": 0 if df_ch1 is None else len(df_ch1),
                    "points_ch2": 0 if df_ch2 is None else len(df_ch2),
                    "wait_after_s": test["wait_after_s"],
                }
            )

            if test["wait_after_s"] > 0:
                print(f"Waiting {test['wait_after_s']:.3f}s ...")
                time.sleep(test["wait_after_s"])

    summary_df = pd.DataFrame(summary_rows)
    combined_df = pd.concat(combined_frames, ignore_index=True) if combined_frames else pd.DataFrame()
    output_path = SAVE_DIR / f"{FILE_STEM}_all_tests.xlsx"
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="Summary", index=False)
        combined_df.to_excel(writer, sheet_name="RawCombined", index=False)
    print(f"Saved combined workbook: {output_path}")


if __name__ == "__main__":
    main()
