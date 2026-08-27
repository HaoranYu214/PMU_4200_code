# -*- coding: utf-8 -*-
"""FTJ Identical V2: standalone fixed pulses with real inter-pulse delay.

Physical purpose:
    Study pulse-to-pulse evolution using repeated fixed-amplitude pulses, as in
    the Identical test. The pulse amplitudes do not increase as they do in
    ISPP.

Difference from ftj_Identical.py (V1):
    V1 builds one large PMU SEQ_LIST containing alternating write/read pulses.
    V2 executes every pulse as a separate PMU test, retrieves its data, turns
    the outputs off, and then waits in Python before starting the next pulse.
    V2 is slower but provides a real, long software-controlled delay and keeps
    every pulse acquisition separate.

Current plan:
    (positive write -> wait -> read -> wait) repeated POSITIVE_REPEAT_COUNT,
    then (negative write -> wait -> read -> wait) repeated
    NEGATIVE_REPEAT_COUNT. The complete positive/negative plan is repeated
    PLAN_REPEAT_COUNT times.

Important:
    The read voltage must be chosen low enough to avoid disturbing or
    reprogramming the FTJ.
"""

from pathlib import Path
import sys
import time

import pandas as pd

PKG_ROOT = Path(__file__).resolve().parents[1]
REPO_ROOT = PKG_ROOT.parent
for path in (PKG_ROOT, REPO_ROOT):
    if str(path) not in sys.path:
        sys.path.insert(0, str(path))

from src.pmu.data_processing import merge_channels, read_both_channels
from debug.waveform_preview import preview_sequence_configs
from src.pmu.pmu_tests import execute_segARB_test, power_off_outputs
from src.pmu.session import PMUSession

INST = "TCPIP0::129.125.87.80::1225::SOCKET"
CH1, CH2 = 1, 2
# SAVE_DIR = Path(r"C:\Users\P317151\Documents\data\18-03-2026\D1\FTJ endurance")
SAVE_DIR = Path(r"D:\Code\data\20260620")
SAVE_DIR.mkdir(parents=True, exist_ok=True)
FILE_STEM = "ftj_identical_v2"

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

WAIT_AFTER_WRITE_POSITIVE_S = 1.0
WAIT_AFTER_READ_S = 1.0
WAIT_AFTER_WRITE_NEGATIVE_S = 1.0

PREVIEW_ONLY = True

WRITE_POSITIVE_SEQ_ID = 1
READ_SEQ_ID = 2
WRITE_NEGATIVE_SEQ_ID = 3

POSITIVE_REPEAT_COUNT = 40
NEGATIVE_REPEAT_COUNT = 40
PLAN_REPEAT_COUNT = 3


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

# Match V1: write pulses are not measured; only the read dwell is measured.
meas_types_write_positive = [0, 0, 0, 0]
meas_start_write_positive = [0.0, 0.0, 0.0, 0.0]
meas_stop_write_positive = time_values_write_positive

meas_types_read = [0, 1, 0, 0]
meas_start_read = [x * 0.5 for x in time_values_read]
meas_stop_read = [x * 0.9 for x in time_values_read]

meas_types_write_negative = [0, 0, 0, 0]
meas_start_write_negative = [0.0, 0.0, 0.0, 0.0]
meas_stop_write_negative = time_values_write_negative

ch1_start_v_write_positive = [0.0, WRITE_POSITIVE_V, WRITE_POSITIVE_V, 0.0]
ch1_stop_v_write_positive = [WRITE_POSITIVE_V, WRITE_POSITIVE_V, 0.0, 0.0]
ch2_start_v_write_positive = [0.0] * 4
ch2_stop_v_write_positive = [0.0] * 4

ch1_start_v_read = [0.0, READ_V, READ_V, 0.0]
ch1_stop_v_read = [READ_V, READ_V, 0.0, 0.0]
ch2_start_v_read = [0.0] * 4
ch2_stop_v_read = [0.0] * 4

ch1_start_v_write_negative = [0.0, WRITE_NEGATIVE_V, WRITE_NEGATIVE_V, 0.0]
ch1_stop_v_write_negative = [WRITE_NEGATIVE_V, WRITE_NEGATIVE_V, 0.0, 0.0]
ch2_start_v_write_negative = [0.0] * 4
ch2_stop_v_write_negative = [0.0] * 4

ch1_write_positive_config = (
    WRITE_POSITIVE_SEQ_ID,
    ch1_start_v_write_positive,
    ch1_stop_v_write_positive,
    time_values_write_positive,
    meas_types_write_positive,
    meas_start_write_positive,
    meas_stop_write_positive,
)
ch2_write_positive_config = (
    WRITE_POSITIVE_SEQ_ID,
    ch2_start_v_write_positive,
    ch2_stop_v_write_positive,
    time_values_write_positive,
    meas_types_write_positive,
    meas_start_write_positive,
    meas_stop_write_positive,
)
ch1_read_config = (
    READ_SEQ_ID,
    ch1_start_v_read,
    ch1_stop_v_read,
    time_values_read,
    meas_types_read,
    meas_start_read,
    meas_stop_read,
)
ch2_read_config = (
    READ_SEQ_ID,
    ch2_start_v_read,
    ch2_stop_v_read,
    time_values_read,
    meas_types_read,
    meas_start_read,
    meas_stop_read,
)
ch1_write_negative_config = (
    WRITE_NEGATIVE_SEQ_ID,
    ch1_start_v_write_negative,
    ch1_stop_v_write_negative,
    time_values_write_negative,
    meas_types_write_negative,
    meas_start_write_negative,
    meas_stop_write_negative,
)
ch2_write_negative_config = (
    WRITE_NEGATIVE_SEQ_ID,
    ch2_start_v_write_negative,
    ch2_stop_v_write_negative,
    time_values_write_negative,
    meas_types_write_negative,
    meas_start_write_negative,
    meas_stop_write_negative,
)

write_positive_entry = {
    "name": "write_positive",
    "seq_configs": {
        CH1: [ch1_write_positive_config],
        CH2: [ch2_write_positive_config],
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
        CH1: [ch1_write_negative_config],
        CH2: [ch2_write_negative_config],
    },
    "wait_after_s": WAIT_AFTER_WRITE_NEGATIVE_S,
}

ONE_PLAN = (
    [write_positive_entry, read_entry] * POSITIVE_REPEAT_COUNT
    + [write_negative_entry, read_entry] * NEGATIVE_REPEAT_COUNT
)
TEST_PLAN = ONE_PLAN * PLAN_REPEAT_COUNT


def preview_waveforms(output_path=None):
    """Preview the three standalone waveforms without connecting to the PMU."""
    configs = [
        ch1_write_positive_config,
        ch1_read_config,
        ch1_write_negative_config,
    ]
    return preview_sequence_configs(configs, output_path, title_prefix="FTJ Identical V2 CH1")


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
    params_df = pd.DataFrame(
        [
            {"name": name, "value": repr(value)}
            for name, value in globals().items()
            if name.isupper()
            or name.startswith(
                (
                    "time_values_",
                    "meas_",
                    "ch1_",
                    "ch2_",
                    "TEST_PLAN",
                    "ONE_PLAN",
                )
            )
        ]
    )
    output_path = SAVE_DIR / f"{FILE_STEM}_all_tests.xlsx"
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="Summary", index=False)
        combined_df.to_excel(writer, sheet_name="RawCombined", index=False)
        params_df.to_excel(writer, sheet_name="Parameters", index=False)
    print(f"Saved combined workbook: {output_path}")


if __name__ == "__main__":
    if PREVIEW_ONLY:
        preview_waveforms()
    else:
        main()
