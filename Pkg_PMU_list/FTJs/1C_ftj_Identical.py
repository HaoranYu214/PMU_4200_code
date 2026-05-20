# -*- coding: utf-8 -*-
"""FTJ identical-pulse script."""

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
SAVE_DIR = Path(r"C:\Users\P317151\Documents\data\19-05-2026\03A6\50um-2\FTJ\Identical")
SAVE_DIR.mkdir(parents=True, exist_ok=True)
FILE_STEM = "ftj_identical"

CURRENT_RANGES = {CH1: 1e-7, CH2: 1e-7}
SEGARB_OPTIONS = {
    "ENABLE_CONNECTION_COMP": False,
    "ENABLE_LOAD_CONFIG": False,
    "LOAD_RESISTANCE": 1,
    "ENABLE_LLEC": False,
}


WRITE_POSITIVE_V = 5
READ_V = -1
WRITE_NEGATIVE_V = -4

WRITE_POSITIVE_DWELL = 1e-5
READ_DWELL = 1e-3
WRITE_NEGATIVE_DWELL = 1e-5
PREVIEW_ONLY = False

WRITE_POSITIVE_SEQ_ID = 1
READ_POSITIVE_SEQ_ID = 2
WRITE_NEGATIVE_SEQ_ID = 3
READ_NEGATIVE_SEQ_ID = 4

POSITIVE_REPEAT_COUNT = 20
NEGATIVE_REPEAT_COUNT = 20


# Shared clock definition for each sequence.
time_values_write_positive = [1e-3, 1e-6, WRITE_POSITIVE_DWELL, 1e-6, 1e-2]
time_values_read = [1e-3, 1e-6, READ_DWELL, 1e-6, 1e-2]
time_values_write_negative = [1e-3, 1e-6, WRITE_NEGATIVE_DWELL, 1e-6, 1e-2]

# Measurement start/stop are percentages of each segment window, so use values in [0, 1].
meas_types_write_positive = [0, 0, 0, 0, 0]
meas_start_write_positive = [0.0, 0.0, 0.0, 0.0, 0.0]
meas_stop_write_positive = time_values_write_positive

meas_types_read = [0, 0, 1, 0, 0]
meas_start_read = [x * 0.5 for x in time_values_read]
meas_stop_read = [x * 0.9 for x in time_values_read]

meas_types_write_negative = [0, 0, 0, 0, 0]
meas_start_write_negative = [0.0, 0.0, 0.0, 0.0, 0.0]
meas_stop_write_negative = time_values_write_negative

# Put voltage arrays in the config area so they are easy to edit.
ch1_start_v_write_positive = [0.0, 0.0, WRITE_POSITIVE_V, WRITE_POSITIVE_V, 0]
ch1_stop_v_write_positive = [0.0, WRITE_POSITIVE_V, WRITE_POSITIVE_V, 0.0, 0]
ch2_start_v_seq1 = [0] * 5
ch2_stop_v_seq1 = [0] * 5

ch1_start_v_read1 = [0.0, 0.0, READ_V, READ_V, 0]
ch1_stop_v_read1 = [0.0, READ_V, READ_V, 0.0, 0]
ch1_start_v_read2 = [0.0, 0.0, -READ_V, -READ_V, 0]
ch1_stop_v_read2 = [0.0, -READ_V, -READ_V, 0.0, 0]


ch2_start_v_seq2 = [0] * 5
ch2_stop_v_seq2 = [0] * 5

ch1_start_v_write_negative = [0.0, 0.0, WRITE_NEGATIVE_V, WRITE_NEGATIVE_V, 0]
ch1_stop_v_write_negative = [0.0, WRITE_NEGATIVE_V, WRITE_NEGATIVE_V, 0.0, 0]
ch2_start_v_seq3 = [0] * 5
ch2_stop_v_seq3 = [0] * 5

ch1_config_write_positive = (WRITE_POSITIVE_SEQ_ID, ch1_start_v_write_positive, ch1_stop_v_write_positive, time_values_write_positive, meas_types_write_positive, meas_start_write_positive, meas_stop_write_positive)
ch2_config_write_positive = (WRITE_POSITIVE_SEQ_ID, ch2_start_v_seq1, ch2_stop_v_seq1, time_values_write_positive, meas_types_write_positive, meas_start_write_positive, meas_stop_write_positive)

ch1_config_read1 = (READ_POSITIVE_SEQ_ID, ch1_start_v_read1, ch1_stop_v_read1, time_values_read, meas_types_read, meas_start_read, meas_stop_read)
ch2_config_read1 = (READ_POSITIVE_SEQ_ID, ch2_start_v_seq2, ch2_stop_v_seq2, time_values_read, meas_types_read, meas_start_read, meas_stop_read)
ch1_config_read2 = (READ_NEGATIVE_SEQ_ID, ch1_start_v_read2, ch1_stop_v_read2, time_values_read, meas_types_read, meas_start_read, meas_stop_read)
ch2_config_read2 = (READ_NEGATIVE_SEQ_ID, ch2_start_v_seq2, ch2_stop_v_seq2, time_values_read, meas_types_read, meas_start_read, meas_stop_read)


ch1_config_write_negative = (WRITE_NEGATIVE_SEQ_ID, ch1_start_v_write_negative, ch1_stop_v_write_negative, time_values_write_negative, meas_types_write_negative, meas_start_write_negative, meas_stop_write_negative)
ch2_config_write_negative = (WRITE_NEGATIVE_SEQ_ID, ch2_start_v_seq3, ch2_stop_v_seq3, time_values_write_negative, meas_types_write_negative, meas_start_write_negative, meas_stop_write_negative)

seq_configs = {
    CH1: [ch1_config_write_positive, ch1_config_read1, ch1_config_read2, ch1_config_write_negative],
    CH2: [ch2_config_write_positive, ch2_config_read1, ch2_config_read2, ch2_config_write_negative],
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

SEQ_PLAN = (
    [(WRITE_POSITIVE_SEQ_ID, 1), (READ_POSITIVE_SEQ_ID, 1), (READ_NEGATIVE_SEQ_ID, 1)] * POSITIVE_REPEAT_COUNT +
    [(WRITE_NEGATIVE_SEQ_ID, 1),  (READ_NEGATIVE_SEQ_ID, 1), (READ_POSITIVE_SEQ_ID, 1)] * NEGATIVE_REPEAT_COUNT +
    [(WRITE_POSITIVE_SEQ_ID, 1), (READ_POSITIVE_SEQ_ID, 1), (READ_NEGATIVE_SEQ_ID, 1)] * POSITIVE_REPEAT_COUNT +
    [(WRITE_NEGATIVE_SEQ_ID, 1), (READ_NEGATIVE_SEQ_ID, 1), (READ_POSITIVE_SEQ_ID, 1)] * NEGATIVE_REPEAT_COUNT
)


SEQ_LIST = {
    CH1: SEQ_PLAN,
    CH2: SEQ_PLAN,
}

def preview_waveform(output_path=None):
    """Preview the generated identical-pulse waveform on CH1."""
    return preview_sequence_configs(
        [ch1_config_write_positive, ch1_config_read1, ch1_config_read2, ch1_config_write_negative],
        output_path,
        title_prefix="FTJ Identical CH1",
    )


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
                or name.startswith(("WRITE_", "READ_", "time_values_", "meas_", "ch1_", "ch2_", "seq_configs", "SEQ_LIST", "SEQ_PLAN"))
            ]
        )

        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            if df_ch1 is not None and not df_ch1.empty:
                df_ch1.to_excel(writer, sheet_name="Channel_1", index=False)
            if df_ch2 is not None and not df_ch2.empty:
                df_ch2.to_excel(writer, sheet_name="Channel_2", index=False)
            params_df.to_excel(writer, sheet_name="Parameters", index=False)

    return {
        "df_ch1": df_ch1,
        "df_ch2": df_ch2,
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
