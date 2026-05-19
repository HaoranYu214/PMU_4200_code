# -*- coding: utf-8 -*-
"""FTJ segARB script with direct seq_configs and shared time arrays."""

from datetime import datetime
from pathlib import Path

import pandas as pd

from src.data_processing import read_both_channels
from debug.waveform_preview import preview_sequence_configs
from src.pmu_tests import execute_segARB_test, power_off_outputs
from src.session import PMUSession

INST = "TCPIP0::129.125.87.80::1225::SOCKET"
CH1, CH2 = 1, 2
SAVE_DIR = Path(r"C:\Users\P317151\Documents\data\18-03-2026\D1\FTJ endurance")
SAVE_DIR.mkdir(parents=True, exist_ok=True)
FILE_STEM = "ftj_test"

CURRENT_RANGES = {CH1: 1e-5, CH2: 1e-5}
SEGARB_OPTIONS = {
    "ENABLE_CONNECTION_COMP": False,
    "ENABLE_LOAD_CONFIG": False,
    "LOAD_RESISTANCE": 1e7,
    "ENABLE_LLEC": True,
}


Vibas_seq1 = 5
Vibas_seq2 = 1
Vibas_seq3 = -1.8

Dwell_seq1 = 1e-6
Dwell_seq2 = 1e-3
Dwell_seq3 = 1e-6


# Shared clock definition for each sequence.
time_values_seq1 = [1e-3, 2e-7, Dwell_seq1, 2e-7, 0.1]
time_values_seq2 = [1e-3, 2e-7, Dwell_seq2, 2e-7, 0.1]
time_values_seq3 = [1e-3, 2e-7, Dwell_seq3, 2e-7, 0.1]

# Measurement start/stop are percentages of each segment window, so use values in [0, 1].
meas_types_seq1 = [0, 0, 0, 0, 0]
meas_start_seq1 = [0.0, 0.0, 0.0, 0.0, 0.0]
meas_stop_seq1 = time_values_seq1

meas_types_seq2 = [0, 0, 1, 0, 0]
meas_start_seq2 = [x * 0.5 for x in time_values_seq2]
meas_stop_seq2  = [x * 0.9 for x in time_values_seq2]

meas_types_seq3 = [0, 0, 0, 0, 0]
meas_start_seq3 = [0.0, 0.0, 0.0, 0.0, 0.0]
meas_stop_seq3 = time_values_seq3

# Put voltage arrays in the config area so they are easy to edit.
ch1_start_v_seq1 = [0.0, 0.0, Vibas_seq1, Vibas_seq1, 0]
ch1_stop_v_seq1 = [0.0, Vibas_seq1, Vibas_seq1, 0.0, 0]
ch2_start_v_seq1 = [0] * 5
ch2_stop_v_seq1 = [0] * 5

ch1_start_v_seq2 = [0.0, 0.0, Vibas_seq2, Vibas_seq2, 0]
ch1_stop_v_seq2 = [0.0, Vibas_seq2, Vibas_seq2, 0.0, 0]
ch2_start_v_seq2 = [0] * 5
ch2_stop_v_seq2 = [0] * 5

ch1_start_v_seq3 = [0.0, 0.0, Vibas_seq3, Vibas_seq3, 0]
ch1_stop_v_seq3 = [0.0, Vibas_seq3, Vibas_seq3, 0.0, 0]
ch2_start_v_seq3 = [0] * 5
ch2_stop_v_seq3 = [0] * 5

ch1_config_seq1 = (1, ch1_start_v_seq1, ch1_stop_v_seq1, time_values_seq1, meas_types_seq1, meas_start_seq1, meas_stop_seq1)
ch2_config_seq1 = (1, ch2_start_v_seq1, ch2_stop_v_seq1, time_values_seq1, meas_types_seq1, meas_start_seq1, meas_stop_seq1)
ch1_config_seq2 = (2, ch1_start_v_seq2, ch1_stop_v_seq2, time_values_seq2, meas_types_seq2, meas_start_seq2, meas_stop_seq2)
ch2_config_seq2 = (2, ch2_start_v_seq2, ch2_stop_v_seq2, time_values_seq2, meas_types_seq2, meas_start_seq2, meas_stop_seq2)
ch1_config_seq3 = (3, ch1_start_v_seq3, ch1_stop_v_seq3, time_values_seq3, meas_types_seq3, meas_start_seq3, meas_stop_seq3)
ch2_config_seq3 = (3, ch2_start_v_seq3, ch2_stop_v_seq3, time_values_seq3, meas_types_seq3, meas_start_seq3, meas_stop_seq3)

seq_configs = {
    CH1: [ch1_config_seq1, ch1_config_seq2, ch1_config_seq3],
    CH2: [ch2_config_seq1, ch2_config_seq2, ch2_config_seq3],
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

SEQ_LIST = {
    CH1: (
        [(1, 1), (2, 1)] * 50 +
        [(3, 1), (2, 1)] * 50 +
        [(1, 1), (2, 1)] * 50 +
        [(3, 1), (2, 1)] * 50 
    ),
    CH2: (
        [(1, 1), (2, 1)] * 50 +
        [(3, 1), (2, 1)] * 50 +
        [(1, 1), (2, 1)] * 50 +
        [(3, 1), (2, 1)] * 50 
    ),
}



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
                or name.startswith(("Vibas_", "Dwell_", "time_values_", "meas_", "ch1_", "ch2_", "seq_configs", "SEQ_LIST"))
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
    run_ftj_test()


if __name__ == "__main__":
    main()
