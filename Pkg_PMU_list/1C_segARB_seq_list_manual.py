# -*- coding: utf-8 -*-
"""Manual segARB script with direct seq_configs and shared time arrays."""

from pathlib import Path

import pandas as pd

from src.data_processing import read_both_channels, save_channels_separate_excel
from src.pmu_tests import execute_segARB_test, power_off_outputs
from src.session import PMUSession

INST = "TCPIP0::129.125.87.80::1225::SOCKET"
CH1, CH2 = 1, 2
SAVE_DIR = Path(r"C:\Users\P317151\Documents\data\PMU\ManualSeqList")
SAVE_DIR.mkdir(parents=True, exist_ok=True)
FILE_STEM = "manual_seq_list_example"

CURRENT_RANGES = {CH1: 1e-4, CH2: 1e-4}


Vibas_seq1 = 4.0
Vibas_seq2 = -2.0
Vibas_seq3 = -4.0

Dwell_seq1 = 1e-5
Dwell_seq2 = 1e-3
Dwell_seq3 = 1e-5


# Shared clock definition for each sequence.
time_values_seq1 = [1e-4, 2e-7, Dwell_seq1, 2e-7, 0.5]
time_values_seq2 = [1e-4, 2e-7, Dwell_seq2, 2e-7, 0.5]
time_values_seq3 = [1e-4, 2e-7, Dwell_seq3, 2e-7, 0.5]

# Measurement start/stop are percentages of each segment window, so use values in [0, 1].
meas_types_seq1 = [0, 0, 1, 0, 0]
meas_start_seq1 = [0.0, 0.0, 0.5, 0.0, 0.0]
meas_stop_seq1 = [1.0, 1.0, 0.9, 1.0, 1.0]

meas_types_seq2 = [0, 0, 1, 0, 0]
meas_start_seq2 = [0.0, 0.0, 0.5, 0.0, 0.0]
meas_stop_seq2 = [1.0, 1.0, 0.9, 1.0, 1.0]

meas_types_seq3 = [0, 0, 1, 0, 0]
meas_start_seq3 = [0.0, 0.0, 0.5, 0.0, 0.0]
meas_stop_seq3 = [1.0, 1.0, 0.9, 1.0, 1.0]

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
SEQ_LIST = {
    CH1: (
        [(1, 1), (2, 1)] * 10 +
        [(3, 1), (2, 1)] * 10
    ),
    CH2: (
        [(1, 1), (2, 1)] * 10 +
        [(3, 1), (2, 1)] * 10
    ),
}


def main():
    """Run the manual segARB sequence list and save raw data."""
    with PMUSession(INST, channels=(CH1, CH2)) as session:
        query = session.query
        execute_segARB_test(
            query,
            channels=[CH1, CH2],
            seq_configs=seq_configs,
            seq_list=SEQ_LIST,
            current_ranges=CURRENT_RANGES,
        )

        df_ch1, df_ch2 = read_both_channels(query, CH1, CH2)
        power_off_outputs(query, (CH1, CH2))

    if df_ch1 is None and df_ch2 is None:
        raise ValueError("No data returned from the manual segARB run.")

    save_channels_separate_excel({1: df_ch1, 2: df_ch2}, str(SAVE_DIR / f"{FILE_STEM}_raw.xlsx"))

    summary_rows = []
    for channel, df in ((CH1, df_ch1), (CH2, df_ch2)):
        count = 0 if df is None else len(df)
        summary_rows.append({"Channel": channel, "Points": count, "SeqList": SEQ_LIST[channel]})
    pd.DataFrame(summary_rows).to_excel(SAVE_DIR / f"{FILE_STEM}_summary.xlsx", index=False)


if __name__ == "__main__":
    main()
