# -*- coding: utf-8 -*-
"""FTJ ISPP V1: separate sequence for every write-voltage step.

ISPP means Incremental Step Pulse Programming:
    the write-pulse amplitude is increased step by step, and a low-voltage
    read pulse checks the FTJ state after every write.

Physical purpose:
    Use incrementally stronger programming pulses to produce controlled,
    gradual state updates and a more nearly linear resistance/conductance
    trajectory. The increasing amplitude compensates for the nonlinear FTJ
    response that often makes identical fixed-amplitude pulses ineffective at
    first and excessively strong later.

    Cycle-to-cycle variation can be studied by repeating the complete
    staircase, but that is an optional analysis rather than the definition or
    primary purpose of ISPP.

Execution pattern:
    write(V1) -> common read -> write(V2) -> common read -> ...

Difference from V2:
    V1 creates one reusable read sequence plus one write sequence for every
    voltage step. This makes the write/read plan easy to rearrange or repeat,
    but the number of sequence IDs grows with POS_STEPS + NEG_STEPS.

    V2 instead packs all positive write/read pairs into one long sequence and
    all negative pairs into a second long sequence. V2 uses fewer sequence IDs,
    but each sequence contains many more segments.

Current measurement choice:
    V1 measures only the common read pulse; write pulses are not measured.
    Use this version when only resistance/state after each write is required,
    or when the sequence-list order needs to remain highly configurable.
"""

from datetime import datetime
from pathlib import Path

import pandas as pd

import sys
PKG_ROOT = Path(__file__).resolve().parents[1]
REPO_ROOT = PKG_ROOT.parent
for path in (PKG_ROOT, REPO_ROOT):
    if str(path) not in sys.path:
        sys.path.insert(0, str(path))

from src.pmu.data_processing import read_both_channels
from src.pmu.pmu_tests import execute_segARB_test, power_off_outputs
from src.pmu.session import PMUSession
from debug.waveform_preview import preview_sequence_configs

INST = "TCPIP0::129.125.87.80::1225::SOCKET"
CH1, CH2 = 1, 2
SAVE_DIR = Path(r"C:\Users\P317151\Documents\data\07-07-2026\FTJ")
# SAVE_DIR = Path(r"D:\Code\data\20260620")
SAVE_DIR.mkdir(parents=True, exist_ok=True)
FILE_STEM = "ftj_ispp_v1"

CURRENT_RANGES = {CH1: 1e-5, CH2: 1e-5}
SEGARB_OPTIONS = {
    "ENABLE_CONNECTION_COMP": False,
    "ENABLE_LOAD_CONFIG": False,
    "LOAD_RESISTANCE": 1e6,
    "ENABLE_LLEC": True,
}


READ_V = 0.1
POS_V_START = 0.5
POS_V_STOP = 2.0
POS_STEPS = 15
NEG_V_START = -0.5
NEG_V_STOP = -2.0
NEG_STEPS = 15

WRITE_DWELL = 1e-6
READ_DWELL = 1e-5

READ_SEQ_ID = 1
WRITE_POSITIVE_SEQ_START_ID = 2
WRITE_NEGATIVE_SEQ_START_ID = WRITE_POSITIVE_SEQ_START_ID + POS_STEPS
PREVIEW_ONLY = True


# Shared clock definition for each sequence.
time_values_write = [1e-3, 2e-7, WRITE_DWELL, 2e-7, 0.1]
time_values_read = [1e-3, 2e-7, READ_DWELL, 2e-7, 0.1]

# Measurement start/stop are times inside each segment window.
meas_types_write = [0, 0, 0, 0, 0]
meas_start_write = [0.0, 0.0, 0.0, 0.0, 0.0]
meas_stop_write = time_values_write

meas_types_read = [0, 0, 1, 0, 0]
meas_start_read = [x * 0.5 for x in time_values_read]
meas_stop_read = [x * 0.9 for x in time_values_read]


def voltage_steps(start, stop, steps):
    """Return inclusive voltage steps from start to stop."""
    if steps <= 1:
        return [stop]
    step = (stop - start) / (steps - 1)
    return [start + i * step for i in range(steps)]


def make_read_configs():
    """Build the common read sequence config."""
    ch1_start_v = [0.0, 0.0, READ_V, READ_V, 0]
    ch1_stop_v = [0.0, READ_V, READ_V, 0.0, 0]
    ch2_start_v = [0] * 5
    ch2_stop_v = [0] * 5
    ch1_config = (READ_SEQ_ID, ch1_start_v, ch1_stop_v, time_values_read, meas_types_read, meas_start_read, meas_stop_read)
    ch2_config = (READ_SEQ_ID, ch2_start_v, ch2_stop_v, time_values_read, meas_types_read, meas_start_read, meas_stop_read)
    return ch1_config, ch2_config


def make_write_configs(seq_start_id, voltages):
    """Build CH1/CH2 write sequence configs for one ISPP voltage ladder."""
    ch1_configs = []
    ch2_configs = []
    for index, voltage in enumerate(voltages):
        seq_id = seq_start_id + index
        ch1_start_v = [0.0, 0.0, voltage, voltage, 0]
        ch1_stop_v = [0.0, voltage, voltage, 0.0, 0]
        ch2_start_v = [0] * 5
        ch2_stop_v = [0] * 5
        ch1_configs.append((seq_id, ch1_start_v, ch1_stop_v, time_values_write, meas_types_write, meas_start_write, meas_stop_write))
        ch2_configs.append((seq_id, ch2_start_v, ch2_stop_v, time_values_write, meas_types_write, meas_start_write, meas_stop_write))
    return ch1_configs, ch2_configs


def make_ispp_plan(seq_start_id, steps):
    """Build [(write_step_seq, 1), (read_seq, 1), ...]."""
    plan = []
    for index in range(steps):
        plan.append((seq_start_id + index, 1))
        plan.append((READ_SEQ_ID, 1))
    return plan


POS_VOLTAGES = voltage_steps(POS_V_START, POS_V_STOP, POS_STEPS)
NEG_VOLTAGES = voltage_steps(NEG_V_START, NEG_V_STOP, NEG_STEPS)

ch1_read_config, ch2_read_config = make_read_configs()
ch1_pos_configs, ch2_pos_configs = make_write_configs(WRITE_POSITIVE_SEQ_START_ID, POS_VOLTAGES)
ch1_neg_configs, ch2_neg_configs = make_write_configs(WRITE_NEGATIVE_SEQ_START_ID, NEG_VOLTAGES)

seq_configs = {
    CH1: [ch1_read_config] + ch1_pos_configs + ch1_neg_configs,
    CH2: [ch2_read_config] + ch2_pos_configs + ch2_neg_configs,
}

# Equivalent to :PMU:SARB:WFM:SEQ:LIST.
# Each tuple is (seq_id, loop_count), so loop_count repeats that seq in hardware.
SEQ_PLAN = (
    make_ispp_plan(WRITE_POSITIVE_SEQ_START_ID, POS_STEPS) +
    make_ispp_plan(WRITE_NEGATIVE_SEQ_START_ID, NEG_STEPS)
)

SEQ_LIST = {
    CH1: SEQ_PLAN,
    
    CH2: SEQ_PLAN,
}


def preview_waveform(output_path=None):
    """Preview the generated ISPP V1 waveform on CH1."""
    return preview_sequence_configs(
        [ch1_read_config] + ch1_pos_configs + ch1_neg_configs,
        output_path,
        title_prefix="FTJ ISPP V1 CH1",
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
                or name.startswith(("WRITE_", "READ_", "POS_", "NEG_", "time_values_", "meas_", "ch1_", "ch2_", "seq_configs", "SEQ_LIST", "SEQ_PLAN"))
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
