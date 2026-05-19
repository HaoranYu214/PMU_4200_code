# -*- coding: utf-8 -*-
"""FTJ RV script with one prepost sequence and one full write+read scan sequence."""

from datetime import datetime
from pathlib import Path
import sys

import pandas as pd

PKG_ROOT = Path(__file__).resolve().parents[1]
if str(PKG_ROOT) not in sys.path:
    sys.path.insert(0, str(PKG_ROOT))

from debug.waveform_preview import preview_sequence_configs
from src.data_processing import read_both_channels
from src.pmu_tests import execute_segARB_test, power_off_outputs
from src.session import PMUSession

INST = "TCPIP0::129.125.87.80::1225::SOCKET"
CH1, CH2 = 1, 2
SAVE_DIR = Path(r"C:\Users\P317151\Documents\data\19-05-2026\Test")
SAVE_DIR.mkdir(parents=True, exist_ok=True)
FILE_STEM = "ftj_rv"

CURRENT_RANGES = {CH1: 1e-5, CH2: 1e-5}
SEGARB_OPTIONS = {
    "ENABLE_CONNECTION_COMP": False,
    "ENABLE_LOAD_CONFIG": False,
    "LOAD_RESISTANCE": 1e7,
    "ENABLE_LLEC": True,
}

OFFSET_V = 0.0
WRITE_V_MAX = 2.0
WRITE_V_STEP = 0.1
READ_V = 0.1
PREPOST_V = -WRITE_V_MAX

PREPOST_DWELL = 1e-6
WRITE_DWELL = 1e-6
READ_DWELL = 1e-5

PREPOST_IDLE_1 = 1e-3
PREPOST_RISE = 2e-7
PREPOST_FALL = 2e-7
PREPOST_IDLE_2 = 1e-3

WRITE_IDLE_1 = 1e-3
WRITE_RISE = 2e-7
WRITE_FALL = 2e-7
WRITE_IDLE_2 = 1e-3

READ_IDLE_1 = 1e-3
READ_RISE = 2e-7
READ_FALL = 2e-7
READ_IDLE_2 = 1e-3

PREPOST_SEQ_ID = 1
SCAN_SEQ_ID = 2
MAX_SEGMENTS_PER_SEQ = 10000

PREVIEW_ONLY = False

time_values_prepost = [PREPOST_IDLE_1, PREPOST_RISE, PREPOST_DWELL, PREPOST_FALL, PREPOST_IDLE_2]
time_values_write = [WRITE_IDLE_1, WRITE_RISE, WRITE_DWELL, WRITE_FALL, WRITE_IDLE_2]
time_values_read = [READ_IDLE_1, READ_RISE, READ_DWELL, READ_FALL, READ_IDLE_2]

meas_types_prepost = [0, 0, 0, 0, 0]
meas_start_prepost = [0.0] * len(time_values_prepost)
meas_stop_prepost = time_values_prepost

meas_types_write = [0, 0, 0, 0, 0]
meas_start_write = [0.0] * len(time_values_write)
meas_stop_write = time_values_write

meas_types_read = [0, 0, 1, 0, 0]
meas_start_read = [0.0, 0.0, READ_DWELL * 0.5, 0.0, 0.0]
meas_stop_read = [time_values_read[0], time_values_read[1], READ_DWELL * 0.9, time_values_read[3], time_values_read[4]]


def build_pulse_block(amplitude, time_values):
    """Return a 5-segment offset -> amplitude -> offset pulse block."""
    start_v = [OFFSET_V, OFFSET_V, amplitude, amplitude, OFFSET_V]
    stop_v = [OFFSET_V, amplitude, amplitude, OFFSET_V, OFFSET_V]
    return start_v, stop_v, list(time_values)


def voltage_sweep_path(vmax, step):
    """Return offset -> +Vmax -> offset -> -Vmax -> offset scan amplitudes."""
    n_steps = max(1, int(round(vmax / step)))
    positive = [round(step * i, 10) for i in range(1, n_steps + 1)]
    back_to_zero = [round(step * i, 10) for i in range(n_steps - 1, -1, -1)]
    negative = [round(-step * i, 10) for i in range(1, n_steps + 1)]
    back_from_negative = [round(-step * i, 10) for i in range(n_steps - 1, -1, -1)]
    return positive + back_to_zero + negative + back_from_negative


def make_prepost_sequence():
    """Build the prepost sequence: one pulse only."""
    start_v, stop_v, time_values = build_pulse_block(PREPOST_V, time_values_prepost)
    ch1_config = (PREPOST_SEQ_ID, start_v, stop_v, time_values, meas_types_prepost, meas_start_prepost, meas_stop_prepost)
    ch2_config = (
        PREPOST_SEQ_ID,
        [0.0] * len(time_values),
        [0.0] * len(time_values),
        time_values,
        meas_types_prepost,
        meas_start_prepost,
        meas_stop_prepost,
    )
    return ch1_config, ch2_config


def make_scan_sequence(voltages):
    """Build one large sequence: write pulse, then small read pulse, for every scan point."""
    ch1_start_v = []
    ch1_stop_v = []
    ch2_start_v = []
    ch2_stop_v = []
    time_values = []
    meas_types = []
    meas_start = []
    meas_stop = []

    for voltage in voltages:
        write_start_v, write_stop_v, write_times = build_pulse_block(voltage, time_values_write)
        ch1_start_v.extend(write_start_v)
        ch1_stop_v.extend(write_stop_v)
        ch2_start_v.extend([0.0] * len(write_times))
        ch2_stop_v.extend([0.0] * len(write_times))
        time_values.extend(write_times)
        meas_types.extend(meas_types_write)
        meas_start.extend(meas_start_write)
        meas_stop.extend(meas_stop_write)

        read_start_v, read_stop_v, read_times = build_pulse_block(READ_V, time_values_read)
        ch1_start_v.extend(read_start_v)
        ch1_stop_v.extend(read_stop_v)
        ch2_start_v.extend([0.0] * len(read_times))
        ch2_stop_v.extend([0.0] * len(read_times))
        time_values.extend(read_times)
        meas_types.extend(meas_types_read)
        meas_start.extend(meas_start_read)
        meas_stop.extend(meas_stop_read)

    if len(time_values) > MAX_SEGMENTS_PER_SEQ:
        raise ValueError(
            f"RV scan sequence has {len(time_values)} segments, "
            f"above MAX_SEGMENTS_PER_SEQ={MAX_SEGMENTS_PER_SEQ}."
        )

    ch1_config = (SCAN_SEQ_ID, ch1_start_v, ch1_stop_v, time_values, meas_types, meas_start, meas_stop)
    ch2_config = (SCAN_SEQ_ID, ch2_start_v, ch2_stop_v, time_values, meas_types, meas_start, meas_stop)
    return ch1_config, ch2_config


SCAN_VOLTAGES = voltage_sweep_path(WRITE_V_MAX, WRITE_V_STEP)

ch1_prepost_config, ch2_prepost_config = make_prepost_sequence()
ch1_scan_config, ch2_scan_config = make_scan_sequence(SCAN_VOLTAGES)

seq_configs = {
    CH1: [ch1_prepost_config, ch1_scan_config],
    CH2: [ch2_prepost_config, ch2_scan_config],
}

SEQ_PLAN = [(PREPOST_SEQ_ID, 1), (SCAN_SEQ_ID, 1)]
SEQ_LIST = {
    CH1: SEQ_PLAN,
    CH2: SEQ_PLAN,
}


def preview_waveform(output_path=None):
    """Preview the generated RV waveform on CH1."""
    return preview_sequence_configs(
        [ch1_prepost_config, ch1_scan_config],
        output_path,
        title_prefix="FTJ RV CH1",
    )


def build_readback_table(df_ch1, df_ch2):
    """Return only the readback points, one row per write voltage."""
    if df_ch1 is None or df_ch2 is None or df_ch1.empty or df_ch2.empty:
        raise ValueError("RV run returned empty data.")

    count = min(len(df_ch1), len(df_ch2), len(SCAN_VOLTAGES))
    rv_df = pd.DataFrame(
        {
            "WriteVoltage": SCAN_VOLTAGES[:count],
            "ReadTimestamp": df_ch1[f"Timestamp {CH1}"].values[:count],
            "ReadVoltage": df_ch1[f"Voltage {CH1}"].values[:count],
            "ReadCurrent": df_ch1[f"Current {CH1}"].values[:count],
        }
    )
    rv_df["Resistance"] = rv_df["ReadVoltage"] / rv_df["ReadCurrent"].replace(0, pd.NA)
    return rv_df


def run_ftj_test(*, save_results=True, save_dir=SAVE_DIR, file_stem=FILE_STEM):
    """Run the FTJ RV test and optionally save readback-only results."""
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

    rv_df = build_readback_table(df_ch1, df_ch2)

    output_path = None
    if save_results:
        save_dir = Path(save_dir)
        save_dir.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = save_dir / f"{file_stem}_{timestamp}_rv.xlsx"
        params_df = pd.DataFrame(
            [
                {"name": name, "value": repr(value)}
                for name, value in globals().items()
                if name.isupper()
                or name.startswith(("time_values_", "meas_", "ch1_", "ch2_", "seq_configs", "SEQ_LIST"))
            ]
        )

        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            rv_df.to_excel(writer, sheet_name="RV_ReadOnly", index=False)
            df_ch1.to_excel(writer, sheet_name="Channel_1_ReadOnly", index=False)
            df_ch2.to_excel(writer, sheet_name="Channel_2_ReadOnly", index=False)
            params_df.to_excel(writer, sheet_name="Parameters", index=False)

    return {
        "df_ch1": df_ch1,
        "df_ch2": df_ch2,
        "rv_df": rv_df,
        "output_path": output_path,
    }


def main():
    """Run the FTJ RV sequence and save readback data."""
    if PREVIEW_ONLY:
        preview_waveform()
        return
    run_ftj_test()


if __name__ == "__main__":
    main()
