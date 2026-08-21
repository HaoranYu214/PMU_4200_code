# -*- coding: utf-8 -*-
"""Segmented SMU voltage sweep: 0 -> V1 -> 0 -> V2 -> 0."""

from datetime import datetime
from pathlib import Path
import sys

import pandas as pd

PKG_ROOT = Path(__file__).resolve().parents[2] / "Pkg_PMU_list"
if str(PKG_ROOT) not in sys.path:
    sys.path.insert(0, str(PKG_ROOT))

from src.smu.data_processing import retrieve_variables, save_workbook
from src.smu.plotting import save_current_plots
from src.smu.session import SMUSession
from src.smu.system_mode import build_segmented_voltage_path, run_list_voltage_sweep


INST = "TCPIP0::129.125.87.80::1225::SOCKET"
SWEEP_CHANNEL = 2
BIAS_CHANNEL = 1
AVAILABLE_CHANNELS = (1, 2, 3, 4)

# The instrument follows these turning points in order.
V1 = 3
V2 = -3.0
TURNING_POINTS = [0.0, V1, 0.0, V2, 0.0]

# One value applies to every segment. A per-segment version is also valid:
# SEGMENT_STEP = [0.05, 0.05, 0.1, 0.1]
SEGMENT_STEP = 0.05

PARAMS = {
    "sweep_current_compliance": 1e-3,
    "bias_voltage": 0.0,
    "bias_current_compliance": 1e-3,
    # Numeric value sends RG after SM DM2. Use "auto" or None to keep defaults.
    # Examples: 1e-12 with a preamp, 100e-9 without a preamp.
    "sweep_current_range": "auto",
    "bias_current_range": "auto",
    "hold_time": 0.0,
    "sweep_delay": 0.02,
    # IT1=Fast, IT2=Normal, IT3=Quiet.
    "integration": "IT3",
    "timeout_s": 300.0,
}

NAMES = {
    "sweep_voltage": "V2",
    "sweep_current": "I2",
    "bias_voltage": "V1",
    "bias_current": "I1",
}

SAVE_DIR = Path(r"C:\Users\P317151\Documents\data\06-07-2026\03C6\R20um3\DC")


def main():
    """Run the segmented list sweep and save measured plus commanded values."""
    sweep_values = build_segmented_voltage_path(TURNING_POINTS, SEGMENT_STEP)
    print(
        f"Segmented sweep: {TURNING_POINTS}, "
        f"{len(sweep_values)} commanded points."
    )

    with SMUSession(INST) as session:
        variables = run_list_voltage_sweep(
            session.query,
            values=sweep_values,
            sweep_channel=SWEEP_CHANNEL,
            bias_channel=BIAS_CHANNEL,
            sweep_voltage_name=NAMES["sweep_voltage"],
            sweep_current_name=NAMES["sweep_current"],
            bias_voltage_name=NAMES["bias_voltage"],
            bias_current_name=NAMES["bias_current"],
            available_channels=AVAILABLE_CHANNELS,
            **PARAMS,
        )
        data = retrieve_variables(
            session.query,
            variables,
            expected_point_count=len(sweep_values),
        )

    # Keep the programmed path next to the measured voltage/current. Series
    # padding makes a point-count mismatch visible instead of hiding it.
    commanded = pd.DataFrame(
        {
            "PointIndex": pd.Series(range(len(sweep_values)), dtype=int),
            "CommandedVoltage": pd.Series(sweep_values, dtype=float),
        }
    )
    data = pd.concat([commanded, data.reset_index(drop=True)], axis=1)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = SAVE_DIR / f"segmented_voltage_sweep_{timestamp}.xlsx"
    saved_parameters = {
        "INST": INST,
        "SWEEP_CHANNEL": SWEEP_CHANNEL,
        "BIAS_CHANNEL": BIAS_CHANNEL,
        "AVAILABLE_CHANNELS": AVAILABLE_CHANNELS,
        "TURNING_POINTS": TURNING_POINTS,
        "SEGMENT_STEP": SEGMENT_STEP,
        "POINT_COUNT": len(sweep_values),
        **NAMES,
        **PARAMS,
    }
    save_workbook(output_path, data, saved_parameters)
    try:
        iv_path, log_path = save_current_plots(
            data,
            output_path,
            voltage_column=NAMES["sweep_voltage"],
            current_column=NAMES["sweep_current"],
        )
        print(f"Saved I-V plot: {iv_path.resolve()}")
        print(f"Saved log(abs(I)) plot: {log_path.resolve()}")
    except Exception as exc:
        print(f"Warning: failed to save current plots: {exc}")
    print(f"Saved segmented SMU sweep: {output_path.resolve()}")


if __name__ == "__main__":
    main()
