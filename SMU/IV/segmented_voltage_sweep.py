# -*- coding: utf-8 -*-
"""Segmented SMU voltage sweep: 0 -> V1 -> 0 -> V2 -> 0."""

from datetime import datetime
from pathlib import Path
import sys

import pandas as pd

SMU_ROOT = Path(__file__).resolve().parents[1]
if str(SMU_ROOT) not in sys.path:
    sys.path.insert(0, str(SMU_ROOT))

from src.data_processing import retrieve_variables, save_workbook
from src.session import SMUSession
from src.system_mode import build_segmented_voltage_path, run_list_voltage_sweep


INST = "TCPIP0::129.125.87.80::1225::SOCKET"
SWEEP_CHANNEL = 2
BIAS_CHANNEL = 1

# The instrument follows these turning points in order.
V1 = 2.0
V2 = -3.0
TURNING_POINTS = [0.0, V1, 0.0, V2, 0.0]

# One value applies to every segment. A per-segment version is also valid:
# SEGMENT_STEP = [0.05, 0.05, 0.1, 0.1]
SEGMENT_STEP = 0.05

PARAMS = {
    "sweep_current_compliance": 1e-3,
    "bias_voltage": 0.0,
    "bias_current_compliance": 1e-3,
    "hold_time": 0.0,
    "sweep_delay": 0.02,
    # IT1=Fast, IT2=Normal, IT3=Quiet.
    "integration": "IT2",
    "timeout_s": 300.0,
}

NAMES = {
    "sweep_voltage": "V2",
    "sweep_current": "I2",
    "bias_voltage": "V1",
    "bias_current": "I1",
}

SAVE_DIR = Path(r"D:\Code\data\SMU\IV")


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
            **PARAMS,
        )
        data = retrieve_variables(session.query, variables)

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
        "TURNING_POINTS": TURNING_POINTS,
        "SEGMENT_STEP": SEGMENT_STEP,
        "POINT_COUNT": len(sweep_values),
        **NAMES,
        **PARAMS,
    }
    save_workbook(output_path, data, saved_parameters)
    print(f"Saved segmented SMU sweep: {output_path.resolve()}")


if __name__ == "__main__":
    main()
