# -*- coding: utf-8 -*-
"""Two-channel SMU linear voltage sweep using the reusable System Mode layer."""

from datetime import datetime
from pathlib import Path
import sys

PKG_ROOT = Path(__file__).resolve().parents[2] / "Pkg_PMU_list"
if str(PKG_ROOT) not in sys.path:
    sys.path.insert(0, str(PKG_ROOT))

from src.smu.data_processing import retrieve_variables, save_workbook
from src.smu.plotting import save_current_plots
from src.smu.session import SMUSession
from src.smu.system_mode import linear_sweep_point_count, run_linear_voltage_sweep


INST = "TCPIP0::129.125.87.80::1225::SOCKET"
SWEEP_CHANNEL = 2
BIAS_CHANNEL = 1
AVAILABLE_CHANNELS = (1, 2, 3, 4)

PARAMS = {
    "start": 0.0,
    "stop": 1.0,
    "step": 0.1,
    "sweep_current_compliance": 1e-3,
    "bias_voltage": 0.0,
    "bias_current_compliance": 1e-3,
    # Numeric value sends RG after SM DM2. Use "auto" or None to keep defaults.
    # Examples: 1e-12 with a preamp, 100e-9 without a preamp.
    "sweep_current_range": "auto",
    "bias_current_range": "auto",
    "hold_time": 0.0,
    "sweep_delay": 0.02,
    # IT1=Fast/0.1 PLC, IT2=Normal/1 PLC, IT3=Quiet/10 PLC.
    # Custom example: "IT4, delay_factor, filter_factor, aperture_plc"
    "integration": "IT2",
    "timeout_s": 300.0,
}

NAMES = {
    "sweep_voltage": "V2",
    "sweep_current": "I2",
    "bias_voltage": "V1",
    "bias_current": "I1",
}

SAVE_DIR = Path(r"C:\Users\P317151\Documents\data\06-07-2026\03C6\R20um1\FE\frequency")


def main():
    """Run the sweep and save data plus the complete parameter table."""
    with SMUSession(INST) as session:
        variables = run_linear_voltage_sweep(
            session.query,
            sweep_channel=SWEEP_CHANNEL,
            bias_channel=BIAS_CHANNEL,
            sweep_voltage_name=NAMES["sweep_voltage"],
            sweep_current_name=NAMES["sweep_current"],
            bias_voltage_name=NAMES["bias_voltage"],
            bias_current_name=NAMES["bias_current"],
            available_channels=AVAILABLE_CHANNELS,
            **PARAMS,
        )
        expected_points = linear_sweep_point_count(
            PARAMS["start"],
            PARAMS["stop"],
            PARAMS["step"],
        )
        data = retrieve_variables(
            session.query,
            variables,
            expected_point_count=expected_points,
        )

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = SAVE_DIR / f"linear_voltage_sweep_{timestamp}.xlsx"
    saved_parameters = {
        "INST": INST,
        "SWEEP_CHANNEL": SWEEP_CHANNEL,
        "BIAS_CHANNEL": BIAS_CHANNEL,
        "AVAILABLE_CHANNELS": AVAILABLE_CHANNELS,
        "EXPECTED_POINT_COUNT": expected_points,
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
    print(f"Saved SMU sweep: {output_path.resolve()}")


if __name__ == "__main__":
    main()
