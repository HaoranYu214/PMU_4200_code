# -*- coding: utf-8 -*-
"""Two-channel SMU linear voltage sweep using the reusable System Mode layer."""

from datetime import datetime
from pathlib import Path
import sys

SMU_ROOT = Path(__file__).resolve().parents[1]
if str(SMU_ROOT) not in sys.path:
    sys.path.insert(0, str(SMU_ROOT))

from src.data_processing import retrieve_variables, save_workbook
from src.session import SMUSession
from src.system_mode import run_linear_voltage_sweep


INST = "TCPIP0::129.125.87.80::1225::SOCKET"
SWEEP_CHANNEL = 2
BIAS_CHANNEL = 1

PARAMS = {
    "start": 0.0,
    "stop": 1.0,
    "step": 0.1,
    "sweep_current_compliance": 1e-3,
    "bias_voltage": 0.0,
    "bias_current_compliance": 1e-3,
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

SAVE_DIR = Path(r"D:\Code\data\SMU\IV")


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
            **PARAMS,
        )
        data = retrieve_variables(session.query, variables)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = SAVE_DIR / f"linear_voltage_sweep_{timestamp}.xlsx"
    saved_parameters = {
        "INST": INST,
        "SWEEP_CHANNEL": SWEEP_CHANNEL,
        "BIAS_CHANNEL": BIAS_CHANNEL,
        **NAMES,
        **PARAMS,
    }
    save_workbook(output_path, data, saved_parameters)
    print(f"Saved SMU sweep: {output_path.resolve()}")


if __name__ == "__main__":
    main()
