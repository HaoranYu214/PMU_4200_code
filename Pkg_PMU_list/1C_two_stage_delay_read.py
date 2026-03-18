# -*- coding: utf-8 -*-
"""Run a write test, wait, then run a separate read test.

This script is for cases where "continuous read right after write" and
"write -> wait -> read" produce different results.
"""

from pathlib import Path
import time

import matplotlib.pyplot as plt
import pandas as pd

from src.data_processing import calculate_polarization, read_both_channels, save_channels_separate_excel
from src.pmu_tests import hy_pv2_segARB, hy_pund_segARB, power_off_outputs
from src.session import PMUSession

INST = "TCPIP0::129.125.87.80::1225::SOCKET"
CH1, CH2 = 1, 2
WAIT_BETWEEN_STAGES_S = 1.0
SAVE_DIR = Path(r"C:\Users\P317151\Documents\data\PMU\TwoStageDelay")
SAVE_DIR.mkdir(parents=True, exist_ok=True)

# Stage 1: the full test that changes the device state.
WRITE_STAGE_NAME = "PUND_write"
WRITE_STAGE_PARAMS = dict(
    rise_time=5e-5,
    rise_point=250,
    Vp=3.0,
    offset=0.0,
    area_cm2=1.2567e-5,
    Irange1=1e-4,
    Irange2=1e-4,
)

# Stage 2: the readout after the delay.
READ_STAGE_NAME = "PV2_read"
READ_STAGE_PARAMS = dict(
    rise_time=5e-5,
    rise_point=200,
    Vp=2.0,
    offset=0.0,
    area_cm2=1.2567e-5,
    Irange1=1e-4,
    Irange2=1e-4,
)


def run_write_stage(query):
    """Run the first stage that programs or stresses the device."""
    hy_pund_segARB(query, CH1, CH2, WRITE_STAGE_PARAMS)
    df_ch1, df_ch2 = read_both_channels(query, CH1, CH2)
    power_off_outputs(query, (CH1, CH2))
    save_channels_separate_excel(
        {1: df_ch1, 2: df_ch2},
        str(SAVE_DIR / f"{WRITE_STAGE_NAME}_raw.xlsx"),
    )
    return df_ch1, df_ch2


def run_read_stage(query):
    """Run the delayed readback stage and save the analyzed output."""
    hy_pv2_segARB(query, CH1, CH2, READ_STAGE_PARAMS)
    df_ch1, df_ch2 = read_both_channels(query, CH1, CH2)
    power_off_outputs(query, (CH1, CH2))
    if df_ch1 is None or df_ch1.empty:
        raise ValueError("Read stage returned no channel 1 data.")

    polarization = calculate_polarization(
        df_ch1[f"Current {CH1}"].values,
        df_ch1[f"Timestamp {CH1}"].values,
        READ_STAGE_PARAMS.get("area_cm2", 1.0),
    )
    df_read = pd.DataFrame(
        {
            "Time": df_ch1[f"Timestamp {CH1}"].values,
            "Voltage": df_ch1[f"Voltage {CH1}"].values,
            "Current": df_ch1[f"Current {CH1}"].values,
            "Polarization": polarization,
        }
    )

    save_channels_separate_excel(
        {1: df_ch1, 2: df_ch2},
        str(SAVE_DIR / f"{READ_STAGE_NAME}_raw.xlsx"),
    )
    df_read.to_excel(SAVE_DIR / f"{READ_STAGE_NAME}_analysis.xlsx", index=False)

    fig, ax = plt.subplots(figsize=(6, 5))
    ax.plot(df_read["Voltage"], df_read["Polarization"], "b-")
    ax.set_xlabel("Voltage (V)")
    ax.set_ylabel("Polarization (uC/cm^2)")
    ax.set_title(f"{READ_STAGE_NAME} after {WAIT_BETWEEN_STAGES_S:.1f}s delay")
    ax.grid(alpha=0.3)
    fig.tight_layout()
    fig.savefig(SAVE_DIR / f"{READ_STAGE_NAME}_loop.png", dpi=300)
    plt.close(fig)
    return df_read


def main():
    """Run the two-stage delayed measurement flow."""
    with PMUSession(INST, channels=(CH1, CH2)) as session:
        query = session.query

        print(f"Stage 1: {WRITE_STAGE_NAME}")
        run_write_stage(query)

        print(f"Waiting {WAIT_BETWEEN_STAGES_S:.1f}s before readback...")
        time.sleep(WAIT_BETWEEN_STAGES_S)

        print(f"Stage 2: {READ_STAGE_NAME}")
        run_read_stage(query)


if __name__ == "__main__":
    main()
