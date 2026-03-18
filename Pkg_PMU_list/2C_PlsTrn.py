# -*- coding: utf-8 -*-
"""Two-channel pulse-train test."""

from pathlib import Path

from src.data_processing import add_resistance_columns, merge_channels, read_both_channels, save_channels_separate_excel
from src.plotting_utils import PlotManager, plot_time_series
from src.pmu_tests import dual_channel_pulse_train, power_off_outputs
from src.session import PMUSession

params = dict(
    CH1_BASE=0.0,
    CH1_AMPLITUDE=-2,
    CH1_PERIOD=2000e-6,
    CH1_WIDTH=800e-6,
    CH1_RISE=50e-6,
    CH1_FALL=50e-6,
    CH1_DELAY=50e-6,
    CH1_RANGE=1e-3,
    CH2_BASE=0.0,
    CH2_AMPLITUDE=0.2,
    CH2_PERIOD=2000e-6,
    CH2_WIDTH=800e-6,
    CH2_RISE=50e-6,
    CH2_FALL=50e-6,
    CH2_DELAY=1050e-6,
    CH2_RANGE=1e-4,
    PULSE_COUNT=100,
    MEASURE_START_D=0.6,
    MEASURE_STOP_D=0.8,
    MEASURE_START_W=0.2,
    MEASURE_STOP_W=0.2,
    ENABLE_LOAD_CONFIG=True,
    LOAD_RESISTANCE=1e6,
    ENABLE_LLEC=False,
    ENABLE_CONNECTION_COMP=False,
    CURRENT_EPS=1e-12,
    RES_MIN=1.0,
    RES_MAX=1e15,
)

INST = "TCPIP0::129.125.87.80::1225::SOCKET"
CH1, CH2 = 1, 2
TEST_MODE = 1
RESISTANCE_SCALE = "linear"

width_us = int(params["CH1_WIDTH"] * 1e6)


with PMUSession(INST, channels=(CH1, CH2)) as session:
    Q = session.query
    print("Running pulse train...")
    dual_channel_pulse_train(Q, CH1, CH2, params, mode=TEST_MODE)
    df1, df2 = read_both_channels(Q, CH1, CH2)
    power_off_outputs(Q, (CH1, CH2))
    if df1 is None or df2 is None or df1.empty or df2.empty:
        raise ValueError("Pulse train returned empty channel data.")

    dfs = {1: df1, 2: df2}
    merged = add_resistance_columns(
        merge_channels(dfs),
        eps=params.get("CURRENT_EPS", 1e-12),
        res_min=params.get("RES_MIN", 1.0),
        res_max=params.get("RES_MAX", 1e15),
    )

    save_dir = Path(r"C:\Users\P317151\Documents\data\FTJ\Refined")
    save_dir.mkdir(parents=True, exist_ok=True)
    fname_base = save_dir / (
        f"2C_PulseTrain_mode{TEST_MODE}_{width_us}us_"
        f"{params['CH1_AMPLITUDE']}V_{params['PULSE_COUNT']}x"
    )
    save_channels_separate_excel(dfs, f"{fname_base}_raw.xlsx")

    with PlotManager(mode="batched", block=True, close_after_show=False) as pm:
        pm.add(
            plot_time_series(
                merged,
                width_us=width_us,
                amp_v=params["CH1_AMPLITUDE"],
                resistance_scale=RESISTANCE_SCALE,
                show=False,
                return_fig=True,
            )
        )
    print("Pulse train complete.")
