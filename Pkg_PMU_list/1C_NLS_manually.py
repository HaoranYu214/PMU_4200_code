# -*- coding: utf-8 -*-
"""Compatibility wrapper for running a single NLS switch test."""

from pathlib import Path

from NLS_1C_switch import run_nls_switch_test
from src.session import PMUSession

INST = "TCPIP0::129.125.87.80::1225::SOCKET"
CH1, CH2 = 1, 2
params = dict(
    offset=0,
    Vp=5,
    Rt_p=5e-5,
    Delaytime=100e-6,
    Vsquare=1.5,
    Rt_s=1e-7,
    Dwell=1e-4,
    Irange1=1e-4,
    Irange2=1e-4,
    MeasureSquare=False,
    area_cm2=7.0686e-6,
)

SAVE_DIR = Path(r"C:\Users\P317151\Documents\data\Jingtian\2025-12-14\BTO\Device2")
fname_prefix = (
    f"NISswitch_Meas{params['Vp']}V_{int(params['Rt_p'] * 1e6)}us_"
    f"with{params['Vsquare']}V_{int(params['Rt_s'] * 1e6)}us"
)

with PMUSession(INST, channels=(CH1, CH2)) as session:
    run_nls_switch_test(session.query, CH1, CH2, params, SAVE_DIR, fname_prefix=fname_prefix)
