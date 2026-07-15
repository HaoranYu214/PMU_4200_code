# -*- coding: utf-8 -*-
"""Central configuration for one-click FTJ route tests.

Edit this file before running ``run_route.py``. Values in each test section
override the constants in the corresponding local test script before waveform
sequences are built.
"""

from pathlib import Path


BASE_SAVE_DIR = Path(r"C:\Users\P317151\Documents\data\06-07-2026\03C6\L40um6\1click")
INST = "TCPIP0::129.125.87.80::1225::SOCKET"
CH1, CH2 = 1, 2

COMMON_CURRENT_RANGES = {CH1: 1e-4, CH2: 1e-4}
COMMON_SEGARB_OPTIONS = {
    "ENABLE_CONNECTION_COMP": False,
    "ENABLE_LOAD_CONFIG": False,
    "LOAD_RESISTANCE": 1e6,
    "ENABLE_LLEC": False,
}

TEST_ORDER = ["rv2", "pwm", "identical", "mrd"]

TEST_CONFIGS = {
    "rv2": {
        "SAVE_DIR": BASE_SAVE_DIR / "RV2",
        "CURRENT_RANGES": {CH1: 1e-5, CH2: 1e-5},
        "OFFSET_V": -2,
        "VP": 4,
        "WRITE_LEVEL_STEP": 0.2,
        "READ": -1,
        "SCAN_CYCLES": 1,
        "WRITE_DWELL": 5e-5,
        "READ_DWELL": 5e-5,
        "WRITE_IDLE_2": 0.5,
        "READ_IDLE_2": 1e-3,
        "PREVIEW_ONLY": False,
    },
    "pwm": {
        "SAVE_DIR": BASE_SAVE_DIR / "PWM",
        "CURRENT_RANGES": COMMON_CURRENT_RANGES,
        "SEGARB_OPTIONS": {**COMMON_SEGARB_OPTIONS, "ENABLE_LLEC": True},
        "WRITE_POSITIVE_V": 6,
        "WRITE_NEGATIVE_V": -6,
        "READ_V": -2,
        "WRITE_BASE_DWELL": 1e-6,
        "WIDTH_MULTIPLIERS": [1, 2, 5, 10, 20, 50, 100, 200, 500, 1000, 2000, 5000],
        "PWM_REPEAT_COUNT": 1,
        "WRITE_POSITIVE_IDLE": 0.5,
        "WRITE_NEGATIVE_IDLE": 0.5,
        "READ_IDLE": 1e-3,
        "PREVIEW_ONLY": False,
    },
    "identical": {
        "SAVE_DIR": BASE_SAVE_DIR / "Identical",
        "CURRENT_RANGES": {CH1: 1e-5, CH2: 1e-5},
        "WRITE_POSITIVE_V": 0.7,
        "WRITE_NEGATIVE_V": -5,
        "READ_V": -1.2,
        "WRITE_POSITIVE_DWELL": 5e-5,
        "WRITE_NEGATIVE_DWELL": 5e-5,
        "READ_DWELL": 5e-5,
        "WRITE_POSITIVE_IDLE": 0.5,
        "WRITE_NEGATIVE_IDLE": 0.5,
        "READ_IDLE": 0.5,
        "POSITIVE_REPEAT_COUNT": 50,
        "NEGATIVE_REPEAT_COUNT": 50,
        "SEQ_CYCLE_COUNT": 2,
        "PREVIEW_ONLY": False,
    },
    "mrd": {
        "SAVE_DIR": BASE_SAVE_DIR / "MRD",
        "CURRENT_RANGES": {CH1: 1e-3, CH2: 1e-3},
        "WRITE_VOLTAGES": [0.1, 0.5, 1, 1.5, 2, 2.5, 3, 3.5, 4, 4.5, 5, 6, 7],
        "READ_V": -2,
        "CYCLES_PER_LEVEL": 1,
        "WRITE_DWELL": 5e-5,
        "READ_DWELL": 5e-5,
        "WRITE_IDLE_2": 0.1,
        "READ_IDLE_2": 0.1,
        "PREVIEW_ONLY": False,
    },
}


def apply_route_config(test_name, namespace):
    """Apply shared and per-test config values into a module namespace."""
    config = {
        "INST": INST,
        "CH1": CH1,
        "CH2": CH2,
        "SEGARB_OPTIONS": COMMON_SEGARB_OPTIONS,
        **TEST_CONFIGS.get(test_name, {}),
    }
    namespace.update(config)
    if "SAVE_DIR" in namespace:
        namespace["SAVE_DIR"] = Path(namespace["SAVE_DIR"])
        namespace["SAVE_DIR"].mkdir(parents=True, exist_ok=True)
    return config
