# -*- coding: utf-8 -*-
"""One-click baseline workflow: run one PV2, then one triangular PUND."""

from __future__ import annotations

import importlib
from pathlib import Path
import sys
import time
import traceback

REPO_ROOT = next(
    parent for parent in Path(__file__).resolve().parents
    if (parent / "pyproject.toml").is_file()
)
SRC_ROOT = REPO_ROOT / "src"
for path in (SRC_ROOT, REPO_ROOT):
    if str(path) not in sys.path:
        sys.path.insert(0, str(path))



# =============================================================================
# USER CONFIGURATION
# =============================================================================

PREVIEW_ONLY = False
STOP_ON_ERROR = True
SETTLE_TIME_S = 0.5

# Change this one path to relocate all PV2 and PUND outputs.
BASE_SAVE_DIR = Path(r"C:\Users\P317151\Documents\data\10-09-2026\04A1_2700_1200_300\R10_1\PVInitial")
SAVE_DIRS = {
    "PV2": BASE_SAVE_DIR,
    "PUND_tri": BASE_SAVE_DIR,
}

INST = "TCPIP0::129.125.87.80::1225::SOCKET"
CH1, CH2 = 1, 2
# DEVICE_AREA_CM2 = (20e-4) ** 2
DEVICE_AREA_CM2 = (10*1e-4)**2*3.14
VP_BOTH = 4
RISE_TIME = 2.5e-4
OFFSET_BOTH = 0


# These dictionaries replace the measurement files' ``params`` dictionaries
# completely.  Therefore Vp/rise_time/delay_time/offset also control the PV2
# conditioning triangle and every PUND preset/P/U/N/D pulse; no waveform part
# falls back to a hidden value from the original entry files.
PV2_PARAMS = {
    "rise_time": RISE_TIME,
    "delay_time": 1e-3,
    "Vp": VP_BOTH,
    "offset": OFFSET_BOTH,
    "area_cm2": DEVICE_AREA_CM2,
    "Irange1": 1e-4,
    "Irange2": 1e-4,
}

PUND_PARAMS = {
    "rise_time": RISE_TIME,
    "delay_time": 1e-3,
    "offset_ramp_time": 1e-4,
    "Vp": VP_BOTH,
    "offset": OFFSET_BOTH,
    "area_cm2": DEVICE_AREA_CM2,
    "Irange1": 1e-5,
    "Irange2": 1e-6,
}

SEGARB_OPTIONS = {
    "ENABLE_CONNECTION_COMP": False,
    "ENABLE_LOAD_CONFIG": True,
    "LOAD_RESISTANCE": 1e6,
    "ENABLE_LLEC": False,
}

# =============================================================================
# WORKFLOW
# =============================================================================

def load_measurements():
    """Load the maintained PV2 and PUND measurement modules without VISA I/O."""
    return {
        "PV2": importlib.import_module("measurements.pmu.fe_cap.PV2"),
        "PUND_tri": importlib.import_module("measurements.pmu.fe_cap.PUND_tri"),
    }


def configure_measurement(
    module,
    *,
    params,
    segarb_options,
    save_dir,
    inst,
    channels,
):
    """Inject the complete run config before any waveform is constructed."""
    module.INST = inst
    module.CH1, module.CH2 = channels
    module.params.clear()
    module.params.update(params)
    module.SEGARB_OPTIONS.clear()
    module.SEGARB_OPTIONS.update(segarb_options)
    module.SAVE_DIR = Path(save_dir)


def run_pv_and_pund(
    *,
    inst=INST,
    channels=(CH1, CH2),
    pv2_params=None,
    pund_params=None,
    segarb_options=None,
    save_dirs=None,
    settle_time_s=SETTLE_TIME_S,
    stop_on_error=STOP_ON_ERROR,
    modules=None,
):
    """Run PV2 followed by PUND_tri."""
    modules = load_measurements() if modules is None else modules
    pv2_params = dict(PV2_PARAMS if pv2_params is None else pv2_params)
    pund_params = dict(PUND_PARAMS if pund_params is None else pund_params)
    segarb_options = dict(
        SEGARB_OPTIONS if segarb_options is None else segarb_options
    )
    save_dirs = dict(SAVE_DIRS if save_dirs is None else save_dirs)

    steps = (
        ("PV2", modules["PV2"], pv2_params),
        ("PUND_tri", modules["PUND_tri"], pund_params),
    )
    for step_number, (name, module, params) in enumerate(steps, start=1):
        save_dir = Path(save_dirs[name])
        save_dir.mkdir(parents=True, exist_ok=True)
        configure_measurement(
            module,
            params=params,
            segarb_options=segarb_options,
            save_dir=save_dir,
            inst=inst,
            channels=tuple(channels),
        )
        print(f"\n=== PV and PUND {step_number}/2: {name} ===")
        try:
            module.main()
        except KeyboardInterrupt:
            raise
        except Exception:
            if stop_on_error:
                raise
            print(traceback.format_exc())
        if step_number == 1 and settle_time_s > 0:
            time.sleep(float(settle_time_s))

    print("PV and PUND complete.")


def preview_pv_and_pund(
    *,
    inst=INST,
    channels=(CH1, CH2),
    pv2_params=None,
    pund_params=None,
    segarb_options=None,
    modules=None,
    show=True,
    title_prefix=None,
):
    """Show the existing PV2 and PUND previews without saving image files."""
    import matplotlib.pyplot as plt

    modules = load_measurements() if modules is None else modules
    pv2_params = dict(PV2_PARAMS if pv2_params is None else pv2_params)
    pund_params = dict(PUND_PARAMS if pund_params is None else pund_params)
    segarb_options = dict(
        SEGARB_OPTIONS if segarb_options is None else segarb_options
    )
    results = []
    for name, module, params in (
        ("PV2", modules["PV2"], pv2_params),
        ("PUND_tri", modules["PUND_tri"], pund_params),
    ):
        configure_measurement(
            module,
            params=params,
            segarb_options=segarb_options,
            save_dir=BASE_SAVE_DIR,
            inst=inst,
            channels=tuple(channels),
        )
        preview_title = (
            None if title_prefix is None else f"{title_prefix} | {name}"
        )
        results.append(
            module.preview_waveform(
                show=False,
                title_prefix=preview_title,
            )
        )
    if show:
        plt.show()
    return results


def main():
    if PREVIEW_ONLY:
        return preview_pv_and_pund()
    return run_pv_and_pund()


if __name__ == "__main__":
    main()
