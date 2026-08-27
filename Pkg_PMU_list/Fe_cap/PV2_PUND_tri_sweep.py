# -*- coding: utf-8 -*-
"""Sweep Vp, triangle frequency, and delay time for PV2 and triangular PUND."""

from __future__ import annotations

from datetime import datetime
from itertools import product
from pathlib import Path
import sys
import time
import traceback

import pandas as pd


PKG_ROOT = Path(__file__).resolve().parents[1]
REPO_ROOT = PKG_ROOT.parent
for path in (PKG_ROOT, REPO_ROOT):
    if str(path) not in sys.path:
        sys.path.insert(0, str(path))
SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

try:
    from . import PUND_tri, PV2
except ImportError:
    import PUND_tri
    import PV2

# Edit these lists to define the Cartesian parameter sweep.
# VP_VALUES = [3.0, 3.5, 4.0, 4.5, 5.0, 5.5, 6.0]
VP_VALUES = [2.5,3,3.5,4,4.5]
FREQUENCY_VALUES_HZ = [125, 250, 500, 1000, 2000, 5000, 10000, 20000, 50000, 100000]
# Use a short nonzero segment for the practical "no delay" case.
DELAY_TIME_VALUES_S = [1e-2]

RUN_PV2 = True
RUN_PUND_TRI = True
STOP_ON_ERROR = False
SETTLE_TIME_S = 0.5
SEGARB_OPTIONS = {
    "ENABLE_CONNECTION_COMP": False,
    "ENABLE_LOAD_CONFIG": True,
    "LOAD_RESISTANCE": 1e3,
    "ENABLE_LLEC": False,
}

SAVE_ROOT = Path(
    r"C:\Users\P317151\Documents\data\25-08-2026\03B4_hZO_2700_800\L20_4\PV and PUND"
)
SUMMARY_CSV = SAVE_ROOT / "sweep_summary_live.csv"


def frequency_to_rise_time(frequency_hz):
    """Convert full 0,+V,-V,0 triangle frequency to one ramp duration."""
    if frequency_hz <= 0:
        raise ValueError("Frequency must be positive.")
    return 1.0 / (4.0 * frequency_hz)


def configure_test(module, *, vp, frequency_hz, delay_time_s, save_dir):
    """Apply one sweep point to a PV2 or triangular-PUND module."""
    module.params["Vp"] = float(vp)
    module.params["rise_time"] = frequency_to_rise_time(frequency_hz)
    module.params["delay_time"] = float(delay_time_s)
    module.SEGARB_OPTIONS.clear()
    module.SEGARB_OPTIONS.update(SEGARB_OPTIONS)
    module.SAVE_DIR = Path(save_dir)
    module.SAVE_DIR.mkdir(parents=True, exist_ok=True)


def snapshot_files(directory):
    """Return the files currently present in a test output directory."""
    directory = Path(directory)
    if not directory.exists():
        return set()
    return {path.resolve() for path in directory.iterdir() if path.is_file()}


def save_live_summary(rows):
    """Persist sweep progress after every attempted test."""
    SAVE_ROOT.mkdir(parents=True, exist_ok=True)
    pd.DataFrame(rows).to_csv(SUMMARY_CSV, index=False)


def run_one_test(module, test_name, vp, frequency_hz, delay_time_s, run_index, total_runs):
    """Configure and run one test, returning one summary row."""
    test_save_dir = SAVE_ROOT / test_name
    configure_test(
        module,
        vp=vp,
        frequency_hz=frequency_hz,
        delay_time_s=delay_time_s,
        save_dir=test_save_dir,
    )

    start_time = datetime.now()
    before_files = snapshot_files(test_save_dir)
    print(
        f"[{run_index}/{total_runs}] {test_name}: "
        f"Vp={vp:g} V, f={frequency_hz:g} Hz, "
        f"delay={delay_time_s:g} s, "
        f"rise={module.params['rise_time']:.3e} s"
    )

    status = "ok"
    error_text = ""
    try:
        module.main()
    except KeyboardInterrupt:
        raise
    except Exception:
        status = "failed"
        error_text = traceback.format_exc()
        print(error_text)

    end_time = datetime.now()
    after_files = snapshot_files(test_save_dir)
    new_files = sorted(str(path) for path in after_files - before_files)
    return {
        "run_index": run_index,
        "test": test_name,
        "status": status,
        "Vp_V": vp,
        "frequency_Hz": frequency_hz,
        "rise_time_s": module.params["rise_time"],
        "delay_time_s": delay_time_s,
        "Irange1_A": module.params["Irange1"],
        "Irange2_A": module.params["Irange2"],
        "load_config_enabled": module.SEGARB_OPTIONS["ENABLE_LOAD_CONFIG"],
        "load_resistance_ohm": module.SEGARB_OPTIONS["LOAD_RESISTANCE"],
        "start_time": start_time,
        "end_time": end_time,
        "duration_s": (end_time - start_time).total_seconds(),
        "output_files": " | ".join(new_files),
        "error": error_text,
    }


def run_sweep():
    """Run the configured Cartesian sweep and save live/final summaries."""
    SAVE_ROOT.mkdir(parents=True, exist_ok=True)
    tests = []
    if RUN_PV2:
        tests.append(("PV2", PV2))
    if RUN_PUND_TRI:
        tests.append(("PUND_tri", PUND_tri))
    if not tests:
        raise ValueError("At least one of RUN_PV2 or RUN_PUND_TRI must be True.")

    sweep_points = list(
        product(VP_VALUES, FREQUENCY_VALUES_HZ, DELAY_TIME_VALUES_S)
    )
    total_runs = len(sweep_points) * len(tests)
    rows = []
    interrupted = False
    run_index = 0

    try:
        for vp, frequency_hz, delay_time_s in sweep_points:
            for test_name, module in tests:
                run_index += 1
                row = run_one_test(
                    module,
                    test_name,
                    vp,
                    frequency_hz,
                    delay_time_s,
                    run_index,
                    total_runs,
                )
                rows.append(row)
                save_live_summary(rows)
                if row["status"] != "ok" and STOP_ON_ERROR:
                    raise RuntimeError(
                        f"{test_name} failed at Vp={vp:g} V, "
                        f"f={frequency_hz:g} Hz, delay={delay_time_s:g} s."
                    )
                if SETTLE_TIME_S > 0:
                    time.sleep(SETTLE_TIME_S)
    except KeyboardInterrupt:
        interrupted = True
        print(f"Sweep interrupted after {run_index}/{total_runs} runs.")
    finally:
        if rows:
            save_live_summary(rows)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            final_path = SAVE_ROOT / f"sweep_summary_{timestamp}.xlsx"
            pd.DataFrame(rows).to_excel(final_path, index=False)
            print(f"Saved sweep summary to {final_path}")

    success_count = sum(row["status"] == "ok" for row in rows)
    status = "interrupted" if interrupted else "complete"
    print(f"Sweep {status}: {success_count}/{len(rows)} runs successful.")
    return pd.DataFrame(rows)


if __name__ == "__main__":
    run_sweep()
