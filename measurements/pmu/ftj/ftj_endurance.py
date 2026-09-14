# -*- coding: utf-8 -*-
"""External-loop endurance runner for any maintained FTJ measurement.

The target measurement owns its waveform. This runner configures that module,
runs it repeatedly, and saves a live summary after every attempt. Keeping each
iteration as a separate PMU execution avoids building an unbounded Segment Arb
sequence while still allowing long endurance experiments.
"""

from __future__ import annotations

from datetime import datetime
import importlib
from pathlib import Path
import sys
import traceback

import pandas as pd

REPO_ROOT = next(
    parent for parent in Path(__file__).resolve().parents
    if (parent / "pyproject.toml").is_file()
)
SRC_ROOT = REPO_ROOT / "src"
for path in (SRC_ROOT, REPO_ROOT):
    if str(path) not in sys.path:
        sys.path.insert(0, str(path))


from keithley4200.output import prepare_output_dir, reserve_summary_stem

# =============================================================================
# USER CONFIGURATION
# =============================================================================

TARGET_MODULE_NAME = "measurements.pmu.ftj.ftj_RV2"
TARGET_PARAM_OVERRIDES = {}

# Leave these as None to retain the target module's own settings.
TARGET_INST = None
TARGET_CHANNELS = None
TARGET_CURRENT_RANGES = None
TARGET_SEGARB_OPTIONS = None

SAVE_DIR = Path(
    r"C:\Users\P317151\Documents\data\06-07-2026\03C6\L40um6\endurance"
)
LOOP_COUNT = 1000
SAVE_EVERY_RUN = True
STOP_ON_ERROR = True
FILE_STEM_PREFIX = "ftj_endurance"


def load_ftj_module(module_name=TARGET_MODULE_NAME):
    """Import one maintained FTJ measurement module without opening VISA."""
    module = importlib.import_module(module_name)
    if not callable(getattr(module, "configure_measurement", None)):
        raise TypeError(f"{module_name} does not provide configure_measurement(...).")
    if not callable(getattr(module, "run_ftj_test", None)):
        raise TypeError(f"{module_name} does not provide run_ftj_test(...).")
    return module


def configure_target(
    module,
    *,
    param_overrides=None,
    inst=None,
    channels=None,
    current_ranges=None,
    segarb_options=None,
    save_dir=None,
):
    """Merge optional endurance overrides, then rebuild the target waveform."""
    configured_params = dict(module.params)
    configured_params.update(param_overrides or {})
    module.configure_measurement(
        params_override=configured_params,
        inst=inst,
        channels=channels,
        current_ranges=current_ranges,
        segarb_options=segarb_options,
        save_dir=save_dir,
    )
    return configured_params


def make_summary_row(run_index, status, start_time, end_time, output_path, error_text):
    return {
        "run_index": run_index,
        "status": status,
        "start_time": start_time,
        "end_time": end_time,
        "duration_s": (end_time - start_time).total_seconds(),
        "output_path": str(output_path) if output_path else "",
        "error": error_text,
    }


def save_live_summary(summary_rows, summary_csv, run_time):
    """Persist progress after every attempt so an interrupted run is auditable."""
    pd.DataFrame(summary_rows).assign(time=run_time).to_csv(summary_csv, index=False)


def run_endurance(
    *,
    module_name=None,
    param_overrides=None,
    inst=None,
    channels=None,
    current_ranges=None,
    segarb_options=None,
    save_dir=None,
    loop_count=None,
    save_every_run=None,
    stop_on_error=None,
    file_stem_prefix=None,
    module=None,
):
    """Configure one FTJ test once, then execute it in an external loop."""
    module_name = TARGET_MODULE_NAME if module_name is None else module_name
    inst = TARGET_INST if inst is None else inst
    channels = TARGET_CHANNELS if channels is None else channels
    current_ranges = (
        TARGET_CURRENT_RANGES if current_ranges is None else current_ranges
    )
    segarb_options = (
        TARGET_SEGARB_OPTIONS if segarb_options is None else segarb_options
    )
    save_dir = SAVE_DIR if save_dir is None else save_dir
    loop_count = LOOP_COUNT if loop_count is None else loop_count
    save_every_run = SAVE_EVERY_RUN if save_every_run is None else save_every_run
    stop_on_error = STOP_ON_ERROR if stop_on_error is None else stop_on_error
    file_stem_prefix = (
        FILE_STEM_PREFIX if file_stem_prefix is None else file_stem_prefix
    )
    module = load_ftj_module(module_name) if module is None else module
    save_dir = prepare_output_dir(save_dir)
    summary_stem, run_time = reserve_summary_stem(save_dir, "endurance_summary")
    summary_csv = Path(f"{summary_stem}_live.csv")
    configure_target(
        module,
        param_overrides=(
            TARGET_PARAM_OVERRIDES if param_overrides is None else param_overrides
        ),
        inst=inst,
        channels=channels,
        current_ranges=current_ranges,
        segarb_options=segarb_options,
        save_dir=save_dir,
    )

    summary_rows = []
    for run_index in range(1, int(loop_count) + 1):
        start_time = datetime.now()
        print(
            f"FTJ endurance run {run_index}/{loop_count} started at "
            f"{start_time:%Y-%m-%d %H:%M:%S}"
        )
        output_path = None
        error_text = ""
        status = "ok"
        try:
            result = module.run_ftj_test(
                save_results=bool(save_every_run),
                save_dir=save_dir,
                file_stem=file_stem_prefix,
            )
            output_path = result.get("output_path")
        except KeyboardInterrupt:
            end_time = datetime.now()
            summary_rows.append(
                make_summary_row(
                    run_index,
                    "interrupted",
                    start_time,
                    end_time,
                    output_path,
                    "KeyboardInterrupt",
                )
            )
            save_live_summary(summary_rows, summary_csv, run_time)
            raise
        except Exception:
            status = "failed"
            error_text = traceback.format_exc()
            print(error_text)

        end_time = datetime.now()
        summary_rows.append(
            make_summary_row(
                run_index, status, start_time, end_time, output_path, error_text
            )
        )
        save_live_summary(summary_rows, summary_csv, run_time)
        print(
            f"FTJ endurance run {run_index}/{loop_count} finished "
            f"with status={status}"
        )
        if status != "ok" and stop_on_error:
            break

    summary_df = pd.DataFrame(summary_rows).assign(time=run_time)
    summary_path = Path(f"{summary_stem}.xlsx")
    summary_df.to_excel(summary_path, index=False)
    print(f"Saved FTJ endurance summary to {summary_path}")
    return {
        "summary_df": summary_df,
        "summary_csv": summary_csv,
        "summary_path": summary_path,
    }


def main():
    return run_endurance()


if __name__ == "__main__":
    main()
