# -*- coding: utf-8 -*-
"""External loop endurance runner for FTJ scripts."""

from datetime import datetime
import importlib.util
from pathlib import Path
import traceback

import sys
PKG_ROOT = Path(__file__).resolve().parents[1]
if str(PKG_ROOT) not in sys.path:
    sys.path.insert(0, str(PKG_ROOT))

import pandas as pd

TARGET_SCRIPT_NAME = "1C_ftj_Identical.py"
FTJ_SCRIPT = Path(__file__).with_name(TARGET_SCRIPT_NAME)
SAVE_DIR = Path(r"C:\Users\P317151\Documents\data\18-03-2026\D1\FTJ_endurance")
SAVE_DIR.mkdir(parents=True, exist_ok=True)

LOOP_COUNT = 3000
SAVE_EVERY_RUN = True
STOP_ON_ERROR = True
FILE_STEM_PREFIX = "ftj_endurance"
SUMMARY_CSV = SAVE_DIR / f"{FILE_STEM_PREFIX}_summary_live.csv"


def load_ftj_module():
    """Load the FTJ test script from its file path."""
    if not FTJ_SCRIPT.exists():
        raise FileNotFoundError(f"FTJ target script not found: {FTJ_SCRIPT}")
    spec = importlib.util.spec_from_file_location("ftj_test_module", FTJ_SCRIPT)
    if spec is None or spec.loader is None:
        raise ImportError(f"Unable to load FTJ script: {FTJ_SCRIPT}")

    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def main():
    """Run the FTJ test in an external loop to bypass the 500-sequence limit."""
    ftj_module = load_ftj_module()
    summary_rows = []

    for run_index in range(1, LOOP_COUNT + 1):
        start_time = datetime.now()
        print(f"FTJ endurance run {run_index}/{LOOP_COUNT} started at {start_time:%Y-%m-%d %H:%M:%S}")
        output_path = None
        error_text = ""
        status = "ok"

        try:
            result = ftj_module.run_ftj_test(
                save_results=SAVE_EVERY_RUN,
                save_dir=SAVE_DIR,
                file_stem=f"{FILE_STEM_PREFIX}_run{run_index:06d}",
            )
            output_path = result.get("output_path")
        except Exception:
            status = "failed"
            error_text = traceback.format_exc()
            print(error_text)
            if STOP_ON_ERROR:
                end_time = datetime.now()
                summary_rows.append(make_summary_row(run_index, status, start_time, end_time, output_path, error_text))
                save_live_summary(summary_rows)
                break
        end_time = datetime.now()
        summary_rows.append(make_summary_row(run_index, status, start_time, end_time, output_path, error_text))
        save_live_summary(summary_rows)
        print(f"FTJ endurance run {run_index}/{LOOP_COUNT} finished with status={status}")

        if status != "ok" and STOP_ON_ERROR:
            break

    summary_df = pd.DataFrame(summary_rows)
    summary_timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    summary_path = SAVE_DIR / f"{FILE_STEM_PREFIX}_summary_{summary_timestamp}.xlsx"
    summary_df.to_excel(summary_path, index=False)
    print(f"Saved FTJ endurance summary to {summary_path}")


def make_summary_row(run_index, status, start_time, end_time, output_path, error_text):
    """Build one summary row for both live and final summaries."""
    return {
        "run_index": run_index,
        "status": status,
        "start_time": start_time,
        "end_time": end_time,
        "duration_s": (end_time - start_time).total_seconds(),
        "output_path": str(output_path) if output_path else "",
        "error": error_text,
    }


def save_live_summary(summary_rows):
    """Persist progress after every completed run."""
    pd.DataFrame(summary_rows).to_csv(SUMMARY_CSV, index=False)


if __name__ == "__main__":
    main()
