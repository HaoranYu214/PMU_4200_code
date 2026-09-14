# -*- coding: utf-8 -*-
"""Run independent dual-polarity FeFET tests over a retention-delay sweep.

Each delay invokes bipolar_program_read.main() as a fresh PMU session.
The dual test remains responsible for the exact positive/negative write and
read parameters; configure those in bipolar_program_read.py.
"""

from pathlib import Path
import sys

import matplotlib.pyplot as plt
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
from measurements.pmu.fet import bipolar_program_read as dual


# Retention delays to test. Every value starts a separate dual test/session.
DELAY_TIMES = (1e-5, 1e-3, 1e-2, 1e-1, 1.0, 10.0, 50.0)

# Store individual runs and the combined summary in one sweep directory.
SWEEP_NAME = "dual_delay_sweep"


def validate_delay_times():
    if not DELAY_TIMES:
        raise ValueError("DELAY_TIMES must contain at least one delay.")
    if any(float(delay) <= 0 for delay in DELAY_TIMES):
        raise ValueError("Every retention delay must be greater than zero.")


def save_summary(summary, output_dir):
    summary_stem, run_time = reserve_summary_stem(output_dir, "delay_summary")
    summary = summary.assign(time=run_time)
    workbook_path = Path(f"{summary_stem}.xlsx")
    summary.to_excel(workbook_path, sheet_name="Delay_Summary", index=False)

    fig, axis = plt.subplots(figsize=(8, 5))
    for polarity, group in summary.groupby("WritePolarity", sort=False):
        ordered = group.sort_values("RequestedDelay_s")
        axis.plot(
            ordered["RequestedDelay_s"],
            ordered["Id"],
            marker="o",
            label=str(polarity),
        )
    axis.set_xscale("log")
    axis.set_xlabel("Write-to-read delay (s)")
    axis.set_ylabel("Ids, PMU CH2 (A)")
    axis.set_title(
        "FeFET dual-polarity retention\n"
        f"Vg_read = {dual.PARAMS['read_gate_voltage']:g} V, "
        f"Vd_read = {dual.PARAMS['read_drain_voltage']:g} V"
    )
    axis.grid(True, which="both", alpha=0.3)
    axis.legend(title="Write polarity")
    fig.tight_layout()
    plot_path = Path(f"{summary_stem}.png")
    fig.savefig(plot_path, dpi=180)
    plt.close(fig)
    print(f"Saved delay summary: {workbook_path.resolve()}")
    print(f"Saved delay plot: {plot_path.resolve()}")
    return workbook_path, plot_path


def main():
    validate_delay_times()
    if dual.PREVIEW_ONLY:
        raise ValueError("Set PREVIEW_ONLY = False in bipolar_program_read.py.")

    output_dir = prepare_output_dir(Path(dual.SAVE_DIR) / SWEEP_NAME)
    original_save_dir = dual.SAVE_DIR
    original_delay = dual.PARAMS["read_delay"]
    original_output_tag = dual.OUTPUT_TAG
    rows = []

    print(
        "Important: the first positive point uses the device state present before "
        "the script starts. Later positive points follow the preceding negative state."
    )
    try:
        dual.SAVE_DIR = output_dir
        for index, delay in enumerate(DELAY_TIMES, start=1):
            delay = float(delay)
            dual.READ_DELAY = delay
            dual.PARAMS["read_delay"] = delay
            dual.OUTPUT_TAG = None  # The standard filename already includes td.
            print(f"\n=== Delay {delay:g} s ({index}/{len(DELAY_TIMES)}) ===")
            result_path = dual.main()
            if result_path is None:
                raise RuntimeError(f"Delay {delay:g} s did not return a result file.")
            frame = pd.read_excel(result_path, sheet_name="FET_Data")
            frame.insert(0, "RequestedDelay_s", delay)
            frame.insert(1, "DelayRunIndex", index)
            frame["SourceWorkbook"] = Path(result_path).name
            rows.append(frame)
    finally:
        dual.SAVE_DIR = original_save_dir
        dual.READ_DELAY = original_delay
        dual.PARAMS["read_delay"] = original_delay
        dual.OUTPUT_TAG = original_output_tag

    summary = pd.concat(rows, ignore_index=True)
    return save_summary(summary, output_dir)


if __name__ == "__main__":
    main()
