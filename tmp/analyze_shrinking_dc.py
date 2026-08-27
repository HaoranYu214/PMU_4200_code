"""Offline analysis for the shrinking-positive-peak segmented DC sweep."""

from __future__ import annotations

import argparse
from pathlib import Path

import matplotlib.pyplot as plt
from matplotlib.colors import Normalize
from matplotlib.cm import ScalarMappable
import numpy as np
import pandas as pd


REQUIRED_COLUMNS = {
    "CommandedVoltage",
    "V1",
    "I1",
    "V2",
    "I2",
    "I1_Status",
    "V1_Status",
    "I2_Status",
    "V2_Status",
}


def split_cycles(raw: pd.DataFrame) -> list[dict]:
    """Split cycles at each zero reached after returning from negative voltage."""
    commanded = raw["CommandedVoltage"].to_numpy(dtype=float)
    cycle_ends = [
        index
        for index in range(1, len(commanded))
        if np.isclose(commanded[index], 0.0, atol=1e-12)
        and commanded[index - 1] < 0.0
    ]
    starts = [0, *cycle_ends[:-1]]
    cycles = []
    for number, (start, end) in enumerate(zip(starts, cycle_ends), start=1):
        frame = raw.iloc[start : end + 1].copy()
        cycles.append(
            {
                "cycle": number,
                "positive_peak_v": float(frame["CommandedVoltage"].max()),
                "frame": frame,
            }
        )
    return cycles


def make_summary(cycles: list[dict]) -> pd.DataFrame:
    rows = []
    for cycle in cycles:
        frame = cycle["frame"]
        negative_peak_rows = frame.loc[
            np.isclose(frame["CommandedVoltage"], -5.0, atol=1e-12)
        ]
        rows.append(
            {
                "Cycle": cycle["cycle"],
                "PositivePeak_V": cycle["positive_peak_v"],
                "PointCount": len(frame),
                "MinMeasuredV1_V": frame["V1"].min(),
                "MaxMeasuredV1_V": frame["V1"].max(),
                "MinI1_A": frame["I1"].min(),
                "MaxI1_A": frame["I1"].max(),
                "MaxAbsI1_A": frame["I1"].abs().max(),
                "I1AtMinus5V_A": (
                    negative_peak_rows["I1"].iloc[0]
                    if not negative_peak_rows.empty
                    else np.nan
                ),
                "I1ComplianceCount": int((frame["I1_Status"] == "C").sum()),
                "I2ComplianceCount": int((frame["I2_Status"] == "C").sum()),
            }
        )
    return pd.DataFrame(rows)


def plot_cycles(cycles: list[dict], output_path: Path, source_name: str) -> None:
    plt.style.use("seaborn-v0_8-whitegrid")
    figure, axes = plt.subplots(1, 2, figsize=(13.2, 5.5), constrained_layout=True)
    cmap = plt.get_cmap("viridis")
    norm = Normalize(vmin=0.0, vmax=5.0)

    for cycle in cycles:
        frame = cycle["frame"]
        peak = cycle["positive_peak_v"]
        color = cmap(norm(peak))
        axes[0].plot(frame["V1"], frame["I1"] * 1e6, color=color, linewidth=1.25)
        axes[1].plot(
            frame["V1"],
            frame["I1"].abs().clip(lower=1e-15),
            color=color,
            linewidth=1.25,
        )

    axes[0].axhline(0.0, color="#666666", linewidth=0.8)
    axes[0].axvline(0.0, color="#666666", linewidth=0.8)
    axes[0].set(
        title="DC I-V loop evolution",
        xlabel="Measured V1 (V)",
        ylabel="I1 (µA)",
    )

    axes[1].axvline(0.0, color="#666666", linewidth=0.8)
    axes[1].set_yscale("log")
    axes[1].set(
        title="Absolute-current evolution",
        xlabel="Measured V1 (V)",
        ylabel="|I1| (A)",
    )
    axes[1].grid(True, which="both", alpha=0.28)

    colorbar = figure.colorbar(
        ScalarMappable(norm=norm, cmap=cmap),
        ax=axes,
        location="right",
        shrink=0.88,
        pad=0.025,
    )
    colorbar.set_label("Positive peak voltage (V)")
    figure.suptitle(
        "Shrinking positive peak with fixed −5 V negative branch\n"
        f"26 cycles, 3901 points — {source_name}",
        fontsize=13,
    )
    figure.savefig(output_path, dpi=300, bbox_inches="tight")
    plt.close(figure)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("workbook", type=Path)
    parser.add_argument("--output-dir", type=Path)
    args = parser.parse_args()

    workbook = args.workbook.resolve()
    output_dir = (args.output_dir or workbook.parent).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    raw = pd.read_excel(workbook, sheet_name="Raw")
    missing = REQUIRED_COLUMNS.difference(raw.columns)
    if missing:
        raise ValueError(f"Raw sheet is missing columns: {sorted(missing)}")
    if len(raw) != 3901:
        raise ValueError(f"Expected 3901 rows, found {len(raw)}.")

    cycles = split_cycles(raw)
    if len(cycles) != 26:
        raise ValueError(f"Expected 26 cycles, found {len(cycles)}.")

    stem = workbook.stem
    plot_path = output_dir / f"{stem}_cycle_evolution.png"
    summary_path = output_dir / f"{stem}_cycle_summary.csv"
    plot_cycles(cycles, plot_path, workbook.name)
    make_summary(cycles).to_csv(summary_path, index=False)
    print(plot_path)
    print(summary_path)


if __name__ == "__main__":
    main()
