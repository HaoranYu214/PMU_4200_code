"""Plot helpers for saved SMU sweep data."""

from __future__ import annotations

from pathlib import Path
import math

import pandas as pd


def save_current_plots(data, output_base, *, voltage_column, current_column):
    """Save I-V and log(abs(I))-V plots next to a measurement file."""
    output_base = Path(output_base)
    voltage = pd.to_numeric(data[voltage_column], errors="coerce")
    current = pd.to_numeric(data[current_column], errors="coerce")
    plot_data = pd.DataFrame({"voltage": voltage, "current": current}).dropna()

    if plot_data.empty:
        raise ValueError(
            f"No finite {voltage_column}/{current_column} data available to plot."
        )

    import matplotlib

    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    iv_path = output_base.with_name(f"{output_base.stem}_iv.png")
    log_path = output_base.with_name(f"{output_base.stem}_log_abs_i.png")

    fig, ax = plt.subplots()
    ax.plot(plot_data["voltage"], plot_data["current"], marker="o", linewidth=1)
    ax.set_xlabel(f"{voltage_column} (V)")
    ax.set_ylabel(f"{current_column} (A)")
    ax.grid(True, alpha=0.3)
    fig.tight_layout()
    fig.savefig(iv_path, dpi=300)
    plt.close(fig)

    log_data = plot_data[plot_data["current"] != 0].copy()
    if log_data.empty:
        raise ValueError(f"All {current_column} values are zero; cannot plot log(abs(I)).")
    log_data["log_abs_current"] = log_data["current"].map(
        lambda value: math.log10(abs(value))
    )

    fig, ax = plt.subplots()
    ax.plot(
        log_data["voltage"],
        log_data["log_abs_current"],
        marker="o",
        linewidth=1,
    )
    ax.set_xlabel(f"{voltage_column} (V)")
    ax.set_ylabel(f"log10(abs({current_column}))")
    ax.grid(True, alpha=0.3)
    fig.tight_layout()
    fig.savefig(log_path, dpi=300)
    plt.close(fig)

    return iv_path, log_path
