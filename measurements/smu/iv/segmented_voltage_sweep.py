# -*- coding: utf-8 -*-
"""Segmented SMU voltage sweep: 0 -> V1 -> 0 -> V2 -> 0."""

from pathlib import Path
import sys

import pandas as pd

REPO_ROOT = next(
    parent for parent in Path(__file__).resolve().parents
    if (parent / "pyproject.toml").is_file()
)
SRC_ROOT = REPO_ROOT / "src"
for path in (SRC_ROOT, REPO_ROOT):
    if str(path) not in sys.path:
        sys.path.insert(0, str(path))

from keithley4200.output import measurement_name, reserve_output_stem
from keithley4200.smu.data_processing import build_plot_data, retrieve_variables, save_workbook
from keithley4200.smu.plotting import save_current_density_plots
from keithley4200.smu.session import SMUSession
from keithley4200.smu.system_mode import build_segmented_voltage_path, run_list_voltage_sweep


INST = "TCPIP0::129.125.87.80::1225::SOCKET"
SWEEP_CHANNEL = 1
BIAS_CHANNEL = 2
# List every installed/mapped SMU so initialization explicitly disables all
# unused channels before defining the active sweep and bias pair.
AVAILABLE_CHANNELS = (1, 2, 3, 4)

# Physical SMU-to-probe wiring for this 4200A. SMU1 and SMU2 pass through
# the RPMs attached to PMU1 channels 1 and 2; SMU3/SMU4 (when used by future
# experiments) are direct probe connections.
SMU_CONNECTIONS = {
    1: "rpm:PMU1-1",
    2: "rpm:PMU1-2",
    3: "direct",
    4: "direct",
}


SAVE_DIR = Path(r"C:\Users\P317151\Documents\data\10-09-2026\04A1_2700_1200_300\R10_1\IV")

DEVICE_AREA_CM2 = (20e-4) ** 2
# DEVICE_AREA_CM2 = (15*1e-4)**2*3.14

# Generic device-check loop. Package/orchestration files may override these
# globals before calling ``main()`` without duplicating the SMU implementation.
POSITIVE_PEAK_V = 6
NEGATIVE_PEAK_V = -6
TURNING_POINTS = [0.0, POSITIVE_PEAK_V, 0.0, NEGATIVE_PEAK_V, 0.0]
SEGMENT_STEP = 0.1

PARAMS = {
    "sweep_current_compliance": 1e-4,
    "bias_voltage": 0.0,
    "bias_current_compliance": 1e-4,
    "sweep_current_range": "auto",
    "bias_current_range": "auto",
    "hold_time": 0.0,
    "sweep_delay": 0.02,
    # IT1=Fast, IT2=Normal, IT3=Quiet.
    "integration": "IT2",
    # None: no overall test deadline; keep polling SP until KXCI completes.
    # Long endurance runs may legitimately take hours or days. Use a positive
    # number only when an explicit wall-clock limit is wanted; "auto" remains
    # available as an opt-in estimate based on points/delay/integration.
    "timeout_s": None,
}

NAMES = {
    "sweep_voltage": "V1",
    "sweep_current": "I1",
    "sweep_current_density": "J1_A_per_cm2",
    "bias_voltage": "V2",
    "bias_current": "I2",
    "bias_current_density": "J2_A_per_cm2",
}


def preview_waveform(output_path=None, *, show=True, title=None):
    """Show the exact segmented voltage list, or save it when requested."""
    import matplotlib.pyplot as plt

    sweep_values = build_segmented_voltage_path(TURNING_POINTS, SEGMENT_STEP)
    figure, axis = plt.subplots(figsize=(9, 4.5))
    axis.plot(range(len(sweep_values)), sweep_values, linewidth=1.4)
    figure_title = title or (
        f"Segmented DC I-V preview ({len(sweep_values)} points)"
    )
    axis.set(
        title=figure_title,
        xlabel="Point index",
        ylabel="Commanded voltage (V)",
    )
    figure.canvas.manager.set_window_title(figure_title)
    axis.grid(alpha=0.3)
    figure.tight_layout()
    if output_path is None:
        if show:
            plt.show()
        return figure

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    figure.savefig(output_path, dpi=200)
    plt.close(figure)
    return output_path


def main():
    """Run the segmented list sweep and save measured plus commanded values."""
    sweep_values = build_segmented_voltage_path(TURNING_POINTS, SEGMENT_STEP)
    if len(sweep_values) > 4096:
        raise ValueError(
            f"Segmented sweep contains {len(sweep_values)} points; "
            "KXCI VL list sweeps are limited to 4096."
        )
    print(
        f"Segmented sweep: {TURNING_POINTS}, "
        f"{len(sweep_values)} commanded points."
    )

    with SMUSession(INST) as session:
        variables = run_list_voltage_sweep(
            session.query,
            values=sweep_values,
            sweep_channel=SWEEP_CHANNEL,
            bias_channel=BIAS_CHANNEL,
            sweep_voltage_name=NAMES["sweep_voltage"],
            sweep_current_name=NAMES["sweep_current"],
            bias_voltage_name=NAMES["bias_voltage"],
            bias_current_name=NAMES["bias_current"],
            available_channels=AVAILABLE_CHANNELS,
            smu_connections=SMU_CONNECTIONS,
            **PARAMS,
        )
        data = retrieve_variables(
            session.query,
            variables,
            expected_point_count=len(sweep_values),
        )

    # Keep the programmed path next to the measured voltage/current. Series
    # padding makes a point-count mismatch visible instead of hiding it.
    commanded = pd.DataFrame(
        {
            "PointIndex": pd.Series(range(len(sweep_values)), dtype=int),
            "CommandedVoltage": pd.Series(sweep_values, dtype=float),
        }
    )
    data = pd.concat([commanded, data.reset_index(drop=True)], axis=1)
    plot_data = build_plot_data(data, area_cm2=DEVICE_AREA_CM2)

    output_stem = reserve_output_stem(SAVE_DIR, measurement_name("IV", max(abs(value) for value in TURNING_POINTS)))
    output_path = Path(f"{output_stem}.xlsx")
    saved_parameters = {
        "INST": INST,
        "DEVICE_AREA_CM2": DEVICE_AREA_CM2,
        "SWEEP_CHANNEL": SWEEP_CHANNEL,
        "BIAS_CHANNEL": BIAS_CHANNEL,
        "AVAILABLE_CHANNELS": AVAILABLE_CHANNELS,
        "SMU_CONNECTIONS": SMU_CONNECTIONS,
        "TURNING_POINTS": TURNING_POINTS,
        "SEGMENT_STEP": SEGMENT_STEP,
        "POINT_COUNT": len(sweep_values),
        **NAMES,
        **PARAMS,
    }
    save_workbook(
        output_path,
        data,
        saved_parameters,
        plot_data=plot_data,
    )
    try:
        jv_path, log_path = save_current_density_plots(
            plot_data,
            output_path,
            voltage_column=NAMES["sweep_voltage"],
            current_density_column=NAMES["sweep_current_density"],
        )
        print(f"Saved J-V plot: {jv_path.resolve()}")
        print(f"Saved log(abs(J)) plot: {log_path.resolve()}")
    except Exception as exc:
        print(f"Warning: failed to save current-density plots: {exc}")
    print(f"Saved segmented SMU sweep: {output_path.resolve()}")
    return {
        "output_path": output_path,
        "jv_plot_path": jv_path if "jv_path" in locals() else None,
        "log_plot_path": log_path if "log_path" in locals() else None,
        "point_count": len(sweep_values),
    }


if __name__ == "__main__":
    main()
