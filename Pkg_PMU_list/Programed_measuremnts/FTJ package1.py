# -*- coding: utf-8 -*-
"""One-click FTJ characterization package.

Sequence: initial PV2 -> PV2/PUND sweep -> pre-IV PV2 -> repeated segmented
DC I-V -> post-IV PV2. Existing test modules perform every measurement; this
file only exposes parameters, applies them, orders stages, and audits outputs.
"""

from __future__ import annotations

from datetime import datetime
import importlib
from itertools import product
from pathlib import Path
import sys
import time
import traceback

import pandas as pd


SCRIPT_DIR = Path(__file__).resolve().parent
PKG_ROOT = SCRIPT_DIR.parent
REPO_ROOT = PKG_ROOT.parent
for path in (PKG_ROOT, REPO_ROOT):
    if str(path) not in sys.path:
        sys.path.insert(0, str(path))


# =============================================================================
# USER CONFIGURATION
# =============================================================================

# Safety default: preview waveforms only. Set False only for a real run.
# Preview mode never creates PMUSession or SMUSession.
PREVIEW_ONLY = True

# Stop the remaining package after a failed/interrupted hardware stage.
STOP_ON_ERROR = True
STAGE_SETTLE_TIME_S = 1.0

INST = "TCPIP0::129.125.87.80::1225::SOCKET"
CH1, CH2 = 1, 2

# Change this one path to relocate the complete package output.
BASE_SAVE_DIR = Path(r"C:\Users\P317151\Documents\data\FTJ_package1")
SAVE_DIRS = {
    "initial_pv2": BASE_SAVE_DIR / "01_initial_PV2",
    "pv_pund_sweep": BASE_SAVE_DIR / "02_PV2_PUND_sweep",
    "pv2_before_iv": BASE_SAVE_DIR / "03_PV2_before_IV",
    "segmented_iv": BASE_SAVE_DIR / "04_segmented_DC_IV",
    "pv2_after_iv": BASE_SAVE_DIR / "05_PV2_after_IV",
}

# Stages can be disabled or reordered here. Names must match STAGE_CONFIGS.
RUN_ORDER = [
    "initial_pv2",
    "pv_pund_sweep",
    "pv2_before_iv",
    "segmented_iv",
    "pv2_after_iv",
]

COMMON_SEGARB_OPTIONS = {
    "ENABLE_CONNECTION_COMP": False,
    "ENABLE_LOAD_CONFIG": True,
    "LOAD_RESISTANCE": 1e3,
    "ENABLE_LLEC": False,
}

# First PV2: initialization plus an initial device-health record.
INITIAL_PV2 = {
    "enabled": True,
    "repeat_count": 1,
    "save_dir": SAVE_DIRS["initial_pv2"],
    "params": {
        "rise_time": 2.5e-5,
        "delay_time": 2.5e-5,
        "Vp": 4.5,
        "offset": 0.0,
        "area_cm2": (20e-4) ** 2,
        "Irange1": 1e-4,
        "Irange2": 1e-4,
    },
    "segarb_options": dict(COMMON_SEGARB_OPTIONS),
}

# Cartesian PV2/PUND sweep. Real-run count equals the three list lengths
# multiplied together, then multiplied by the number of enabled test types.
PV_PUND_SWEEP = {
    "enabled": True,
    "run_pv2": True,
    "run_pund_tri": True,
    "vp_values": [2.5, 3.0, 3.5, 4.0, 4.5],
    "frequency_values_hz": [
        125,
        250,
        500,
        1000,
        2000,
        5000,
        10000,
        20000,
        50000,
        100000,
    ],
    "delay_time_values_s": [1e-2],
    "settle_time_s": 0.5,
    "save_root": SAVE_DIRS["pv_pund_sweep"],
    "pv2_base_params": {
        "rise_time": 2.5e-5,  # overwritten from frequency_values_hz
        "delay_time": 1e-2,   # overwritten from delay_time_values_s
        "Vp": 4.5,            # overwritten from vp_values
        "offset": 0.0,
        "area_cm2": (20e-4) ** 2,
        "Irange1": 1e-4,
        "Irange2": 1e-4,
    },
    "pund_base_params": {
        "rise_time": 2.5e-5,  # overwritten from frequency_values_hz
        "delay_time": 1e-2,   # overwritten from delay_time_values_s
        "offset_ramp_time": 1e-4,
        "Vp": 4.5,            # overwritten from vp_values
        "offset": 0.0,
        "area_cm2": (20e-4) ** 2,
        "Irange1": 1e-5,
        "Irange2": 1e-6,
    },
    "segarb_options": dict(COMMON_SEGARB_OPTIONS),
}

# PV2 immediately before DC I-V. It is deliberately independent of the
# post-IV dictionary so either check can be adjusted explicitly.
PV2_BEFORE_IV = {
    "enabled": True,
    "repeat_count": 1,
    "save_dir": SAVE_DIRS["pv2_before_iv"],
    "params": {
        "rise_time": 2.5e-5,
        "delay_time": 2.5e-5,
        "Vp": 4.5,
        "offset": 0.0,
        "area_cm2": (20e-4) ** 2,
        "Irange1": 1e-4,
        "Irange2": 1e-4,
    },
    "segarb_options": dict(COMMON_SEGARB_OPTIONS),
}

# Segmented DC I-V. Each repeat is a separate System Mode execution/workbook.
# timeout_s=None means no total-duration deadline; SP is polled until complete.
SEGMENTED_IV = {
    "enabled": True,
    "repeat_count": 3,
    "settle_time_s": 1.0,
    "save_dir": SAVE_DIRS["segmented_iv"],
    "turning_points": [0.0, 5.0, 0.0, -5.0, 0.0],
    "segment_step": 0.1,
    "sweep_channel": 1,
    "bias_channel": 2,
    "available_channels": (1, 2, 3, 4),
    # Current wiring: SMU1/2 through RPM PMU1-1/1-2; SMU3/4 direct.
    "smu_connections": {
        1: "rpm:PMU1-1",
        2: "rpm:PMU1-2",
        3: "direct",
        4: "direct",
    },
    "params": {
        "sweep_current_compliance": 1e-3,
        "bias_voltage": 0.0,
        "bias_current_compliance": 1e-3,
        "sweep_current_range": "auto",
        "bias_current_range": "auto",
        "hold_time": 0.0,
        "sweep_delay": 0.02,
        "integration": "IT2",
        "timeout_s": None,
    },
    "names": {
        "sweep_voltage": "V1",
        "sweep_current": "I1",
        "bias_voltage": "V2",
        "bias_current": "I2",
    },
}

PV2_AFTER_IV = {
    "enabled": True,
    "repeat_count": 1,
    "save_dir": SAVE_DIRS["pv2_after_iv"],
    "params": {
        "rise_time": 2.5e-5,
        "delay_time": 2.5e-5,
        "Vp": 4.5,
        "offset": 0.0,
        "area_cm2": (20e-4) ** 2,
        "Irange1": 1e-4,
        "Irange2": 1e-4,
    },
    "segarb_options": dict(COMMON_SEGARB_OPTIONS),
}

STAGE_CONFIGS = {
    "initial_pv2": INITIAL_PV2,
    "pv_pund_sweep": PV_PUND_SWEEP,
    "pv2_before_iv": PV2_BEFORE_IV,
    "segmented_iv": SEGMENTED_IV,
    "pv2_after_iv": PV2_AFTER_IV,
}


# =============================================================================
# ORCHESTRATION (normally no edits are needed below this line)
# =============================================================================

def load_test_modules():
    """Import existing tests lazily; importing this package never opens VISA."""
    return {
        "pv2": importlib.import_module("Pkg_PMU_list.Fe_cap.PV2"),
        "pund": importlib.import_module("Pkg_PMU_list.Fe_cap.PUND_tri"),
        "sweep": importlib.import_module(
            "Pkg_PMU_list.Fe_cap.PV2_PUND_tri_sweep"
        ),
        "iv": importlib.import_module("SMU.IV.segmented_voltage_sweep"),
    }


def _replace_dict(target, values):
    target.clear()
    target.update(values)


def configure_pv2(module, config):
    module.INST = INST
    module.CH1, module.CH2 = CH1, CH2
    _replace_dict(module.params, config["params"])
    _replace_dict(module.SEGARB_OPTIONS, config["segarb_options"])
    module.SAVE_DIR = Path(config["save_dir"])


def configure_sweep(module, config):
    module.VP_VALUES = list(config["vp_values"])
    module.FREQUENCY_VALUES_HZ = list(config["frequency_values_hz"])
    module.DELAY_TIME_VALUES_S = list(config["delay_time_values_s"])
    module.RUN_PV2 = bool(config["run_pv2"])
    module.RUN_PUND_TRI = bool(config["run_pund_tri"])
    module.STOP_ON_ERROR = bool(STOP_ON_ERROR)
    module.SETTLE_TIME_S = float(config["settle_time_s"])
    module.SAVE_ROOT = Path(config["save_root"])
    module.SUMMARY_CSV = module.SAVE_ROOT / "sweep_summary_live.csv"
    _replace_dict(module.SEGARB_OPTIONS, config["segarb_options"])

    for child, base_params in (
        (module.PV2, config["pv2_base_params"]),
        (module.PUND_tri, config["pund_base_params"]),
    ):
        child.INST = INST
        child.CH1, child.CH2 = CH1, CH2
        _replace_dict(child.params, base_params)
        _replace_dict(child.SEGARB_OPTIONS, config["segarb_options"])


def configure_iv(module, config):
    module.INST = INST
    module.SWEEP_CHANNEL = int(config["sweep_channel"])
    module.BIAS_CHANNEL = int(config["bias_channel"])
    module.AVAILABLE_CHANNELS = tuple(config["available_channels"])
    module.SMU_CONNECTIONS = dict(config["smu_connections"])
    module.TURNING_POINTS = [float(value) for value in config["turning_points"]]
    module.SEGMENT_STEP = config["segment_step"]
    module.PARAMS = dict(config["params"])
    module.NAMES = dict(config["names"])
    module.SAVE_DIR = Path(config["save_dir"])


def validate_package_config(modules):
    unknown = [name for name in RUN_ORDER if name not in STAGE_CONFIGS]
    if unknown:
        raise ValueError(f"RUN_ORDER contains unknown stages: {unknown}")
    if len(set(RUN_ORDER)) != len(RUN_ORDER):
        raise ValueError("RUN_ORDER must not contain duplicate stages.")

    for name in ("initial_pv2", "pv2_before_iv", "pv2_after_iv"):
        if int(STAGE_CONFIGS[name]["repeat_count"]) <= 0:
            raise ValueError(f"{name}.repeat_count must be positive.")
    if int(SEGMENTED_IV["repeat_count"]) <= 0:
        raise ValueError("SEGMENTED_IV.repeat_count must be positive.")

    sweep_values = modules["iv"].build_segmented_voltage_path(
        SEGMENTED_IV["turning_points"], SEGMENTED_IV["segment_step"]
    )
    if len(sweep_values) > 4096:
        raise ValueError(
            f"SEGMENTED_IV creates {len(sweep_values)} points; KXCI allows 4096."
        )

    sweep = PV_PUND_SWEEP
    if not sweep["vp_values"] or not sweep["frequency_values_hz"]:
        raise ValueError("PV/PUND voltage and frequency lists must not be empty.")
    if not sweep["delay_time_values_s"]:
        raise ValueError("PV/PUND delay list must not be empty.")
    if not sweep["run_pv2"] and not sweep["run_pund_tri"]:
        raise ValueError("PV/PUND sweep must enable PV2 and/or PUND_tri.")
    sweep_parameter_points = (
        len(sweep["vp_values"])
        * len(sweep["frequency_values_hz"])
        * len(sweep["delay_time_values_s"])
    )
    return {
        "iv_point_count": len(sweep_values),
        "sweep_parameter_points": sweep_parameter_points,
    }


def snapshot_files(directory):
    directory = Path(directory)
    if not directory.exists():
        return set()
    return {path.resolve() for path in directory.rglob("*") if path.is_file()}


def run_pv2_stage(module, config):
    configure_pv2(module, config)
    for run_index in range(1, int(config["repeat_count"]) + 1):
        print(f"PV2 repeat {run_index}/{config['repeat_count']}")
        module.main()


def run_sweep_stage(module, config):
    configure_sweep(module, config)
    result = module.run_sweep()
    expected_runs = (
        len(config["vp_values"])
        * len(config["frequency_values_hz"])
        * len(config["delay_time_values_s"])
        * (int(config["run_pv2"]) + int(config["run_pund_tri"]))
    )
    if len(result) != expected_runs:
        raise RuntimeError(
            f"PV2/PUND sweep is incomplete: {len(result)}/{expected_runs} runs recorded."
        )
    failed = result[result["status"] != "ok"] if not result.empty else result
    if not failed.empty:
        raise RuntimeError(f"PV2/PUND sweep contains {len(failed)} failed runs.")


def run_iv_stage(module, config):
    configure_iv(module, config)
    for run_index in range(1, int(config["repeat_count"]) + 1):
        print(f"Segmented DC I-V repeat {run_index}/{config['repeat_count']}")
        module.main()
        if run_index < int(config["repeat_count"]):
            time.sleep(float(config["settle_time_s"]))


def _stage_output_root(stage_name, config):
    key = "save_root" if stage_name == "pv_pund_sweep" else "save_dir"
    return Path(config[key])


def save_package_summary(rows):
    BASE_SAVE_DIR.mkdir(parents=True, exist_ok=True)
    live_path = BASE_SAVE_DIR / "package_summary_live.csv"
    pd.DataFrame(rows).to_csv(live_path, index=False)
    return live_path


def run_package():
    """Run enabled stages and preserve a live package-level audit trail."""
    modules = load_test_modules()
    counts = validate_package_config(modules)
    enabled_sweep_tests = int(PV_PUND_SWEEP["run_pv2"]) + int(
        PV_PUND_SWEEP["run_pund_tri"]
    )
    print(
        f"Package preflight: segmented I-V={counts['iv_point_count']} points/run; "
        f"PV/PUND sweep={counts['sweep_parameter_points'] * enabled_sweep_tests} runs."
    )

    runners = {
        "initial_pv2": lambda config: run_pv2_stage(modules["pv2"], config),
        "pv_pund_sweep": lambda config: run_sweep_stage(modules["sweep"], config),
        "pv2_before_iv": lambda config: run_pv2_stage(modules["pv2"], config),
        "segmented_iv": lambda config: run_iv_stage(modules["iv"], config),
        "pv2_after_iv": lambda config: run_pv2_stage(modules["pv2"], config),
    }
    rows = []
    for stage_number, stage_name in enumerate(RUN_ORDER, start=1):
        config = STAGE_CONFIGS[stage_name]
        if not config.get("enabled", True):
            print(f"Skipping disabled stage: {stage_name}")
            continue

        root = _stage_output_root(stage_name, config)
        before = snapshot_files(root)
        started = datetime.now()
        status = "ok"
        error_text = ""
        print(f"\n=== Stage {stage_number}/{len(RUN_ORDER)}: {stage_name} ===")
        try:
            runners[stage_name](config)
        except KeyboardInterrupt:
            status = "interrupted"
            error_text = "KeyboardInterrupt"
        except Exception:
            status = "failed"
            error_text = traceback.format_exc()
            print(error_text)
        ended = datetime.now()
        after = snapshot_files(root)
        rows.append(
            {
                "stage_number": stage_number,
                "stage": stage_name,
                "status": status,
                "start_time": started,
                "end_time": ended,
                "duration_s": (ended - started).total_seconds(),
                "output_root": str(root.resolve()),
                "new_files": " | ".join(str(path) for path in sorted(after - before)),
                "error": error_text,
            }
        )
        save_package_summary(rows)

        if status != "ok" and (STOP_ON_ERROR or status == "interrupted"):
            raise RuntimeError(
                f"Package stopped after {stage_name}: {status}. "
                f"See {BASE_SAVE_DIR / 'package_summary_live.csv'}"
            )
        if (
            status == "ok"
            and STAGE_SETTLE_TIME_S > 0
            and stage_number < len(RUN_ORDER)
        ):
            time.sleep(float(STAGE_SETTLE_TIME_S))

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    final_path = BASE_SAVE_DIR / f"package_summary_{timestamp}.xlsx"
    pd.DataFrame(rows).to_excel(final_path, index=False)
    print(f"FTJ package complete. Summary: {final_path.resolve()}")
    return pd.DataFrame(rows)


def preview_package():
    """Save representative PMU previews plus the exact segmented-IV path."""
    import matplotlib.pyplot as plt

    modules = load_test_modules()
    counts = validate_package_config(modules)
    preview_root = BASE_SAVE_DIR / "00_previews"
    preview_root.mkdir(parents=True, exist_ok=True)
    outputs = []

    for stage_name, config in (
        ("initial_pv2", INITIAL_PV2),
        ("pv2_before_iv", PV2_BEFORE_IV),
        ("pv2_after_iv", PV2_AFTER_IV),
    ):
        if not config.get("enabled", True):
            continue
        configure_pv2(modules["pv2"], config)
        output = preview_root / f"{stage_name}.png"
        modules["pv2"].preview_waveform(output)
        outputs.append(output)

    if PV_PUND_SWEEP.get("enabled", True):
        configure_sweep(modules["sweep"], PV_PUND_SWEEP)
        combinations = list(
            product(
                PV_PUND_SWEEP["vp_values"],
                PV_PUND_SWEEP["frequency_values_hz"],
                PV_PUND_SWEEP["delay_time_values_s"],
            )
        )
        representatives = [combinations[0]]
        if combinations[-1] != combinations[0]:
            representatives.append(combinations[-1])
        for index, (vp, frequency, delay) in enumerate(representatives, start=1):
            for test_name, child, enabled in (
                ("PV2", modules["sweep"].PV2, PV_PUND_SWEEP["run_pv2"]),
                ("PUND_tri", modules["sweep"].PUND_tri, PV_PUND_SWEEP["run_pund_tri"]),
            ):
                if not enabled:
                    continue
                modules["sweep"].configure_test(
                    child,
                    vp=vp,
                    frequency_hz=frequency,
                    delay_time_s=delay,
                    save_dir=preview_root,
                )
                output = preview_root / (
                    f"sweep_{index}_{test_name}_Vp{vp:g}V_f{frequency:g}Hz.png"
                )
                child.preview_waveform(output)
                outputs.append(output)

    if SEGMENTED_IV.get("enabled", True):
        path = modules["iv"].build_segmented_voltage_path(
            SEGMENTED_IV["turning_points"], SEGMENTED_IV["segment_step"]
        )
        figure, axis = plt.subplots(figsize=(9, 4.5))
        axis.plot(range(len(path)), path, linewidth=1.4)
        axis.set(
            title=f"Segmented DC I-V preview ({len(path)} points per repeat)",
            xlabel="Point index",
            ylabel="Commanded voltage (V)",
        )
        axis.grid(alpha=0.3)
        figure.tight_layout()
        output = preview_root / "segmented_iv.png"
        figure.savefig(output, dpi=200)
        plt.close(figure)
        outputs.append(output)

    print(
        f"Preview complete: {len(outputs)} files; "
        f"segmented I-V={counts['iv_point_count']} points/run."
    )
    return outputs


def main():
    if PREVIEW_ONLY:
        return preview_package()
    return run_package()


if __name__ == "__main__":
    main()
