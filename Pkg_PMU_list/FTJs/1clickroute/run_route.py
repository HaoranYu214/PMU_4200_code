# -*- coding: utf-8 -*-
"""Run selected FTJ tests from this one-click route folder."""

from __future__ import annotations

import argparse
import importlib
import importlib.util
import sys
import time
from pathlib import Path

from route_config import TEST_ORDER


SCRIPT_DIR = Path(__file__).resolve().parent
PKG_ROOT = SCRIPT_DIR.parents[1]
if str(PKG_ROOT) not in sys.path:
    sys.path.insert(0, str(PKG_ROOT))
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))


TEST_MODULES = {
    "rv2": "ftj_RV2.py",
    "pwm": "ftj_PWM.py",
    "identical": "ftj_Identical_V1.py",
    # "mrd": "ftj_MRD.py",
}


def load_test_module(test_name):
    script_path = SCRIPT_DIR / TEST_MODULES[test_name]
    module_name = f"oneclick_{test_name}"
    spec = importlib.util.spec_from_file_location(module_name, script_path)
    if spec is None or spec.loader is None:
        raise ImportError(f"Cannot load {script_path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    spec.loader.exec_module(module)
    return module


def run_test(test_name, *, save_results=True):
    module = load_test_module(test_name)
    print(f"\n=== Running {test_name} ({TEST_MODULES[test_name]}) ===")
    start = time.time()
    if getattr(module, "PREVIEW_ONLY", False):
        result = module.preview_waveform()
    else:
        result = module.run_ftj_test(save_results=save_results)
    elapsed = time.time() - start
    output_path = result.get("output_path") if isinstance(result, dict) else None
    if output_path:
        print(f"Saved: {output_path}")
    print(f"Finished {test_name} in {elapsed:.1f}s")
    return result


def parse_args():
    parser = argparse.ArgumentParser(description="Run one-click FTJ route tests.")
    parser.add_argument(
        "tests",
        nargs="*",
        choices=list(TEST_MODULES),
        help="Tests to run. Defaults to TEST_ORDER in route_config.py.",
    )
    parser.add_argument("--no-save", action="store_true", help="Run without saving Excel outputs.")
    return parser.parse_args()


def main():
    args = parse_args()
    tests = args.tests or TEST_ORDER
    for test_name in tests:
        run_test(test_name, save_results=not args.no_save)


if __name__ == "__main__":
    main()
