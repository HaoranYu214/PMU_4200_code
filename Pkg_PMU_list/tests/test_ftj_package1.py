from pathlib import Path
import importlib.util
from types import SimpleNamespace
import sys
import tempfile
import unittest
from unittest import mock

import pandas as pd


REPO_ROOT = Path(__file__).resolve().parents[2]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from src.smu.system_mode import build_segmented_voltage_path, run_list_voltage_sweep


PACKAGE_PATH = (
    REPO_ROOT / "Pkg_PMU_list" / "Programed_measuremnts" / "FTJ package1.py"
)


def load_package():
    spec = importlib.util.spec_from_file_location("ftj_package1_test", PACKAGE_PATH)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class FakeKxci:
    def __init__(self):
        self.commands = []

    def __call__(self, command):
        self.commands.append(command)
        return "ACK"


class FtjPackage1Tests(unittest.TestCase):
    def test_defaults_to_preview_and_current_counts_are_valid(self):
        package = load_package()
        modules = {
            "iv": SimpleNamespace(
                build_segmented_voltage_path=build_segmented_voltage_path
            )
        }
        counts = package.validate_package_config(modules)

        self.assertTrue(package.PREVIEW_ONLY)
        self.assertEqual(counts["iv_point_count"], 201)
        self.assertEqual(counts["sweep_parameter_points"], 50)

    def test_package_applies_configs_and_runs_five_stages_in_order(self):
        package = load_package()
        calls = []

        pv2 = SimpleNamespace(
            params={},
            SEGARB_OPTIONS={},
            main=lambda: calls.append("pv2"),
        )
        sweep_pv2 = SimpleNamespace(params={}, SEGARB_OPTIONS={})
        sweep_pund = SimpleNamespace(params={}, SEGARB_OPTIONS={})
        sweep = SimpleNamespace(
            PV2=sweep_pv2,
            PUND_tri=sweep_pund,
            SEGARB_OPTIONS={},
            run_sweep=lambda: (
                calls.append("sweep")
                or pd.DataFrame(
                    [{"status": "ok"}] * 100
                )
            ),
        )
        iv = SimpleNamespace(
            build_segmented_voltage_path=build_segmented_voltage_path,
            main=lambda: calls.append("iv"),
        )
        fake_modules = {"pv2": pv2, "pund": sweep_pund, "sweep": sweep, "iv": iv}

        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            package.BASE_SAVE_DIR = root
            package.STAGE_SETTLE_TIME_S = 0
            package.SEGMENTED_IV["settle_time_s"] = 0
            for name, config in package.STAGE_CONFIGS.items():
                key = "save_root" if name == "pv_pund_sweep" else "save_dir"
                config[key] = root / name

            with mock.patch.object(package, "load_test_modules", return_value=fake_modules):
                result = package.run_package()

        self.assertEqual(
            calls,
            ["pv2", "sweep", "pv2", "iv", "iv", "iv", "pv2"],
        )
        self.assertEqual(result["stage"].tolist(), package.RUN_ORDER)
        self.assertTrue((result["status"] == "ok").all())
        self.assertEqual(iv.TURNING_POINTS, [0.0, 5.0, 0.0, -5.0, 0.0])
        self.assertIsNone(iv.PARAMS["timeout_s"])

    def test_preview_generates_waveforms_without_opening_sessions(self):
        package = load_package()

        def forbid_session(*_args, **_kwargs):
            raise AssertionError("Preview attempted to create an instrument session.")

        with tempfile.TemporaryDirectory() as temp_dir:
            package.BASE_SAVE_DIR = Path(temp_dir)
            modules = package.load_test_modules()
            modules["pv2"].PMUSession = forbid_session
            modules["sweep"].PV2.PMUSession = forbid_session
            modules["sweep"].PUND_tri.PMUSession = forbid_session
            modules["iv"].SMUSession = forbid_session
            with mock.patch.object(package, "load_test_modules", return_value=modules):
                outputs = package.preview_package()
            self.assertEqual(len(outputs), 8)
            self.assertTrue(all(path.exists() for path in outputs))

    def test_oversized_list_is_rejected_before_any_hardware_command(self):
        fake = FakeKxci()
        with self.assertRaisesRegex(ValueError, "4096"):
            run_list_voltage_sweep(
                fake,
                values=[0.0] * 4097,
                sweep_channel=1,
                bias_channel=2,
            )
        self.assertEqual(fake.commands, [])


if __name__ == "__main__":
    unittest.main()
