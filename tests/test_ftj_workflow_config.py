from importlib import reload
from pathlib import Path
from types import SimpleNamespace
import inspect
import sys
import tempfile
import unittest
from unittest import mock


REPO_ROOT = Path(__file__).resolve().parents[1]
SRC_ROOT = REPO_ROOT / "src"
for path in (SRC_ROOT, REPO_ROOT):
    if str(path) not in sys.path:
        sys.path.insert(0, str(path))

from keithley4200.pmu.pmu_tests import (
    MAX_SEGMENTS_PER_SEQUENCE,
    _normalize_seg_arb_measurements,
)
from measurements.pmu.ftj import (
    ftj_Identical_V1,
    ftj_Identical_V2,
    ftj_ISPP_V1,
    ftj_ISPP_V2,
    ftj_MRD,
    ftj_PWM,
    ftj_RV1,
    ftj_RV2,
    ftj_endurance,
)
from workflows import ftj_package1


def assert_configs_are_pmu_valid(test_case, module):
    for configs in module.seq_configs.values():
        total_segments = sum(len(config[3]) for config in configs)
        test_case.assertLessEqual(total_segments, MAX_SEGMENTS_PER_SEQUENCE)
        for config in configs:
            _normalize_seg_arb_measurements(
                config[3], config[4], config[5], config[6]
            )


class FtjWorkflowConfigTests(unittest.TestCase):
    def test_every_ftj_entry_has_one_rebuildable_parameter_interface(self):
        modules = (
            ftj_RV1,
            ftj_RV2,
            ftj_PWM,
            ftj_Identical_V1,
            ftj_Identical_V2,
            ftj_ISPP_V1,
            ftj_ISPP_V2,
            ftj_MRD,
        )
        for imported_module in modules:
            with self.subTest(module=imported_module.__name__):
                module = reload(imported_module)
                self.assertTrue(module.params)
                self.assertIn("base_v", module.params)
                module.configure_measurement(params_override=dict(module.params))
                self.assertEqual(module.BASE_V, float(module.params["base_v"]))
                for config in module.seq_configs[module.CH1]:
                    self.assertEqual(config[1][0], module.BASE_V)
                    self.assertEqual(config[2][-1], module.BASE_V)
                assert_configs_are_pmu_valid(self, module)

    def test_rv2_rebuilds_offset_scan_read_level_and_channels(self):
        module = reload(ftj_RV2)
        configured = dict(module.params)
        configured.update(
            base_v=0.25,
            offset_v=-1.5,
            vp=1.0,
            write_level_step=0.5,
            read_v=-0.25,
            scan_cycles=2,
        )
        module.configure_measurement(
            params_override=configured,
            channels=(3, 4),
            current_ranges={3: 1e-5, 4: 2e-5},
        )

        self.assertEqual(module.SCAN_LEVELS, [1.0, 0.5, 0.0, -0.5, -1.0, -0.5, 0.0, 0.5, 1.0, 0.5, 0.0, -0.5, -1.0, -0.5, 0.0, 0.5, 1.0])
        self.assertEqual(module.SCAN_VOLTAGES[0], -0.5)
        self.assertEqual(module.SCAN_VOLTAGES[4], -2.5)
        first_scan = module.ch1_scan_configs[0]
        self.assertEqual(first_scan[1][0], 0.25)
        self.assertEqual(first_scan[1][1], -0.5)
        self.assertEqual(first_scan[1][5], -0.25)
        self.assertEqual(first_scan[2][-1], 0.25)
        self.assertEqual(set(module.seq_configs), {3, 4})
        self.assertEqual(module.CURRENT_RANGES, {3: 1e-5, 4: 2e-5})
        assert_configs_are_pmu_valid(self, module)

    def test_pwm_rebuilds_widths_repeat_metadata_and_sequence_list(self):
        module = reload(ftj_PWM)
        configured = dict(module.params)
        configured.update(
            write_base_dwell=2e-6,
            width_multipliers=[1, 3, 10],
            repeat_count=2,
            read_v=-0.8,
        )
        module.configure_measurement(params_override=configured)

        for actual, expected in zip(module.WRITE_WIDTHS, [2e-6, 6e-6, 20e-6]):
            self.assertAlmostEqual(actual, expected)
        self.assertEqual(len(module.PWM_SINGLE_RUN_STEPS), 6)
        self.assertEqual(len(module.PWM_STEPS), 12)
        self.assertEqual(module.SEQ_LIST[module.CH1], [(module.FULL_PWM_SEQ_ID, 2)])
        self.assertFalse(module.SEGARB_OPTIONS["ENABLE_LLEC"])
        assert_configs_are_pmu_valid(self, module)

    def test_identical_rebuilds_counts_voltages_and_stored_segments(self):
        module = reload(ftj_Identical_V1)
        configured = dict(module.params)
        configured.update(
            write_positive_v=1.2,
            write_negative_v=-3.4,
            read_v=-0.6,
            positive_repeat_count=2,
            negative_repeat_count=3,
            sequence_cycle_count=2,
        )
        module.configure_measurement(params_override=configured)

        self.assertEqual(len(module.SINGLE_CYCLE_PLAN), 10)
        self.assertEqual(len(module.SEQ_PLAN), 20)
        self.assertEqual(len(module.seq_configs[module.CH1]), 2)
        self.assertIn(1.2, module.ch1_expanded_configs[0][1])
        self.assertIn(-3.4, module.ch1_expanded_configs[0][1])
        assert_configs_are_pmu_valid(self, module)

    def test_oversized_workflow_config_is_rejected_during_rebuild(self):
        module = reload(ftj_Identical_V1)
        configured = dict(module.params)
        configured.update(
            positive_repeat_count=100,
            negative_repeat_count=100,
            sequence_cycle_count=2,
        )
        with self.assertRaisesRegex(ValueError, "stored Segment Arb segments"):
            module.configure_measurement(params_override=configured)

    def test_mrd_rebuilds_voltage_levels_cycles_and_measurement_windows(self):
        module = reload(ftj_MRD)
        configured = dict(module.params)
        configured.update(
            base_v=0.2,
            reference_v=-4.0,
            write_voltages=[0.5, 1.0],
            read_v=-0.75,
            cycles_per_level=3,
        )
        module.configure_measurement(params_override=configured)

        self.assertEqual(module.SEQ_PLAN, [(1, 3), (2, 3)])
        self.assertEqual(module.ch1_configs[0][1][0], 0.2)
        self.assertEqual(module.ch1_configs[0][1][1], -4.0)
        self.assertEqual(module.ch1_configs[0][2][-1], 0.2)
        expected = module.expected_cycle_table(module.SEQ_METADATA)
        self.assertEqual(len(expected), 12)
        self.assertEqual(expected["ReadType"].tolist()[:2], ["RefStateRead", "AfterWriteRead"])
        assert_configs_are_pmu_valid(self, module)

    def test_runtime_save_defaults_are_not_captured_at_import(self):
        for module in (
            ftj_RV1,
            ftj_RV2,
            ftj_PWM,
            ftj_Identical_V1,
            ftj_Identical_V2,
            ftj_ISPP_V1,
            ftj_ISPP_V2,
            ftj_MRD,
        ):
            parameters = inspect.signature(module.run_ftj_test).parameters
            self.assertIsNone(parameters["save_dir"].default)
            self.assertIsNone(parameters["file_stem"].default)

    def test_package_imports_maintained_modules_and_applies_full_config(self):
        self.assertEqual(
            ftj_package1.FTJ_TESTS["rv2"]["module"],
            "measurements.pmu.ftj.ftj_RV2",
        )
        module = reload(ftj_RV2)
        ftj_package1.configure_test("rv2", module)
        self.assertEqual(module.params, ftj_package1.FTJ_TESTS["rv2"]["params"])
        self.assertEqual(module.params["base_v"], 0.0)
        self.assertEqual(module.SAVE_DIR, ftj_package1.FTJ_TESTS["rv2"]["save_dir"])
        assert_configs_are_pmu_valid(self, module)

    def test_package_runs_configured_stages_in_order_without_hardware(self):
        calls = []

        class FakeMeasurement:
            def __init__(self, name):
                self.name = name

            def configure_measurement(self, **kwargs):
                calls.append(("configure", self.name, kwargs))

            def run_ftj_test(self, **kwargs):
                calls.append(("run", self.name, kwargs))
                return {"output_path": None}

        modules = {
            test_name: FakeMeasurement(test_name)
            for test_name in ftj_package1.FTJ_TESTS
        }
        original_settle = ftj_package1.STAGE_SETTLE_TIME_S
        try:
            ftj_package1.STAGE_SETTLE_TIME_S = 0
            with tempfile.TemporaryDirectory() as temp_dir:
                batch_base = Path(temp_dir)
                test_configs = {
                    name: {**config, "save_dir": batch_base / name}
                    for name, config in ftj_package1.FTJ_TESTS.items()
                }
                with mock.patch.object(ftj_package1, "BASE_SAVE_DIR", batch_base), mock.patch.object(
                    ftj_package1, "FTJ_TESTS", test_configs
                ):
                    results = ftj_package1.run_package(modules=modules)
        finally:
            ftj_package1.STAGE_SETTLE_TIME_S = original_settle

        self.assertEqual(list(results), ftj_package1.RUN_ORDER)
        self.assertEqual(
            [name for action, name, _kwargs in calls if action == "run"],
            ftj_package1.RUN_ORDER,
        )
        configured = {
            name: kwargs
            for action, name, kwargs in calls
            if action == "configure"
        }
        for test_name in ftj_package1.RUN_ORDER:
            self.assertEqual(
                configured[test_name]["params_override"],
                ftj_package1.FTJ_TESTS[test_name]["params"],
            )

    def test_package_previews_all_stages_with_one_final_show(self):
        calls = []

        class FakeMeasurement:
            def __init__(self, name):
                self.name = name

            def configure_measurement(self, **kwargs):
                calls.append(("configure", self.name, kwargs))

            def preview_waveform(self, **kwargs):
                calls.append(("preview", self.name, kwargs))
                return self.name

        modules = {
            test_name: FakeMeasurement(test_name)
            for test_name in ftj_package1.FTJ_TESTS
        }
        with mock.patch("matplotlib.pyplot.show") as show_mock:
            figures = ftj_package1.preview_package(modules=modules)

        self.assertEqual(figures, ftj_package1.RUN_ORDER)
        preview_calls = [entry for entry in calls if entry[0] == "preview"]
        self.assertEqual([entry[1] for entry in preview_calls], ftj_package1.RUN_ORDER)
        self.assertTrue(all(entry[2]["show"] is False for entry in preview_calls))
        show_mock.assert_called_once_with()

    def test_endurance_merges_overrides_and_uses_runtime_save_path(self):
        calls = []
        fake = SimpleNamespace(
            params={"vp": 4.0, "read_v": -1.0},
            configure_measurement=lambda **kwargs: calls.append(("configure", kwargs)),
            run_ftj_test=lambda **kwargs: (
                calls.append(("run", kwargs)) or {"output_path": None}
            ),
        )
        with tempfile.TemporaryDirectory() as temp_dir:
            result = ftj_endurance.run_endurance(
                module=fake,
                param_overrides={"vp": 3.0},
                save_dir=Path(temp_dir),
                loop_count=2,
                save_every_run=False,
                file_stem_prefix="offline",
            )

        configured = calls[0][1]["params_override"]
        self.assertEqual(configured, {"vp": 3.0, "read_v": -1.0})
        run_calls = [kwargs for kind, kwargs in calls if kind == "run"]
        self.assertEqual(len(run_calls), 2)
        self.assertEqual(run_calls[0]["file_stem"], "offline")
        self.assertEqual(Path(run_calls[0]["save_dir"]), Path(temp_dir))
        self.assertIn("time", result["summary_df"].columns)
        self.assertEqual(result["summary_df"]["status"].tolist(), ["ok", "ok"])


if __name__ == "__main__":
    unittest.main()
