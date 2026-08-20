import ast
from contextlib import redirect_stdout
import io
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch


PKG_ROOT = Path(__file__).resolve().parents[1]
NLS_ROOT = PKG_ROOT / "NLS"
for path in (PKG_ROOT, NLS_ROOT):
    if str(path) not in sys.path:
        sys.path.insert(0, str(path))

import NLS_1C_switch as nls_switch
from src import pmu_tests


REQUIRED_OPTIONS = {
    "ENABLE_CONNECTION_COMP",
    "ENABLE_LOAD_CONFIG",
    "LOAD_RESISTANCE",
    "ENABLE_LLEC",
}


def assignment_literal(tree, name):
    for node in tree.body:
        if not isinstance(node, ast.Assign):
            continue
        if any(isinstance(target, ast.Name) and target.id == name for target in node.targets):
            return ast.literal_eval(node.value)
    raise AssertionError(f"{name} assignment is missing")


class NlsSequenceTests(unittest.TestCase):
    def test_promoted_sequence_matches_the_legacy_hardware_helper(self):
        with patch.object(pmu_tests, "execute_segARB_test") as execute:
            with redirect_stdout(io.StringIO()):
                pmu_tests.hy_NISswitch_segARB(
                    lambda _command: None,
                    1,
                    2,
                    nls_switch.PARAMS,
                )
        legacy_configs = execute.call_args.args[2]
        self.assertEqual(
            nls_switch.make_nls_seq_configs(1, 2, nls_switch.PARAMS),
            legacy_configs,
        )

    def test_sequence_has_the_expected_two_aligned_14_segment_channels(self):
        configs = nls_switch.make_nls_seq_configs(1, 2, nls_switch.PARAMS)
        self.assertEqual(set(configs), {1, 2})
        self.assertEqual(len(configs[1]), 1)
        self.assertEqual(len(configs[2]), 1)
        for channel in (1, 2):
            config = configs[channel][0]
            self.assertEqual(len(config[1]), 14)
            self.assertEqual(len(config[2]), 14)
            self.assertEqual(len(config[3]), 14)
            self.assertEqual(len(config[4]), 14)

    def test_preview_saves_without_opening_a_pmu_session(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "nls_preview.png"
            nls_switch.preview_nls_waveform(nls_switch.PARAMS, output)
            self.assertTrue(output.exists())
            self.assertGreater(output.stat().st_size, 0)


class NlsEntryWiringTests(unittest.TestCase):
    def test_single_entry_exposes_options_and_passes_them_to_execute(self):
        path = NLS_ROOT / "NLS_1C_switch.py"
        source = path.read_text(encoding="utf-8")
        tree = ast.parse(source, filename=str(path))
        options = assignment_literal(tree, "SEGARB_OPTIONS")
        self.assertTrue(REQUIRED_OPTIONS.issubset(options))

        execute_calls = [
            node
            for node in ast.walk(tree)
            if isinstance(node, ast.Call)
            and isinstance(node.func, ast.Name)
            and node.func.id == "execute_segARB_test"
        ]
        self.assertEqual(len(execute_calls), 1)
        self.assertIn("options", {keyword.arg for keyword in execute_calls[0].keywords})
        self.assertIn("def preview_nls_waveform", source)
        self.assertIn("PREVIEW_ONLY", source)

    def test_sweep_entry_exposes_options_and_preview(self):
        path = NLS_ROOT / "NLS_1C_switch_list.py"
        source = path.read_text(encoding="utf-8")
        tree = ast.parse(source, filename=str(path))
        options = assignment_literal(tree, "SEGARB_OPTIONS")
        self.assertTrue(REQUIRED_OPTIONS.issubset(options))
        self.assertIn("segarb_options=SEGARB_OPTIONS", source)
        self.assertIn("def preview_sweep_waveform", source)
        self.assertIn("PREVIEW_ONLY", source)


if __name__ == "__main__":
    unittest.main()
