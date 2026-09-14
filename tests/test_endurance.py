from pathlib import Path
import sys
import unittest

import numpy as np
import pandas as pd


REPO_ROOT = Path(__file__).resolve().parents[1]
SRC_ROOT = REPO_ROOT / "src"
for path in (SRC_ROOT, REPO_ROOT):
    if str(path) not in sys.path:
        sys.path.insert(0, str(path))

from measurements.pmu.fe_cap import endurance
from keithley4200.pmu.pmu_tests import (
    MAX_SEGMENTS_PER_SEQUENCE,
    _validate_segment_arb_storage,
    configure_segARB_sequence,
)


class SegmentArbMeasurementWindowTests(unittest.TestCase):
    def test_multiple_sequences_share_limit_per_channel_not_between_channels(self):
        segment_count = MAX_SEGMENTS_PER_SEQUENCE // 2
        sequence = (1, [0.0] * segment_count, [0.0] * segment_count,
                    [1e-5] * segment_count)
        second_sequence = (2, sequence[1], sequence[2], sequence[3])

        # Both channels may independently use all 2048 stored segments.
        _validate_segment_arb_storage(
            {1: [sequence, second_sequence], 2: [sequence, second_sequence]}
        )

        extra = (3, [0.0], [0.0], [1e-5])
        with self.assertRaisesRegex(ValueError, "per-channel 4225-PMU limit"):
            _validate_segment_arb_storage(
                {1: [sequence, second_sequence, extra], 2: [sequence]}
            )

    def test_sequence_above_hardware_segment_limit_is_rejected(self):
        commands = []
        values = [0.0] * (MAX_SEGMENTS_PER_SEQUENCE + 1)
        with self.assertRaisesRegex(ValueError, "4225-PMU limit"):
            configure_segARB_sequence(
                commands.append,
                1,
                1,
                values,
                values,
                [1e-5] * len(values),
                [0] * len(values),
            )
        self.assertEqual(commands, [])

    def test_unmeasured_segments_are_sent_with_zero_windows(self):
        commands = []
        configure_segARB_sequence(
            commands.append,
            1,
            1,
            [0.0, 0.0],
            [0.0, 1.0],
            [1e-5, 2e-5],
            [0, 2],
        )

        self.assertIn(
            ":PMU:SARB:SEQ:MEAS:START 1, 1, 0.00e+00, 0.00e+00",
            commands,
        )
        self.assertIn(
            ":PMU:SARB:SEQ:MEAS:STOP 1, 1, 0.00e+00, 2.00e-05",
            commands,
        )

    def test_invalid_measured_window_is_rejected_before_commands_are_sent(self):
        commands = []
        with self.assertRaisesRegex(ValueError, "outside its"):
            configure_segARB_sequence(
                commands.append,
                1,
                1,
                [0.0],
                [1.0],
                [1e-5],
                [2],
                [0.0],
                [1.0],
            )
        self.assertEqual(commands, [])


class EnduranceConfigurationTests(unittest.TestCase):
    def test_cycle_segments_are_explicitly_unmeasured(self):
        for configs in endurance.make_cycle_seq_configs().values():
            config = configs[0]
            self.assertEqual(config[4], [0, 0, 0, 0])
            self.assertEqual(config[5], [0.0, 0.0, 0.0, 0.0])
            self.assertEqual(config[6], [0.0, 0.0, 0.0, 0.0])

    def test_cycle_targets_are_converted_to_increments(self):
        self.assertEqual(
            endurance.build_cycle_schedule([1, 10, 100, 1000]),
            [(1, 1), (10, 9), (100, 90), (1000, 900)],
        )
        full_schedule = endurance.build_cycle_schedule(endurance.cycle_counts)
        self.assertEqual(sum(increment for _target, increment in full_schedule), 10_000_000)

    def test_cycle_targets_must_increase(self):
        with self.assertRaisesRegex(ValueError, "strictly increasing"):
            endurance.build_cycle_schedule([1, 10, 10])

    def test_pv2_measures_only_the_two_readback_loops(self):
        config = endurance.make_pv2_seq_configs()[endurance.CH1][0]
        self.assertEqual(config[4], [0, 0, 0, 0, 0, 2, 2, 2, 2, 2])
        self.assertEqual(config[3][4], endurance.params_pv2["delay_time"])

    def test_cycle_block_emits_valid_zero_windows_and_increment(self):
        commands = []

        def fake_query(command):
            commands.append(command)
            return "0"

        endurance.run_cycle_block(fake_query, 9)

        self.assertIn(":PMU:SARB:WFM:SEQ:LIST 1, 1, 9", commands)
        measurement_stop_commands = [
            command for command in commands if ":PMU:SARB:SEQ:MEAS:STOP" in command
        ]
        self.assertEqual(len(measurement_stop_commands), 2)
        for command in measurement_stop_commands:
            self.assertTrue(command.endswith("0.00e+00, 0.00e+00, 0.00e+00, 0.00e+00"))

    def test_pv2_analysis_splits_and_integrates_each_loop(self):
        loop_voltage = np.array([0.0, 1.0, 0.0, -1.0, 0.0])
        voltage = np.concatenate([loop_voltage, loop_voltage])
        point_count = len(voltage)
        time = np.arange(point_count, dtype=float) * 1e-6
        current = np.linspace(1e-6, 2e-6, point_count)
        status = np.zeros(point_count, dtype=int)
        ch1 = pd.DataFrame(
            {
                "Voltage 1": voltage,
                "Current 1": current,
                "Timestamp 1": time,
                "Status 1": status,
            }
        )
        ch2 = pd.DataFrame(
            {
                "Voltage 2": np.zeros(point_count),
                "Current 2": -current,
                "Timestamp 2": time,
                "Status 2": status,
            }
        )

        result = endurance.analyze_pv2_readback(ch1, ch2)

        self.assertEqual(len(result["i1_delay"]), 5)
        self.assertEqual(len(result["i1_no_delay"]), 5)
        self.assertEqual(len(result["i2_delay"]), 5)
        self.assertEqual(len(result["i2_no_delay"]), 5)
        self.assertTrue(np.isfinite(result["i2_loops"].to_numpy()).all())


if __name__ == "__main__":
    unittest.main()
