from pathlib import Path
import sys
import unittest
from unittest import mock


PKG_ROOT = Path(__file__).resolve().parents[1]
if str(PKG_ROOT) not in sys.path:
    sys.path.insert(0, str(PKG_ROOT))

from src.smu.data_processing import retrieve_variables
from src.smu.session import SMUSession
from src.smu.system_mode import (
    configure_measurement_list,
    execute_and_wait,
    linear_sweep_point_count,
    run_linear_voltage_sweep,
    validate_linear_sweep,
)
from src.smu.user_mode import source_voltage


class FakeKxci:
    def __init__(self, *, error="No error. (0)", status="1", buffers=None):
        self.commands = []
        self.error = error
        self.status = status
        self.buffers = buffers or {}

    def __call__(self, command):
        self.commands.append(command)
        if command == ":ERROR:LAST:GET":
            return self.error
        if command == "SP":
            return self.status
        if command.startswith("DO '"):
            variable = command.split("'", 2)[1]
            return self.buffers.get(variable, "")
        return "ACK"


class SmuSystemModeTests(unittest.TestCase):
    def test_linear_sweep_disables_every_available_channel_before_defining_use(self):
        fake = FakeKxci()

        variables = run_linear_voltage_sweep(
            fake,
            sweep_channel=2,
            bias_channel=1,
            available_channels=(1, 2, 3, 4),
            start=0,
            stop=1,
            step=0.1,
            timeout_s=1,
        )

        disable_positions = [fake.commands.index(f"CH{channel}") for channel in (1, 2, 3, 4)]
        first_definition = min(
            index
            for index, command in enumerate(fake.commands)
            if command.startswith(("CH1,", "CH2,"))
        )
        self.assertLess(max(disable_positions), first_definition)
        self.assertIn("ST 1, 1", fake.commands)
        self.assertIn("ST 2, 1", fake.commands)
        self.assertIn(":ERROR:LAST:CLEAR", fake.commands)
        self.assertIn(":ERROR:LAST:GET", fake.commands)
        self.assertIn("ME1", fake.commands)
        self.assertEqual(variables, ["ISWEEP", "VSWEEP", "IBIAS", "VBIAS"])

    def test_setup_error_aborts_and_disables_channels(self):
        fake = FakeKxci(error="KXCI command error. (-992)")

        with self.assertRaisesRegex(RuntimeError, "-992"):
            run_linear_voltage_sweep(
                fake,
                sweep_channel=2,
                bias_channel=1,
                available_channels=(1, 2, 3, 4),
                timeout_s=1,
            )

        self.assertNotIn("ME1", fake.commands)
        self.assertIn("ME4", fake.commands)
        for channel in (1, 2, 3, 4):
            self.assertIn(f"CH{channel}", fake.commands)

    def test_timeout_sends_me4(self):
        fake = FakeKxci(status="16")
        with mock.patch(
            "src.smu.system_mode.time.monotonic",
            side_effect=(0.0, 1.0),
        ):
            with self.assertRaises(TimeoutError):
                execute_and_wait(
                    fake,
                    timeout_s=0.5,
                    poll_interval_s=0,
                )
        self.assertEqual(fake.commands[-2:], ["MD", "ME4"])

    def test_sweep_point_limit_is_checked_before_hardware(self):
        self.assertEqual(linear_sweep_point_count(0, 1, 0.1), 11)
        with self.assertRaisesRegex(ValueError, "1024"):
            validate_linear_sweep(0, 2, 0.001)

    def test_measurement_names_allow_only_documented_suffix_extension(self):
        fake = FakeKxci()
        configure_measurement_list(fake, ("ABCDEF", "ABCDEFT"))
        self.assertIn("LI 'ABCDEF', 'ABCDEFT'", fake.commands)
        with self.assertRaisesRegex(ValueError, "T/S suffix"):
            configure_measurement_list(fake, ("ABCDEFG",))


class SmuDataTests(unittest.TestCase):
    def test_completed_buffers_must_match_expected_point_count(self):
        fake = FakeKxci(
            buffers={
                "V1": "N 0.0,N 0.5,N 1.0",
                "I1": "N 1e-9,N 2e-9,N 3e-9",
            }
        )
        data = retrieve_variables(
            fake,
            ("V1", "I1"),
            expected_point_count=3,
        )
        self.assertEqual(len(data), 3)
        self.assertEqual(list(data["V1_Status"]), ["N", "N", "N"])

    def test_empty_or_mismatched_buffers_fail(self):
        empty = FakeKxci(buffers={"V1": "", "I1": ""})
        with self.assertRaisesRegex(ValueError, "empty buffers"):
            retrieve_variables(empty, ("V1", "I1"))

        mismatch = FakeKxci(
            buffers={
                "V1": "N 0.0,N 1.0",
                "I1": "N 1e-9",
            }
        )
        with self.assertRaisesRegex(ValueError, "lengths do not match"):
            retrieve_variables(mismatch, ("V1", "I1"))

    def test_buffer_request_and_expected_count_must_be_unambiguous(self):
        fake = FakeKxci(buffers={"V1": "N 0.0"})
        with self.assertRaisesRegex(ValueError, "at least one"):
            retrieve_variables(fake, ())
        with self.assertRaisesRegex(ValueError, "duplicate"):
            retrieve_variables(fake, ("V1", "V1"))
        with self.assertRaisesRegex(ValueError, "positive integer"):
            retrieve_variables(fake, ("V1",), expected_point_count=1.5)

    def test_compliance_status_is_visible(self):
        fake = FakeKxci(buffers={"I1": "N 1e-9,C 2e-9"})
        with self.assertWarnsRegex(RuntimeWarning, "compliance"):
            data = retrieve_variables(fake, ("I1",), expected_point_count=2)
        self.assertEqual(list(data["I1_Status"]), ["N", "C"])


class SmuUserModeTests(unittest.TestCase):
    def test_source_voltage_checks_kxci_error_queue(self):
        fake = FakeKxci()
        source_voltage(fake, 1, 0.1, 1e-3, range_code=0)
        self.assertEqual(fake.commands, ["DV1, 0, 0.1, 0.001", ":ERROR:LAST:GET"])

    def test_fractional_range_code_is_rejected_instead_of_truncated(self):
        fake = FakeKxci()
        with self.assertRaisesRegex(ValueError, "integer"):
            source_voltage(fake, 1, 0.1, 1e-3, range_code=1.5)
        self.assertEqual(fake.commands, [])


class SmuSessionTests(unittest.TestCase):
    def test_connection_setup_failure_closes_all_visa_resources(self):
        class BadInstrument:
            @property
            def write_termination(self):
                return None

            @write_termination.setter
            def write_termination(self, value):
                raise RuntimeError("termination setup failed")

        client = mock.Mock()
        client._instrument_object = BadInstrument()
        with mock.patch("src.smu.session.Communications", return_value=client):
            session = SMUSession("TCPIP0::example::SOCKET")
            with self.assertRaisesRegex(RuntimeError, "termination setup failed"):
                session.connect()
        client.close.assert_called_once_with()
        self.assertIsNone(session.client)


if __name__ == "__main__":
    unittest.main()
