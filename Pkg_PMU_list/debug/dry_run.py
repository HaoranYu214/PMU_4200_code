# -*- coding: utf-8 -*-
r"""Generic dry-run runner for PMU test scripts.

Usage:
    python Pkg_PMU_list\debug\dry_run.py
    python Pkg_PMU_list\debug\dry_run.py Pkg_PMU_list\1C_ftj_test.py
    python Pkg_PMU_list\debug\dry_run.py --no-save Pkg_PMU_list\1C_ftj_test.py

The target script is executed, but instrument communication is intercepted and
printed instead of sent to the real PMU.
"""

from __future__ import annotations

import argparse
import runpy
import sys
import time
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
DEFAULT_SCRIPT = ROOT / "1C_ftj_test.py"

# Edit this path when you want to click "Run" on dry_run.py from the IDE.
TARGET_SCRIPT = DEFAULT_SCRIPT
NO_SAVE = True


class DryRunState:
    """Shared dry-run state for printed commands and fake query responses."""

    def __init__(self) -> None:
        self.command_index = 0

    def print_command(self, command: str) -> None:
        self.command_index += 1
        print(f"{self.command_index:03d}: {command}")

    def query_response(self, command: str) -> str:
        if ":PMU:TEST:STATUS?" in command:
            return "0"
        if ":PMU:DATA:COUNT?" in command:
            return "4"
        if ":PMU:DATA:GET" in command:
            return (
                "0.0,1e-6,0.0,0;"
                "0.5,2e-6,1e-6,0;"
                "1.0,3e-6,2e-6,0;"
                "0.0,1e-6,3e-6,0"
            )
        return "0"


class DummyInstrument:
    """Minimal stand-in for the VISA instrument object."""

    def __init__(self) -> None:
        self.timeout = 20000
        self.write_termination = "\0"
        self.read_termination = "\0"
        self.send_end = True

    def close(self) -> None:
        return


class DryRunCommunications:
    """Drop-in replacement for Communications that only prints commands."""

    def __init__(self, instrument_resource_string=None, state: DryRunState | None = None):
        self._instrument_resource_string = instrument_resource_string
        self._instrument_object = DummyInstrument()
        self._state = state or DryRunState()

    def connect(self, instrument_resource_string=None, timeout=None):
        if instrument_resource_string is not None:
            self._instrument_resource_string = instrument_resource_string
        self._state.print_command(
            f"# CONNECT {self._instrument_resource_string} timeout={timeout if timeout is not None else 20000}"
        )

    def disconnect(self):
        self._state.print_command("# DISCONNECT")

    def write(self, command: str):
        self._state.print_command(command)

    def read(self):
        return ""

    def query(self, command: str):
        self._state.print_command(command)
        return self._state.query_response(command)


class DryRunSession:
    """Drop-in replacement for PMUSession."""

    def __init__(self, instrument_resource, channels=(), timeout=None, write_termination="\0", read_termination="\0"):
        self.instrument_resource = instrument_resource
        self.channels = tuple(channels)
        self.timeout = timeout
        self.write_termination = write_termination
        self.read_termination = read_termination
        self._state = DryRunState()
        self.client = DryRunCommunications(instrument_resource, state=self._state)

    def __enter__(self):
        self.client.connect(timeout=self.timeout)
        self.client._instrument_object.write_termination = self.write_termination
        self.client._instrument_object.read_termination = self.read_termination
        self._state.print_command(
            f"# SET_TERMINATION write={self.write_termination!r} read={self.read_termination!r}"
        )
        return self

    def __exit__(self, exc_type, exc, tb):
        for channel in self.channels:
            self.client.query(f":PMU:OUTPUT:STATE {channel}, 0")
        self.client.disconnect()
        return False

    @property
    def query(self):
        return self.client.query


def install_dry_run_hooks(no_save=False):
    """Patch imported modules so test scripts print instead of connecting."""
    sys.path.insert(0, str(ROOT))
    import src.data_processing as data_processing
    import src.instrcomms as instrcomms
    import src.session as session
    import matplotlib.figure as mpl_figure
    import pandas as pd

    state = DryRunState()

    class SharedDryRunCommunications(DryRunCommunications):
        def __init__(self, instrument_resource_string=None):
            super().__init__(instrument_resource_string=instrument_resource_string, state=state)

    class SharedDryRunSession(DryRunSession):
        def __init__(self, instrument_resource, channels=(), timeout=None, write_termination="\0", read_termination="\0"):
            self.instrument_resource = instrument_resource
            self.channels = tuple(channels)
            self.timeout = timeout
            self.write_termination = write_termination
            self.read_termination = read_termination
            self._state = state
            self.client = SharedDryRunCommunications(instrument_resource)

    def fake_read_both_channels(_query, ch1, ch2, debug=False):
        import pandas as pd

        df1 = pd.DataFrame(
            {
                f"Voltage {ch1}": [0.0, 0.5, 1.0, 0.0],
                f"Current {ch1}": [1e-6, 2e-6, 3e-6, 1e-6],
                f"Timestamp {ch1}": [0.0, 1e-6, 2e-6, 3e-6],
                f"Status {ch1}": [0, 0, 0, 0],
            }
        )
        df2 = pd.DataFrame(
            {
                f"Voltage {ch2}": [0.0, 0.0, 0.0, 0.0],
                f"Current {ch2}": [1e-7, 1e-7, 1e-7, 1e-7],
                f"Timestamp {ch2}": [0.0, 1e-6, 2e-6, 3e-6],
                f"Status {ch2}": [0, 0, 0, 0],
            }
        )
        return df1, df2

    instrcomms.Communications = SharedDryRunCommunications
    session.Communications = SharedDryRunCommunications
    session.PMUSession = SharedDryRunSession
    data_processing.read_both_channels = fake_read_both_channels

    if no_save:
        class DummyExcelWriter:
            def __init__(self, path, *args, **kwargs):
                self.path = path

            def __enter__(self):
                print(f"# SKIP_SAVE ExcelWriter -> {self.path}")
                return self

            def __exit__(self, exc_type, exc, tb):
                return False

        def fake_save_channels_separate_excel(dfs, path):
            print(f"# SKIP_SAVE save_channels_separate_excel -> {path}")
            return True

        def fake_save_csv(df, path, sep=None, for_excel=True):
            print(f"# SKIP_SAVE save_csv -> {path}")
            return True

        def fake_save_excel(df, path):
            print(f"# SKIP_SAVE save_excel -> {path}")
            return True

        def fake_fig_save(self, fname, *args, **kwargs):
            print(f"# SKIP_SAVE figure.savefig -> {fname}")

        def fake_df_to_excel(self, excel_writer, *args, **kwargs):
            print(f"# SKIP_SAVE DataFrame.to_excel -> {excel_writer}")

        def fake_df_to_csv(self, path_or_buf=None, *args, **kwargs):
            print(f"# SKIP_SAVE DataFrame.to_csv -> {path_or_buf}")

        data_processing.save_channels_separate_excel = fake_save_channels_separate_excel
        data_processing.save_csv = fake_save_csv
        data_processing.save_excel = fake_save_excel
        mpl_figure.Figure.savefig = fake_fig_save

        pd.ExcelWriter = DummyExcelWriter
        pd.DataFrame.to_excel = fake_df_to_excel
        pd.DataFrame.to_csv = fake_df_to_csv

    time.sleep = lambda *_args, **_kwargs: None


def parse_args():
    """Parse the optional target script path."""
    parser = argparse.ArgumentParser(description="Dry-run any PMU test script.")
    parser.add_argument(
        "--no-save",
        action="store_true",
        help="Do not write any Excel/CSV/image outputs during the dry run.",
    )
    parser.add_argument(
        "script",
        nargs="?",
        default=None,
        help="Path to the test script to execute in dry-run mode.",
    )
    return parser.parse_args()


def dry_run_script(script, no_save=True):
    """Execute any PMU test script under dry-run hooks."""
    script_path = Path(script).resolve()
    if not script_path.exists():
        raise FileNotFoundError(f"Script not found: {script_path}")

    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    install_dry_run_hooks(no_save=no_save)
    print(f"DRY RUN: {script_path}")
    if no_save:
        print("DRY RUN MODE: saves disabled")
    runpy.run_path(str(script_path), run_name="__main__")


def main():
    """CLI entrypoint. If no CLI script is passed, use the editable default target."""
    args = parse_args()
    script = Path(args.script) if args.script else TARGET_SCRIPT
    no_save = args.no_save if args.script else NO_SAVE
    dry_run_script(script, no_save=no_save)


if __name__ == "__main__":
    main()

    
dry_run_script("Pkg_PMU_list/1C_two_stage_delay_read.py", no_save=True)
