# PMU reusable layer

- `session.py`: safe PMU connection and output shutdown lifecycle.
- `pmu_tests.py`: reusable 4225-PMU and Segment Arb commands.
- `current_range.py`: fixed-range assessment and automatic retry helpers.
- `data_processing.py`: PMU buffer parsing and analysis helpers.
- `plotting_utils.py`: common PMU plotting utilities.
- `fet_three_terminal_common.py`: shared execution helpers for three-terminal
  FET pulse experiments.

The PMU package imports the common VISA implementation from
`src/transport.py`. Experiment-specific pulse sequences stay in runnable
entries under `Pkg_PMU_list`.
