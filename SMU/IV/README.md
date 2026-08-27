# SMU I-V experiments

Scripts in this folder keep experiment parameters near the top while reusing
connection, command, execution, and data helpers from
`src/smu` alongside the PMU library in `src/pmu`.

- `linear_voltage_sweep.py`: System Mode two-channel voltage sweep.
- `segmented_voltage_sweep.py`: arbitrary turning-point path such as
  `0 -> V1 -> 0 -> V2 -> 0`.
- `user_mode_spot.py`: User Mode source-voltage/measure-current spot test.

System Mode entries expose `AVAILABLE_CHANNELS`. List every SMU installed and
mapped in KCon: the library disables that complete set before defining the
active sweep/bias channels. Completed buffers must be nonempty, equal-length,
and match the programmed point count before they are saved.

Every runnable entry also exposes `SMU_CONNECTIONS`. Use `"direct"` for an
SMU wired straight to a probe, or `"rpm:PMUN-C"` for an SMU whose probe path
passes through the RPM attached to that PMU channel. The current instrument
uses RPM paths for SMU1/SMU2 (`PMU1-1`/`PMU1-2`) and direct paths for
SMU3/SMU4.

Saved workbooks use three worksheets in a stable order:

1. `Raw`: commanded values, every measured buffer, and the original KXCI
   status columns.
2. `PlotData`: `CommandedVoltage`, `V1`, `I1`, `AbsI1`, `V2`, `I2`, and
   `AbsI2`, with no status columns, for plotting and scripted extraction.
3. `Parameters`: the complete experiment configuration.

The logarithmic current PNG plots `abs(I)` on a logarithmic Y axis. Major tick
labels display physical current values such as `1e-7` and `1e-6` A rather than
the numerical `log10` values `-7` and `-6`.
