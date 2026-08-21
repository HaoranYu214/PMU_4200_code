# SMU I-V experiments

Scripts in this folder keep experiment parameters near the top while reusing
connection, command, execution, and data helpers from
`Pkg_PMU_list/src/smu` alongside the PMU library.

- `linear_voltage_sweep.py`: System Mode two-channel voltage sweep.
- `segmented_voltage_sweep.py`: arbitrary turning-point path such as
  `0 -> V1 -> 0 -> V2 -> 0`.
- `user_mode_spot.py`: User Mode source-voltage/measure-current spot test.

System Mode entries expose `AVAILABLE_CHANNELS`. List every SMU installed and
mapped in KCon: the library disables that complete set before defining the
active sweep/bias channels. Completed buffers must be nonempty, equal-length,
and match the programmed point count before they are saved.
