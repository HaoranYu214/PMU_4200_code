# SMU I-V experiments

Scripts in this folder keep experiment parameters near the top while reusing
connection, command, execution, and data helpers from
`src/keithley4200/smu` alongside the PMU library in
`src/keithley4200/pmu`.

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
2. `PlotData`: `CommandedVoltage`, then `V1`, `I1`, `J1_A_per_cm2`,
   `AbsI1`, `AbsJ1_A_per_cm2`, followed by the corresponding V2/I2/J2
   columns. It contains no status columns and is intended for plotting and
   scripted extraction.
3. `Parameters`: the complete experiment configuration.

Every IV entry exposes `DEVICE_AREA_CM2`. Raw current remains in amperes; the
processed current density is `J = I / DEVICE_AREA_CM2` in A/cm². Default PNGs
plot J-V and log(abs(J))-V. Major log tick labels display physical values such
as `1e-2` and `1e-1` A/cm² rather than their numerical log10 exponents.

## Memristor 分段扫描

[segmented_voltage_sweep_Memristor.py](segmented_voltage_sweep_Memristor.py) 逐次改变正峰值，配合固定负峰值。POSITIVE_PEAK_STEP 决定峰值之间的变化，SEGMENT_STEP 决定每段内的采样步长。最终展开列表每次最多 4096 点。

参数依据：[手册限制与模式速查](../../../reference/manuals/PARAMETER_LIMITS.md)。
