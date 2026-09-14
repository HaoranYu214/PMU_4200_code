# Keithley4200Measurement

Python measurement library and experiment collection for the Keithley
4200A-SCS. The repository deliberately separates reusable instrument code,
editable experiment recipes, multi-step workflows, and vendor references.

## Layout

- `src/keithley4200/pmu`: reusable PMU commands, sessions, processing, plotting.
- `src/keithley4200/smu`: reusable SMU system/user-mode commands and RPM routing.
- `measurements/pmu`: editable PMU experiment recipes grouped by device/test type.
  See the [FET entry guide](measurements/pmu/fet/README.md) for pulse, program/read, and sweep entry points.
- `measurements/smu`: editable SMU experiment recipes.
- `workflows`: multi-test measurement packages and one-click routes.
- `src/keithley4200/tools`: offline preview and dry-run helpers.
- `reference`: official examples and manuals; not maintained as production code.
- `tests`: hardware-free unit tests using fake instrument responses.

Experiment files remain directly runnable. Each one adds this repository's
`src` directory to `sys.path`, so an editable install is optional. For normal
library use, install once with `python -m pip install -e .` and import from
`keithley4200`.

Running an experiment can contact the instrument. Unit tests and waveform
preview helpers are hardware-free.


## Saved results

Automatic measurement outputs use short names, with voltage before timing:

```text
PV2_03.000000V_tr250us_td1000us_r001.xlsx
PV2_03.000000V_tr250us_td1000us_r001_i1.png
PV2_03.000000V_tr250us_td1000us_r001_i2.png
PV2_03.000000V_tr250us_td1000us_r002.xlsx
PUNDtri_03.500000V_tr250us_td1000us_r001.xlsx
PUND_03.500000V_tr250us_td1000us_tw50us_r001.xlsx
```

- `tr`: rise time; `td`: delay; `tw`: dwell/pulse width. Times use microseconds,
  including fractions such as `0.1us`. PV2 and PUND always include delay.
- Voltage uses at least two integer digits and six fixed decimal places, so
  positive labels below 100 V sort as `03.000000V, 03.005000V, 03.010000V`.
  Negative bias labels retain the minus sign; alphabetical ordering is not a
  signed numerical sort, and magnitudes of 100 V or more need numeric sorting.
- Each acquisition reserves the first available `r001, r002, ...` for the entire
  workbook/image group. Existing companions count as occupied, even if the
  workbook is missing. Full parameters and an ISO `saved_at` timestamp remain in
  the parameter-bearing workbooks; display labels are not a parameter archive.
- Reservations use exclusive file creation in `.reservations/`, coordinating
  separate processes using the shared helper. Keep this internal ledger with
  the data directory. Interrupted or failed runs can leave gaps intentionally.
- Standalone measurements and batch workflows use their configured directories
  directly. No automatic time subdirectory is added; existing test/stage folders
  are retained. Summary outputs include a per-invocation time, for example
  `map_summary_20260911_103449_r001.xlsx` and its matching `_live.csv`.
  A `time` column records the same ISO timestamp in the summary. Live updates
  reuse that run's CSV; a later invocation reserves a new summary name. The
  numeric suffix also prevents collisions between same-second runs.
- Existing data is not renamed. Script/module names and measurement waveforms
  are unchanged. Explicit low-level save/preview functions still honor the exact
  path supplied by their caller; automatic naming belongs at the acquisition
  entry point.

The common helpers live in `src/keithley4200/output.py`. Reserve a stem once with
`reserve_output_stem(directory, measurement_name(...))`, then append extensions
or plot suffixes to that same stem for every companion output.

## Offline dry runs

`python src/keithley4200/tools/dry_run.py --no-save measurements/pmu/fe_cap/PV2.py` intercepts
instrument communication and generates synthetic data from the configured
Segment Arb or pulse commands. Segment Arb skips unmeasured segments, retains
their elapsed time, returns one point per spot measurement, and samples each
waveform segment at 32 points. The ordinary data reader still parses the
responses, including block reads and pulse High/Low fields.

These data test software flow; they do not simulate device physics or the
instrument sample rate. Averaged acquisition modes and acquisitions exceeding
65,536 synthetic points raise explicit errors. Run the helper in a separate
process: its communication, sleep, and optional save hooks are process-wide.

After an editable install, the equivalent module command is
`python -m keithley4200.tools.dry_run --no-save measurements/pmu/fe_cap/PV2.py`.
Preview helpers are imported from `keithley4200.tools.waveform_preview`.
Explicit relative script paths are resolved from the current working directory;
the default RV2 script is located relative to the source checkout. A packaged
installation without the experiment files requires an explicit script path.

## 中文使用导航

- [测量入口](measurements/README.md)：选测试、改器件参数。
- [工作流](workflows/README.md)：批量扫描和多步测量。
- [公共源码](src/README.md)：包结构及工具。
- [参数与模式速查](reference/manuals/PARAMETER_LIMITS.md)：电压、电流档、SMU 限流、0/1/2 等模式码及来源页码。
- [测试说明](tests/README.md)：离线验证方式。

参数注释是使用说明，不会替代验证器件安全范围；本次文档整理不改变实验参数。

参数依据：[手册限制与模式速查](reference/manuals/PARAMETER_LIMITS.md)。
