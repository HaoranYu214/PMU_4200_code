# PMU reusable layer

- `session.py`: safe PMU connection and output shutdown lifecycle.
- `pmu_tests.py`: reusable 4225-PMU and Segment Arb commands.
- `current_range.py`: fixed-range assessment and automatic retry helpers.
- `data_processing.py`: PMU buffer parsing and analysis helpers.
- `plotting_utils.py`: common PMU plotting utilities.
- `fet_three_terminal_common.py`: shared execution helpers for three-terminal
  FET pulse experiments.

The PMU package imports the common VISA implementation from
`keithley4200.transport`. Experiment-specific pulse sequences stay in
runnable entries under `measurements/pmu`.

## 使用与参数

推荐调用顺序是建立 PMUSession → 构造序列 → execute_segARB_test → 等待结束/读取 → 关输出 → 保存分析。current_range 的自动重试会重新输出整个波形。电流测量档不是限流；段数校验也不代表全部电压/时间边界已经校验。模式码、硬件条件及代码限制见下方速查。

参数依据：[手册限制与模式速查](../../../reference/manuals/PARAMETER_LIMITS.md)。
