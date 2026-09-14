# Shared Keithley source library

The reusable instrument code is divided by hardware command family:

- `pmu/`: 4225-PMU sessions, Segment Arb execution, data processing,
  current-range helpers, plotting, and shared FET pulse helpers.
- `smu/`: SMU System Mode, User Mode, RPM routing, data processing, plotting,
  and sessions.
- `transport.py`: the single shared PyVISA communication implementation used
  by both packages.

Runnable experiment parameters and waveforms remain outside `src`; this
directory contains reusable mechanisms rather than experiment protocols.

## 其他公共模块

- [output.py](output.py)：统一文件标签、原子占号和汇总时间。
- [tools](tools/README.md)：波形预览与 dry-run，已包含在安装包内。

__init__.py 标记可导入的包；一般不需要直接运行。实验入口只组合这些公共能力，具体波形仍由 measurements 定义。

参数依据：[手册限制与模式速查](../../reference/manuals/PARAMETER_LIMITS.md)。
