# PMU 实验

| 目录 | 实验用途 |
|---|---|
| [fe_cap](fe_cap/README.md) | PV、三角/方波 PUND、NLS、疲劳循环 |
| [ftj](ftj/README.md) | 写入后读阻，改变写电压、脉宽或脉冲次数 |
| [fet](fet/README.md) | 两路脉冲，以及 Gate/Drain 配合的 FeFET 写读 |
| [programmed](programmed/README.md) | FORC、NLS 参数扫描 |

大部分入口使用 Segment Arb；fet/pulse_train.py 和 pulse_sweep.py 使用普通脉冲模式，两者的时间定义不同。Irange/CURRENT_RANGES 是电流测量档，不是 SMU 式电流限流。需要按 RPM 和 10/40 V 档选择；默认 INIT 为 10 V 档。

公共参数表还解释了模式 0–4、LOAD/LLEC、保存开关和软件等待。

参数依据：[手册限制与模式速查](../../reference/manuals/PARAMETER_LIMITS.md)。
