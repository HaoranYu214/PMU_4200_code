# Ferroelectric capacitor tests

[English](#english) | [中文](#中文)

## English

| File | Purpose and parameters to adjust |
|---|---|
| [PV2.py](PV2.py) | Two triangular P-V loops after conditioning; Vp, rise_time, delay_time, offset, area_cm2 |
| [PUND_tri.py](PUND_tri.py) | Triangular P/U/N/D subtraction; rise_time is one edge's duration and delay_time is the interval between pulses |
| [PUND_Squr.py](PUND_Squr.py) | Square/trapezoidal PUND; dwell_time is the plateau segment duration |
| [NLS_manually.py](NLS_manually.py) | Square write pulse with triangular readout; Vsquare, Dwell, Rt_s, MeasureSquare |
| [endurance.py](endurance.py) | Unmeasured endurance cycles, with PV2/PUND readback at cumulative cycle_counts milestones |

PV2, PUND_tri, and PUND_Squr use automatic range retries: each acquisition uses a fixed current range; an unsuitable range triggers a new acquisition at another range. Irange1/2 are the initial ranges, and the accepted ranges are saved with the measurement. The default candidate ranges target a 10 V RPM.

Some PV2 segments last 2 × rise_time. Since an individual segment is limited to 1 s, rise_time is also constrained to at most 0.5 s. PUND dwell is a plateau duration, not the width parameter of standard PULSE:TIMES.

area_cm2 is the effective device area in cm². It controls charge/polarization conversion, not instrument output. MeasureSquare=False disables acquisition of the square segment without removing the square pulse.

Parameter source: [manual limits and mode reference (Chinese)](../../../reference/manuals/PARAMETER_LIMITS.md).

### Fixed NLS readout interval

`TotalDelay` (seconds, default 0.11) runs from the end of the preset/prepost falling edge to the start of the first half-PUND read rising edge. `Dwell` is the disturb flat-top duration. An unmeasured baseline hold is inserted after the disturb pulse: `PaddingDelay = TotalDelay - 2*Delaytime - 2*Rt_s - Dwell`. The baseline is `offset` (currently 0 V). Both existing waits and disturb edges count toward the total; `Delaytime` between the two half-PUND read pulses is unchanged.

Insufficient total time is rejected before measurement. Zero padding is omitted; nonzero padding must meet the default 10 V range segment duration limit of 20 ns–1 s. Hardware timing is limited by its 10 ns resolution. Parameter sheets and the batch summary save `TotalDelay` and `PaddingDelay`. Fixing this interval controls elapsed time since preset, while the time from the end of disturbance to readout still varies with pulse width.

### Retention: split executions
- [retentionPV.py](retentionPV.py): preset, output-off wait, then two continuous PV loops; two executions.
- [retentionPUND.py](retentionPUND.py): separate Preset/P/U/N/D executions with `delay_time` between successive pulses; five executions.
- Waveforms remain in each entry. [_retention.py](_retention.py) supplies local execution, timing, checkpoint and preview support without changes to src.

These dedicated entries always split; they do not switch modes at 0.1 s. Defaults are a 1 s delay and preview-only mode. Review voltage, fixed current ranges, area and save directory before setting `PREVIEW_ONLY=False`. There is no automatic range retry. This version requires zero offset; waiting uses output-off, not guaranteed 0 V drive or verified SSR isolation.

Deadlines use the previous EXECUTE send time plus programmed duration. Transfer and configuration consume the wait budget. Unknown communication/startup latency remains: `estimated_delay_s` and `EstimatedGlobalTime_s` are host estimates, not measured hardware timing. `overrun_s` records known lateness. Raw local timestamps and Stage/Execution labels remain intact. Raw data and execution logs are checkpointed before analysis, including data already retrieved before failure/interruption. Preview shows separate executions and software gaps. PV retains PV2 integration; PUND pairs labeled branches and interpolates unequal sample counts in normalized time.


Parameter API: [standalone and workflow configuration](../README.md#standalone-and-workflow-configuration).

## 中文

### 铁电电容测试

| 文件 | 用途及应修改的参数 |
|---|---|
| [PV2.py](PV2.py) | 预处理后测两圈三角 P-V；Vp、rise_time、delay_time、offset、area_cm2 |
| [PUND_tri.py](PUND_tri.py) | 三角脉冲 P/U/N/D 差分；rise_time 为单边时间，delay_time 为脉冲间等待 |
| [PUND_Squr.py](PUND_Squr.py) | 方波/梯形脉冲 PUND；额外的 dwell_time 是平台段时间 |
| [NLS_manually.py](NLS_manually.py) | 方形写入脉冲配合三角读出；Vsquare、Dwell、Rt_s、MeasureSquare |
| [endurance.py](endurance.py) | 无测量的疲劳循环，到累计 cycle_counts 节点执行 PV2/PUND 读回 |

PV2、PUND_tri、PUND_Squr 调用自动改档：每次用固定电流档采集，不合适时换档重测。Irange1/2 是初始档，实际接受的档位随测量记录保存。其候选范围默认面向 10 V RPM。

PV2 的部分实际段长为 2×rise_time，单段最长 1 s，因此 rise_time 上限还受 0.5 s 约束。PUND 的 dwell 是平台段，不是普通 PULSE:TIMES 的 width。

area_cm2 是器件有效面积，单位 cm²；它决定电荷/极化换算，不改变仪器输出。MeasureSquare=False 仅关闭方形段的采集，不取消方形脉冲。

参数依据：[手册限制与模式速查](../../../reference/manuals/PARAMETER_LIMITS.md)。

### NLS 固定读出间隔

`TotalDelay`（秒，默认 0.11）从 preset/prepost 下降沿结束计时，到 half-PUND 第一个读取上升沿开始。方波平台为 `Dwell`；在方波后增加不采集的基线等待段：`PaddingDelay = TotalDelay - 2*Delaytime - 2*Rt_s - Dwell`。基线为 `offset`，当前为 0 V；原有两段等待和扰动边沿均计入总间隔，half-PUND 两次读取之间的 `Delaytime` 不变。

总间隔不足会在测试前报错；补偿为零时省略新增段，非零补偿须满足默认 10 V 档 20 ns–1 s 段长限制。实际仪器时间精度受 10 ns 分辨率限制。参数表保存 `TotalDelay` 和 `PaddingDelay`；批量扫描汇总表也保存这两项。固定此间隔可控制 preset 后的总等待时间，但不同脉宽仍会对应不同的扰动结束后等待时间。

参数接口：[独立运行与工作流配置](../README.md#独立运行与工作流配置)。

## Retention：分次执行
- [retentionPV.py](retentionPV.py)：Preset → 关闭输出等待 → 连续双 PV 回线，共两次执行。
- [retentionPUND.py](retentionPUND.py)：Preset、P、U、N、D 分别执行，共五次；每对相邻脉冲之间按 `delay_time` 等待。
- 波形仍在各入口中定义；[_retention.py](_retention.py) 仅为这两个入口提供配置、计时、原始数据保存和预览支持，不改变 src。

独立入口固定采用软件拆分，不按 0.1 s 阈值逐点切换。默认 `delay_time=1.0` 秒、`PREVIEW_ONLY=True`；确认电压、量程、面积及输出目录后，将 `PREVIEW_ONLY=False` 运行。量程固定，不自动重测。初版要求 `offset=0`，两次执行之间为 `output_off`，不承诺驱动 0 V 或已验证的 SSR 浮空。

等待基准是上一次 EXECUTE 发送时刻加该次波形时长，至下一次 EXECUTE 发送时刻；数据读取、关闭输出和配置耗时计入等待预算。通信与硬件启动延迟无法消除，因此 `ExecutionTiming` 中的 `estimated_delay_s` 是主机估计，`overrun_s` 记录已知超时量，不能作为精确硬件时间。每批原始 `Timestamp` 保持局部值，`Stage`/`Execution` 标记来源，`EstimatedGlobalTime_s` 单独标注估计全局时间。

原始通道、参数及执行记录先保存，再进行分析；中断或失败时保留已经读回的数据。预览将各次执行分开显示，并标注软件等待。PV 使用原 PV2 的回线分析；PUND 按脉冲标签和分支配对，对点数不同的分支使用归一化时间插值。
