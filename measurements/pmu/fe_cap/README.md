# 铁电电容测试

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
