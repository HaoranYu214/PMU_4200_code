# FORC 与 NLS 扫描

| 文件 | 用途 |
|---|---|
| [FORC.py](FORC.py) | 每个反转电压单独执行一条一阶反转曲线 |
| [FORC_1excute.py](FORC_1excute.py) | 将多条反转曲线合并为一次 SARB 执行，再分离分析 |
| [NLS_1C_switch.py](NLS_1C_switch.py) | 提供可配置的单次 NLS 写入/读回及预览 |
| [NLS_1C_switch_list.py](NLS_1C_switch_list.py) | 调用前者扫描 Dwell_list × Vsquare_list |

FORC 的 Vmax 是相对 offset 的幅值；反转电压满足 -Vmax ≤ Vr < Vmax 是算法约束。full_sweep_time 按电压跨度换算分支时间，所以靠近 +Vmax 的曲线可能出现很短的实际段。合并版本还须符合单次段数与采样点预算。

NLS 的 Dwell 是方形写入平台时间，Vsquare 为写入电平；MeasureSquare 控制这部分是否采集，关闭采集仍输出脉冲。

参数依据：[手册限制与模式速查](../../../reference/manuals/PARAMETER_LIMITS.md)。
