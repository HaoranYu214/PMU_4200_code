# FTJ 写读实验

共同思路是施加写入脉冲，再用 read_v 读取状态；平台段测量用于计算电阻/电导。params 是配置源，运行时通过 configure_measurement / _rebuild_runtime_config 更新波形，避免只改派生变量。

| 文件 | 扫描对象或版本区别 |
|---|---|
| [ftj_RV1.py](ftj_RV1.py) | 从 0 开始，先正后负的写入电压路径 |
| [ftj_RV2.py](ftj_RV2.py) | 从 +Vp 到 -Vp 再回 +Vp，支持 scan_cycles |
| [ftj_ISPP_V1.py](ftj_ISPP_V1.py) | 写电压阶梯；每个写电压独立序列，复用读序列 |
| [ftj_ISPP_V2.py](ftj_ISPP_V2.py) | 把每个极性的完整阶梯打包进一个序列 |
| [ftj_Identical_V1.py](ftj_Identical_V1.py) | 固定写幅值，累积脉冲写读；计划在 SARB 内执行 |
| [ftj_Identical_V2.py](ftj_Identical_V2.py) | 拆成多次执行，wait_after_* 通过 Python 等待 |
| [ftj_PWM.py](ftj_PWM.py) | 固定写电压，扫描 write_base_dwell × width_multipliers |
| [ftj_MRD.py](ftj_MRD.py) | 每个写电压重复 reset/read/write/read，观察状态分布 |
| [ftj_endurance.py](ftj_endurance.py) | 外层重复指定 FTJ 测试，逐次更新汇总表 |

base_v 是脉冲间保持电平；RV 的 offset_v 会移动写入电压窗口，不能只检查 vp 的绝对值。read_v 也是真实施加到器件上的测试条件。

*_dwell、*_rise、*_fall、*_idle 是 SARB 段时间；wait_after_* 是软件等待，没有单段 1 s 上限，但包含软件调度/通信开销。重复次数还会影响测量点数及段预算。SAVE_EVERY_RUN=False 仍保留 endurance 汇总，不保存每次目标测试的完整输出。

参数依据：[手册限制与模式速查](../../../reference/manuals/PARAMETER_LIMITS.md)。
