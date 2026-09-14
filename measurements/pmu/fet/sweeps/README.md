# FET 参数扫描

- [delay_sweep.py](delay_sweep.py)：逐个 DELAY_TIMES 调用基础正负写读测试，每次重新写入，汇总 Id 随等待时间的变化。
- [read_bias_delay_sweep.py](read_bias_delay_sweep.py)：组合扫描 Vg、Vd、delay，并对每组 Vg/Vd 追加重复写读，再生成状态区分指标及排序。

两者调用 [bipolar_program_read.py](../bipolar_program_read.py)，写入电压、正负重复次数、循环次数继承自该入口。扫描不是对同一次写入状态持续读回；不应把每个 delay 自动理解为只有一对数据。

DELAY_TIMES > 1 s 会走基础入口的软件等待分支；精确间隔还含重配置开销。VG_READ_VALUES/VD_READ_VALUES 同样受 PMU 电压档和器件允许读取条件限制。

参数依据：[手册限制与模式速查](../../../../reference/manuals/PARAMETER_LIMITS.md)。
