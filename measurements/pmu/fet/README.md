# PMU FET 测试入口

本目录按测试用途命名。每个脚本可以直接运行；导入模块不会启动测量。

| 入口 | 用途 | 主要参数 |
|---|---|---|
| [pulse_train.py](pulse_train.py) | 两路固定幅值脉冲串，可通过两路延迟设置错开的脉冲 | CH1/CH2 幅值、脉宽、延迟、PULSE_COUNT |
| [pulse_sweep.py](pulse_sweep.py) | CH1 扫描脉冲幅值，CH2 输出固定幅值脉冲串 | CH1_START/STOP/STEP、CH2_AMPLITUDE |
| [program_read.py](program_read.py) | Gate 写入一串脉冲，等待后设置读取 Vg/Vd 并测量，循环执行 | WRITE_VOLTAGE、PULSES_PER_TRAIN、READ_DELAY、SEQ_CYCLE_COUNT |
| [bipolar_program_read.py](bipolar_program_read.py) | 分别执行正、负两组写入条件下的写入—等待—读取 | POS/NEG 写入参数及重复次数、读取 Vg/Vd、READ_DELAY |
| [sweeps/delay_sweep.py](sweeps/delay_sweep.py) | 调用 bipolar_program_read，扫描写入至读取的等待时间并汇总 | DELAY_TIMES |
| [sweeps/read_bias_delay_sweep.py](sweeps/read_bias_delay_sweep.py) | 扫描读取 Vg、Vd、delay，比较状态电流并排序；每对偏压另做重复写读测试 | VG_READ_VALUES、VD_READ_VALUES、DELAY_TIMES、FEFET_REPEAT_N/T |

## 写读顺序与配置

前两个入口使用通用 CH1/CH2 配置，没有固定 Gate/Drain 的角色。
后两个基础写读入口使用 Segment Arb，Gate/Drain 由 GATE_CH 和 DRAIN_CH 指定。

program_read 的默认参数结构提供一组写入电压；bipolar_program_read 提供独立的
正、负写入配置。这一区别与使用几个 PMU 通道无关，两者都使用两路 PMU。

bipolar_program_read 在每轮中先完成所有正向写读重复，再完成所有负向写读重复。
每次写读包含一串写入脉冲和一次读取；脉冲串长度、写读重复次数、整个过程的循环
次数是三个不同参数。

两个扫描入口从 bipolar_program_read 导入配置，并调用其 main()。它们继承基础
脚本中的写入条件、正负写读重复次数和循环次数；每个 delay 不一定只得到一对读数。
每个 delay 都重新执行写入，不是对同一次写入连续跟踪不同等待时间。

修改基础写读参数应在对应入口顶部进行；扫描取值在 sweeps 中配置。
PARAMS 是导入时生成的参数快照，扫描脚本会临时更新它，并在结束时恢复。

## 输出

基础写读测试保存原始通道数据、FET_Data、参数及相关图形；波形预览由入口中的
选项控制。扫描另外生成汇总表和图。输出位置沿用各脚本的 SAVE_DIR/SWEEP_NAME
设置，文件继续使用公共命名和编号规则。

重命名后的脚本路径如上；IDE 运行配置或库外脚本中的旧路径需要相应更新。

## 参数与输出开关

CURRENT_RANGES 是 Gate/Drain 电流测量档，不是限流；SOURCE_COMPLIANCE 仅在使用 Source SMU 时限制该 SMU 电流。器件允许的 Gate/Drain 电压需按器件确定。

pulse_train/pulse_sweep 中 TEST_MODE=1 表示定点，2 表示波形；0–4 是仪器采集模式，不是保存模式。SAVE_WAVEFORM_PREVIEW 是额外波形图开关，PREVIEW_ONLY 是预览/实测选择。

参数依据：[手册限制与模式速查](../../../reference/manuals/PARAMETER_LIMITS.md)。
