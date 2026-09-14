# 多步测量工作流

| 文件 | 执行内容 |
|---|---|
| [pv_and_pund.py](pv_and_pund.py) | 顺序执行 PV2 和三角 PUND，可重复调用 |
| [pv2_pund_map.py](pv2_pund_map.py) | 扫描幅值、三角波频率和 delay 的组合 |
| [package1.py](package1.py) | 按 RUN_ORDER 串联 PV/PUND、map 和 SMU 分段 I-V |
| [ftj_package1.py](ftj_package1.py) | 按 RUN_ORDER 配置并运行 RV2/PWM/Identical/MRD |

工作流自己的参数表会覆盖下层入口的相应设置；批量测量时优先改这里的配置。STOP_ON_ERROR=True 表示遇到错误停止后续阶段，False 表示允许按流程继续并记录失败。SETTLE_TIME_S 是 Python 阶段间等待，不是仪器脉冲延迟。

map 中 frequency 转换为 rise_time=1/(4f)，仅表示三角波的标称频率；delay 和其他预处理段会增加实际总时间。所有展开后的真实时间段仍受手册边界约束。

保存按配置目录执行，汇总表含 time；r001/r002 是同名输出组编号。RUN_*、PREVIEW_ONLY 等布尔值在参数表中表示开关，不能和仪器模式码混为一谈。

参数依据：[手册限制与模式速查](../reference/manuals/PARAMETER_LIMITS.md)。
