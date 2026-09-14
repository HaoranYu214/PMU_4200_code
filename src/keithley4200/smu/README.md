# SMU reusable layer

This folder separates the two KXCI SMU command families:

- `system_mode.py`: Clarius-style sweeps using `DE`, `CH`, `SS`, `VR/IR`,
  `HT`, `DT`, `IT`, `ME`, `SP`, and `DO`.
- `user_mode.py`: direct spot control using `US`, `DV/DI`, and `TI/TV`.

SMU and PMU sessions share the maintained VISA implementation in
`../transport.py`; the vendor `instrcomms.py` remains under `reference` only.

Do not mix CVU `:CVU:SPEED` commands with SMU `IT` commands. Ethernet tests
poll `SP`; `DR` is for GPIB Data Ready service requests.

The initialization helpers send `EM 1,0`, clear the KXCI error queue, reset the
instrument, and disable every explicitly configured available channel before
defining active channels. Active System Mode channels use automatic standby;
exceptions and interrupts trigger best-effort `ME4` abort and channel disable.

`routing.py` validates explicit per-channel probe connections. RPM-backed
channels are selected with `RP PMUN-C, 2` only after `*RST`; direct channels
are untouched. After output standby/shutdown, switched RPMs are returned with
`RP PMUN-C, 0`. Runnable experiments expose this map as `SMU_CONNECTIONS` so
the physical topology is not hidden in the reusable layer.

Do not restore an RPM after an uncertain shutdown. The run helpers restore
Pulse mode only after successful completion with automatic SMU standby. Error,
timeout, and interrupt paths abort/disable best-effort but intentionally leave
the RPM routed to the SMU for diagnosis.

The API remains close to the official examples while adding the validation and
cleanup behavior needed for real experiments. Runnable entries remain under
`measurements/smu/iv`.

Completed System Mode data is downloaded with `DO`. `RD` is reserved for
real-time retrieval when the expected point count is already known; a returned
`0` means that point is not ready, not that the data buffer has ended.

Numeric `RG` values specify the lowest autoranged measurement range, not a
fixed range. Use `"auto"` or `None` when preamplifier availability is unknown.

`data_processing.py` saves untouched measurements/statuses to `Raw` and a
fixed numeric extraction table to `PlotData`. `plotting.py` plots absolute
current on a true logarithmic axis so tick labels remain physical amperes.

## 常改参数

System Mode 的 DT 是每点等待，HT 是整次扫描前等待，IT 是积分时间。RG 是自动量程下限，电流 compliance 是另一个参数。user_mode 的 DV range 参数是代码编号，不是直接填伏特值。详见下方手册速查。

参数依据：[手册限制与模式速查](../../../reference/manuals/PARAMETER_LIMITS.md)。
