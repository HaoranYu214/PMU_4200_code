# Keithley 4200A SMU tests

The runnable experiments are in `iv/`. Their reusable KXCI implementation
lives in `src/keithley4200/smu/`.

Layout:

- `iv/`: editable experiment entries with parameters, safe session handling,
  validated data retrieval, and result saving.
- `../../reference/manuals/smu/`: the local Keithley KXCI programming manual.
- `../../reference/official_examples/smu/`: vendor examples retained for command-format
  comparison only. They are not production-safe experiment entries.

Runtime PMU and SMU connections use `src/keithley4200/transport.py`. The
vendor `instrcomms.py` in `reference/official_examples/smu/System Mode/` is kept
as an unmodified reference and is not imported by runnable experiments.

Before running an experiment, set `AVAILABLE_CHANNELS` to every SMU installed
and mapped in KCon. The library disables all of those channels first, then
defines only the requested sweep and bias channels.

Set `SMU_CONNECTIONS` to the physical probe wiring. On the current system,
SMU1 and SMU2 pass through the RPMs on PMU1 channels 1 and 2, while SMU3 and
SMU4 go directly to probes:

```python
SMU_CONNECTIONS = {
    1: "rpm:PMU1-1",
    2: "rpm:PMU1-2",
    3: "direct",
    4: "direct",
}
```

For active `rpm:` entries, the reusable layer selects RPM mode 2 after
`*RST` and before source setup (blue LED). It returns those RPMs to pulse mode
0 only after the SMUs enter standby or are explicitly powered off. Direct
entries do not send any RPM commands.

RPM safety rule: never switch the relay while an attached source may be
energized. A normally completed System Mode run uses `ST channel,1`, waits for
completion, then restores the RPM to pulse mode. User Mode explicitly powers
off with `DVchannel` before restoration. On timeout, command error, interrupt,
or failed shutdown, the library deliberately leaves the RPM in SMU mode rather
than risk switching an active source; inspect the instrument before recovery.

Manual locations in `reference/manuals/smu/4200A-KXCI-907-01D_May_2024.pdf`:

- Printed pages 7-67 to 7-68: `RP HRID, mode`, blue-SMU example, and the
  instruction to set outputs to 0 V before returning the RPM to pulse.
- Printed pages 7-22 to 7-23: equivalent `:PMU:RPM:CONFIGURE HRID, mode`
  command and all accepted modes (`0` PMU, `1` CV 2-wire, `2` SMU,
  `3` CV 4-wire).
- Printed page 5-14: `ST channel,1` automatic standby behavior.

Changing physical RPM/PMU cables is different from electronic `RP` switching:
power down the 4200A and disconnect mains before plugging or unplugging RPM
hardware, as required by the 4225-RPM hardware instructions.

For Ethernet KXCI, configure the command/string terminator to `None` in KCon;
the Python session sends and reads the required null (`\0`) terminator. Keep
the KXCI reading delimiter consistent with the comma-separated buffer format
used by `DO`.

## 参数先看什么

AVAILABLE_CHANNELS 和 SMU_CONNECTIONS 必须匹配实际设备/接线。电压、CURRENT_COMPLIANCE 和积分参数按具体 SMU 型号选择。不要把硬件最大范围直接当作器件的安全测试范围。

参数依据：[手册限制与模式速查](../../reference/manuals/PARAMETER_LIMITS.md)。
