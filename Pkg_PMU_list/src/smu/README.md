# SMU reusable layer in the PMU package

This folder separates the two KXCI SMU command families:

- `system_mode.py`: Clarius-style sweeps using `DE`, `CH`, `SS`, `VR/IR`,
  `HT`, `DT`, `IT`, `ME`, `SP`, and `DO`.
- `user_mode.py`: direct spot control using `US`, `DV/DI`, and `TI/TV`.

Do not mix CVU `:CVU:SPEED` commands with SMU `IT` commands. Ethernet tests
poll `SP`; `DR` is for GPIB Data Ready service requests.

The initialization helpers send `EM 1,0`, clear the KXCI error queue, reset the
instrument, and disable every explicitly configured available channel before
defining active channels. Active System Mode channels use automatic standby;
exceptions and interrupts trigger best-effort `ME4` abort and channel disable.

The API remains close to the official examples while adding the validation and
cleanup behavior needed for real experiments. Runnable entries remain under
the repository-level `SMU/IV` directory.

Completed System Mode data is downloaded with `DO`. `RD` is reserved for
real-time retrieval when the expected point count is already known; a returned
`0` means that point is not ready, not that the data buffer has ended.

Numeric `RG` values specify the lowest autoranged measurement range, not a
fixed range. Use `"auto"` or `None` when preamplifier availability is unknown.
