# SMU reusable layer

This folder separates the two KXCI SMU command families:

- `system_mode.py`: Clarius-style sweeps using `DE`, `CH`, `SS`, `VR/IR`,
  `HT`, `DT`, `IT`, `ME`, `SP`, and `RD`.
- `user_mode.py`: direct spot control using `US`, `DV/DI`, and `TI/TV`.

Do not mix CVU `:CVU:SPEED` commands with SMU `IT` commands. Ethernet tests
poll `SP`; `DR` is for GPIB Data Ready service requests.

The first version intentionally keeps the API close to the official examples.
More specialized IV/family-of-curves reshaping can be added in `SMU/IV`.
