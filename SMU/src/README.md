# SMU reusable layer

This folder separates the two KXCI SMU command families:

- `system_mode.py`: Clarius-style sweeps using `DE`, `CH`, `SS`, `VR/IR`,
  `HT`, `DT`, `IT`, `ME`, `SP`, and `RD`.
- `user_mode.py`: direct spot control using `US`, `DV/DI`, and `TI/TV`.

Do not mix CVU `:CVU:SPEED` commands with SMU `IT` commands. Ethernet tests
poll `SP`; `DR` is for GPIB Data Ready service requests.

The initialization helpers send `EM 1,0`, selecting the 4200A command set for
the current session only. This makes custom `IT4,X,Y,Z` behavior deterministic
without changing the persistent KCon setting.

The first version intentionally keeps the API close to the official examples.
More specialized IV/family-of-curves reshaping can be added in `SMU/IV`.

Completed System Mode data is downloaded with `DO`. `RD` is reserved for
real-time retrieval when the expected point count is already known; a returned
`0` means that point is not ready, not that the data buffer has ended.
