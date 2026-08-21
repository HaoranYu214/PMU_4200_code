# Keithley 4200A SMU tests

The runnable experiments are in `IV/`. Their reusable KXCI implementation now
lives beside the PMU library in `../Pkg_PMU_list/src/smu/`.

Layout:

- `IV/`: editable experiment entries with parameters, safe session handling,
  validated data retrieval, and result saving.
- `reference/manual/`: the local Keithley KXCI programming manual.
- `reference/official_examples/`: vendor examples retained for command-format
  comparison only. They are not production-safe experiment entries.

Before running an experiment, set `AVAILABLE_CHANNELS` to every SMU installed
and mapped in KCon. The library disables all of those channels first, then
defines only the requested sweep and bias channels.

For Ethernet KXCI, configure the command/string terminator to `None` in KCon;
the Python session sends and reads the required null (`\0`) terminator. Keep
the KXCI reading delimiter consistent with the comma-separated buffer format
used by `DO`.
