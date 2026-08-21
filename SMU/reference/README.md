# SMU reference material

This directory is reference-only.

- `manual/4200A-KXCI-907-01D_May_2024.pdf` is the Keithley KXCI programming
  manual used to validate command syntax and safety behavior.
- `official_examples/System Mode/` and `official_examples/User Mode/` contain
  vendor examples. They intentionally stay close to the original samples and
  must not be treated as safe runnable library entries.

The official examples generally execute at module top level, omit robust
timeout/abort handling, and may rely on KCon defaults. Their Ethernet addresses
remain documentation-only `192.0.2.0` placeholders. Use `SMU/IV/` instead.

