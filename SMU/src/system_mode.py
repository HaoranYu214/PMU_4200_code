"""4200A SMU System Mode configuration and execution helpers.

System Mode is used for Clarius-style sweeps. It is separate from User Mode
DV/DI/TI/TV commands.
"""

from __future__ import annotations

import time


SOURCE_VOLTAGE = 1
SOURCE_CURRENT = 2

FUNCTION_VAR1 = 1
FUNCTION_VAR2 = 2
FUNCTION_CONSTANT = 3
FUNCTION_VAR1_RATIO = 4


def initialize_system_mode(query):
    """Clear buffered readings, reset instruments, and enter channel definition."""
    query("BC")
    query("*RST")
    query("DE")


def define_channel(
    query,
    channel,
    voltage_name,
    current_name,
    *,
    source_mode=SOURCE_VOLTAGE,
    source_function=FUNCTION_CONSTANT,
):
    query(
        f"CH{channel}, '{voltage_name}', '{current_name}', "
        f"{source_mode}, {source_function}"
    )


def configure_constant_voltage(query, channel, voltage, current_compliance):
    query(f"VC{channel}, {voltage}, {current_compliance}")


def configure_constant_current(query, channel, current, voltage_compliance):
    query(f"IC{channel}, {current}, {voltage_compliance}")


def configure_linear_voltage_sweep(
    query,
    *,
    start,
    stop,
    step,
    current_compliance,
    sweep_variable=1,
):
    query(f"VR{sweep_variable}, {start}, {stop}, {step}, {current_compliance}")


def configure_linear_current_sweep(
    query,
    *,
    start,
    stop,
    step,
    voltage_compliance,
    sweep_variable=1,
):
    query(f"IR{sweep_variable}, {start}, {stop}, {step}, {voltage_compliance}")


def configure_list_voltage_sweep(
    query,
    values,
    *,
    current_compliance,
    sweep_variable=1,
    mode=1,
):
    value_text = ",".join(f"{float(value):.12g}" for value in values)
    query(f"VL{sweep_variable},{mode},{current_compliance},{value_text}")


def configure_timing(
    query,
    *,
    hold_time=0.0,
    sweep_delay=0.0,
    integration="IT1",
):
    """Configure hold, per-point delay, and SMU integration settings.

    integration accepts IT1/IT2/IT3, or a complete custom command such as
    ``IT4, delay_factor, filter_factor, aperture_plc``.
    """
    query(f"HT {hold_time}")
    query(f"DT {sweep_delay}")
    integration = str(integration).strip()
    if not integration.upper().startswith("IT"):
        raise ValueError("integration must be an IT command.")
    query(integration)


def configure_measurement_list(query, variables, *, display_mode=2):
    variable_text = ", ".join(f"'{variable}'" for variable in variables)
    query(f"SM DM{display_mode}")
    query(f"LI {variable_text}")


def execute_and_wait(
    query,
    *,
    execution_mode=1,
    timeout_s=300.0,
    poll_interval_s=0.2,
):
    """Start a System Mode test and wait for SP to report completion."""
    query("MD")
    query(f"ME{execution_mode}")
    deadline = time.monotonic() + timeout_s
    while True:
        response = query("SP").strip()
        try:
            status = int(response)
        except ValueError:
            status = None
        if status in (0, 1):
            return status
        if time.monotonic() >= deadline:
            raise TimeoutError(
                f"SMU System Mode test did not finish within {timeout_s:g} s; "
                f"last SP response={response!r}."
            )
        time.sleep(poll_interval_s)


def run_linear_voltage_sweep(
    query,
    *,
    sweep_channel,
    bias_channel,
    sweep_voltage_name="VSWEEP",
    sweep_current_name="ISWEEP",
    bias_voltage_name="VBIAS",
    bias_current_name="IBIAS",
    start=0.0,
    stop=1.0,
    step=0.1,
    sweep_current_compliance=1e-3,
    bias_voltage=0.0,
    bias_current_compliance=1e-3,
    hold_time=0.0,
    sweep_delay=0.0,
    integration="IT1",
    return_variables=None,
    timeout_s=300.0,
):
    """Configure and execute a two-channel linear voltage sweep."""
    initialize_system_mode(query)
    define_channel(
        query,
        bias_channel,
        bias_voltage_name,
        bias_current_name,
        source_mode=SOURCE_VOLTAGE,
        source_function=FUNCTION_CONSTANT,
    )
    define_channel(
        query,
        sweep_channel,
        sweep_voltage_name,
        sweep_current_name,
        source_mode=SOURCE_VOLTAGE,
        source_function=FUNCTION_VAR1,
    )
    query("SS")
    configure_linear_voltage_sweep(
        query,
        start=start,
        stop=stop,
        step=step,
        current_compliance=sweep_current_compliance,
    )
    configure_constant_voltage(
        query,
        bias_channel,
        bias_voltage,
        bias_current_compliance,
    )
    configure_timing(
        query,
        hold_time=hold_time,
        sweep_delay=sweep_delay,
        integration=integration,
    )
    if return_variables is None:
        return_variables = [
            sweep_current_name,
            sweep_voltage_name,
            bias_current_name,
            bias_voltage_name,
        ]
    configure_measurement_list(query, return_variables)
    execute_and_wait(query, execution_mode=1, timeout_s=timeout_s)
    return list(return_variables)
