"""4200A SMU System Mode configuration and execution helpers.

System Mode is used for Clarius-style sweeps. It is separate from User Mode
DV/DI/TI/TV commands.
"""

from __future__ import annotations

import time
from collections.abc import Sequence


SOURCE_VOLTAGE = 1
SOURCE_CURRENT = 2

FUNCTION_VAR1 = 1
FUNCTION_VAR2 = 2
FUNCTION_CONSTANT = 3
FUNCTION_VAR1_RATIO = 4


def initialize_system_mode(query):
    """Select 4200A commands, clear/reset, and enter channel definition."""
    query("EM 1,0")
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
    channel,
    current_compliance,
    mode=1,
):
    """Configure a voltage list on the actual SMU/VS channel number."""
    values = [float(value) for value in values]
    if not 1 <= int(channel) <= 9:
        raise ValueError("channel must be between 1 and 9.")
    if not values:
        raise ValueError("values must contain at least one voltage.")
    if len(values) > 4096:
        raise ValueError("KXCI list sweeps are limited to 4096 points.")
    value_text = ",".join(f"{float(value):.12g}" for value in values)
    query(f"VL{int(channel)},{mode},{current_compliance},{value_text}")


def build_segmented_voltage_path(turning_points, step):
    """Build a list sweep through several turning points.

    Example:
        ``turning_points=[0, 2, 0, -3, 0]`` produces
        ``0 -> 2 -> 0 -> -3 -> 0``.

    ``step`` may be one positive number shared by all segments, or one
    positive number per segment. Segment endpoints are always included and
    shared endpoints are not duplicated.
    """
    def clean(value):
        return float(f"{float(value):.12g}")

    turning_points = [clean(value) for value in turning_points]
    if len(turning_points) < 2:
        raise ValueError("turning_points must contain at least two voltages.")

    segment_count = len(turning_points) - 1
    if isinstance(step, Sequence) and not isinstance(step, (str, bytes)):
        steps = [float(value) for value in step]
        if len(steps) != segment_count:
            raise ValueError(
                f"Expected {segment_count} segment steps, received {len(steps)}."
            )
    else:
        steps = [float(step)] * segment_count

    values = [turning_points[0]]
    for start, stop, segment_step in zip(
        turning_points[:-1],
        turning_points[1:],
        steps,
    ):
        if segment_step <= 0:
            raise ValueError("All segment steps must be positive.")
        if start == stop:
            continue

        direction = 1.0 if stop > start else -1.0
        next_value = start + direction * segment_step
        tolerance = max(abs(start), abs(stop), segment_step, 1.0) * 1e-12
        while (
            next_value < stop - tolerance
            if direction > 0
            else next_value > stop + tolerance
        ):
            values.append(clean(next_value))
            next_value += direction * segment_step
        values.append(clean(stop))

    return values


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
        data_ready = status is not None and bool(status & 0b00000001)
        busy = status is not None and bool(status & 0b00010000)
        if data_ready and not busy:
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


def run_list_voltage_sweep(
    query,
    *,
    values,
    sweep_channel,
    bias_channel,
    sweep_voltage_name="VSWEEP",
    sweep_current_name="ISWEEP",
    bias_voltage_name="VBIAS",
    bias_current_name="IBIAS",
    sweep_current_compliance=1e-3,
    bias_voltage=0.0,
    bias_current_compliance=1e-3,
    hold_time=0.0,
    sweep_delay=0.0,
    integration="IT1",
    return_variables=None,
    timeout_s=300.0,
):
    """Configure and execute an arbitrary two-channel voltage list sweep."""
    values = [float(value) for value in values]
    if not values:
        raise ValueError("values must contain at least one voltage.")

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
    configure_list_voltage_sweep(
        query,
        values,
        channel=sweep_channel,
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
