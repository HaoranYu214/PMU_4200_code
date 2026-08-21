"""4200A SMU System Mode configuration and execution helpers.

System Mode is used for Clarius-style sweeps. It is separate from User Mode
DV/DI/TI/TV commands.
"""

from __future__ import annotations

import time
from collections.abc import Sequence

from .common import (
    clear_kxci_error,
    normalize_channels,
    raise_for_kxci_error,
    validate_channel,
)


SOURCE_VOLTAGE = 1
SOURCE_CURRENT = 2

FUNCTION_VAR1 = 1
FUNCTION_VAR2 = 2
FUNCTION_CONSTANT = 3
FUNCTION_VAR1_RATIO = 4
DEFAULT_AVAILABLE_CHANNELS = (1, 2, 3, 4)


def disable_system_channels(query, channels):
    """Disable explicitly listed System Mode channels on the DE page."""
    channels = normalize_channels(channels, name="available_channels")
    query("DE")
    for channel in channels:
        query(f"CH{channel}")
    return channels


def initialize_system_mode(query, *, available_channels=DEFAULT_AVAILABLE_CHANNELS):
    """Reset System Mode and disable every known SMU before redefining use."""
    query("EM 1,0")
    clear_kxci_error(query)
    query("BC")
    query("*RST")
    return disable_system_channels(query, available_channels)


def configure_auto_standby(query, channels, *, enabled=True):
    """Put active SMUs in standby automatically when a test completes."""
    for channel in normalize_channels(channels, name="active_channels"):
        query(f"ST {channel}, {1 if enabled else 0}")


def shutdown_system_mode(query, channels):
    """Best-effort abort plus channel disable for error and interrupt paths."""
    commands = ["MD", "ME4", "DE"]
    try:
        channels = normalize_channels(channels, name="available_channels")
    except (TypeError, ValueError):
        channels = ()
    commands.extend(f"CH{channel}" for channel in channels)
    for command in commands:
        try:
            query(command)
        except BaseException:
            pass


def define_channel(
    query,
    channel,
    voltage_name,
    current_name,
    *,
    source_mode=SOURCE_VOLTAGE,
    source_function=FUNCTION_CONSTANT,
):
    channel = validate_channel(channel)
    for name, value in (("voltage_name", voltage_name), ("current_name", current_name)):
        if not value or len(str(value)) > 6:
            raise ValueError(f"{name} must contain 1 to 6 characters.")
    if source_mode not in (SOURCE_VOLTAGE, SOURCE_CURRENT, 3):
        raise ValueError("source_mode must be voltage, current, or common.")
    if source_function not in (
        FUNCTION_VAR1,
        FUNCTION_VAR2,
        FUNCTION_CONSTANT,
        FUNCTION_VAR1_RATIO,
    ):
        raise ValueError("source_function must be a valid KXCI CH function.")
    query(
        f"CH{channel}, '{voltage_name}', '{current_name}', "
        f"{source_mode}, {source_function}"
    )


def configure_constant_voltage(query, channel, voltage, current_compliance):
    channel = validate_channel(channel)
    query(f"VC{channel}, {voltage}, {current_compliance}")


def configure_constant_current(query, channel, current, voltage_compliance):
    channel = validate_channel(channel)
    query(f"IC{channel}, {current}, {voltage_compliance}")


def configure_current_range(query, channel, current_range):
    """Set the lowest current measurement range for one SMU channel.

    Use ``None`` or ``"auto"`` to leave the instrument in its default autorange
    behavior. Numeric values send the System Mode ``RG`` command.
    """
    if current_range is None:
        return
    channel = validate_channel(channel)
    if isinstance(current_range, str):
        if current_range.strip().lower() in ("auto", "default", ""):
            return
        current_range = float(current_range)
    current_range = float(current_range)
    if current_range <= 0:
        raise ValueError("current_range must be positive, 'auto', or None.")
    query(f"RG {channel}, {current_range:.9g}")


def configure_linear_voltage_sweep(
    query,
    *,
    start,
    stop,
    step,
    current_compliance,
    sweep_variable=1,
):
    validate_linear_sweep(start, stop, step)
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
    validate_linear_sweep(start, stop, step)
    query(f"IR{sweep_variable}, {start}, {stop}, {step}, {voltage_compliance}")


def linear_sweep_point_count(start, stop, step):
    """Return the KXCI VAR1 point count defined by the programming manual."""
    start = float(start)
    stop = float(stop)
    step = float(step)
    if step == 0:
        raise ValueError("Linear sweep step must be nonzero.")
    return int(abs((stop - start) / step) + 1.5)


def validate_linear_sweep(start, stop, step):
    """Reject linear sweeps that KXCI cannot represent safely."""
    point_count = linear_sweep_point_count(start, stop, step)
    if point_count < 1:
        raise ValueError("Linear sweep must contain at least one point.")
    if point_count > 1024:
        raise ValueError(
            f"KXCI VAR1 sweeps are limited to 1024 points; requested {point_count}."
        )
    return point_count


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
    channel = validate_channel(channel)
    if not values:
        raise ValueError("values must contain at least one voltage.")
    if len(values) > 4096:
        raise ValueError("KXCI list sweeps are limited to 4096 points.")
    value_text = ",".join(f"{float(value):.12g}" for value in values)
    if mode not in (0, 1):
        raise ValueError("List sweep mode must be 0 (subordinate) or 1 (master).")
    query(f"VL{channel},{mode},{current_compliance},{value_text}")


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
    hold_time = float(hold_time)
    sweep_delay = float(sweep_delay)
    if hold_time < 0:
        raise ValueError("hold_time must be nonnegative.")
    if not 0 <= sweep_delay <= 6.553:
        raise ValueError("sweep_delay must be between 0 and 6.553 seconds.")
    query(f"HT {hold_time:g}")
    query(f"DT {sweep_delay:g}")
    integration = str(integration).strip()
    if not integration.upper().startswith("IT"):
        raise ValueError("integration must be an IT command.")
    query(integration)


def configure_measurement_list(
    query,
    variables,
    *,
    display_mode=2,
    current_ranges=None,
):
    variables = [str(variable) for variable in variables]
    if not variables or len(variables) > 6:
        raise ValueError("KXCI LI requires between 1 and 6 variables.")
    invalid_variables = [
        variable
        for variable in variables
        if not variable
        or (
            len(variable) > 6
            and not (
                len(variable) == 7
                and variable[-1].upper() in {"T", "S"}
                and len(variable[:-1]) <= 6
            )
        )
    ]
    if invalid_variables:
        raise ValueError(
            "KXCI measurement names must contain at most 6 characters; "
            "a seventh character is allowed only for a T/S suffix."
        )
    variable_text = ", ".join(f"'{variable}'" for variable in variables)
    query(f"SM DM{display_mode}")
    if current_ranges:
        for channel, current_range in current_ranges:
            configure_current_range(query, channel, current_range)
    query(f"LI {variable_text}")


def execute_and_wait(
    query,
    *,
    execution_mode=1,
    timeout_s=300.0,
    poll_interval_s=0.2,
):
    """Start a System Mode test and wait for SP to report completion."""
    timeout_s = float(timeout_s)
    poll_interval_s = float(poll_interval_s)
    if timeout_s <= 0:
        raise ValueError("timeout_s must be positive.")
    if poll_interval_s < 0:
        raise ValueError("poll_interval_s must be nonnegative.")
    query("MD")
    query(f"ME{execution_mode}")
    deadline = time.monotonic() + timeout_s
    try:
        while True:
            response = query("SP").strip()
            try:
                status = int(response)
            except ValueError:
                status = None
                raise_for_kxci_error(query, context="SMU execution")
            data_ready = status is not None and bool(status & 0b00000001)
            busy = status is not None and bool(status & 0b00010000)
            if data_ready and not busy:
                raise_for_kxci_error(query, context="SMU execution")
                return status
            if time.monotonic() >= deadline:
                raise TimeoutError(
                    f"SMU System Mode test did not finish within {timeout_s:g} s; "
                    f"last SP response={response!r}."
                )
            time.sleep(poll_interval_s)
    except BaseException:
        for command in ("MD", "ME4"):
            try:
                query(command)
            except BaseException:
                pass
        raise


def _validate_two_channel_run(sweep_channel, bias_channel, available_channels):
    available_channels = normalize_channels(
        available_channels,
        name="available_channels",
    )
    sweep_channel = validate_channel(sweep_channel, name="sweep_channel")
    bias_channel = validate_channel(bias_channel, name="bias_channel")
    if sweep_channel == bias_channel:
        raise ValueError("sweep_channel and bias_channel must be different.")
    missing = {sweep_channel, bias_channel}.difference(available_channels)
    if missing:
        raise ValueError(
            f"Active channels {sorted(missing)} are missing from available_channels."
        )
    return sweep_channel, bias_channel, available_channels


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
    sweep_current_range=None,
    bias_current_range=None,
    hold_time=0.0,
    sweep_delay=0.0,
    integration="IT1",
    return_variables=None,
    timeout_s=300.0,
    available_channels=DEFAULT_AVAILABLE_CHANNELS,
):
    """Configure and execute a two-channel linear voltage sweep."""
    sweep_channel, bias_channel, available_channels = _validate_two_channel_run(
        sweep_channel,
        bias_channel,
        available_channels,
    )
    validate_linear_sweep(start, stop, step)
    try:
        initialize_system_mode(query, available_channels=available_channels)
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
        configure_auto_standby(query, (sweep_channel, bias_channel))
        if return_variables is None:
            return_variables = [
                sweep_current_name,
                sweep_voltage_name,
                bias_current_name,
                bias_voltage_name,
            ]
        configure_measurement_list(
            query,
            return_variables,
            current_ranges=[
                (sweep_channel, sweep_current_range),
                (bias_channel, bias_current_range),
            ],
        )
        raise_for_kxci_error(query, context="Linear voltage sweep setup")
        execute_and_wait(query, execution_mode=1, timeout_s=timeout_s)
        return list(return_variables)
    except BaseException:
        shutdown_system_mode(query, available_channels)
        raise


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
    sweep_current_range=None,
    bias_current_range=None,
    hold_time=0.0,
    sweep_delay=0.0,
    integration="IT1",
    return_variables=None,
    timeout_s=300.0,
    available_channels=DEFAULT_AVAILABLE_CHANNELS,
):
    """Configure and execute an arbitrary two-channel voltage list sweep."""
    values = [float(value) for value in values]
    if not values:
        raise ValueError("values must contain at least one voltage.")

    sweep_channel, bias_channel, available_channels = _validate_two_channel_run(
        sweep_channel,
        bias_channel,
        available_channels,
    )
    try:
        initialize_system_mode(query, available_channels=available_channels)
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
        configure_auto_standby(query, (sweep_channel, bias_channel))
        if return_variables is None:
            return_variables = [
                sweep_current_name,
                sweep_voltage_name,
                bias_current_name,
                bias_voltage_name,
            ]
        configure_measurement_list(
            query,
            return_variables,
            current_ranges=[
                (sweep_channel, sweep_current_range),
                (bias_channel, bias_current_range),
            ],
        )
        raise_for_kxci_error(query, context="List voltage sweep setup")
        execute_and_wait(query, execution_mode=1, timeout_s=timeout_s)
        return list(return_variables)
    except BaseException:
        shutdown_system_mode(query, available_channels)
        raise
