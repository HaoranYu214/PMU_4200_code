"""4200A SMU User Mode spot-source and measurement helpers."""

from __future__ import annotations

from .data_processing import parse_kxci_reading


def initialize_user_mode(query):
    query("EM 1,0")
    query("BC")
    query("*RST")
    query("US")


def source_voltage(query, channel, voltage, current_compliance, *, range_code=0):
    """Configure an SMU channel with DV in User Mode."""
    query(f"DV{channel}, {range_code}, {voltage}, {current_compliance}")


def source_current(query, channel, current, voltage_compliance, *, range_code=0):
    """Configure an SMU channel with DI in User Mode."""
    query(f"DI{channel}, {range_code}, {current}, {voltage_compliance}")


def measure_current(query, channel):
    value, _status = parse_kxci_reading(query(f"TI{channel}"))
    return value


def measure_voltage(query, channel):
    value, _status = parse_kxci_reading(query(f"TV{channel}"))
    return value


def power_off_voltage_source(query, channel):
    query(f"DV{channel}")


def power_off_current_source(query, channel):
    query(f"DI{channel}")


def power_off_user_channels(query, voltage_channels=(), current_channels=()):
    """Best-effort User Mode shutdown for channels configured by DV or DI."""
    for channel in voltage_channels:
        try:
            power_off_voltage_source(query, channel)
        except Exception:
            pass
    for channel in current_channels:
        try:
            power_off_current_source(query, channel)
        except Exception:
            pass
