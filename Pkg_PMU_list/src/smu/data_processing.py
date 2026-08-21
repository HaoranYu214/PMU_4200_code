"""SMU buffer retrieval and tabular data helpers."""

from __future__ import annotations

from pathlib import Path
import math
import operator
import re
import warnings

import pandas as pd


_NUMBER_AT_END = re.compile(
    r"([+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?)\s*$"
)
_READING_SEPARATOR = re.compile(r"[,\x0e]+")


def parse_kxci_reading(token):
    """Parse one SMU reading and preserve its N/C status prefix.

    System Mode commonly returns ``N/C`` plus a numeric value. User Mode adds
    instrument and measure-mode letters before the number. The numeric value is
    therefore parsed from the end instead of passing the whole string to
    ``float``.
    """
    token = str(token).strip()
    if not token:
        raise ValueError("KXCI returned an empty reading.")
    match = _NUMBER_AT_END.search(token)
    if match is None:
        raise ValueError(f"Unable to parse KXCI reading: {token!r}")
    status = token[0].upper() if token[0].isalpha() else ""
    return float(match.group(1)), status


def retrieve_variable(query, variable):
    """Download one completed System Mode buffer using DO."""
    response = query(f"DO '{variable}'").strip()
    if not response:
        return [], []
    values = []
    statuses = []
    for token in _READING_SEPARATOR.split(response):
        token = token.strip()
        if not token:
            continue
        value, status = parse_kxci_reading(token)
        values.append(value)
        statuses.append(status)
    return values, statuses


def retrieve_variables(
    query,
    variables,
    *,
    include_status=True,
    expected_point_count=None,
    strict=True,
    fail_on_compliance=False,
):
    """Download and validate completed System Mode buffers."""
    variables = [str(variable) for variable in variables]
    if not variables:
        raise ValueError("variables must contain at least one KXCI buffer name.")
    if len(set(variables)) != len(variables):
        raise ValueError("variables must not contain duplicate KXCI buffer names.")
    if expected_point_count is not None:
        try:
            expected_point_count = operator.index(expected_point_count)
        except TypeError as exc:
            raise ValueError("expected_point_count must be a positive integer.") from exc
        if expected_point_count < 1:
            raise ValueError("expected_point_count must be a positive integer.")
    columns = {}
    lengths = {}
    compliance_hits = {}
    for variable in variables:
        values, statuses = retrieve_variable(query, variable)
        lengths[variable] = len(values)
        if any(not math.isfinite(value) or abs(value) >= 1e36 for value in values):
            raise ValueError(f"{variable} contains invalid/overflow KXCI readings.")
        compliance_count = sum(status.upper() == "C" for status in statuses)
        if compliance_count:
            compliance_hits[variable] = compliance_count
        columns[variable] = pd.Series(values, dtype=float)
        if include_status and any(statuses):
            columns[f"{variable}_Status"] = pd.Series(statuses, dtype="string")

    if strict:
        empty = [variable for variable, length in lengths.items() if length == 0]
        if empty:
            raise ValueError(f"KXCI returned empty buffers for: {', '.join(empty)}.")
        if len(set(lengths.values())) > 1:
            raise ValueError(f"KXCI buffer lengths do not match: {lengths}.")
        if expected_point_count is not None:
            wrong = {
                variable: length
                for variable, length in lengths.items()
                if length != expected_point_count
            }
            if wrong:
                raise ValueError(
                    f"Expected {expected_point_count} points per variable; received {wrong}."
                )

    if compliance_hits:
        message = f"KXCI compliance status detected: {compliance_hits}."
        if fail_on_compliance:
            raise RuntimeError(message)
        warnings.warn(message, RuntimeWarning, stacklevel=2)
    return pd.DataFrame(columns)


def retrieve_realtime_point(query, variable, index):
    """Retrieve one RD point.

    A raw response of ``0`` means the point is not ready yet. It is not an
    end-of-buffer marker and must not be used to discover the result length.
    """
    response = query(f"RD '{variable}', {index}").strip()
    if response in ("", "0"):
        return None
    return parse_kxci_reading(response)


def retrieve_realtime_variables(
    query,
    variables,
    *,
    point_count,
    timestamp_variable,
    poll_interval_s=0.05,
    timeout_s=300.0,
):
    """Retrieve a known number of points in real time using RD.

    The timestamp is polled first, following the manual's Example 4. A nonzero
    timestamp indicates that all requested values for that point are ready.
    """
    import time

    rows = []
    deadline = time.monotonic() + timeout_s
    for index in range(1, int(point_count) + 1):
        while True:
            timestamp = retrieve_realtime_point(query, timestamp_variable, index)
            if timestamp is not None:
                break
            if time.monotonic() >= deadline:
                raise TimeoutError(
                    f"SMU point {index}/{point_count} was not ready within "
                    f"{timeout_s:g} s."
                )
            time.sleep(poll_interval_s)

        row = {
            timestamp_variable: timestamp[0],
            f"{timestamp_variable}_Status": timestamp[1],
        }
        for variable in variables:
            reading = retrieve_realtime_point(query, variable, index)
            if reading is None:
                raise RuntimeError(
                    f"{variable} point {index} was unavailable after its "
                    "timestamp became ready."
                )
            row[variable] = reading[0]
            row[f"{variable}_Status"] = reading[1]
        rows.append(row)
    return pd.DataFrame(rows)


def save_workbook(path, data, parameters=None):
    """Save SMU data and optional parameters to one Excel workbook."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        data.to_excel(writer, sheet_name="Data", index=False)
        if parameters is not None:
            parameter_df = pd.DataFrame(
                [{"name": key, "value": repr(value)} for key, value in parameters.items()]
            )
            parameter_df.to_excel(writer, sheet_name="Parameters", index=False)
    return path
