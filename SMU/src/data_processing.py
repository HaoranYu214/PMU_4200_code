"""SMU buffer retrieval and tabular data helpers."""

from __future__ import annotations

from pathlib import Path
import re

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


def retrieve_variables(query, variables, *, include_status=True):
    """Download completed System Mode buffers into a padded DataFrame."""
    columns = {}
    for variable in variables:
        values, statuses = retrieve_variable(query, variable)
        columns[variable] = pd.Series(values, dtype=float)
        if include_status and any(statuses):
            columns[f"{variable}_Status"] = pd.Series(statuses, dtype="string")
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
