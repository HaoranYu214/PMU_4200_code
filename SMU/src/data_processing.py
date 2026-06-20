"""SMU buffer retrieval and tabular data helpers."""

from __future__ import annotations

from pathlib import Path

import pandas as pd


def retrieve_variable(query, variable, *, start_index=1, max_points=None):
    """Retrieve one System Mode variable using indexed RD queries."""
    values = []
    index = start_index
    while max_points is None or len(values) < max_points:
        response = query(f"RD '{variable}', {index}").strip()
        if response == "0" or response == "":
            break
        values.append(float(response))
        index += 1
    return values


def retrieve_variables(query, variables, *, max_points=None):
    """Retrieve multiple named variables and return a padded DataFrame."""
    series = {
        variable: pd.Series(
            retrieve_variable(query, variable, max_points=max_points),
            dtype=float,
        )
        for variable in variables
    }
    return pd.DataFrame(series)


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
