# -*- coding: utf-8 -*-
"""Batch parameter sweep for the NLS switch test."""

from pathlib import Path
import time

import numpy as np
import pandas as pd

from NLS_1C_switch import run_nls_switch_test
from src.session import PMUSession

INST = "TCPIP0::129.125.87.80::1225::SOCKET"
CH1, CH2 = 1, 2

BASE_PARAMS = dict(
    offset=0,
    Vp=2,
    Rt_p=5e-5,
    Delaytime=100e-6,
    Rt_s=1e-7,
    Dwell=1e-6,
    Irange1=1e-4,
    Irange2=1e-4,
    MeasureSquare=False,
    area_cm2=7.0686e-6,
)

SAVE_DIR = Path(r"C:\Users\P317151\Documents\data\Jingtian\2025-12-14\BTO\Device3\NLS_Sweep")
Dwell_list = np.logspace(-7, -1, 31)
Vsquare_list = np.arange(0, 2.0, 0.1)


def run_sweep():
    """Run a Dwell x Vsquare sweep and save a summary workbook."""
    results = []
    summary_data = []
    total_tests = len(Dwell_list) * len(Vsquare_list)
    test_idx = 0
    interrupted = False

    with PMUSession(INST, channels=(CH1, CH2)) as session:
        query = session.query
        try:
            for dwell in Dwell_list:
                for vsquare in Vsquare_list:
                    test_idx += 1
                    print(f"[{test_idx}/{total_tests}] Dwell={dwell:.1e}s, Vsquare={vsquare:.2f}V")
                    params = BASE_PARAMS.copy()
                    params["Vsquare"] = vsquare
                    params["Dwell"] = dwell

                    try:
                        result = run_nls_switch_test(query, CH1, CH2, params, SAVE_DIR)
                        result["params"] = params.copy()
                        results.append(result)
                        if "df_vp_ch1" in result and result["df_vp_ch1"] is not None:
                            pol = result["df_vp_ch1"]["Polarization"].values
                            pr_value = -pol[0] + pol[-1] if len(pol) > 0 else None
                        else:
                            pr_value = None
                        summary_data.append({"Vsquare": vsquare, "Dwell": dwell, "Pr": pr_value})
                    except KeyboardInterrupt:
                        raise
                    except Exception as exc:
                        print(f"  Test failed: {exc}")
                        results.append({"params": params.copy(), "success": False, "error": str(exc)})
                        summary_data.append({"Vsquare": vsquare, "Dwell": dwell, "Pr": None})
                    time.sleep(0.5)
        except KeyboardInterrupt:
            interrupted = True
            print(f"User interrupted after {test_idx}/{total_tests} test points.")

    if summary_data:
        df_summary = pd.DataFrame(summary_data)
        SAVE_DIR.mkdir(parents=True, exist_ok=True)
        df_summary.to_excel(SAVE_DIR / "total.xlsx", index=False)

    status = "interrupted" if interrupted else "complete"
    success_count = sum(1 for result in results if result.get("success", False))
    print(f"Sweep {status}: {success_count}/{len(results)} successful.")
    return results


if __name__ == "__main__":
    run_sweep()
