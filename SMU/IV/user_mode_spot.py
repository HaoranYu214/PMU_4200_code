# -*- coding: utf-8 -*-
"""User Mode spot I-V example with explicit compliance and shutdown."""

from pathlib import Path
import sys
import time

REPO_ROOT = Path(__file__).resolve().parents[2]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from src.smu.session import SMUSession
from src.smu.user_mode import (
    initialize_user_mode,
    measure_current,
    power_off_voltage_source,
    restore_user_mode_rpms,
    source_voltage,
)


INST = "TCPIP0::129.125.87.80::1225::SOCKET"
CHANNEL = 1
SOURCE_VOLTAGE = 0.1
CURRENT_COMPLIANCE = 1e-3
VOLTAGE_RANGE_CODE = 0
SETTLE_TIME_S = 0.1

# Physical SMU-to-probe wiring for this 4200A.
SMU_CONNECTIONS = {
    1: "rpm:PMU1-1",
    2: "rpm:PMU1-2",
    3: "direct",
    4: "direct",
}


def main():
    with SMUSession(INST) as session:
        query = session.query
        rpm_targets = initialize_user_mode(
            query,
            active_channels=(CHANNEL,),
            smu_connections=SMU_CONNECTIONS,
        )
        try:
            source_voltage(
                query,
                CHANNEL,
                SOURCE_VOLTAGE,
                CURRENT_COMPLIANCE,
                range_code=VOLTAGE_RANGE_CODE,
            )
            time.sleep(SETTLE_TIME_S)
            current = measure_current(query, CHANNEL)
            print(f"CH{CHANNEL}: V={SOURCE_VOLTAGE:g} V, I={current:.6e} A")
        finally:
            # Do not restore an RPM relay unless DV shutdown succeeds. If
            # shutdown raises, leave the RPM blue/SMU-routed for safe diagnosis.
            power_off_voltage_source(query, CHANNEL)
            restore_user_mode_rpms(query, rpm_targets)


if __name__ == "__main__":
    main()
