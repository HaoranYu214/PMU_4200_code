# -*- coding: utf-8 -*-
"""User Mode spot I-V example with explicit compliance and shutdown."""

from pathlib import Path
import sys
import time

SMU_ROOT = Path(__file__).resolve().parents[1]
if str(SMU_ROOT) not in sys.path:
    sys.path.insert(0, str(SMU_ROOT))

from src.session import SMUSession
from src.user_mode import (
    initialize_user_mode,
    measure_current,
    power_off_user_channels,
    source_voltage,
)


INST = "TCPIP0::129.125.87.80::1225::SOCKET"
CHANNEL = 1
SOURCE_VOLTAGE = 0.1
CURRENT_COMPLIANCE = 1e-3
VOLTAGE_RANGE_CODE = 0
SETTLE_TIME_S = 0.1


def main():
    with SMUSession(INST) as session:
        query = session.query
        initialize_user_mode(query)
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
            power_off_user_channels(query, voltage_channels=(CHANNEL,))


if __name__ == "__main__":
    main()
