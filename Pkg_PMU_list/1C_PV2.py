# -*- coding: utf-8 -*-
"""
PV2 测试（精简版）：仅真机 + 封装调用
"""
from src.data_processing import save_channels_separate_excel, read_both_channels, calculate_polarization
from src.pmu_tests import hy_pv2_segARB, power_off_outputs
from src.instrcomms import Communications
import matplotlib.pyplot as plt
import pandas as pd
from pathlib import Path

INST = "TCPIP0::129.125.87.80::1225::SOCKET"
CH1, CH2 = 1, 2
params = dict(
    rise_time=5e-5,
    rise_point=200,
    Vp=5,
    offset=0,
    area_cm2=7.0686e-6,
    Irange1=1e-4,
    Irange2=1e-4,
)

SAVE_DIR = Path(r"C:\Users\P317151\Documents\data\Jingtian\2025-12-14\BTO\Device2")
SAVE_DIR.mkdir(parents=True, exist_ok=True)
fname_base = SAVE_DIR / f"PV2_{int(params['rise_time']*1e6)}us_{params['Vp']}V"

k = Communications(INST)
k.connect()
k._instrument_object.write_termination = "\0"
k._instrument_object.read_termination = "\0"
Q = k.query

try:
    print("▶️ 运行 PV2 ...")
    hy_pv2_segARB(Q, CH1, CH2, params)
    df_ch1, df_ch2 = read_both_channels(Q, CH1, CH2)
    power_off_outputs(Q, (CH1, CH2))
    if df_ch1 is None or df_ch1.empty:
        raise ValueError("PV2读取数据为空")
    P = calculate_polarization(
        df_ch1[f'Current {CH1}'].values,
        df_ch1[f'Timestamp {CH1}'].values,
        params.get('area_cm2', 1.0)
    )
    df_pv2 = pd.DataFrame({
        'Time': df_ch1[f'Timestamp {CH1}'].values,
        'Voltage': df_ch1[f'Voltage {CH1}'].values,
        'Current': df_ch1[f'Current {CH1}'].values,
        'Polarization': P
    })
    save_channels_separate_excel({1: df_ch1, 2: df_ch2}, f"{fname_base}_raw.xlsx")
    df_pv2.to_excel(f"{fname_base}_pv2.xlsx", index=False)

    # 简单PV绘图
    fig, ax = plt.subplots(figsize=(6,5))
    ax.plot(df_pv2['Voltage'], df_pv2['Polarization'], 'b-')
    ax.set_xlabel("Voltage (V)")
    ax.set_ylabel("Polarization (μC/cm²)")
    ax.set_title("PV2 Loop")
    ax.grid(alpha=0.3)
    fig.tight_layout()
    fig.savefig(f"{fname_base}_loop.png", dpi=300)
    print("✅ 完成")
except Exception as e:
    print(f"❌ 运行出错: {e}")
    raise
finally:
    try:
        k.disconnect()
    except Exception as e:
        print(f"⚠️ 断开连接时出错: {e}")
