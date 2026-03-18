# -*- coding: utf-8 -*-
"""
双通道 扫描 + 训练（精简版）
"""
from src.plotting_utils import PlotManager, plot_time_series
from src.data_processing import save_channels_separate_excel, read_both_channels, merge_channels, add_resistance_columns
from src.pmu_tests import dual_channel_sweep_train, power_off_outputs
from src.instrcomms import Communications
from pathlib import Path

params = dict(
    CH1_START=-2, CH1_STOP=2, CH1_STEP=0.5, CH1_VBASE=0.0, CH1_DUALSWEEP=1,
    CH1_PERIOD=2000e-6, CH1_WIDTH=750e-6, CH1_RISE=50e-6, CH1_FALL=50e-6, CH1_DELAY=100e-6,
    CH1_RANGE=1e-3,
    CH2_BASE=0, CH2_AMPLITUDE=0.5,
    CH2_PERIOD=2000e-6, CH2_WIDTH=750e-6, CH2_RISE=50e-6, CH2_FALL=50e-6, CH2_DELAY=1100e-6,
    CH2_RANGE=1e-7,
    PULSE_COUNT=1,
    MEASURE_START_D=0.6, MEASURE_STOP_D=0.8,
    MEASURE_START_W=0.2, MEASURE_STOP_W=0.2,
    ENABLE_LOAD_CONFIG=False, LOAD_RESISTANCE=1e6,
    ENABLE_LLEC=False, ENABLE_CONNECTION_COMP=False,
    CURRENT_EPS=1e-12, RES_MIN=1.0, RES_MAX=1e15
)
TEST_MODE = 2
INST = "TCPIP0::129.125.87.80::1225::SOCKET"
CH1, CH2 = 1, 2
RESISTANCE_SCALE = 'linear'
SAVE_DIR = Path(r"C:\Users\P317151\Documents\data\FTJ\Refined")
SAVE_DIR.mkdir(parents=True, exist_ok=True)

width_us = int(params["CH1_WIDTH"] * 1e6)
fname_base = f"2C_SweepTrain_mode{TEST_MODE}_{width_us}us_{params['CH1_START']}to{params['CH1_STOP']}V"

k = Communications(INST)
k.connect()
k._instrument_object.write_termination = "\0"
k._instrument_object.read_termination = "\0"
Q = k.query

try:
    print("▶️ 运行扫描+训练...")
    dual_channel_sweep_train(Q, CH1, CH2, params, mode=TEST_MODE)
    df1, df2 = read_both_channels(Q, CH1, CH2)
    power_off_outputs(Q, (CH1, CH2))
    if df1 is None or df2 is None or df1.empty or df2.empty:
        raise ValueError("扫描测试读取数据为空")
    dfs = {1: df1, 2: df2}
    merged = add_resistance_columns(
        merge_channels(dfs),
        eps=params.get('CURRENT_EPS', 1e-12),
        res_min=params.get('RES_MIN', 1.0),
        res_max=params.get('RES_MAX', 1e15)
    )
    save_channels_separate_excel(dfs, str(SAVE_DIR / f"{fname_base}_raw.xlsx"))
    with PlotManager(mode='batched', block=True, close_after_show=False) as pm:
        pm.add(plot_time_series(merged, width_us=width_us,
                                amp_v=params['CH1_STOP'],
                                resistance_scale=RESISTANCE_SCALE,
                                show=False, return_fig=True))
    print("✅ 完成")
except Exception as e:
    print(f"❌ 运行出错: {e}")
    raise
finally:
    try:
        k.disconnect()
    except Exception as e:
        print(f"⚠️ 断开连接时出错: {e}")
