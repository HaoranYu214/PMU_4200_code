# -*- coding: utf-8 -*-
"""
双通道 4225-PMU 脉冲训练测试（精简版）
- 仅真机路径
- 扁平流程：测试 -> 读取 -> 关输出 -> 分析 -> 绘图
"""

# 导入自定义模块
from src.plotting_utils import PlotManager, plot_time_series
from src.data_processing import save_channels_separate_excel, read_both_channels, merge_channels, add_resistance_columns
from src.pmu_tests import dual_channel_pulse_train, power_off_outputs
from src.instrcomms import Communications
from pathlib import Path

# ==================== 配置区 ====================
# ==================== 配置区 ====================

# 测试参数字典
params = dict(
    # # CH1
    # CH1_BASE=0.0, CH1_AMPLITUDE=-5, 
    # CH1_PERIOD=200e-6, CH1_WIDTH=50e-6, CH1_RISE=5e-6, CH1_FALL=5e-6, CH1_DELAY=10e-6, 
    # CH1_RANGE=1e-3,
    # # CH2
    # CH2_BASE=0.0, CH2_AMPLITUDE=0.1,
    # CH2_PERIOD=200e-6, CH2_WIDTH=80e-6, CH2_RISE=10e-6, CH2_FALL=10e-6, CH2_DELAY=90e-6, 
    # CH2_RANGE=1e-4,

    # CH1
    CH1_BASE=0.0, CH1_AMPLITUDE=-2, 
    CH1_PERIOD=2000e-6, CH1_WIDTH=800e-6, CH1_RISE=50e-6, CH1_FALL=50e-6, CH1_DELAY=50e-6, 
    CH1_RANGE=1e-3,
    # CH2
    CH2_BASE=0.0, CH2_AMPLITUDE=0.2,
    CH2_PERIOD=2000e-6, CH2_WIDTH=800e-6, CH2_RISE=50e-6, CH2_FALL=50e-6, CH2_DELAY=1050e-6, 
    CH2_RANGE=1e-4,



    # 测量与burst
    PULSE_COUNT=100, 
    # Discrete mode
    MEASURE_START_D=0.6, MEASURE_STOP_D=0.8,
    # Waveform mode
    MEASURE_START_W=0.2, MEASURE_STOP_W=0.2,
    ENABLE_LOAD_CONFIG=True, LOAD_RESISTANCE=1e6,
    ENABLE_LLEC=False, ENABLE_CONNECTION_COMP=False,
    # 保存/分析参数-3
    CURRENT_EPS=1e-12, RES_MIN=1.0, RES_MAX=1e15
)

# 连接配置
INST = "TCPIP0::129.125.87.80::1225::SOCKET"  # 你的仪器地址
CH1, CH2 = 1, 2

# 测试模式选择 - 便捷开关！
TEST_MODE = 1  # 选择：
                   # 'D' 或 1: Spot mean discrete (离散点测量)
                   # 'WFM' 或 2: Waveform discrete (离散波形测量) 
                   # 0: No measurements (无测量)
                   # 3: Spot mean average (平均点测量)
                   # 4: Waveform average (平均波形测量)

# 绘图选项
RESISTANCE_SCALE = 'linear'  # 电阻坐标模式：'log'=对数坐标，'linear'=线性坐标

# ==================== main ====================
width_us = int(params["CH1_WIDTH"] * 1e6)
amp_v    = params["CH1_AMPLITUDE"]

# 连接
k = Communications(INST)
k.connect()
k._instrument_object.write_termination = "\0"
k._instrument_object.read_termination = "\0"
Q = k.query

try:
    print("▶️ 运行脉冲训练...")
    dual_channel_pulse_train(Q, CH1, CH2, params, mode=TEST_MODE)
    df1, df2 = read_both_channels(Q, CH1, CH2)
    power_off_outputs(Q, (CH1, CH2))
    if df1 is None or df2 is None or df1.empty or df2.empty:
        raise ValueError("脉冲训练读取数据为空")
    dfs = {1: df1, 2: df2}
    merged = add_resistance_columns(
        merge_channels(dfs),
        eps=params.get('CURRENT_EPS', 1e-12),
        res_min=params.get('RES_MIN', 1.0),
        res_max=params.get('RES_MAX', 1e15)
    )
    # 保存分通道
    SAVE_DIR = Path(r"C:\Users\P317151\Documents\data\FTJ\Refined")
    SAVE_DIR.mkdir(parents=True, exist_ok=True)
    fname_base = SAVE_DIR / f"2C_PulseTrain_mode{TEST_MODE}_{width_us}us_{params['CH1_AMPLITUDE']}V_{params['PULSE_COUNT']}x"
    save_channels_separate_excel(dfs, f"{fname_base}_raw.xlsx")
    # 绘图
    with PlotManager(mode='batched', block=True, close_after_show=False) as pm:
        pm.add(plot_time_series(merged, width_us=width_us,
                                amp_v=params['CH1_AMPLITUDE'],
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

# 摘要
if merged is not None and not merged.empty:
    print("\n📊 最终数据摘要：")
    print("  形状   :", merged.shape)
    for ch in (1, 2):
        for col in (f"Voltage {ch}", f"Current {ch}", f"Resistance {ch}"):
            if col in merged.columns:
                s = merged[col].dropna()
                if not s.empty:
                    print(f"  {col}: min={s.min():.3e}, max={s.max():.3e}")
else:
    print("⚠️ 无有效数据")
