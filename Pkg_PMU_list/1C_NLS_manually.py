# -*- coding: utf-8 -*-
"""
NIS Switch 测试：铁电隧穿结电阻切换测量
"""
from src.data_processing import save_channels_separate_excel, calculate_polarization
from src.pmu_tests import hy_NISswitch_segARB, power_off_outputs
from src.data_processing import read_both_channels
from src.instrcomms import Communications
import matplotlib.pyplot as plt
import pandas as pd
import numpy as np
from pathlib import Path

INST = "TCPIP0::129.125.87.80::1225::SOCKET"
CH1, CH2 = 1, 2
params = dict(
    offset=0,
    Vp=5,
    Rt_p=5e-5,
    Delaytime=100e-6,

    Vsquare=1.5,
    Rt_s=1e-7,
    Dwell=1e-4,

    Irange1=1e-4,
    Irange2=1e-4,

    MeasureSquare= False,
    area_cm2=7.0686e-6,
)

SAVE_DIR = Path(r"C:\Users\P317151\Documents\data\Jingtian\2025-12-14\BTO\Device2")
SAVE_DIR.mkdir(parents=True, exist_ok=True)
fname_base = SAVE_DIR / f"NISswitch_Meas{params['Vp']}V_{int(params['Rt_p']*1e6)}us_with{params['Vsquare']}V_{int(params['Rt_s']*1e6)}us"

k = Communications(INST)
k.connect()
k._instrument_object.write_termination = "\0"
k._instrument_object.read_termination = "\0"
Q = k.query

try:
    print("▶️ 运行 NIS Switch ...")
    hy_NISswitch_segARB(Q, CH1, CH2, params)
    df_ch1, df_ch2 = read_both_channels(Q, CH1, CH2)
    power_off_outputs(Q, (CH1, CH2))
    if df_ch1 is None or df_ch2 is None or df_ch1.empty or df_ch2.empty:
        raise ValueError("NIS Switch读取数据为空")
    
    # 保存原始数据
    save_channels_separate_excel({1: df_ch1, 2: df_ch2}, f"{fname_base}_raw.xlsx")
    
    if params['MeasureSquare']:
        # 方波测量模式：绘制 V-t 和 I-t 双轴图
        fig, axes = plt.subplots(2, 1, figsize=(12, 8), sharex=True)
        
        # CH1
        t1 = df_ch1[f'Timestamp {CH1}']
        v1 = df_ch1[f'Voltage {CH1}']
        i1 = df_ch1[f'Current {CH1}']
        
        ax1_v = axes[0]
        ax1_i = ax1_v.twinx()
        ax1_v.plot(t1, v1, 'b-', linewidth=0.8, label='Voltage')
        ax1_i.plot(t1, i1*1e6, 'r-', linewidth=0.8, label='Current')
        ax1_v.set_ylabel("Voltage (V)", color='b')
        ax1_i.set_ylabel("Current (μA)", color='r')
        ax1_v.set_title("CH1 Waveform")
        ax1_v.grid(alpha=0.3)
        
        # CH2
        t2 = df_ch2[f'Timestamp {CH2}']
        v2 = df_ch2[f'Voltage {CH2}']
        i2 = df_ch2[f'Current {CH2}']
        
        ax2_v = axes[1]
        ax2_i = ax2_v.twinx()
        ax2_v.plot(t2, v2, 'b-', linewidth=0.8, label='Voltage')
        ax2_i.plot(t2, i2*1e6, 'r-', linewidth=0.8, label='Current')
        ax2_v.set_xlabel("Time (s)")
        ax2_v.set_ylabel("Voltage (V)", color='b')
        ax2_i.set_ylabel("Current (μA)", color='r')
        ax2_v.set_title("CH2 Waveform")
        ax2_v.grid(alpha=0.3)
        
        fig.suptitle("NIS Switch - Waveform Check", fontsize=14)
        fig.tight_layout()
        fig.savefig(f"{fname_base}_waveform.png", dpi=300)
        print(f"✅ 方波测量模式完成 (CH1: {len(df_ch1)} pts, CH2: {len(df_ch2)} pts)")
        
    else:
        # 非方波测量模式：只测量后4段（2个三角波），每段点数 = 总点数 / 4
        area_cm2 = params.get('area_cm2', 1.0)
        
        def process_channel(df, ch_num):
            """处理单通道：4段数据，前2段 vs 后2段做差分"""
            v = df[f'Voltage {ch_num}'].values
            i = df[f'Current {ch_num}'].values
            t = df[f'Timestamp {ch_num}'].values
            
            n = len(v)
            seg_pts = n // 4
            
            # 前2段 (三角波1) vs 后2段 (三角波2)
            v_tri1 = v[:2*seg_pts]
            i_tri1 = i[:2*seg_pts]
            t_tri1 = t[:2*seg_pts]
            
            i_tri2 = i[2*seg_pts:4*seg_pts]
            
            # 差分
            min_len = min(len(i_tri1), len(i_tri2))
            i_diff = i_tri1[:min_len] - i_tri2[:min_len]
            v_diff = v_tri1[:min_len]
            dt = params['Rt_p'] * 4 / n
            t_diff = np.arange(1, min_len + 1) * dt  # [dt, 2*dt, 3*dt, ...]
            
            # 极化
            P = calculate_polarization(i_diff, t_diff, area_cm2)
            
            return pd.DataFrame({
                'Time': t_diff, 'Voltage': v_diff,
                'DiffCurrent': i_diff, 'Polarization': P
            })
        
        df_vp_ch1 = process_channel(df_ch1, CH1)
        df_vp_ch2 = process_channel(df_ch2, CH2)
        
        # 保存
        df_vp_ch1.to_excel(f"{fname_base}_vp_ch1.xlsx", index=False)
        df_vp_ch2.to_excel(f"{fname_base}_vp_ch2.xlsx", index=False)
        
        # 绘制 V-P 曲线
        fig, axes = plt.subplots(1, 2, figsize=(12, 5))
        
        axes[0].plot(df_vp_ch1['Voltage'], df_vp_ch1['Polarization'], 'b-', linewidth=1)
        axes[0].set_xlabel("Voltage (V)")
        axes[0].set_ylabel("Polarization (μC/cm²)")
        axes[0].set_title("CH1 V-P (Tri1 - Tri2)")
        axes[0].grid(alpha=0.3)
        
        axes[1].plot(df_vp_ch2['Voltage'], df_vp_ch2['Polarization'], 'r-', linewidth=1)
        axes[1].set_xlabel("Voltage (V)")
        axes[1].set_ylabel("Polarization (μC/cm²)")
        axes[1].set_title("CH2 V-P (Tri1 - Tri2)")
        axes[1].grid(alpha=0.3)
        
        fig.suptitle("NIS Switch - Differential Polarization", fontsize=14)
        fig.tight_layout()
        fig.savefig(f"{fname_base}_vp.png", dpi=300)
        print(f"✅ 完成 (CH1: {len(df_ch1)} pts → {len(df_vp_ch1)} diff pts)")
except Exception as e:
    print(f"❌ 运行出错: {e}")
    raise
finally:
    try:
        k.disconnect()
    except Exception as e:
        print(f"⚠️ 断开连接时出错: {e}")
