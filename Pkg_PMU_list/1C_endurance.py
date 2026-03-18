# -*- coding: utf-8 -*-
"""
Endurance 疲劳测试（精简版）
循环执行：n × cycle + PV2 read + PUND read
每次循环独立保存数据，避免数据点超限
"""
from src.data_processing import save_channels_separate_excel
from src.pmu_tests import hy_pv2_segARB, hy_pund_segARB, execute_segARB_test, power_off_outputs
from src.data_processing import read_both_channels, calculate_polarization, analyze_pund_diff
from src.instrcomms import Communications
from pathlib import Path
import matplotlib.pyplot as plt
import pandas as pd

# ==================== 配置区 ====================
INST = "TCPIP0::129.125.87.80::1225::SOCKET"
CH1, CH2 = 1, 2

# 三个独立的参数字典
params_cycle = dict(
    rise_time_cycle=10e-6,  # 快速循环
    Vc=3.0,
    offset_c=0,
    Irange1=1e-3,
    Irange2=1e-3,
)

params_pv2 = dict(
    rise_time=50e-6,  # 慢速读取
    Vp=2.0,
    offset=0,
    area_cm2=1.2567e-3,
    Irange1=1e-3,
    Irange2=1e-3,
)

params_pund = dict(
    rise_time=50e-6,
    Vp=1.5,
    offset=0,
    area_cm2=1.2567e-3,
    Irange1=1e-3,
    Irange2=1e-3,
)

# 循环次数列表（对数台阶）
cycle_counts = [1, 10, 100, 1000, 10000, 1e5, 1e6, 1e7]  # 可自定义：[1, 10, 100, 1000, 10000]

# 保存路径
SAVE_DIR = Path(r"C:\Users\P317151\Documents\data\FTJ\Refined\Endurance")
SAVE_DIR.mkdir(parents=True, exist_ok=True)

# ==================== 主程序 ====================
k = Communications(INST)
k.connect()
k._instrument_object.write_termination = "\0"
k._instrument_object.read_termination = "\0"
Q = k.query

try:
    print(f"🔄 开始 Endurance 测试，总共 {len(cycle_counts)} 个台阶")
    print(f"   循环次数: {cycle_counts}")
    
    for i, n_cycles in enumerate(cycle_counts, 1):
        print(f"\n{'='*60}")
        print(f"📍 台阶 {i}/{len(cycle_counts)}: {n_cycles} 次循环")
        print(f"{'='*60}")
        
        try:
            # 1. 执行 n 次循环（不测量）
            print(f"🔁 执行 {int(n_cycles)} 次循环...")
            current_ranges = {CH1: params_cycle['Irange1'], CH2: params_cycle['Irange2']}
            rst_c = params_cycle['rise_time_cycle']
            Vc = params_cycle['Vc']
            offset_c = params_cycle.get('offset_c', 0)
            start_v = [offset_c, -Vc+offset_c, offset_c, Vc+offset_c]
            stop_v  = [-Vc+offset_c, offset_c, Vc+offset_c, offset_c]
            time_v  = [rst_c, rst_c, rst_c, rst_c]
            meas_types = [0, 0, 0, 0]
            ch1_config = (1, start_v, stop_v, time_v, meas_types, [0.0]*4, [1]*4)
            ch2_config = (1, [0.0]*4, [0.0]*4, time_v, meas_types, [0.0]*4, [1]*4)
            seq_configs = {CH1: [ch1_config], CH2: [ch2_config]}
            seq_list = {CH1: [(1, int(n_cycles))], CH2: [(1, int(n_cycles))]}
            execute_segARB_test(Q, [CH1, CH2], seq_configs, seq_list=seq_list, current_ranges=current_ranges)
            power_off_outputs(Q, (CH1, CH2))
            print(f"   ✅ {int(n_cycles)} 次循环完成")
            
            # 2. PV2 读取
            print(f"📊 执行 PV2 读取...")
            hy_pv2_segARB(Q, CH1, CH2, params_pv2)
            pv2_ch1, pv2_ch2 = read_both_channels(Q, CH1, CH2)
            power_off_outputs(Q, (CH1, CH2))
            if pv2_ch1 is None or pv2_ch1.empty:
                raise ValueError("PV2读取数据为空")
            pv2_p = calculate_polarization(
                pv2_ch1[f'Current {CH1}'].values,
                pv2_ch1[f'Timestamp {CH1}'].values,
                params_pv2.get('area_cm2', 1.0)
            )
            pv2_df = pd.DataFrame({
                'Time': pv2_ch1[f'Timestamp {CH1}'].values,
                'Voltage': pv2_ch1[f'Voltage {CH1}'].values,
                'Current': pv2_ch1[f'Current {CH1}'].values,
                'Polarization': pv2_p
            })
            
            # 保存 PV2 数据
            fname_pv2 = SAVE_DIR / f"Endurance_PV2_after_{int(n_cycles)}cycles"
            save_channels_separate_excel(
                {1: pv2_ch1, 2: pv2_ch2},
                f"{fname_pv2}_raw.xlsx"
            )
            pv2_df.to_excel(f"{fname_pv2}_analysis.xlsx", index=False)
            
            # 简单 PV2 绘图
            fig, ax = plt.subplots(figsize=(6, 5))
            ax.plot(pv2_df['Voltage'], pv2_df['Polarization'], 'b-')
            ax.set_xlabel("Voltage (V)")
            ax.set_ylabel("Polarization (μC/cm²)")
            ax.set_title(f"PV2 after {int(n_cycles)} cycles")
            ax.grid(alpha=0.3)
            fig.tight_layout()
            fig.savefig(f"{fname_pv2}_loop.png", dpi=300)
            plt.close(fig)
            
            # 3. PUND 读取
            print(f"📊 执行 PUND 读取...")
            hy_pund_segARB(Q, CH1, CH2, params_pund)
            df_ch1, df_ch2 = read_both_channels(Q, CH1, CH2)
            power_off_outputs(Q, (CH1, CH2))
            if df_ch1 is None or df_ch2 is None or df_ch1.empty or df_ch2.empty:
                raise ValueError("PUND读取数据为空")
            pund_result = analyze_pund_diff(df_ch1, df_ch2, params_pund)
            
            # 保存 PUND 数据
            fname_pund = SAVE_DIR / f"Endurance_PUND_after_{int(n_cycles)}cycles"
            save_channels_separate_excel(
                {1: df_ch1, 2: df_ch2},
                f"{fname_pund}_raw.xlsx"
            )
            pund_result['df_total'].to_excel(f"{fname_pund}_total.xlsx", index=False)
            pund_result['pund_diff'].to_excel(f"{fname_pund}_diff.xlsx", index=False)
            
            # PUND 差分绘图
            fig, ax = plt.subplots(figsize=(7, 5))
            for seg in ['P', 'U', 'N', 'D']:
                sub = pund_result['pund_diff'][pund_result['pund_diff']['Segment'] == seg]
                ax.plot(sub['Voltage'], sub['Polarization'], '.', label=seg, markersize=4)
            ax.set_xlabel("Voltage (V)")
            ax.set_ylabel("Polarization (μC/cm²)")
            ax.set_title(f"PUND after {int(n_cycles)} cycles")
            ax.legend()
            ax.grid(alpha=0.3)
            fig.tight_layout()
            fig.savefig(f"{fname_pund}_loop.png", dpi=300)
            plt.close(fig)
            
            print(f"✅ 台阶 {i} 完成")
            
        except Exception as e:
            print(f"❌ 台阶 {i} 失败: {e}")
            continue
    
    print(f"\n{'='*60}")
    print("✅ Endurance 测试全部完成")
    print(f"📁 数据保存路径: {SAVE_DIR}")
    
finally:
    try:
        k.disconnect()
    except Exception as e:
        print(f"⚠️ 断开连接时出错: {e}")
