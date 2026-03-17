# -*- coding: utf-8 -*-
"""
Endurance 疲劳测试（精简版）
循环执行：n × cycle + PV2 read + PUND read
每次循环独立保存数据，避免数据点超限
"""
from src.data_processing import save_channels_separate_excel
from src.pmu_tests import run_cycle_and_read, run_pv2_and_read, run_pund_and_read
from src.instrcomms import Communications
from pathlib import Path
import matplotlib.pyplot as plt

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
            run_cycle_and_read(Q, CH1, CH2, params_cycle, n_cycles=int(n_cycles))
            print(f"   ✅ {int(n_cycles)} 次循环完成")
            
            # 2. PV2 读取
            print(f"📊 执行 PV2 读取...")
            pv2_data = run_pv2_and_read(Q, CH1, CH2, params_pv2)
            
            # 保存 PV2 数据
            fname_pv2 = SAVE_DIR / f"Endurance_PV2_after_{int(n_cycles)}cycles"
            save_channels_separate_excel(
                {1: pv2_data['df_ch1'], 2: pv2_data['df_ch2']},
                f"{fname_pv2}_raw.xlsx"
            )
            pv2_data['df_pv2'].to_excel(f"{fname_pv2}_analysis.xlsx", index=False)
            
            # 简单 PV2 绘图
            fig, ax = plt.subplots(figsize=(6, 5))
            ax.plot(pv2_data['df_pv2']['Voltage'], 
                   pv2_data['df_pv2']['Polarization'], 'b-')
            ax.set_xlabel("Voltage (V)")
            ax.set_ylabel("Polarization (μC/cm²)")
            ax.set_title(f"PV2 after {int(n_cycles)} cycles")
            ax.grid(alpha=0.3)
            fig.tight_layout()
            fig.savefig(f"{fname_pv2}_loop.png", dpi=300)
            plt.close(fig)
            
            # 3. PUND 读取
            print(f"📊 执行 PUND 读取...")
            pund_result = run_pund_and_read(Q, CH1, CH2, params_pund)
            df_ch1, df_ch2 = pund_result['df_ch1'], pund_result['df_ch2']
            
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