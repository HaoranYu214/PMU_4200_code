# -*- coding: utf-8 -*-
"""
NLS Switch 测试 - 封装版
"""
from src.data_processing import save_channels_separate_excel, calculate_polarization
from src.pmu_tests import hy_NISswitch_segARB, power_off_outputs
from src.data_processing import read_both_channels
from src.instrcomms import Communications
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from pathlib import Path


def run_nls_switch_test(Q, CH1, CH2, params, save_dir, fname_prefix=None):
    """
    执行单次 NLS Switch 测试
    
    Args:
        Q: 查询函数
        CH1, CH2: 通道号
        params: 测试参数字典
        save_dir: 保存目录 (Path 或 str)
        fname_prefix: 文件名前缀 (可选，默认自动生成)
    
    Returns:
        dict: {'df_ch1', 'df_ch2', 'df_vp_ch1', 'df_vp_ch2', 'success'}
    """
    save_dir = Path(save_dir)
    save_dir.mkdir(parents=True, exist_ok=True)
    
    if fname_prefix is None:
        fname_prefix = f"NLSswitch_Meas{params['Vp']}V_{int(params['Rt_p']*1e6)}us_with{params['Vsquare']:.2f}V_Dwell{params['Dwell']:.1e}s"
    
    fname_base = save_dir / fname_prefix
    
    print(f"▶️ 运行 NLS Switch (Vp={params['Vp']}V, Vsquare={params['Vsquare']:.2f}V, Dwell={params['Dwell']:.1e}s)...")
    hy_NISswitch_segARB(Q, CH1, CH2, params)
    df_ch1, df_ch2 = read_both_channels(Q, CH1, CH2)
    power_off_outputs(Q, (CH1, CH2))
    if df_ch1 is None or df_ch2 is None or df_ch1.empty or df_ch2.empty:
        raise ValueError("NIS Switch读取数据为空")
    
    result = {'df_ch1': df_ch1, 'df_ch2': df_ch2, 'success': True}
    
    if params['MeasureSquare']:
        # 方波测量模式：只保存原始数据
        excel_path = f"{fname_base}.xlsx"
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            df_ch1.to_excel(writer, sheet_name='Raw_CH1', index=False)
            df_ch2.to_excel(writer, sheet_name='Raw_CH2', index=False)
        print(f"✅ 数据已保存: {excel_path}")
        
        # 绘制波形
        fig, axes = plt.subplots(2, 1, figsize=(12, 8), sharex=True)
        
        t1 = df_ch1[f'Timestamp {CH1}']
        v1 = df_ch1[f'Voltage {CH1}']
        i1 = df_ch1[f'Current {CH1}']
        
        ax1_v = axes[0]
        ax1_i = ax1_v.twinx()
        ax1_v.plot(t1, v1, 'b-', linewidth=0.8)
        ax1_i.plot(t1, i1*1e6, 'r-', linewidth=0.8)
        ax1_v.set_ylabel("Voltage (V)", color='b')
        ax1_i.set_ylabel("Current (μA)", color='r')
        ax1_v.set_title("CH1 Waveform")
        ax1_v.grid(alpha=0.3)
        
        t2 = df_ch2[f'Timestamp {CH2}']
        v2 = df_ch2[f'Voltage {CH2}']
        i2 = df_ch2[f'Current {CH2}']
        
        ax2_v = axes[1]
        ax2_i = ax2_v.twinx()
        ax2_v.plot(t2, v2, 'b-', linewidth=0.8)
        ax2_i.plot(t2, i2*1e6, 'r-', linewidth=0.8)
        ax2_v.set_xlabel("Time (s)")
        ax2_v.set_ylabel("Voltage (V)", color='b')
        ax2_i.set_ylabel("Current (μA)", color='r')
        ax2_v.set_title("CH2 Waveform")
        ax2_v.grid(alpha=0.3)
        
        fig.suptitle("NIS Switch - Waveform Check", fontsize=14)
        fig.tight_layout()
        fig.savefig(f"{fname_base}_waveform.png", dpi=300)
        plt.close(fig)
        print(f"✅ 方波测量模式完成 (CH1: {len(df_ch1)} pts, CH2: {len(df_ch2)} pts)")
        
    else:
        # 非方波测量模式：差分极化分析
        area_cm2 = params.get('area_cm2', 1.0)
        
        def process_channel(df, ch_num):
            """处理单通道：4段数据，前2段 vs 后2段做差分"""
            v = df[f'Voltage {ch_num}'].values
            i = df[f'Current {ch_num}'].values
            
            n = len(v)
            seg_pts = n // 4
            
            v_tri1 = v[:2*seg_pts]
            i_tri1 = i[:2*seg_pts]
            i_tri2 = i[2*seg_pts:4*seg_pts]
            
            min_len = min(len(i_tri1), len(i_tri2))
            i_diff = i_tri1[:min_len] - i_tri2[:min_len]
            v_diff = v_tri1[:min_len]
            dt = params['Rt_p'] * 4 / n
            t_diff = np.arange(1, min_len + 1) * dt
            
            P = calculate_polarization(i_diff, t_diff, area_cm2)
            
            return pd.DataFrame({
                'Time': t_diff, 'Voltage': v_diff,
                'DiffCurrent': i_diff, 'Polarization': P
            })
        
        df_vp_ch1 = process_channel(df_ch1, CH1)
        df_vp_ch2 = process_channel(df_ch2, CH2)
        
        result['df_vp_ch1'] = df_vp_ch1
        result['df_vp_ch2'] = df_vp_ch2
        
        # 保存到同一个 Excel 的不同 sheet
        excel_path = f"{fname_base}.xlsx"
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            df_ch1.to_excel(writer, sheet_name='Raw_CH1', index=False)
            df_ch2.to_excel(writer, sheet_name='Raw_CH2', index=False)
            df_vp_ch1.to_excel(writer, sheet_name='VP_CH1', index=False)
            df_vp_ch2.to_excel(writer, sheet_name='VP_CH2', index=False)
        print(f"✅ 数据已保存: {excel_path}")
        
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
        plt.close(fig)
        print(f"✅ 完成 (CH1: {len(df_ch1)} pts → {len(df_vp_ch1)} diff pts)")
    
    return result


# ==================== 单次测试入口 ====================
if __name__ == "__main__":
    INST = "TCPIP0::129.125.87.80::1225::SOCKET"
    CH1, CH2 = 1, 2
    
    params = dict(
        offset=0,
        Vp=2,
        Rt_p=5e-5,
        Delaytime=100e-6,
        Vsquare=1,
        Rt_s=1e-7,
        Dwell=1e-6,
        Irange1=1e-4,
        Irange2=1e-4,
        MeasureSquare=False,
        area_cm2=7.0686e-6,
    )
    
    SAVE_DIR = Path(r"C:\Users\P317151\Documents\data\Jingtian\2025-12-14\BTO\Device3")
    
    k = Communications(INST)
    k.connect()
    k._instrument_object.write_termination = "\0"
    k._instrument_object.read_termination = "\0"
    Q = k.query
    
    try:
        run_nls_switch_test(Q, CH1, CH2, params, SAVE_DIR)
    except Exception as e:
        print(f"❌ 运行出错: {e}")
        raise
    finally:
        try:
            k.disconnect()
        except Exception as e:
            print(f"⚠️ 断开连接时出错: {e}")
