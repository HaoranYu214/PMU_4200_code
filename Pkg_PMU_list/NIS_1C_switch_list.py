# -*- coding: utf-8 -*-
"""
NLS Switch 批量扫描测试
对不同参数组合进行连续测试
"""
from src.instrcomms import Communications
from pathlib import Path
import time
import pandas as pd

# 导入封装好的测试函数
from NLS_1C_switch import run_nls_switch_test


# ==================== 配置 ====================
INST = "TCPIP0::129.125.87.80::1225::SOCKET"
CH1, CH2 = 1, 2

# 基础参数（不变的部分）
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

# 保存目录
SAVE_DIR = Path(r"C:\Users\P317151\Documents\data\Jingtian\2025-12-14\BTO\Device3\NLS_Sweep")

# ==================== 扫描参数 ====================
# Dwell: 从 1e-7s 到 1s，每10倍取10个点
import numpy as np
Dwell_list = np.logspace(-7, -1, 31)  # 5个数量级 × 10点/数量级 + 1 = 51个点

# Vsquare: 从 0V 到 2V，步进 0.1V
Vsquare_list = np.arange(0, 2.0, 0.1)  # [0, 0.1, 0.2, ..., 2.0]

# 方式2：自定义参数组合列表 (如果需要更灵活的扫描)
# PARAM_COMBOS = 
#     {'Vp': 2, 'Rt_p': 5e-5, 'Vsquare': 1},
#     {'Vp': 3, 'Rt_p': 1e-4, 'Vsquare': 1.5},
#     ...
# ]


# ==================== 执行扫描 ====================
def run_sweep():
    """执行参数扫描: Dwell x Vsquare
    
    按 Ctrl+C 可以安全中断，已完成的测试结果会保留。
    """
    k = Communications(INST)
    k.connect()
    k._instrument_object.write_termination = "\0"
    k._instrument_object.read_termination = "\0"
    Q = k.query
    
    results = []
    summary_data = []  # 汇总数据: Vsquare, Dwell, Pr
    total_tests = len(Dwell_list) * len(Vsquare_list)
    test_idx = 0
    interrupted = False
    
    print(f"🚀 开始 NLS Switch 扫描测试，共 {total_tests} 组参数")
    print(f"   Dwell: {len(Dwell_list)} 个点 ({Dwell_list[0]:.1e}s ~ {Dwell_list[-1]:.1e}s)")
    print(f"   Vsquare: {len(Vsquare_list)} 个点 ({Vsquare_list[0]:.2f}V ~ {Vsquare_list[-1]:.2f}V)")
    print(f"   💡 按 Ctrl+C 可安全中断")
    print("=" * 60)
    
    try:
        for Dwell in Dwell_list:
            for Vsquare in Vsquare_list:
                test_idx += 1
                print(f"\n[{test_idx}/{total_tests}] Dwell={Dwell:.1e}s, Vsquare={Vsquare:.2f}V")
                
                # 构建本次测试参数
                params = BASE_PARAMS.copy()
                params['Vsquare'] = Vsquare
                params['Dwell'] = Dwell
                
                try:
                    result = run_nls_switch_test(Q, CH1, CH2, params, SAVE_DIR)
                    result['params'] = params.copy()
                    results.append(result)
                    
                    # 提取 Pr 值 (Polarization 第一个值 - 最后一个值)
                    if 'df_vp_ch1' in result and result['df_vp_ch1'] is not None:
                        p = result['df_vp_ch1']['Polarization'].values
                        Pr = - p[0] + p[-1] if len(p) > 0 else None
                    else:
                        Pr = None
                    
                    summary_data.append({
                        'Vsquare': Vsquare,
                        'Dwell': Dwell,
                        'Pr': Pr
                    })
                    
                except KeyboardInterrupt:
                    raise  # 重新抛出以便外层捕获
                except Exception as e:
                    print(f"  ❌ 测试失败: {e}")
                    results.append({'params': params.copy(), 'success': False, 'error': str(e)})
                    summary_data.append({
                        'Vsquare': Vsquare,
                        'Dwell': Dwell,
                        'Pr': None
                    })
                
                # 测试间隔
                time.sleep(0.5)
    
    except KeyboardInterrupt:
        interrupted = True
        print(f"\n\n⚠️ 用户中断！已完成 {test_idx}/{total_tests} 组测试")
    
    finally:
        print("\n" + "=" * 60)
        success_count = sum(1 for r in results if r.get('success', False))
        status = "中断" if interrupted else "完成"
        print(f"✅ 扫描{status}: {success_count}/{len(results)} 成功")
        
        # 保存汇总数据
        if summary_data:
            df_summary = pd.DataFrame(summary_data)
            summary_path = SAVE_DIR / "total.xlsx"
            df_summary.to_excel(summary_path, index=False)
            print(f"📊 汇总数据已保存: {summary_path}")
        
        try:
            k.disconnect()
            print("🔌 仪器已断开连接")
        except Exception as e:
            print(f"⚠️ 断开连接时出错: {e}")
    
    return results


if __name__ == "__main__":
    run_sweep()