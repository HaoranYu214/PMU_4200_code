# -*- coding: utf-8 -*-
# PMU data-processing helpers.
"""
数据处理模块 - 数据读取、保存、合并、计算
"""

import numpy as np
import pandas as pd
import locale
import csv


def _channel_data_columns(ch, field_count, pulse_iv=None):
    """Return columns for waveform/segARB or pulse-I-V records."""
    standard = [f"Voltage {ch}", f"Current {ch}", f"Timestamp {ch}", f"Status {ch}"]
    if field_count == 8:
        return [
            f"Voltage High {ch}", f"Current High {ch}",
            f"Timestamp High {ch}", f"Status High {ch}",
            f"Voltage Low {ch}", f"Current Low {ch}",
            f"Timestamp Low {ch}", f"Status Low {ch}",
        ]
    if field_count != 4:
        raise ValueError(
            f"CH{ch} returned {field_count} values per point; expected 4 or 8."
        )
    if pulse_iv == (False, True):
        return [
            f"Voltage Low {ch}", f"Current Low {ch}",
            f"Timestamp Low {ch}", f"Status Low {ch}",
        ]
    return standard


def read_channel_data(Q, ch, block=2048, debug=False, pulse_iv=None):
    """按块读取单通道，返回 [Voltage ch, Current ch, Timestamp ch, Status ch]；无数据返回 None。
    
    debug=True 时打印原始响应的前几个数据点，用于检查精度问题。
    """
    count = int(Q(f":PMU:DATA:COUNT? {ch}").strip())
    if count == 0:
        return None
    if debug:
        print(f"🔍 [DEBUG] CH{ch} 数据点数: {count}")
    
    if pulse_iv is not None:
        if len(pulse_iv) != 2:
            raise ValueError("pulse_iv must be (acquire_high, acquire_low).")
        pulse_iv = tuple(bool(value) for value in pulse_iv)
        if pulse_iv == (False, False):
            raise ValueError("Pulse I-V must acquire at least High or Low data.")

    cols = None
    chunks = []
    for start in range(0, count, block):
        resp = Q(f":PMU:DATA:GET {ch}, {start}, {block}")
        if not resp:
            continue
        
        # Debug: 打印第一块的原始响应
        if debug and start == 0:
            print(f"🔍 [DEBUG] CH{ch} 原始响应长度: {len(resp)}")
            print(f"🔍 [DEBUG] CH{ch} 前500字符:\n{resp[:500]}")
            # 解析第一个数据点看看
            first_row = resp.split(";")[0] if ";" in resp else resp
            print(f"🔍 [DEBUG] CH{ch} 第一个数据点原始: '{first_row}'")
        
        rows = [seg.split(",") for seg in resp.split(";") if seg.strip()]
        if not rows:
            continue
        field_counts = {len(row) for row in rows}
        if len(field_counts) != 1:
            raise ValueError(
                f"CH{ch} returned inconsistent record widths: {sorted(field_counts)}."
            )
        chunk_cols = _channel_data_columns(ch, field_counts.pop(), pulse_iv=pulse_iv)
        if cols is None:
            cols = chunk_cols
        elif cols != chunk_cols:
            raise ValueError(f"CH{ch} data layout changed while reading blocks.")
        chunks.append(pd.DataFrame(rows, columns=cols))
    
    if not chunks:
        return None
    
    result = pd.concat(chunks, ignore_index=True)
    numeric_columns = [
        column for column in result.columns
        if column.startswith(("Voltage ", "Current ", "Timestamp "))
    ]
    result[numeric_columns] = result[numeric_columns].astype(float)
    result = result.reset_index(drop=True)
    if pulse_iv is not None:
        result.attrs["pulse_iv_acquire_high"] = pulse_iv[0]
        result.attrs["pulse_iv_acquire_low"] = pulse_iv[1]
    
    if debug:
        print(f"🔍 [DEBUG] CH{ch} 解析后前3行:\n{result.head(3).to_string()}")
    
    return result


def read_both_channels(Q, ch1, ch2, debug=False, pulse_iv=None):
    """读取双通道数据（若某通道无数据返回 None）"""
    df1 = read_channel_data(Q, ch1, debug=debug, pulse_iv=pulse_iv)
    df2 = read_channel_data(Q, ch2, debug=debug, pulse_iv=pulse_iv)
    return df1, df2


def select_pulse_iv_level(df, ch, level="High"):
    """Return one pulse-I-V level using the standard four-column schema."""
    if df is None or df.empty:
        return df
    level = level.title()
    if level not in ("High", "Low"):
        raise ValueError("level must be 'High' or 'Low'.")

    source_columns = [
        f"Voltage {level} {ch}",
        f"Current {level} {ch}",
        f"Timestamp {level} {ch}",
        f"Status {level} {ch}",
    ]
    if all(column in df.columns for column in source_columns):
        result = df[source_columns].copy()
        result.columns = [
            f"Voltage {ch}",
            f"Current {ch}",
            f"Timestamp {ch}",
            f"Status {ch}",
        ]
        result.attrs.update(df.attrs)
        result.attrs["pulse_iv_selected_level"] = level
        return result

    standard_columns = [
        f"Voltage {ch}",
        f"Current {ch}",
        f"Timestamp {ch}",
        f"Status {ch}",
    ]
    if all(column in df.columns for column in standard_columns):
        return df.copy()
    raise ValueError(f"CH{ch} does not contain {level} pulse-I-V data.")


def merge_channels(dfs: dict):
    """dfs={1:df1或None, 2:df2或None} → 横向拼列；长度不同自动补NaN。"""
    d1, d2 = dfs.get(1), dfs.get(2)
    if d1 is None and d2 is None:
        return pd.DataFrame()
    if d1 is None:
        return d2.copy()
    if d2 is None:
        return d1.copy()
    return pd.concat([d1.reset_index(drop=True), d2.reset_index(drop=True)], axis=1)


def add_resistance_columns(df, eps=1e-12, res_min=1.0, res_max=1e15):
    """在 df 内联添加 'Resistance 1/2'（如对应V/I存在）。"""
    if df is None or df.empty:
        return df
    for ch in (1, 2):
        vcol, icol, rcol = f"Voltage {ch}", f"Current {ch}", f"Resistance {ch}"
        if vcol in df.columns and icol in df.columns:
            i = df[icol].astype(float)
            valid_i = np.abs(i) >= eps
            r = pd.Series(np.where(valid_i, df[vcol] / i, np.nan), index=df.index)
            keep = (np.abs(r) >= res_min) & (np.abs(r) <= res_max)
            df[rcol] = r.where(keep, np.nan)
    return df


def calculate_polarization(current_data, time_data, area_cm2):
    """计算极化值，返回单位 μC/cm²"""
    charge = np.zeros_like(current_data, dtype=float)
    for i in range(1, len(current_data)):
        dt = time_data[i] - time_data[i-1]
        charge[i] = charge[i-1] + dt * current_data[i]
    midpoint = len(charge) // 2
    center_offset = (charge[0] + charge[midpoint]) / 2.0 if len(charge) > 1 else 0.0
    polarization = (charge - center_offset) / area_cm2 * 1e6
    return polarization


def analyze_pund_diff(df_ch1, df_ch2, params):
    """
    PUND差分分析。
    返回: dict(df_total, pund_diff, meta)
    """
    if df_ch1 is None or df_ch2 is None or df_ch1.empty or df_ch2.empty:
        raise ValueError("PUND原始数据为空")

    def _detect_ch(df, kind='Voltage'):
        for c in df.columns:
            if c.startswith(f'{kind} '):
                try:
                    return int(c.split(' ')[1])
                except:
                    continue
        raise ValueError(f"无法在列中识别通道号: {list(df.columns)}")

    ch1 = _detect_ch(df_ch1, 'Voltage')
    ch2 = _detect_ch(df_ch2, 'Voltage')

    v_total = df_ch1[f'Voltage {ch1}'].values - df_ch2[f'Voltage {ch2}'].values
    i_total = -df_ch2[f'Current {ch2}'].values
    t_total = df_ch1[f'Timestamp {ch1}'].values
    df_total = pd.DataFrame({'Time': t_total, 'Voltage': v_total, 'Current': i_total})
    
    total_points = len(i_total)
    seg_points = total_points // 22
    if seg_points == 0:
        raise ValueError("PUND数据点数不足，无法按22段进行差分")

    pund_pairs = [(6, 10, 'P'), (8, 12, 'U'), (14, 18, 'N'), (16, 20, 'D')]
    v_collect, i_collect, labels = [], [], []
    
    for seg_a, seg_b, label in pund_pairs:
        a0, a1 = seg_a * seg_points, (seg_a + 1) * seg_points
        b0, b1 = seg_b * seg_points, (seg_b + 1) * seg_points
        v_seg = v_total[a0:a1]
        i_diff = i_total[a0:a1] - i_total[b0:b1]
        if len(v_seg) == 0 or len(i_diff) == 0:
            continue
        v_collect.append(v_seg)
        i_collect.append(i_diff)
        labels.extend([label] * len(v_seg))
    
    if not v_collect:
        raise ValueError("PUND差分阶段提取失败")
    
    v_all = np.concatenate(v_collect)
    i_all = np.concatenate(i_collect)
    dt_est = (t_total[-1] - t_total[0]) / max(len(i_all)-1, 1)
    time_local = np.arange(len(i_all)) * dt_est
    P = calculate_polarization(i_all, time_local, params.get('area_cm2', 1.0))
    
    diff_df = pd.DataFrame({
        'Time': time_local, 'Voltage': v_all, 'DiffCurrent': i_all,
        'Polarization': P, 'Segment': labels
    })
    return {'df_total': df_total, 'pund_diff': diff_df, 
            'meta': {'points_per_segment': seg_points, 'pairs': pund_pairs}}


def analyze_nis_switch(df_ch1, df_ch2, params):
    """
    NIS Switch 数据分析 - 简化版
    
    直接返回原始数据，由调用方（1C_NISswitch.py）根据 MeasureSquare 决定如何处理。
    
    返回: dict(df_ch1, df_ch2, meta)
    """
    if df_ch1 is None or df_ch2 is None or df_ch1.empty or df_ch2.empty:
        raise ValueError("NIS Switch 原始数据为空")

    measure_square = params.get('MeasureSquare', True)
    
    return {
        'df_ch1': df_ch1,
        'df_ch2': df_ch2,
        'meta': {
            'measure_square': measure_square,
            'points_ch1': len(df_ch1),
            'points_ch2': len(df_ch2)
        }
    }


# === 保存功能 ===

def save_channels_separate_excel(dfs: dict, path):
    """将双通道数据保存到Excel的不同sheet中"""
    if not dfs or all(df is None or df.empty for df in dfs.values()):
        print(f"⚠️ 无数据，跳过保存：{path}")
        return False
    
    try:
        if not path.lower().endswith(".xlsx"):
            path = path + ".xlsx"
        
        with pd.ExcelWriter(path, engine='openpyxl') as writer:
            saved_sheets = 0
            for ch, df in dfs.items():
                if df is not None and not df.empty:
                    sheet_name = f"Channel_{ch}"
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    print(f"✅ 已保存 CH{ch} 到 {sheet_name}: {df.shape}")
                    saved_sheets += 1
            
            if saved_sheets > 0:
                print(f"✅ Excel文件已保存：{path} (共{saved_sheets}个sheet)")
                return True
            else:
                print(f"⚠️ 没有有效数据保存到：{path}")
                return False
    except Exception as e:
        print(f"❌ 保存Excel失败: {e}")
        return False


def save_csv(df, path, sep=None, for_excel=True):
    """保存 DataFrame 到 CSV；在小数点为','的地区默认用';'分隔，适配 Excel。"""
    if df is None or df.empty:
        print(f"⚠️ 无数据，跳过保存：{path}")
        return False

    if sep is None:
        try:
            decimal_point = locale.localeconv().get('decimal_point', '.')
        except Exception:
            decimal_point = '.'
        sep = ';' if decimal_point == ',' else ','

    encoding = 'utf-8-sig' if for_excel else 'utf-8'
    print(f"📊 保存数据: 形状 {df.shape}, 列数 {len(df.columns)}, 分隔符='{sep}'")

    try:
        df.to_csv(path, index=False, sep=sep, encoding=encoding,
                  float_format='%.6e', na_rep='NaN', quoting=csv.QUOTE_MINIMAL)
        print(f"✅ 已保存：{path}")
        
        test_df = pd.read_csv(path, nrows=1, sep=sep)
        if len(test_df.columns) == len(df.columns):
            print(f"✅ 验证成功: {len(df.columns)} 列正确保存")
        else:
            print(f"⚠️ 列数不匹配: 原始 {len(df.columns)}, 读取 {len(test_df.columns)}")
        return True
    except Exception as e:
        print(f"❌ 保存失败: {e}")
        return False


def save_excel(df, path):
    """保存单个 DataFrame 到 Excel"""
    if df is None or df.empty:
        print(f"⚠️ 无数据，跳过保存：{path}")
        return False
    try:
        if not path.lower().endswith(".xlsx"):
            path = path + ".xlsx"
        df.to_excel(path, index=False)
        print(f"✅ 已保存 Excel：{path}")
        return True
    except Exception as e:
        print(f"❌ 保存失败: {e}")
        return False


def print_data_summary(df):
    """打印数据摘要信息"""
    if df is not None and not df.empty:
        print("\n📊 数据摘要：")
        print("  形状   :", df.shape)
        print("  列名   :", list(df.columns))
        for ch in (1, 2):
            for col in (f"Voltage {ch}", f"Current {ch}", f"Resistance {ch}"):
                if col in df.columns:
                    s = df[col].dropna()
                    if not s.empty:
                        print(f"  {col}: min={s.min():.3e}, max={s.max():.3e}")
    else:
        print("⚠️ 无有效数据")


# === 模拟数据生成 ===

def generate_mock_dual(n_points=2000, duration=1e-3):
    """生成模拟双通道数据"""
    def triangle_wave(t, cycles=4, duration=1e-3, amp=5.0):
        phase = (cycles * t / duration) % 1.0
        tri01 = 2.0*np.abs(phase - 0.5)
        return (1 - 2*tri01) * amp

    t = np.linspace(0, duration, n_points)
    cycles = 4
    v1 = triangle_wave(t, cycles=cycles, duration=duration, amp=5.0)
    v2 = triangle_wave((t + duration/(4*cycles)) % duration, cycles=cycles, duration=duration, amp=2.5)

    shift = int(n_points/(4*cycles))
    t_shift = np.roll(t, shift)
    exp_scale = (np.exp(np.linspace(0, 12, n_points)) - 1) / (np.e**12 - 1)
    env = 1e-12 + exp_scale * (1e-3 - 1e-12)
    i1 = env * np.sin(2*np.pi*cycles * t_shift / duration)
    i2 = env[::-1] * np.cos(2*np.pi*cycles * t_shift / duration)

    s1 = np.zeros(n_points, dtype=int)
    s2 = np.zeros(n_points, dtype=int)
    return pd.DataFrame({
        "Voltage 1": v1, "Current 1": i1, "Timestamp 1": t, "Status 1": s1,
        "Voltage 2": v2, "Current 2": i2, "Timestamp 2": t, "Status 2": s2,
    })
