# -*- coding: utf-8 -*-
"""
绘图工具模块 - 专门处理 Keithley 4200A 测试数据的可视化
包含：PlotManager, 各种绘图函数
"""

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick
from datetime import datetime
from pathlib import Path


class PlotManager:
    """
    统一的绘图管理器
    mode:
      - 'batched'  ：先都画完，退出 with 时一次性 show（阻塞与否由 block 决定）
      - 'headless' ：不展示（适合服务器/批处理）
      - 'save'     ：不展示，直接保存到指定目录（自动带时间戳）
    """
    def __init__(self, mode='batched', block=True, close_after_show=False,
                 save_dir="outputs", save_ext="png", save_dpi=150,
                 run_tag=None, filename_prefix="fig"):
        self.mode = mode
        self.block = block
        self.close_after_show = close_after_show
        self._figs = []
        self._interactive_prev = plt.isinteractive()

        # 保存相关
        self.save_dir = Path(save_dir)
        self.save_ext = save_ext
        self.save_dpi = save_dpi
        self.run_tag = run_tag or datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        self.filename_prefix = filename_prefix
        self._save_counter = 0

    def __enter__(self):
        if self.mode == 'batched':
            plt.ion()     # 仅批量展示时开交互，避免中途阻塞
        else:  # 'save' / 'headless'
            plt.ioff()    # 关闭交互，不会弹窗
        return self

    def add(self, fig, name=None):
        if fig is None:
            return
        if self.mode == 'save':
            self._save_counter += 1
            stem = name or f"{self.filename_prefix}_{self._save_counter}"
            (self.save_dir / self.run_tag).mkdir(parents=True, exist_ok=True)
            path = self.save_dir / self.run_tag / f"{stem}_{self.run_tag}.{self.save_ext}"
            fig.savefig(path, dpi=self.save_dpi, bbox_inches="tight", pad_inches=0.05)
            print(f"💾 已保存图像：{path}")
            plt.close(fig)
        else:
            self._figs.append(fig)

    def __exit__(self, exc_type, exc, tb):
        try:
            if self.mode == 'batched':
                if self.block:
                    plt.ioff()                 # 关交互 → 阻塞 show
                    plt.show(block=True)
                else:
                    plt.ion()                  # 开交互 → 非阻塞 show
                    plt.show(block=False)
                if self.close_after_show:
                    for f in self._figs:
                        try: plt.close(f)
                        except: pass
            elif self.mode == 'headless':
                for f in self._figs:
                    try: plt.close(f)
                    except: pass
            # save 模式在 add 时已保存+关闭
        finally:
            if self._interactive_prev:
                plt.ion()
            else:
                plt.ioff()


def apply_symlog_with_ticks(ax, data, linthresh=1e-12, max_decades=8, unit=""):
    """
    对 ax 应用 symlog，并为 data 手动生成对称对数刻度（含 0 与线性阈值）。
    unit: 刻度标签后缀 (例如 "A", "Ω")
    """
    data = np.asarray(data)
    data = data[np.isfinite(data)]
    if data.size == 0:
        ax.set_yscale("linear")
        return

    dmax = np.nanmax(np.abs(data))
    if not np.isfinite(dmax) or dmax == 0:
        ax.set_yscale("linear")
        return

    ax.set_yscale("symlog", linthresh=linthresh, linscale=1.0)

    hi_exp = int(np.ceil(np.log10(dmax)))
    lo_exp = int(np.floor(np.log10(linthresh)))
    if hi_exp - lo_exp > max_decades:
        lo_exp = hi_exp - max_decades

    decades = 10.0 ** np.arange(lo_exp, hi_exp + 1)
    neg_ticks = (-decades[decades > linthresh])[::-1].tolist()
    core_ticks = [-linthresh, 0.0, linthresh]
    pos_ticks = decades[decades > linthresh].tolist()
    ticks = neg_ticks + core_ticks + pos_ticks

    ax.set_yticks(ticks)
    ax.yaxis.set_major_formatter(
        mtick.FuncFormatter(lambda v, p: "0" if v == 0 else f"{v:.0e}{unit}")
    )
    ax.set_ylim(-1.2 * dmax, 1.2 * dmax)


def plot_time_series(df, channels=(1, 2), width_us=None, amp_v=None, 
                    resistance_scale='log', show=False, return_fig=True):
    """
    统一的时序绘图函数 - 自动检测数据类型并绘制电压/电流/电阻
    
    Args:
        df: 数据DataFrame
        channels: 通道列表，默认 (1, 2)
        width_us: 脉宽(μs)，用于标题显示
        amp_v: 幅度(V)，用于标题显示
        resistance_scale: 电阻坐标模式，'log'=对数坐标，'linear'=线性坐标
        show: 是否立即显示
        return_fig: 是否返回Figure对象
    """
    if df is None or df.empty:
        print("⚠️ 无数据，跳过绘图")
        return None

    # 检测可用的数据类型
    data_types = []
    if any(f"Voltage {ch}" in df.columns for ch in channels):
        data_types.append('voltage')
    if any(f"Current {ch}" in df.columns for ch in channels):
        data_types.append('current')
    if any(f"Resistance {ch}" in df.columns for ch in channels):
        data_types.append('resistance')
    
    if not data_types:
        print("⚠️ 未找到可绘制的数据列")
        return None

    # 创建子图
    n_plots = len(data_types)
    fig, axes = plt.subplots(n_plots, 1, figsize=(12, 4 * n_plots))
    if n_plots == 1:
        axes = [axes]

    # 标题
    suffix = []
    if width_us is not None: suffix.append(f"{width_us} µs")
    if amp_v is not None:    suffix.append(f"{amp_v} V")
    title = "Dual-channel Measurements" + (" - " + ", ".join(suffix) if suffix else "")
    fig.suptitle(title, fontsize=14, fontweight='bold')

    # 颜色和标记
    colors = {1: 'darkred', 2: 'darkblue'}
    markers = {'voltage': 'o', 'current': '^', 'resistance': 'd'}
    units = {'voltage': 'V', 'current': 'A', 'resistance': 'Ω'}

    for i, data_type in enumerate(data_types):
        ax = axes[i]
        
        for ch in channels:
            col_data = f"{data_type.title()} {ch}"
            col_time = f"Timestamp {ch}"
            
            if col_data in df.columns and col_time in df.columns:
                data = df[col_data]
                time = df[col_time]
                
                # 过滤有效数据
                if data_type == 'resistance':
                    valid = np.isfinite(data)
                    if valid.any():
                        data = data[valid]
                        time = time[valid]
                    else:
                        continue
                
                # 绘制
                ax.plot(time, data, 
                       color=colors[ch], linewidth=1.5, 
                       marker=markers[data_type], markersize=3,
                       markerfacecolor=colors[ch], markeredgecolor=colors[ch],
                       label=f'Channel {ch} {data_type.title()}', alpha=0.8)
        
        # 设置轴标签和格式
        unit = units[data_type]
        ax.set_ylabel(f"{data_type.title()} ({unit})")
        ax.set_title(f"{data_type.title()} vs Time")
        
        # 电阻坐标处理（可选对数/线性）
        if data_type == 'resistance' and resistance_scale == 'log':
            all_res_data = []
            for ch in channels:
                col = f"Resistance {ch}"
                if col in df.columns:
                    valid_data = df[col].dropna()
                    if not valid_data.empty:
                        all_res_data.extend(valid_data.tolist())
            
            if all_res_data:
                apply_symlog_with_ticks(ax, all_res_data, linthresh=1.0, unit="Ω")
        
        ax.grid(True, alpha=0.3)
        ax.legend()
        
        # 最后一个子图添加时间轴标签
        if i == len(data_types) - 1:
            ax.set_xlabel("Time (s)")

    plt.tight_layout()
    if show: plt.show()
    return fig if return_fig else None


def plot_iv_characteristics(df, channels=(1, 2), width_us=None, amp_v=None,
                           current_linthresh=1e-12, show=False, return_fig=True):
    """
    I-V特性曲线绘图
    """
    if df is None or df.empty:
        print("⚠️ 无数据，跳过I-V图")
        return None

    fig, ax = plt.subplots(1, 1, figsize=(10, 8))
    
    # 标题
    suffix = []
    if width_us is not None: suffix.append(f"{width_us} µs")
    if amp_v is not None:    suffix.append(f"{amp_v} V")
    title = "I-V Characteristics" + (" - " + ", ".join(suffix) if suffix else "")
    fig.suptitle(title, fontsize=14, fontweight='bold')

    colors = {1: 'red', 2: 'blue'}
    
    for ch in channels:
        v_col = f"Voltage {ch}"
        i_col = f"Current {ch}"
        
        if v_col in df.columns and i_col in df.columns:
            voltage = df[v_col].dropna()
            current = df[i_col].dropna()
            
            if not voltage.empty and not current.empty:
                ax.plot(voltage, current, 
                       color=colors[ch], linewidth=2, marker='o', markersize=4,
                       label=f'Channel {ch}', alpha=0.7)

    ax.set_xlabel("Voltage (V)")
    ax.set_ylabel("Current (A)")
    
    # 电流轴使用对数坐标
    all_currents = []
    for ch in channels:
        i_col = f"Current {ch}"
        if i_col in df.columns:
            all_currents.extend(df[i_col].dropna().tolist())
    
    if all_currents:
        apply_symlog_with_ticks(ax, all_currents, linthresh=current_linthresh, unit="A")
    
    ax.grid(True, alpha=0.3)
    ax.legend()
    
    plt.tight_layout()
    if show: plt.show()
    return fig if return_fig else None


# 为了向后兼容，保留原来的函数名
def plot_currents(df, width_us=None, amp_v=None, show=False, return_fig=True):
    """向后兼容的电流绘图函数"""
    return plot_time_series(df, channels=(1, 2), width_us=width_us, amp_v=amp_v, 
                           show=show, return_fig=return_fig)

def plot_voltages(df, width_us=None, amp_v=None, show=False, return_fig=True):
    """向后兼容的电压绘图函数"""
    return plot_time_series(df, channels=(1, 2), width_us=width_us, amp_v=amp_v, 
                           show=show, return_fig=return_fig)

def plot_resistances(df, width_us=None, amp_v=None, res_min=1.0, 
                    resistance_scale='log', show=False, return_fig=True):
    """向后兼容的电阻绘图函数"""
    return plot_time_series(df, channels=(1, 2), width_us=width_us, amp_v=amp_v, 
                           resistance_scale=resistance_scale, show=show, return_fig=return_fig)
