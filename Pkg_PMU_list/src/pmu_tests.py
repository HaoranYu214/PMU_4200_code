# -*- coding: utf-8 -*-
"""
PMU核心测试模块 - segARB配置、执行、设备控制
"""

import time


# === 设备控制 ===

def power_off_outputs(Q, channels):
    """测试数据读取后立刻关闭输出（幂等）"""
    for ch in channels:
        try:
            Q(f":PMU:OUTPUT:STATE {ch}, 0")
        except Exception:
            pass


# === segARB 核心配置 ===

def configure_segARB_sequence(Q, ch, seq_id, start_voltages, stop_voltages, time_values, 
                              meas_types=None, meas_start=None, meas_stop=None):
    """通用segARB序列配置函数"""
    n_segments = len(start_voltages)
    if meas_types is None:
        meas_types = [2] * n_segments
    if meas_start is None:
        meas_start = [0.0] * n_segments
    if meas_stop is None:
        meas_stop = list(time_values)

    start_v_str = ", ".join(map(str, start_voltages))
    stop_v_str = ", ".join(map(str, stop_voltages))
    time_str = ", ".join([f"{t:.2e}" for t in time_values])
    meas_type_str = ", ".join(map(str, meas_types))
    meas_start_str = ", ".join([f"{t:.2e}" for t in meas_start])
    meas_stop_str = ", ".join([f"{t:.2e}" for t in meas_stop])

    Q(f":PMU:SARB:SEQ:STARTV {ch}, {seq_id}, {start_v_str}")
    Q(f":PMU:SARB:SEQ:STOPV {ch}, {seq_id}, {stop_v_str}")
    Q(f":PMU:SARB:SEQ:TIME {ch}, {seq_id}, {time_str}")
    Q(f":PMU:SARB:SEQ:MEAS:TYPE {ch}, {seq_id}, {meas_type_str}")
    Q(f":PMU:SARB:SEQ:MEAS:START {ch}, {seq_id}, {meas_start_str}")
    Q(f":PMU:SARB:SEQ:MEAS:STOP {ch}, {seq_id}, {meas_stop_str}")


def auto_align_channels(seq_configs):
    """智能通道对齐：当一通道复杂时序，另一通道恒压时，自动扩展对齐"""
    seq_ids = set()
    for _, configs in seq_configs.items():
        for config in configs:
            seq_ids.add(config[0])

    for seq_id in seq_ids:
        ch_configs = {}
        for ch, configs in seq_configs.items():
            for config in configs:
                if config[0] == seq_id:
                    ch_configs[ch] = config
                    break
        if len(ch_configs) <= 1:
            continue

        channels = list(ch_configs.keys())
        ref_ch = max(ch_configs.keys(), key=lambda c: len(ch_configs[c][3]))
        ref_time = ch_configs[ref_ch][3]

        for ch in channels:
            if ch == ref_ch:
                continue
            start_v, stop_v = ch_configs[ch][1], ch_configs[ch][2]
            is_const = all(abs(a-b) < 1e-12 for a, b in zip(start_v, stop_v))
            if not is_const:
                raise ValueError(f"CH{ch} 非恒压序列，且与CH{ref_ch}段数不一致，无法自动对齐")
            v = start_v[0] if len(start_v) else 0.0
            new_config = (ch_configs[ch][0], [v]*len(ref_time), [v]*len(ref_time), ref_time) + ch_configs[ch][4:]
            for i, cfg in enumerate(seq_configs[ch]):
                if cfg[0] == seq_id:
                    seq_configs[ch][i] = new_config
                    break
    return seq_configs


def execute_segARB_test(Q, channels, seq_configs, seq_list=None, current_ranges=None):
    """
    通用segARB测试执行函数
    Args:
        Q: 查询函数
        channels: 通道列表 [CH1, CH2, ...]
        seq_configs: 序列配置字典 {ch: [(seq_id, start_v, stop_v, time_v), ...]}
        seq_list: 序列执行列表 {ch: [(seq_id, loop_count), ...]} (默认执行seq1一次)
        wait_completion: 是否等待测试完成
        current_ranges: 电流测量范围字典 {ch: range_value} (在初始化后设置)
    
    注意: 同一PMU上的所有通道必须使用相同的时间序列，因为它们共享同一个时钟
    """
    Q(":PMU:INIT 1")

    # 配置RPM,链接RPM到PMU通道
    for ch in channels:
        Q(f":PMU:RPM:CONFIGURE PMU1-{ch}, 0")

    # ⚡ 设置测量范围 (必须在 :PMU:INIT 之后设置)
    # 语法: :PMU:MEASURE:RANGE ch, IRangeType, IMeasRange
    #   IRangeType: 0=Autorange, 1=Limited autorange, 2=Fixed range
    #   ⚠️ Segment Arb 模式必须使用 Fixed range (type=2)
    # 
    # 可用电流范围 (取决于PMU/RPM型号):
    #   40V PMU:  0.8A, 0.01A, 0.0001A
    #   10V PMU:  0.2A, 0.01A
    #   40V RPM:  0.8A, 0.01A, 0.0001A
    #   10V RPM:  0.2A, 0.01A, 0.001A, 0.0001A, 0.00001A, 0.000001A, 0.0000001A

    if current_ranges:
        for ch, i_range in current_ranges.items():
            Q(f":PMU:MEASURE:RANGE {ch}, 2, {i_range}")
            print(f"   CH{ch} 电流范围 (固定): {i_range:.2e} A")

    seq_configs = auto_align_channels(seq_configs)

    for ch, configs in seq_configs.items():
        for cfg in configs:
            seq_id = cfg[0]
            start_v, stop_v, time_v = cfg[1], cfg[2], cfg[3]
            meas_types = cfg[4] if len(cfg) > 4 else None
            meas_start = cfg[5] if len(cfg) > 5 else None
            meas_stop = cfg[6] if len(cfg) > 6 else None
            configure_segARB_sequence(Q, ch, seq_id, start_v, stop_v, time_v,
                                      meas_types, meas_start, meas_stop)

    if seq_list is None:
        seq_list = {ch: [(1, 1)] for ch in channels}

    for ch, exec_list in seq_list.items():
        if exec_list:
            seq_str = ", ".join([f"{sid}, {loop}" for sid, loop in exec_list])
            Q(f":PMU:SARB:WFM:SEQ:LIST {ch}, {seq_str}")

    for ch in channels:
        Q(f":PMU:OUTPUT:STATE {ch}, 1")
    Q(":PMU:EXECUTE")

    print("等待 segARB 测试完成...")
    while True:
        try:
            status_str = Q(":PMU:TEST:STATUS?")
            if not status_str or status_str.strip() == "":
                time.sleep(0.1)
                continue
            status_str = status_str.strip().replace("ACK", "").strip()
            if status_str:
                try:
                    if int(status_str) == 0:
                        print("segARB 测试完成")
                        break
                except ValueError:
                    print(f"⚠️ 无法解析状态: '{status_str}'")
        except Exception as e:
            print(f"⚠️ 查询状态出错: {e}")
        time.sleep(0.3)


# === 底层测试函数 ===

def _get_mode_num(mode):
    """解析测量模式"""
    mode_map = {
        'D': 1, 'DISCRETE': 1, 'SPOT': 1,
        'WFM': 2, 'WAVEFORM': 2, 'WAVE': 2,
        'NONE': 0, 'NO': 0,
        'AVG': 3, 'AVERAGE': 3, 'SPOT_AVG': 3,
        'WFM_AVG': 4, 'WAVEFORM_AVG': 4, 'WAVE_AVG': 4,
        0: 0, 1: 1, 2: 2, 3: 3, 4: 4
    }
    return mode_map[mode.upper()] if isinstance(mode, str) else mode_map[mode]


def dual_channel_pulse_train(Q, CH1, CH2, p, mode='D'):
    """双通道脉冲训练测试"""
    mode_num = _get_mode_num(mode)
    names = {0:"无测量",1:"点测量-离散模式",2:"波形测量-离散模式",3:"点测量-平均模式",4:"波形测量-平均模式"}
    print(f"🔧 配置双通道脉冲训练测试 - 模式{mode_num}: {names[mode_num]}")
    
    Q(":PMU:INIT 0")
    Q(f":PMU:RPM:CONFIGURE PMU1-{CH1}, 0")
    Q(f":PMU:RPM:CONFIGURE PMU1-{CH2}, 0")

    if p.get("ENABLE_CONNECTION_COMP", False):
        Q(f":PMU:CONNECTION:COMP {CH1}, 1, 1")
        Q(f":PMU:CONNECTION:COMP {CH2}, 1, 1")
    if p.get("ENABLE_LOAD_CONFIG", False):
        R = p.get("LOAD_RESISTANCE", 1e6)
        Q(f":PMU:LOAD {CH1}, {R}")
        Q(f":PMU:LOAD {CH2}, {R}")
    if p.get("ENABLE_LLEC", False):
        Q(f":PMU:LLEC:CONFIGURE {CH1}, 1")
        Q(f":PMU:LLEC:CONFIGURE {CH2}, 1")

    Q(f":PMU:MEASURE:RANGE {CH1}, 2, {p['CH1_RANGE']}")
    Q(f":PMU:PULSE:TRAIN {CH1}, {p['CH1_BASE']}, {p['CH1_AMPLITUDE']}")
    Q(f":PMU:PULSE:TIMES {CH1}, {p['CH1_PERIOD']}, {p['CH1_WIDTH']}, {p['CH1_RISE']}, {p['CH1_FALL']}, {p['CH1_DELAY']}")
    Q(f":PMU:MEASURE:RANGE {CH2}, 2, {p['CH2_RANGE']}")
    Q(f":PMU:PULSE:TRAIN {CH2}, {p['CH2_BASE']}, {p['CH2_AMPLITUDE']}")
    Q(f":PMU:PULSE:TIMES {CH2}, {p['CH2_PERIOD']}, {p['CH2_WIDTH']}, {p['CH2_RISE']}, {p['CH2_FALL']}, {p['CH2_DELAY']}")

    Q(f":PMU:MEASURE:MODE {mode_num}")
    if mode_num in [1, 3]:
        Q(f":PMU:MEASURE:PIV {CH1}, 1, 0")
        Q(f":PMU:MEASURE:PIV {CH2}, 1, 0")
        Q(f":PMU:TIMES:PIV {CH1}, {p['MEASURE_START_D']}, {p['MEASURE_STOP_D']}")
        Q(f":PMU:TIMES:PIV {CH2}, {p['MEASURE_START_D']}, {p['MEASURE_STOP_D']}")
    elif mode_num in [2, 4]:
        Q(f":PMU:TIMES:WAVEFORM {CH1}, {p['MEASURE_START_W']}, {p['MEASURE_STOP_W']}")
        Q(f":PMU:TIMES:WAVEFORM {CH2}, {p['MEASURE_START_W']}, {p['MEASURE_STOP_W']}")

    Q(f":PMU:PULSE:BURST:COUNT {p['PULSE_COUNT']}")
    Q(f":PMU:OUTPUT:STATE {CH1}, 1")
    Q(f":PMU:OUTPUT:STATE {CH2}, 1")
    Q(":PMU:EXECUTE")
    
    while True:
        if int(Q(":PMU:TEST:STATUS?")) == 0:
            break
        time.sleep(0.3)


def dual_channel_sweep_train(Q, CH1, CH2, p, mode='D'):
    """双通道扫描+训练测试"""
    mode_num = _get_mode_num(mode)
    names = {0:"无测量",1:"点测量-离散模式",2:"波形测量-离散模式",3:"点测量-平均模式",4:"波形测量-平均模式"}
    print(f"🔧 配置双通道扫描+训练测试 - 模式{mode_num}: {names[mode_num]}")
    
    Q(":PMU:INIT 0")
    Q(f":PMU:RPM:CONFIGURE PMU1-{CH1}, 0")
    Q(f":PMU:RPM:CONFIGURE PMU1-{CH2}, 0")

    if p.get("ENABLE_CONNECTION_COMP", False):
        Q(f":PMU:CONNECTION:COMP {CH1}, 1, 1")
        Q(f":PMU:CONNECTION:COMP {CH2}, 1, 1")
    if p.get("ENABLE_LOAD_CONFIG", False):
        R = p.get("LOAD_RESISTANCE", 1e6)
        Q(f":PMU:LOAD {CH1}, {R}")
        Q(f":PMU:LOAD {CH2}, {R}")
    if p.get("ENABLE_LLEC", False):
        Q(f":PMU:LLEC:CONFIGURE 2, 1")

    Q(f":PMU:MEASURE:MODE {mode_num}")
    Q(f":PMU:MEASURE:RANGE {CH2}, 2, {p['CH2_RANGE']}")
    Q(f":PMU:PULSE:TRAIN {CH2}, {p['CH2_BASE']}, {p['CH2_AMPLITUDE']}")
    Q(f":PMU:PULSE:TIMES {CH2}, {p['CH2_PERIOD']}, {p['CH2_WIDTH']}, {p['CH2_RISE']}, {p['CH2_FALL']}, {p['CH2_DELAY']}")
    Q(f":PMU:MEASURE:RANGE {CH1}, 2, {p['CH1_RANGE']}")
    Q(f":PMU:SWEEP:PULSE:AMPLITUDE {CH1}, {p['CH1_START']}, {p['CH1_STOP']}, {p['CH1_STEP']}, {p['CH1_VBASE']}, {p['CH1_DUALSWEEP']}")
    Q(f":PMU:PULSE:TIMES {CH1}, {p['CH1_PERIOD']}, {p['CH1_WIDTH']}, {p['CH1_RISE']}, {p['CH1_FALL']}, {p['CH1_DELAY']}")
    
    Q(f":PMU:MEASURE:MODE {mode_num}")
    if mode_num in [1, 3]:
        Q(f":PMU:MEASURE:PIV {CH1}, 1, 0")
        Q(f":PMU:MEASURE:PIV {CH2}, 1, 0")
        Q(f":PMU:TIMES:PIV {CH1}, {p['MEASURE_START_D']}, {p['MEASURE_STOP_D']}")
        Q(f":PMU:TIMES:PIV {CH2}, {p['MEASURE_START_D']}, {p['MEASURE_STOP_D']}")
    elif mode_num in [2, 4]:
        Q(f":PMU:TIMES:WAVEFORM {CH1}, {p['MEASURE_START_W']}, {p['MEASURE_STOP_W']}")
        Q(f":PMU:TIMES:WAVEFORM {CH2}, {p['MEASURE_START_W']}, {p['MEASURE_STOP_W']}")
    
    Q(f":PMU:PULSE:BURST:COUNT {p['PULSE_COUNT']}")
    Q(f":PMU:OUTPUT:STATE {CH1}, 1")
    Q(f":PMU:OUTPUT:STATE {CH2}, 1")
    Q(":PMU:EXECUTE")
    
    while True:
        if int(Q(":PMU:TEST:STATUS?")) == 0:
            break
        time.sleep(0.3)


# === segARB 专用测试 ===

def hy_pv2_segARB(Q, CH1, CH2, params):
    """PV2测试 - segARB模式"""
    print("🔧 配置 PV2 segARB 测试...")
    current_ranges = {CH1: params['Irange1'], CH2: params['Irange2']}
    rise_time = params['rise_time']
    Vp = params['Vp']
    offset = params['offset']

    start_voltages = [0, offset, offset, -Vp+offset, offset,
                      offset, Vp+offset, -Vp+offset, Vp+offset, -Vp+offset]
    stop_voltages  = [offset, offset, -Vp+offset, offset, offset,
                      Vp+offset, -Vp+offset, Vp+offset, -Vp+offset, offset]
    time_values = [rise_time, rise_time, rise_time, rise_time, 2*rise_time,
                   rise_time, 2*rise_time, 2*rise_time, 2*rise_time, rise_time]
    meas_types = [2] * 10


    ch1_config = (1, start_voltages, stop_voltages, time_values, meas_types)
    ch2_config = (1, [0.0]*10, [0.0]*10, time_values, meas_types)
    seq_configs = {CH1: [ch1_config], CH2: [ch2_config]}
    execute_segARB_test(Q, [CH1, CH2], seq_configs, current_ranges=current_ranges)


def hy_pund_segARB(Q, CH1, CH2, params):
    """PUND测试 - segARB模式"""
    print("🔧 配置 PUND segARB 测试...")
    current_ranges = {CH1: params['Irange1'], CH2: params['Irange2']}
    rise_time = params['rise_time']
    Vp = params['Vp']
    offset = params['offset']

    start_voltages = [
        0, 0, offset, -Vp+offset, -Vp+offset, offset, offset, Vp+offset, Vp+offset,
        offset, offset, Vp+offset, Vp+offset, offset, offset, -Vp+offset, -Vp+offset,
        offset, offset, -Vp+offset, -Vp+offset, offset
    ]
    stop_voltages = [
        0, offset, -Vp+offset, -Vp+offset, offset, offset, Vp+offset, Vp+offset,
        offset, offset, Vp+offset, Vp+offset, offset, offset, -Vp+offset, -Vp+offset,
        offset, offset, -Vp+offset, -Vp+offset, offset, offset
    ]
    time_values = [rise_time] * 22
    meas_types = [2] * 22

    ch1_config = (1, start_voltages, stop_voltages, time_values, meas_types)
    ch2_config = (1, [0.0]*22, [0.0]*22, time_values, meas_types)
    seq_configs = {CH1: [ch1_config], CH2: [ch2_config]}
    execute_segARB_test(Q, [CH1, CH2], seq_configs, current_ranges=current_ranges)


def hy_NISswitch_segARB(Q, CH1, CH2, params):
    """NIS Switch测试 - segARB模式
    用于铁电隧穿结的非易失性电阻切换测量
    """
    print("🔧 配置 NIS Switch segARB 测试...")
    current_ranges = {CH1: params['Irange1'], CH2: params['Irange2']}
    MeasureSquare = params.get('MeasureSquare', True)
    offset = params['offset']
    Vp = params['Vp']
    Rt_p = params['Rt_p']
    Delaytime = params['Delaytime']
    Vsquare = params['Vsquare']
    Rt_s = params['Rt_s']
    Dwell = params['Dwell']

    # 14段序列：极化脉冲 + 方波读取
    start_voltages = [
        0, 0, offset, -Vp+offset, offset,           # 0-4: 初始化 + 负极化脉冲
        offset, Vsquare+offset, Vsquare+offset,     # 5-7: 方波读取 (HRS)
        offset,                                      # 8: 延迟
        offset, Vp+offset, offset,                  # 9-11: 正极化脉冲
        offset, Vp+offset,                          # 12-13: 第二次正极化
    ]
    stop_voltages = [
        0, offset, -Vp+offset, offset, offset,      # 0-4
        Vsquare+offset, Vsquare+offset, offset,     # 5-7
        offset,                                      # 8
        Vp+offset, offset, offset,                  # 9-11
        Vp+offset, offset,                          # 12-13
    ]
    time_values = [
        Rt_p, Rt_p, Rt_p, Rt_p, Delaytime,          # 0-4
        Rt_s, Dwell, Rt_s,                          # 5-7: 方波
        Delaytime,                                   # 8
        Rt_p, Rt_p, Delaytime,                      # 9-11
        Rt_p, Rt_p,                                 # 12-13
    ]
    
    # 测量类型：只在方波读取段测量
    Msq = 2 if MeasureSquare else 0
    meas_types = [0, 0, 0, 0, 0, Msq, Msq, Msq, 0, 2, 2, 0, 2, 2]


    ch1_config = (1, start_voltages, stop_voltages, time_values, meas_types)
    ch2_config = (1, [0.0]*14, [0.0]*14, time_values, meas_types)
    seq_configs = {CH1: [ch1_config], CH2: [ch2_config]}
    execute_segARB_test(Q, [CH1, CH2], seq_configs, current_ranges=current_ranges)


def build_endurance_exec_list(max_repeat, decade=10, seq_ids=(1, 2, 3)):
    """构建Endurance执行列表：1、10、100、...次循环"""
    exec_list = []
    n = 1
    while n <= max_repeat:
        exec_list.append((seq_ids[0], n))
        exec_list.append((seq_ids[1], 1))
        exec_list.append((seq_ids[2], 1))
        n *= decade
    return exec_list


def hy_Endurance_segARB(Q, CH1, CH2, params, max_repeat):
    """Endurance测试 - 3序列版本"""
    print("🔧 配置 Endurance segARB 测试（3序列版本）...")
    current_ranges = {CH1: params['Irange1'], CH2: params['Irange2']}
    
    rst_c = params['rise_time_cycle']; Vc = params['Vc']; offset_c = params['offset_c']
    rst_e = params['rise_time_PV']; Ve = params['Ve']; offset_e = params['offset_e']
    rst_p = params['rise_time_PUND']; Vp = params['Vp']; offset_p = params['offset_p']

    # seq1: 循环三角波（不测量）
    start_cycle = [offset_c, -Vc+offset_c, offset_c, Vc+offset_c]
    stop_cycle  = [-Vc+offset_c, offset_c, Vc+offset_c, offset_c]
    time_cycle  = [rst_c] * 4
    ch1_cfg_cycle = (1, start_cycle, stop_cycle, time_cycle, [0]*4, [0.0]*4, [1]*4)
    ch2_cfg_cycle = (1, [0.0]*4, [0.0]*4, time_cycle, [0]*4, [0.0]*4, [1]*4)

    # seq2: 三角波读取（测量）
    start_tri = [offset_e, -Ve+offset_e, offset_e, Ve+offset_e]
    stop_tri  = [-Ve+offset_e, offset_e, Ve+offset_e, offset_e]
    time_tri  = [rst_e] * 4
    ch1_cfg_tri = (2, start_tri, stop_tri, time_tri, [2]*4, [0.0]*4, [1]*4)
    ch2_cfg_tri = (2, [0.0]*4, [0.0]*4, time_tri, [2]*4, [0.0]*4, [1]*4)

    # seq3: PUND读取（测量）
    start_pund = [
        0, 0, offset_p, -Vp+offset_p, -Vp+offset_p, offset_p, offset_p, Vp+offset_p, Vp+offset_p,
        offset_p, offset_p, Vp+offset_p, Vp+offset_p, offset_p, offset_p, -Vp+offset_p, -Vp+offset_p,
        offset_p, offset_p, -Vp+offset_p, -Vp+offset_p, offset_p
    ]
    stop_pund = [
        0, offset_p, -Vp+offset_p, -Vp+offset_p, offset_p, offset_p, Vp+offset_p, Vp+offset_p,
        offset_p, offset_p, Vp+offset_p, Vp+offset_p, offset_p, offset_p, -Vp+offset_p, -Vp+offset_p,
        offset_p, offset_p, -Vp+offset_p, -Vp+offset_p, offset_p, offset_p
    ]
    time_pund = [rst_p] * 22
    ch1_cfg_pund = (3, start_pund, stop_pund, time_pund, [2]*22, [0.0]*22, [1]*22)
    ch2_cfg_pund = (3, [0.0]*22, [0.0]*22, time_pund, [2]*22, [0.0]*22, [1]*22)

    seq_configs = {
        CH1: [ch1_cfg_cycle, ch1_cfg_tri, ch1_cfg_pund],
        CH2: [ch2_cfg_cycle, ch2_cfg_tri, ch2_cfg_pund],
    }

    exec_list = build_endurance_exec_list(max_repeat=max_repeat, decade=10, seq_ids=(1, 2, 3))
    seq_list = {CH1: exec_list, CH2: exec_list}
    
    print(f"   执行计划: {exec_list}")
    print(f"   seq1(循环，不测): ±{Vc}V, {rst_c*1e6:.1f}μs, offset={offset_c}V")
    print(f"   seq2(PV读，测量): ±{Ve}V, {rst_e*1e6:.1f}μs, offset={offset_e}V")
    print(f"   seq3(PUND读，测量): ±{Vp}V, {rst_p*1e6:.1f}μs, offset={offset_p}V")

    execute_segARB_test(Q, [CH1, CH2], seq_configs, seq_list=seq_list, current_ranges=current_ranges)


# === 便捷构建助手 ===
