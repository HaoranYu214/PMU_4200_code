# -*- coding: utf-8 -*-
# 4225-PMU command and test helpers.
"""
PMU核心测试模块 - segARB配置、执行、设备控制
"""

import time


# 4225-PMU Segment ARB hardware/KXCI limits shared by every test entry.
# Keep these here instead of copying numeric literals into individual tests.
#
# Project limit: 2048 defined segments per channel across all sequences; loops reuse stored definitions.
# KXCI Rev. D explicitly documents 2048 DATA:GET points per block (7-7), not this storage limit.
# The storage limit retains the project setting; confirm its hardware basis in the Pulse Card user manual.
MAX_SEGMENTS_PER_SEQUENCE = 2048

# KXCI can return at most 65,536 generated measurement points for one test
# (per channel). If the requested/default rate would exceed this buffer, KXCI
# automatically lowers the effective sample rate.
KXCI_MAX_DATA_POINTS = 65_536

# :PMU:INIT resets :PMU:SAMPLE:RATE to 200 MSa/s. KXCI may still adjust that
# rate upward for the shortest measurement interval or downward to remain
# within KXCI_MAX_DATA_POINTS, so callers must not assume an exact point count.
KXCI_DEFAULT_SAMPLE_RATE = 200e6


# === 设备控制 ===

def power_off_outputs(Q, channels):
    """测试数据读取后立刻关闭输出（幂等）"""
    for ch in channels:
        try:
            Q(f":PMU:OUTPUT:STATE {ch}, 0")
        except Exception:
            pass


# === segARB 核心配置 ===

# Shared options (KXCI Rev. D, May 2024; printed page numbers):
# SOURCE:RANGE selects 10 or 40 V; INIT defaults to 10 V (7-24).
# LOAD is a DUT impedance estimate, 1-1e7 ohm (7-11), not current compliance.
# SAMPLE_RATE: 1e3-200e6 samples/s; KXCI may adjust it to fit the measurement
# windows and 65536-point budget (7-23/7-24).
# ENABLE_* flags enable the named feature. Enabling connection compensation
# uses previously measured calibration data; it does not perform calibration.
# LLEC applies iterative pulses before acquisition (7-10); confirm support for
# the intended SARB mode rather than assuming every accepted option is supported.
def _apply_common_pmu_options(Q, channels, options=None):
    """Apply optional PMU setup shared by pulse and segARB tests."""
    if not options:
        return
    sample_rate = options.get("SAMPLE_RATE")
    if sample_rate is not None:
        sample_rate = float(sample_rate)
        if not 1e3 <= sample_rate <= 200e6:
            raise ValueError("PMU SAMPLE_RATE must be between 1e3 and 200e6 samples/s.")
        Q(f":PMU:SAMPLE:RATE {sample_rate:.12g}")
    if options.get("ENABLE_CONNECTION_COMP", False):
        comp_channels = options.get("CONNECTION_COMP_CHANNELS", channels)
        for ch in comp_channels:
            if ch not in channels:
                raise ValueError(f"Connection-comp channel {ch} is not an active PMU channel.")
            Q(f":PMU:CONNECTION:COMP {ch}, 1, 1")
    if options.get("ENABLE_LOAD_CONFIG", False):
        resistances = options.get("LOAD_RESISTANCES")
        default_resistance = options.get("LOAD_RESISTANCE", 1e6)
        for ch in channels:
            resistance = (
                resistances.get(ch, default_resistance)
                if resistances is not None
                else default_resistance
            )
            Q(f":PMU:LOAD {ch}, {resistance}")
    if options.get("ENABLE_LLEC", False):
        llec_channels = options.get("LLEC_CHANNELS", channels)
        for ch in llec_channels:
            if ch not in channels:
                raise ValueError(f"LLEC channel {ch} is not an active PMU channel.")
            Q(f":PMU:LLEC:CONFIGURE {ch}, 1")


# KXCI 7-40/7-48/7-52: at most 128 values per SARB array command; append further chunks with :ADD.
SARB_ARRAY_CHUNK = 128


def _send_sarb_array(Q, command, ch, seq_id, values, *, formatter=str):
    """Send a Segment Arb array, appending extra chunks with :ADD commands."""
    if not values:
        return

    chunks = [values[index : index + SARB_ARRAY_CHUNK] for index in range(0, len(values), SARB_ARRAY_CHUNK)]
    first_values = ", ".join(formatter(value) for value in chunks[0])
    Q(f"{command} {ch}, {seq_id}, {first_values}")

    add_command = f"{command}:ADD"
    for chunk in chunks[1:]:
        chunk_values = ", ".join(formatter(value) for value in chunk)
        Q(f"{add_command} {ch}, {seq_id}, {chunk_values}")


def _normalize_seg_arb_measurements(time_values, meas_types=None, meas_start=None, meas_stop=None):
    """Return PMU-valid measurement arrays for one Segment Arb sequence.

    The PMU requires unmeasured segments to use a zero start/stop window. For
    measured segments, the window must be positive and remain inside the
    corresponding segment duration.
    """
    segment_count = len(time_values)
    durations = [float(value) for value in time_values]
    modes = [2] * segment_count if meas_types is None else list(meas_types)
    starts = [0.0] * segment_count if meas_start is None else list(meas_start)
    stops = (
        [duration if mode != 0 else 0.0 for duration, mode in zip(durations, modes)]
        if meas_stop is None
        else list(meas_stop)
    )

    arrays = {
        "meas_types": modes,
        "meas_start": starts,
        "meas_stop": stops,
    }
    for name, values in arrays.items():
        if len(values) != segment_count:
            raise ValueError(
                f"Segment Arb {name} has {len(values)} entries; "
                f"expected {segment_count}."
            )

    for index, duration in enumerate(durations):
        if duration <= 0:
            raise ValueError(f"Segment Arb segment {index} duration must be positive.")

        mode = int(modes[index])
        if mode not in (0, 1, 2, 3, 4):
            raise ValueError(f"Segment Arb segment {index} has invalid measure type {modes[index]}.")
        modes[index] = mode

        if mode == 0:
            # Keithley error -823 is raised when a measurement window is sent
            # for a segment whose measurement type is NONE.
            starts[index] = 0.0
            stops[index] = 0.0
            continue

        start = float(starts[index])
        stop = float(stops[index])
        tolerance = max(abs(duration) * 1e-12, 1e-18)
        if start < 0 or stop <= start or stop > duration + tolerance:
            raise ValueError(
                f"Segment Arb segment {index} measurement window "
                f"[{start:g}, {stop:g}] s is outside its {duration:g} s duration."
            )
        starts[index] = start
        stops[index] = min(stop, duration)

    return modes, starts, stops


# Segment Arb parameter limits (KXCI 7-40, 7-42, 7-48, 7-50, 7-52):
# Voltage endpoints: -10..10 V on the 10 V range, -40..40 V on the 40 V range.
# Include offsets when checking actual levels; adjacent segment voltages must join.
# Each segment: 20 ns-1 s at 10 V, 50 ns-1 s at 40 V, with 10 ns resolution.
# Measurement windows are relative seconds within that segment, not percentages.
# This helper requires 0 <= start < stop <= duration for measured segments;
# unmeasured segments use start=stop=0. Sequence IDs are 1-512; hardware loops
# are 1-1e12 (7-56). Not every manual boundary is validated by this helper.
# These are instrument command bounds, not a verified safe operating area for a DUT.
def configure_segARB_sequence(Q, ch, seq_id, start_voltages, stop_voltages, time_values,
                              meas_types=None, meas_start=None, meas_stop=None):
    """通用segARB序列配置函数，支持长数组自动用 :ADD 续传。"""
    n_segments = len(start_voltages)
    if n_segments > MAX_SEGMENTS_PER_SEQUENCE:
        raise ValueError(
            f"Segment Arb sequence has {n_segments} segments; the 4225-PMU "
            f"limit is {MAX_SEGMENTS_PER_SEQUENCE}."
        )
    array_lengths = {
        "stop_voltages": len(stop_voltages),
        "time_values": len(time_values),
    }
    for name, length in array_lengths.items():
        if length != n_segments:
            raise ValueError(
                f"Segment Arb {name} has {length} entries; expected {n_segments}."
            )
    meas_types, meas_start, meas_stop = _normalize_seg_arb_measurements(
        time_values,
        meas_types,
        meas_start,
        meas_stop,
    )

    _send_sarb_array(Q, ":PMU:SARB:SEQ:STARTV", ch, seq_id, start_voltages)
    _send_sarb_array(Q, ":PMU:SARB:SEQ:STOPV", ch, seq_id, stop_voltages)
    _send_sarb_array(Q, ":PMU:SARB:SEQ:TIME", ch, seq_id, time_values, formatter=lambda value: f"{value:.2e}")
    _send_sarb_array(Q, ":PMU:SARB:SEQ:MEAS:TYPE", ch, seq_id, meas_types)
    _send_sarb_array(Q, ":PMU:SARB:SEQ:MEAS:START", ch, seq_id, meas_start, formatter=lambda value: f"{value:.2e}")
    _send_sarb_array(Q, ":PMU:SARB:SEQ:MEAS:STOP", ch, seq_id, meas_stop, formatter=lambda value: f"{value:.2e}")


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
            channel_time = ch_configs[ch][3]
            # Two dynamic channels are already aligned when their segment
            # durations match.  Do not require either waveform to be constant.
            if len(channel_time) == len(ref_time) and all(
                abs(a - b) < 1e-15 for a, b in zip(channel_time, ref_time)
            ):
                continue
            start_v, stop_v = ch_configs[ch][1], ch_configs[ch][2]
            is_const = all(abs(a-b) < 1e-12 for a, b in zip(start_v, stop_v))
            if not is_const:
                raise ValueError(f"CH{ch} 非恒压序列，且与CH{ref_ch}段数不一致，无法自动对齐")
            v = start_v[0] if len(start_v) else 0.0
            # An expanded hold waveform is deliberately unmeasured.
            zeros = [0.0] * len(ref_time)
            new_config = (
                ch_configs[ch][0],
                [v] * len(ref_time),
                [v] * len(ref_time),
                list(ref_time),
                [0] * len(ref_time),
                zeros.copy(),
                zeros.copy(),
            )
            for i, cfg in enumerate(seq_configs[ch]):
                if cfg[0] == seq_id:
                    seq_configs[ch][i] = new_config
                    break
    return seq_configs


def _validate_segment_arb_storage(seq_configs):
    """Validate the 2048 stored-segment allowance independently per channel."""
    for ch, configs in seq_configs.items():
        total_segments = sum(len(config[3]) for config in configs)
        if total_segments > MAX_SEGMENTS_PER_SEQUENCE:
            raise ValueError(
                f"CH{ch} defines {total_segments} stored Segment Arb segments "
                f"across {len(configs)} sequences; the per-channel 4225-PMU "
                f"limit is {MAX_SEGMENTS_PER_SEQUENCE}."
            )


def validate_segment_arb_configs(seq_configs):
    """Validate complete Segment Arb configs without sending any commands."""
    _validate_segment_arb_storage(seq_configs)
    for channel, configs in seq_configs.items():
        for config in configs:
            if len(config) < 4:
                raise ValueError(f"CH{channel} Segment Arb config is incomplete.")
            seq_id, start_v, stop_v, time_values = config[:4]
            segment_count = len(time_values)
            if len(start_v) != segment_count or len(stop_v) != segment_count:
                raise ValueError(
                    f"CH{channel} seq {seq_id} has mismatched voltage/time arrays."
                )
            _normalize_seg_arb_measurements(
                time_values,
                config[4] if len(config) > 4 else None,
                config[5] if len(config) > 5 else None,
                config[6] if len(config) > 6 else None,
            )


def execute_segARB_test(Q, channels, seq_configs, seq_list=None, current_ranges=None, options=None):
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
    # Alignment can expand a constant companion channel. Validate the aligned
    # result before changing any instrument state.
    seq_configs = auto_align_channels(seq_configs)
    validate_segment_arb_configs(seq_configs)

    # INIT 1 selects Segment Arb; 0 selects standard pulse mode. INIT resets ranges/sample rate (7-9, 7-23).
    Q(":PMU:INIT 1")

    # 配置RPM,链接RPM到PMU通道
    for ch in channels:
        Q(f":PMU:RPM:CONFIGURE PMU1-{ch}, 0")

    _apply_common_pmu_options(Q, channels, options=options)

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
            print(f"   CH{ch} current range (fixed): {i_range:.2e} A")

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

    print("Waiting for Segment Arb test completion...")
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
                        print("Segment Arb test complete.")
                        break
                except ValueError:
                    print(f"WARNING: Could not parse PMU test status: '{status_str}'")
        except Exception as e:
            print(f"WARNING: PMU test status query failed: {e}")
        time.sleep(0.3)


# === 底层测试函数 ===

# Standard pulse timing differs from SARB dwell segments (KXCI 7-17..7-19):
# PMU period: 60 ns-1 s at 10 V; 500 ns-1 s at 40 V. All channels share it.
# Width: >=40 ns at 10 V, >=250 ns at 40 V, and <=period-10 ns.
# Rise/fall: 20 ns-33 ms at 10 V; 50 ns-33 ms at 40 V.
# Also require width > (rise+fall)/2, rise <= width, and
# period-delay-width-(rise+fall)/2 > 40 ns. Burst count: 1-10000.
# PULSE:TRAIN base-to-peak span is limited to 10/40 V respectively (7-21).
# A measurement range setting does not provide SMU-style current compliance.
def _get_mode_num(mode):
    """解析测量模式"""
    # Acquisition modes (7-13/7-44): 0 none, 1 spot discrete, 2 waveform discrete,
    # 3 spot average, 4 waveform average. Mode 0 disables acquisition, not pulse output.
    mode_map = {
        'D': 1, 'DISCRETE': 1, 'SPOT': 1,
        'WFM': 2, 'WAVEFORM': 2, 'WAVE': 2,
        'NONE': 0, 'NO': 0,
        'AVG': 3, 'AVERAGE': 3, 'SPOT_AVG': 3,
        'WFM_AVG': 4, 'WAVEFORM_AVG': 4, 'WAVE_AVG': 4,
        0: 0, 1: 1, 2: 2, 3: 3, 4: 4
    }
    return mode_map[mode.upper()] if isinstance(mode, str) else mode_map[mode]


# ACQUIRE_HIGH/LOW select returned pulse levels; at least one must be enabled.
# PIV start/stop fractions are 0-1 (7-32). WAVEFORM pre/post fractions are also
# 0-1 but describe capture before/after the pulse, not start/stop coordinates;
# the pre-capture interval cannot exceed pulse delay (7-34).
# Standard pulse sweep dualSweep: 0 one-way, 1 forward-and-return (7-29).
def _configure_pulse_iv_acquisition(Q, channels, params):
    """Configure High/Low data returned by spot pulse-I-V modes."""
    acquire_high = bool(params.get("ACQUIRE_HIGH", True))
    acquire_low = bool(params.get("ACQUIRE_LOW", False))
    if not acquire_high and not acquire_low:
        raise ValueError("At least one of ACQUIRE_HIGH or ACQUIRE_LOW must be enabled.")
    for ch in channels:
        Q(f":PMU:MEASURE:PIV {ch}, {int(acquire_high)}, {int(acquire_low)}")
    return acquire_high, acquire_low
