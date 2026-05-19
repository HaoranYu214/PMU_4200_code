# -*- coding: utf-8 -*-
"""Helpers for previewing generated segARB sequence configs."""

from pathlib import Path


MEASURE_MODE_LABELS = {
    0: "off",
    1: "spot discrete",
    2: "waveform discrete",
    3: "spot average",
    4: "waveform average",
}


def _voltage_at(start_v, stop_v, segment_time, offset_time):
    """Linearly interpolate voltage inside one Segment Arb segment."""
    if segment_time <= 0:
        return stop_v
    fraction = max(0.0, min(1.0, offset_time / segment_time))
    return start_v + (stop_v - start_v) * fraction


def preview_sequence_configs(configs, output_path=None, *, title_prefix="", dpi=180):
    """Preview one or more segARB sequence configs.

    Supported config tuple shapes:
    (seq_id, start_v, stop_v, time_values)
    (seq_id, start_v, stop_v, time_values, meas_types)
    (seq_id, start_v, stop_v, time_values, meas_types, meas_start, meas_stop)

    If output_path is None, show the figure interactively. Otherwise save it.
    """
    import matplotlib.pyplot as plt

    if not configs:
        raise ValueError("No sequence configs were provided for preview.")

    fig, axes = plt.subplots(len(configs), 1, figsize=(12, max(3.5, 3.5 * len(configs))), sharex=False)
    if len(configs) == 1:
        axes = [axes]

    colors = ["tab:blue", "tab:red", "tab:green", "tab:purple", "tab:orange"]
    meas_styles = {
        1: ("black", "spot discrete"),
        2: ("tab:orange", "waveform discrete"),
        3: ("tab:purple", "spot average"),
        4: ("tab:green", "waveform average"),
    }
    for index, (ax, config) in enumerate(zip(axes, configs)):
        seq_id, start_v, stop_v, times = config[:4]
        meas_types = config[4] if len(config) > 4 else [2] * len(times)
        meas_start = config[5] if len(config) > 5 else [0.0] * len(times)
        meas_stop = config[6] if len(config) > 6 else list(times)
        x = []
        y = []
        meas_lines = {mode: ([], []) for mode in meas_styles}
        spot_x = []
        spot_y = []
        t = 0.0

        for sv, ev, dt, mt, ms, me in zip(start_v, stop_v, times, meas_types, meas_start, meas_stop):
            x.extend([t, t + dt])
            y.extend([sv, ev])
            if mt != 0:
                ms = max(0.0, min(dt, ms))
                me = max(0.0, min(dt, me))
                start_y = _voltage_at(sv, ev, dt, ms)
                stop_y = _voltage_at(sv, ev, dt, me)
                if mt in (1, 3):
                    spot_t = t + (ms + me) / 2
                    spot_x.append(spot_t)
                    spot_y.append(_voltage_at(sv, ev, dt, (ms + me) / 2))
                if mt in meas_lines:
                    meas_x, meas_y = meas_lines[mt]
                    meas_x.extend([t + ms, t + me, None])
                    meas_y.extend([start_y, stop_y, None])
            t += dt

        color = colors[index % len(colors)]
        ax.plot(x, y, color=color, linewidth=1.2)
        for mode, (mode_label_color, mode_label) in meas_styles.items():
            meas_x, meas_y = meas_lines[mode]
            if meas_x:
                ax.plot(meas_x, meas_y, color=mode_label_color, linewidth=3, alpha=0.55, label=mode_label)
        if spot_x:
            ax.scatter(spot_x, spot_y, color="black", s=18, zorder=3, label="spot sample")
        prefix = f"{title_prefix} " if title_prefix else ""
        ax.set_title(f"{prefix}seq {seq_id}, {len(times)} segments")
        ax.set_ylabel("Voltage (V)")
        ax.grid(alpha=0.3)
        handles, labels = ax.get_legend_handles_labels()
        if handles:
            ax.legend(handles, labels, loc="upper right")

    axes[-1].set_xlabel("Time (s)")
    fig.tight_layout()
    if output_path is None:
        plt.show()
        return None

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(output_path, dpi=dpi)
    plt.close(fig)
    print(f"Saved waveform preview to {output_path}")
    return output_path


def preview_from_pmu_test(test_func, ch1, ch2, params, output_path=None, *, channel=None, title_prefix="", dpi=180):
    """Preview seq configs generated inside a pmu_tests helper.

    This captures the seq_configs passed to execute_segARB_test without opening
    an instrument connection.
    """
    import contextlib
    import io

    import src.pmu_tests as pmu_tests

    captured = {}
    original_execute = pmu_tests.execute_segARB_test

    def capture_execute(_query, _channels, seq_configs, seq_list=None, current_ranges=None, options=None):
        captured["seq_configs"] = seq_configs
        captured["seq_list"] = seq_list
        captured["current_ranges"] = current_ranges
        captured["options"] = options

    try:
        pmu_tests.execute_segARB_test = capture_execute
        with contextlib.redirect_stdout(io.StringIO()):
            test_func(lambda _command: "", ch1, ch2, params)
    finally:
        pmu_tests.execute_segARB_test = original_execute

    if "seq_configs" not in captured:
        raise RuntimeError("No seq_configs were captured from the test function.")

    preview_channel = ch1 if channel is None else channel
    return preview_sequence_configs(
        captured["seq_configs"][preview_channel],
        output_path,
        title_prefix=title_prefix,
        dpi=dpi,
    )
