#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
buffered_flame_measure_eval.py
────────────────────────────────────────────────────────
1. GUI で外部トリガ周波数を入力
2. 周期 × SAFETY_MARGIN (ms) を積分時間に設定（下限 0.01 ms=10 µs）
3. Flame を External Level Trigger で N スキャン取得
4. 全データを RAM にバッファ
5. 測定後にフォルダへ一括保存
6. 同フォルダ内で時刻精度を評価 & グラフ保存
"""

from seabreeze.spectrometers import Spectrometer
from datetime import datetime, timezone
from pathlib import Path
import tkinter as tk
from tkinter import simpledialog, messagebox
import numpy as np
import json, sys, textwrap

# 解析で使用
import matplotlib.pyplot as plt
from dateutil.parser import isoparse             # pip install python-dateutil

# ---------- ユーザ設定 ----------
N_SCANS       = 100          # 取得スキャン数
TRIGGER_MODE  = 3            # External Hardware Level
SAFETY_MARGIN = 0.70         # 積分時間 = 周期 × 0.70
SAVE_WL_FILE  = True
MIN_INTEG_MS  = 0.01         # 10 µs
# -------------------------------


# ──────────────────────────────────────────────────────
# GUI UTILS
# ──────────────────────────────────────────────────────
def ask_frequency() -> float:
    root = tk.Tk(); root.withdraw()
    try:
        f = simpledialog.askfloat(
            "外部トリガ周波数の入力", "外部トリガ周波数 [Hz]：",
            minvalue=0.1, maxvalue=50_000, initialvalue=1_000.0
        )
    except tk.TclError:
        sys.exit("GUI が起動できませんでした。X11/Wayland を確認してください。")
    if f is None: sys.exit("キャンセルされたため終了します。")
    return f


def confirm_settings(freq_hz: float, integ_ms: float):
    root = tk.Tk(); root.withdraw()
    msg = f"外部トリガ周波数: {freq_hz:.3f} Hz\n積分時間: {integ_ms:.3f} ms\n\n開始しますか？"
    if not messagebox.askyesno("設定確認", msg):
        sys.exit("ユーザが中止しました。")


# ──────────────────────────────────────────────────────
# TIMING-EVAL UTIL
# ──────────────────────────────────────────────────────
def evaluate_timing(ts_list, ideal_hz, out_dir: Path):
    """timestamps(list[str]), ideal freq → txt+png を out_dir に保存"""
    ideal_ms = 1_000 / ideal_hz
    t0 = isoparse(ts_list[0])
    rel_ms = np.array([(isoparse(ts) - t0).total_seconds()*1_000 for ts in ts_list])
    interval_ms = np.diff(rel_ms)
    err_ms = interval_ms - ideal_ms

    mean, std = err_ms.mean(), err_ms.std(ddof=1)
    absmax = np.abs(err_ms).max()
    lost = (interval_ms > 1.5*ideal_ms).sum()

    summary = textwrap.dedent(f"""
        理想周期           : {ideal_ms:.3f} ms ({ideal_hz:.3f} Hz)
        サンプル数         : {len(interval_ms)}
        平均誤差           : {mean:+.4f} ms
        標準偏差 (RMSジッタ): {std:.4f} ms
        最大|誤差|          : {absmax:.4f} ms
        周期抜け推定本数    : {lost}
    """).strip()

    # ─ グラフ
    fig,(ax1,ax2)=plt.subplots(2,1,figsize=(8,6),tight_layout=True)
    ax1.plot(interval_ms,".-")
    ax1.axhline(ideal_ms,color="k",lw=.8,ls="--")
    ax1.set_ylabel("Interval (ms)")
    ax1.set_title("Scan intervals vs ideal")

    ax2.hist(err_ms,bins=30)
    ax2.set_xlabel("Error from ideal (ms)")
    ax2.set_ylabel("Count")
    ax2.set_title("Histogram of interval error")

    (out_dir/"timing_plot.png").write_bytes(fig_to_png_bytes(fig))
    plt.close(fig)

    (out_dir/"timing_report.txt").write_text(summary, encoding="utf-8")
    print("\n"+summary+"\n")
    print("→ timing_report.txt / timing_plot.png を保存しました。")


def fig_to_png_bytes(fig):
    import io
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150)
    return buf.getvalue()


# ──────────────────────────────────────────────────────
# MAIN
# ──────────────────────────────────────────────────────
def main():
    # 1) 周波数入力 & 積分時間計算 (ms 表示)
    freq_hz = ask_frequency()
    period_ms = 1_000 / freq_hz
    integ_ms = max(MIN_INTEG_MS, period_ms * SAFETY_MARGIN)
    confirm_settings(freq_hz, integ_ms)
    integ_us = int(round(integ_ms * 1_000))         # µs へ変換

    # 2) スペクトロメータ設定
    spec = Spectrometer.from_first_available()
    spec.integration_time_micros(integ_us)
    spec.trigger_mode(TRIGGER_MODE)

    wavelengths = spec.wavelengths().tolist()
    n_px = len(wavelengths)

    # 3) RAM バッファ
    spectra    = np.empty((N_SCANS, n_px), dtype=np.float32)
    timestamps = []

    # 4) 測定ループ
    print(f"測定開始… (integration {integ_ms:.3f} ms)")
    for i in range(N_SCANS):
        spectra[i] = spec.intensities()
        timestamps.append(datetime.now(timezone.utc).isoformat())
        if (i+1)%10==0 or i==N_SCANS-1:
            print(f"{i+1}/{N_SCANS} スキャン取得")

    spec.close()
    print("測定終了。ディスクへ保存中…")

    # 5) 保存フォルダ作成 & JSON 一括書込
    run_dir = Path(datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S_%fZ"))
    run_dir.mkdir()
    if SAVE_WL_FILE:
        (run_dir/"wavelengths_nm.json").write_text(
            json.dumps(wavelengths, indent=2), encoding="utf-8")

    for i, ts in enumerate(timestamps):
        data = {
            "scan_index": i,
            "timestamp_utc": ts,
            "wavelengths_nm": wavelengths,
            "intensities": spectra[i].tolist(),
        }
        (run_dir/f"scan_{i:03d}.json").write_text(
            json.dumps(data, indent=2), encoding="utf-8")

    print(f"保存完了 → {run_dir.resolve()}")

    # 6) 時刻精度の評価
    evaluate_timing(timestamps, freq_hz, run_dir)


if __name__ == "__main__":
    main()
