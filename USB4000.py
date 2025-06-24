#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
buffered_flame_measure_ms.py  ―  積分時間を ms で表示
────────────────────────────────────────────────────────
1. 周波数入力（GUI）
2. 積分時間 = 周期 × SAFETY_MARGIN → ms で表示
3. Flame を Ext-Level Trigger で N スキャン取得
4. データは RAM にバッファし、測定後に一括保存
"""

from seabreeze.spectrometers import Spectrometer
from datetime import datetime, timezone
from pathlib import Path
import tkinter as tk
from tkinter import simpledialog, messagebox
import numpy as np
import json, sys

# ---------- ユーザ設定 ----------
N_SCANS       = 100          # 取得スキャン数
TRIGGER_MODE  = 3            # External Hardware Level
SAFETY_MARGIN = 0.20         # 積分時間 = 周期 × 0.20
SAVE_WL_FILE  = True
MIN_INTEG_MS  = 0.01         # 最小 0.01 ms = 10 µs
# -------------------------------


def ask_frequency() -> float:
    """外部トリガ周波数 [Hz] を GUI で取得。キャンセルで終了。"""
    root = tk.Tk(); root.withdraw()
    try:
        f = simpledialog.askfloat(
            title="外部トリガ周波数の入力",
            prompt="外部トリガ周波数 [Hz]：",
            minvalue=0.1, maxvalue=50_000, initialvalue=1_000.0
        )
    except tk.TclError:
        sys.exit("GUI が起動できませんでした。X11/Wayland を確認してください。")
    if f is None:
        sys.exit("キャンセルされたため終了します。")
    return f


def confirm_settings(freq_hz: float, integ_ms: float):
    root = tk.Tk(); root.withdraw()
    msg = (
        f"外部トリガ周波数: {freq_hz:.3f} Hz\n"
        f"積分時間: {integ_ms:.3f} ms\n\n開始しますか？"
    )
    if not messagebox.askyesno("設定確認", msg):
        sys.exit("ユーザが中止しました。")


def main() -> None:
    # ── 1) 周波数入力 & 積分時間(ms) 計算 ────────────────────
    freq_hz   = ask_frequency()
    period_ms = 1_000 / freq_hz                       # 周期 [ms]
    integ_ms  = max(MIN_INTEG_MS, period_ms * SAFETY_MARGIN)
    confirm_settings(freq_hz, integ_ms)

    # µs に変換してドライバへ渡す
    integ_us = int(round(integ_ms * 1_000))

    # ── 2) スペクトロメータ初期化 ─────────────────────────────
    spec = Spectrometer.from_first_available()
    spec.integration_time_micros(integ_us)
    spec.trigger_mode(TRIGGER_MODE)

    wavelengths = spec.wavelengths().tolist()
    n_px = len(wavelengths)

    # ── 3) RAM バッファ確保 ────────────────────────────────
    spectra    = np.empty((N_SCANS, n_px), dtype=np.float32)
    timestamps = []

    # ── 4) 測定ループ ─────────────────────────────────────
    print(f"測定開始… (integration {integ_ms:.3f} ms)")
    for i in range(N_SCANS):
        spectra[i] = spec.intensities()
        timestamps.append(datetime.now(timezone.utc).isoformat())
        if (i + 1) % 10 == 0 or i == N_SCANS - 1:
            print(f"{i+1}/{N_SCANS} スキャン取得")

    spec.close()
    print("測定終了。ディスクへ保存中…")

    # ── 5) 一括保存 ───────────────────────────────────────
    run_id = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S_%fZ")
    run_dir = Path(run_id); run_dir.mkdir()
    print(f"保存フォルダ: {run_dir.resolve()}")

    if SAVE_WL_FILE:
        with (run_dir / "wavelengths_nm.json").open("w", encoding="utf-8") as f:
            json.dump(wavelengths, f, indent=2)

    for i, ts in enumerate(timestamps):
        scan = {
            "scan_index": i,
            "timestamp_utc": ts,
            "wavelengths_nm": wavelengths,
            "intensities": spectra[i].tolist()
        }
        with (run_dir / f"scan_{i:03d}.json").open("w", encoding="utf-8") as f:
            json.dump(scan, f, indent=2)

    print("保存完了。お疲れさまでした！")


if __name__ == "__main__":
    main()
