#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
pump_on_off_deltaA_with_timing.py
────────────────────────────────────────────────────────
• 2 台の分光器 (A = 参照, B = 試料) を並列駆動
• ポンプ OFF → N_SCANS ショット取得して <A_off>
• ユーザーがポンプを ON に切替
• ポンプ ON  → N_SCANS ショット取得して <A_on>
• ΔA = <A_on> – <A_off> を計算・保存・プロット
• 各ショットで測定開始時刻差 Δt〔ms〕を計算し
  時系列＋ヒストグラムを保存 & ポップアップ表示
"""

from __future__ import annotations

import concurrent.futures as cf
import io
import json
import sys
import textwrap
from datetime import datetime, timezone
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
from dateutil.parser import isoparse
from seabreeze.spectrometers import Spectrometer, list_devices

# ───────────── ユーザー設定 ─────────────
N_SCANS  = 200       # ポンプ ON / OFF それぞれのショット数
INTEG_MS = 5.0       # 積分時間 [ms]
SAVE_DIR = None      # None → raw_data/日時フォルダを自動生成
# ────────────────────────────────────


def now_iso() -> str:
    """UTC ISO-8601 文字列 (µs 精度) を返す"""
    return datetime.now(timezone.utc).isoformat(timespec="microseconds")


def grab_once(spec: Spectrometer) -> tuple[str, np.ndarray]:
    """1 スキャン取得して (timestamp, intensities) を返す"""
    ts = now_iso()
    return ts, spec.intensities().astype(np.float32)


def acquire_state(
    spec_ref: Spectrometer,
    spec_probe: Spectrometer,
    label: str,
) -> tuple[np.ndarray, np.ndarray, list[str], list[str]]:
    """
    ポンプ状態 label(ON/OFF) で N_SCANS 取得。
    戻り値: (wavelengths, A_mean, ts_ref_list, ts_probe_list)
    """
    wl = spec_ref.wavelengths()
    a_stack = np.empty((N_SCANS, wl.size), dtype=np.float32)
    ts_ref, ts_prb = [], []

    with cf.ThreadPoolExecutor(max_workers=2) as exe:
        for i in range(N_SCANS):
            fut_ref = exe.submit(grab_once, spec_ref)
            fut_prb = exe.submit(grab_once, spec_probe)

            ts_r, i_ref = fut_ref.result()
            ts_p, i_prb = fut_prb.result()

            ts_ref.append(ts_r)
            ts_prb.append(ts_p)

            with np.errstate(divide="ignore"):
                a_stack[i] = -np.log10(i_prb / i_ref)

            if (i + 1) % 10 == 0 or i == N_SCANS - 1:
                print(f"[{label}] {i + 1}/{N_SCANS} shots")

    return wl, a_stack.mean(0), ts_ref, ts_prb


def fig_bytes(fig) -> bytes:
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150)
    return buf.getvalue()


def main() -> None:
    # 1) 分光器 2 台を取得
    devs = list_devices()
    if len(devs) < 2:
        sys.exit("⚠  分光器が 2 台検出されません。USB 接続を確認してください。")

    ref = Spectrometer.from_serial_number(devs[0].serial_number)   # 参照
    prb = Spectrometer.from_serial_number(devs[1].serial_number)   # 試料

    # 2) 積分時間設定 (Normal: trigger_mode = 0)
    integ_us = round(INTEG_MS * 1_000)
    for s in (ref, prb):
        s.integration_time_micros(integ_us)

    # 3) ポンプ OFF 測定
    input("▶ ポンプ光を **OFF** にして Enter を押してください …")
    wl, a_off, ts_off_ref, ts_off_prb = acquire_state(ref, prb, "OFF")

    # 4) ポンプ ON 測定
    input("▶ ポンプ光を **ON** にして Enter を押してください …")
    _,  a_on,  ts_on_ref,  ts_on_prb  = acquire_state(ref, prb, "ON")

    ref.close()
    prb.close()

    # 5) ΔA と Δt 計算
    delta_a = a_on - a_off

    t0 = isoparse(ts_off_ref[0])
    def iso_list_to_ms(lst: list[str]) -> np.ndarray:
        return np.array([(isoparse(x) - t0).total_seconds() * 1_000 for x in lst])

    dt_off = iso_list_to_ms(ts_off_prb) - iso_list_to_ms(ts_off_ref)
    dt_on  = iso_list_to_ms(ts_on_prb)  - iso_list_to_ms(ts_on_ref)

    # 6) 保存フォルダ作成 (raw_data/…)
    if SAVE_DIR:
        run_dir = Path(SAVE_DIR)
    else:
        root = Path("raw_data")
        root.mkdir(exist_ok=True)
        run_dir = root / datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S_%fZ")
    run_dir.mkdir(exist_ok=False)

    # 7) ファイル保存
    (run_dir / "deltaA_mean.json").write_text(
        json.dumps(
            {
                "wavelengths_nm": wl.tolist(),
                "deltaA": delta_a.tolist(),
                "A_on": a_on.tolist(),
                "A_off": a_off.tolist(),
            },
            indent=2,
        ),
        encoding="utf-8",
    )

    np.savez(
        run_dir / "timing_and_absorbance.npz",
        wavelengths_nm=wl,
        dt_on=dt_on,
        dt_off=dt_off,
        A_on=a_on,
        A_off=a_off,
    )

    # 8-1) ΔA スペクトルプロット
    fig_a, ax_a = plt.subplots(figsize=(8, 4))
    ax_a.plot(wl, delta_a, label="ΔA = A_on − A_off")
    ax_a.set_xlabel("Wavelength (nm)")
    ax_a.set_ylabel("ΔA")
    ax_a.set_title("Pump-Induced Absorbance Change")
    ax_a.grid(alpha=0.3)
    ax_a.legend(loc="best")
    (run_dir / "deltaA_plot.png").write_bytes(fig_bytes(fig_a))

    # 8-2) Δt プロット (時系列 + ヒストグラム)
    fig_t, (ax_t1, ax_t2) = plt.subplots(2, 1, figsize=(8, 6), tight_layout=True)
    ax_t1.plot(dt_off, ".-", label="OFF")
    ax_t1.plot(dt_on,  ".-", label="ON")
    ax_t1.set_ylabel("Δt  (ms)\nProbe − Ref")
    ax_t1.set_title("Start-time lag per shot")
    ax_t1.grid(alpha=0.3)
    ax_t1.legend()

    all_dt = np.concatenate([dt_off, dt_on])
    ax_t2.hist(all_dt, bins=30)
    ax_t2.set_xlabel("Δt  (ms)")
    ax_t2.set_ylabel("Count")
    ax_t2.set_title("Histogram of Δt (both states)")
    ax_t2.grid(alpha=0.3)

    (run_dir / "timing_plot.png").write_bytes(fig_bytes(fig_t))

    # 8-3) 画面表示
    plt.show()

    # 9) テキストレポート
    report = textwrap.dedent(
        f"""
        ▼ Pump ON / OFF ΔA ＆ Timing レポート
        --------------------------------------------
        ショット数 (各状態) : {N_SCANS}
        積分時間            : {INTEG_MS:.3f} ms
        Δt 平均 ±SD (OFF)  : {dt_off.mean():+.4f} ± {dt_off.std(ddof=1):.4f} ms
        Δt 平均 ±SD (ON)   : {dt_on.mean():+.4f} ± {dt_on.std(ddof=1):.4f} ms
        Spectrometer Ref   : {ref.serial_number}
        Spectrometer Probe : {prb.serial_number}
        保存先             : {run_dir.resolve()}
        """
    ).strip()

    (run_dir / "deltaA_report.txt").write_text(report + "\n", encoding="utf-8")
    print("\n" + report + "\n")


if __name__ == "__main__":
    main()
