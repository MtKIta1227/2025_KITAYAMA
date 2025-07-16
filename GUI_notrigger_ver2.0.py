#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
pump_on_off_deltaA_with_timing_gui_v7.py
────────────────────────────────────────────────────────
元のスクリプトを PyQt5 ベースの GUI アプリケーションに変換（改訂版）。

【v7での変更点】
• 「生スペクトル」タブのプロット内容を変更。
• 参照光と試料光それぞれのポンプON/OFF時の強度スペクトルと
  吸光度スペクトルを表示（計4つのサブプロット）。
• 元々あった「生スペクトル」タブの（Probe/Refから計算した）
  吸光度プロットは削除。
"""

from __future__ import annotations

import concurrent.futures as cf
import json
import sys
import textwrap
import traceback
from datetime import datetime, timezone
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
from dateutil.parser import isoparse
from matplotlib.backends.backend_qt5agg import (
    FigureCanvasQTAgg as FigureCanvas,
    NavigationToolbar2QT as NavigationToolbar,
)
from PyQt5.QtCore import QObject, QThread, pyqtSignal
from PyQt5.QtWidgets import (
    QApplication,
    QDoubleSpinBox,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QSpinBox,
    QTabWidget,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)
from seabreeze.spectrometers import Spectrometer, list_devices

# ───────────── 初期設定 ─────────────
N_SCANS_DEFAULT = 200
INTEG_MS_DEFAULT = 5.0
SAVE_DIR = None
# ────────────────────────────────────


def now_iso() -> str:
    """UTC ISO-8601 文字列 (µs 精度) を返す"""
    return datetime.now(timezone.utc).isoformat(timespec="microseconds")


class AcquisitionWorker(QObject):
    """分光器データ取得を担うワーカースレッド"""
    progress = pyqtSignal(int, str)
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)

    def __init__(self, specs: dict, n_scans: int, label: str):
        super().__init__()
        self.spec_ref = specs["ref"]
        self.spec_probe = specs["probe"]
        self.n_scans = n_scans
        self.label = label
        self.is_running = True

    def _grab_once(self, spec: Spectrometer) -> tuple:
        """1 スキャン取得して (timestamp, intensities) を返す"""
        ts = now_iso()
        return ts, spec.intensities().astype(np.float32)

    def run(self):
        """N_SCANS 分のデータを取得し、各種平均値を計算する"""
        try:
            wl = self.spec_ref.wavelengths()
            a_stack = np.empty((self.n_scans, wl.size), dtype=np.float32)
            i_ref_stack = np.empty_like(a_stack)
            i_prb_stack = np.empty_like(a_stack)
            ts_ref, ts_prb = [], []

            with cf.ThreadPoolExecutor(max_workers=2) as exe:
                for i in range(self.n_scans):
                    if not self.is_running:
                        break
                    fut_ref = exe.submit(self._grab_once, self.spec_ref)
                    fut_prb = exe.submit(self._grab_once, self.spec_probe)
                    ts_r, i_ref = fut_ref.result()
                    ts_p, i_prb = fut_prb.result()
                    ts_ref.append(ts_r)
                    ts_prb.append(ts_p)
                    i_ref_stack.append(i_ref)
                    i_prb_stack.append(i_prb)
                    with np.errstate(divide="ignore", invalid="ignore"):
                        a_stack.append(-np.log10(i_prb / i_ref))
                    self.progress.emit(i + 1, self.label)

            if not self.is_running:
                return

            result = {
                "label": self.label,
                "wavelengths": wl,
                "A_mean": np.mean(a_stack, axis=0).astype(np.float32),
                "I_ref_mean": np.mean(i_ref_stack, axis=0).astype(np.float32),
                "I_probe_mean": np.mean(i_prb_stack, axis=0).astype(np.float32),
                "ts_ref_list": ts_ref,
                "ts_probe_list": ts_prb,
            }
            self.finished.emit(result)
        except Exception:
            self.error.emit(f"測定中にエラーが発生しました:\n{traceback.format_exc()}")

    def stop(self):
        self.is_running = False


class MainWindow(QMainWindow):
    """メインウィンドウクラス"""
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Pump-Probe ΔA Measurement v7.0 (Raw Abs)")
        self.setGeometry(100, 100, 1200, 900)

        # --- 状態・データ管理用変数 ---
        self.spectrometers = {}
        self.worker = None; self.thread = None
        self.wl = None
        self.a_off = self.a_on = None
        self.i_ref_off = self.i_ref_on = None
        self.i_prb_off = self.i_prb_on = None
        self.ts_off_ref = self.ts_off_prb = None
        self.ts_on_ref = self.ts_on_prb = None
        self.dt_off = self.dt_on = None

        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QHBoxLayout(main_widget)

        # === 左パネル: 設定と制御 ===
        left_panel = QVBoxLayout()
        main_layout.addLayout(left_panel, 1)

        settings_group = QGroupBox("設定")
        settings_layout = QFormLayout()
        self.n_scans_spinbox = QSpinBox()
        self.n_scans_spinbox.setRange(1, 10000); self.n_scans_spinbox.setValue(N_SCANS_DEFAULT)
        self.integ_ms_spinbox = QDoubleSpinBox()
        self.integ_ms_spinbox.setRange(1.0, 10000.0); self.integ_ms_spinbox.setValue(INTEG_MS_DEFAULT)
        self.integ_ms_spinbox.setDecimals(1); self.integ_ms_spinbox.setSingleStep(0.1)
        settings_layout.addRow("スキャン回数 (N_SCANS):", self.n_scans_spinbox)
        settings_layout.addRow("積分時間 (ms):", self.integ_ms_spinbox)
        settings_group.setLayout(settings_layout)
        left_panel.addWidget(settings_group)

        control_group = QGroupBox("制御")
        control_layout = QVBoxLayout()
        self.init_button = QPushButton("1. 分光器を初期化")
        self.reset_button = QPushButton("接続をリセット")
        self.pump_off_button = QPushButton("2. ポンプ OFF 測定を開始")
        self.pump_on_button = QPushButton("3. ポンプ ON 測定を開始")
        self.progress_bar = QProgressBar()
        control_layout.addWidget(self.init_button); control_layout.addWidget(self.pump_off_button)
        control_layout.addWidget(self.pump_on_button); control_layout.addWidget(self.reset_button)
        control_layout.addWidget(self.progress_bar)
        control_group.setLayout(control_layout)
        left_panel.addWidget(control_group)

        log_group = QGroupBox("ログ")
        log_layout = QVBoxLayout()
        self.log_text = QTextEdit(); self.log_text.setReadOnly(True)
        log_layout.addWidget(self.log_text)
        log_group.setLayout(log_layout)
        left_panel.addWidget(log_group, 1)

        # === 右パネル: プロットとレポート ===
        right_panel = QVBoxLayout()
        main_layout.addLayout(right_panel, 3)
        self.tabs = QTabWidget()

        # --- ΔAプロットタブ ---
        plot_widget_a = QWidget(); plot_layout_a = QVBoxLayout(plot_widget_a)
        self.fig_a, self.ax_a = plt.subplots()
        self.canvas_a = FigureCanvas(self.fig_a); toolbar_a = NavigationToolbar(self.canvas_a, self)
        plot_layout_a.addWidget(toolbar_a); plot_layout_a.addWidget(self.canvas_a)
        self.tabs.addTab(plot_widget_a, "ΔA スペクトル")

        # --- Δtプロットタブ ---
        plot_widget_t = QWidget(); plot_layout_t = QVBoxLayout(plot_widget_t)
        self.fig_t, (self.ax_t1, self.ax_t2) = plt.subplots(2, 1, tight_layout=True)
        self.canvas_t = FigureCanvas(self.fig_t); toolbar_t = NavigationToolbar(self.canvas_t, self)
        plot_layout_t.addWidget(toolbar_t); plot_layout_t.addWidget(self.canvas_t)
        self.tabs.addTab(plot_widget_t, "タイミング (Δt)")

        # --- 生スペクトルタブ (4行1列) ---
        plot_widget_raw = QWidget(); plot_layout_raw = QVBoxLayout(plot_widget_raw)
        self.fig_raw, self.axes_raw = plt.subplots(4, 1, sharex=True, tight_layout=True, figsize=(8, 12))
        self.canvas_raw = FigureCanvas(self.fig_raw); toolbar_raw = NavigationToolbar(self.canvas_raw, self)
        plot_layout_raw.addWidget(toolbar_raw); plot_layout_raw.addWidget(self.canvas_raw)
        self.tabs.addTab(plot_widget_raw, "生スペクトル")

        # --- レポートタブ ---
        self.report_text = QTextEdit(); self.report_text.setReadOnly(True)
        self.report_text.setFontFamily("monospace")
        self.tabs.addTab(self.report_text, "レポート")

        right_panel.addWidget(self.tabs)

        # --- シグナルとスロットの接続 ---
        self.init_button.clicked.connect(self.initialize_spectrometers)
        self.pump_off_button.clicked.connect(self.start_off_measurement)
        self.pump_on_button.clicked.connect(self.start_on_measurement)
        self.reset_button.clicked.connect(self.reset_connection)

        self._set_ui_state("initial"); self._log("アプリケーションを開始しました。")

    def _set_ui_state(self, state: str):
        """UIのボタン等の有効/無効を状態に応じて切り替える"""
        states = {
            "initial": (True, False, False, False, True, True),
            "ready_for_off": (False, True, True, False, True, True),
            "measuring": (False, False, False, False, False, False),
            "ready_for_on": (False, True, False, True, False, False),
        }
        s = states.get(state, (True, False, False, False, True, True))
        self.init_button.setEnabled(s["initial"]); self.reset_button.setEnabled(s["reset"])
        self.pump_off_button.setEnabled(s["off"]); self.pump_on_button.setEnabled(s["on"])
        self.n_scans_spinbox.setEnabled(s["n_scans"]); self.integ_ms_spinbox.setEnabled(s["integ_ms"])
        if state == "initial": self.progress_bar.setValue(0)

    def _log(self, message: str):
        """ログウィンドウにタイムスタンプ付きでメッセージを追記する"""
        self.log_text.append(f"[{datetime.now().strftime('%H:%M:%S')}] {message}")

    def initialize_spectrometers(self):
        """分光器を検索し、接続を確立する"""
        try:
            self._log("分光器を検索中...")
            QApplication.processEvents()
            devs = list_devices()
            if len(devs) < 2: raise RuntimeError("分光器が2台検出されません。")
            self.spectrometers["ref"] = Spectrometer.from_serial_number(devs["ref"].serial_number)
            self.spectrometers["probe"] = Spectrometer.from_serial_number(devs["probe"].serial_number)
            self._log(f"参照用分光器: {devs['ref'].serial_number}")
            self._log(f"試料用分光器: {devs['probe'].serial_number}")
            self._set_ui_state("ready_for_off"); self._log("初期化完了。ポンプOFF測定待機中...")
            self.tabs.setCurrentIndex(0); self.report_text.clear()
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"初期化に失敗しました:\n{e}"); self.reset_connection()

    def start_off_measurement(self):
        """ポンプOFF測定を開始する"""
        try:
            integ_us = round(self.integ_ms_spinbox.value() * 1_000)
            for s in self.spectrometers.values(): s.integration_time_micros(integ_us)
            self._log(f"積分時間を {self.integ_ms_spinbox.value():.1f} ms に設定しました。")
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"積分時間の設定に失敗しました:\n{e}"); return
        reply = QMessageBox.question(self, "確認", "ポンプ光が **OFF** になっていることを確認してください。\n測定を開始しますか？",
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes: self._log("ポンプOFF測定を開始します。"); self._start_acquisition("OFF")

    def start_on_measurement(self):
        """ポンプON測定を開始する"""
        reply = QMessageBox.question(self, "確認", "ポンプ光を **ON** に切り替えてください。\n測定を開始しますか？",
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes: self._log("ポンプON測定を開始します。"); self._start_acquisition("ON")

    def _start_acquisition(self, label: str):
        """ワーカースレッドを起動してデータ取得を開始する"""
        self._set_ui_state("measuring")
        n_scans = self.n_scans_spinbox.value()
        self.progress_bar.setMaximum(n_scans)
        self.thread = QThread(); self.worker = AcquisitionWorker(self.spectrometers, n_scans, label)
        self.worker.moveToThread(self.thread); self.thread.started.connect(self.worker.run)
        self.worker.finished.connect(self.on_acquisition_finished)
        self.worker.error.connect(self.on_acquisition_error)
        self.worker.progress.connect(self.update_progress)
        self.thread.start()

    def update_progress(self, count, label):
        """プログレスバーとログを更新する"""
        if count % 10 == 0 or count == self.n_scans_spinbox.value():
            self._log(f"[{label}] {count}/{self.n_scans_spinbox.value()} ショット完了")
        self.progress_bar.setValue(count)

    def on_acquisition_error(self, error_msg):
        """データ取得中にエラーが発生した場合の処理"""
        QMessageBox.critical(self, "測定エラー", error_msg); self._log(error_msg)
        self.thread.quit(); self.thread.wait(); self._set_ui_state("ready_for_off")

    def on_acquisition_finished(self, result):
        """データ取得が完了した際の処理"""
        label = result["label"]
        if label == "OFF":
            self.wl = result["wavelengths"]; self.a_off = result["A_mean"]
            self.ts_off_ref = result["ts_ref_list"]; self.ts_off_prb = result["ts_probe_list"]
            self.i_ref_off = result["I_ref_mean"]; self.i_prb_off = result["I_probe_mean"]
            self._set_ui_state("ready_for_on"); self._log("ポンプOFF測定完了。ポンプON測定待機中...")
        elif label == "ON":
            self.a_on = result["A_mean"]
            self.ts_on_ref = result["ts_ref_list"]; self.ts_on_prb = result["ts_probe_list"]
            self.i_ref_on = result["I_ref_mean"]; self.i_prb_on = result["I_probe_mean"]
            self._log("ポンプON測定完了。"); self.process_and_display_results()
        self.thread.quit(); self.thread.wait()

    def process_and_display_results(self):
        """すべてのデータが揃った後、計算、プロット、保存を行う"""
        self._log("最終結果を計算・表示しています...")

        def calc_dt(ts_prb_list, ts_ref_list):
            t_prb = np.array([isoparse(t) for t in ts_prb_list])
            t_ref = np.array([isoparse(t) for t in ts_ref_list])
            return (t_prb - t_ref).astype('timedelta64["us"]').astype(float) / 1000

        self.dt_off = calc_dt(self.ts_off_prb, self.ts_off_ref)
        self.dt_on = calc_dt(self.ts_on_prb, self.ts_on_ref)

        root = Path("raw_data"); root.mkdir(exist_ok=True)
        run_dir = root / datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S_%fZ")
        run_dir.mkdir(exist_ok=False); self._log(f"結果を {run_dir.resolve()} に保存します。")

        self.plot_results()
        self.save_files(run_dir)
        report = self.generate_report(run_dir)
        self.report_text.setText(report)
        (run_dir / "deltaA_report.txt").write_text(report + "\n", encoding="utf-8")
        self._log("=" * 20 + "\n" + report + "\n" + "=" * 20)
        self._log("すべての処理が完了しました。")
        self._set_ui_state("ready_for_off"); self.tabs.setCurrentIndex(2)  # 生スペクトルタブをデフォルト表示

    def save_files(self, run_dir: Path):
        """測定結果をファイルに保存する"""
        delta_a_conventional = self.a_on - self.a_off
        with np.errstate(divide="ignore", invalid="ignore"):
            delta_a_probe = -np.log10(self.i_prb_on / self.i_prb_off)
            delta_a_ref = -np.log10(self.i_ref_on / self.i_ref_off)

        np.savez(run_dir / "full_data.npz",
                 wavelengths_nm=self.wl, dt_on=self.dt_on, dt_off=self.dt_off,
                 A_on=self.a_on, A_off=self.a_off,
                 I_ref_on=self.i_ref_on, I_ref_off=self.i_ref_off,
                 I_prb_on=self.i_prb_on, I_prb_off=self.i_prb_off,
                 delta_A_conventional=delta_a_conventional,
                 delta_A_probe=delta_a_probe,
                 delta_A_ref=delta_a_ref)
        self.fig_a.savefig(run_dir / "deltaA_plot.png", dpi=150)
        self.fig_t.savefig(run_dir / "timing_plot.png", dpi=150)
        self.fig_raw.savefig(run_dir / "raw_spectra_plot.png", dpi=150)
        self._log("NPZ と PNG ファイルを保存しました。")

    def plot_results(self):
        """すべてのグラフを更新する"""
        # --- ΔAプロット ---
        self.ax_a.clear()
        with np.errstate(divide="ignore", invalid="ignore"):
            delta_a_conventional = self.a_on - self.a_off
            delta_a_probe = -np.log10(self.i_prb_on / self.i_prb_off)
            delta_a_ref = -np.log10(self.i_ref_on / self.i_ref_off)

        self.ax_a.plot(self.wl, delta_a_probe, label="ΔA_probe", lw=2)
        self.ax_a.plot(self.wl, delta_a_ref, label="ΔA_ref", lw=2)
        self.ax_a.plot(self.wl, delta_a_conventional, label="ΔA_conv", ls='--', color='gray')

        self.ax_a.set_xlabel("Wavelength (nm)"); self.ax_a.set_ylabel("Absorbance Change (ΔA)")
        self.ax_a.set_title("Pump-Induced Absorbance Change")
        self.ax_a.grid(alpha=0.3); self.ax_a.legend(loc="best"); self.canvas_a.draw()

        # --- Δtプロット ---
        self.ax_t1.clear(); self.ax_t2.clear()
        self.ax_t1.plot(self.dt_off, ".-", label="OFF"); self.ax_t1.plot(self.dt_on, ".-", label="ON")
        self.ax_t1.set_ylabel("Δt (ms)\nProbe − Ref"); self.ax_t1.set_title("Start-time lag per shot")
        self.ax_t1.grid(alpha=0.3); self.ax_t1.legend()
        all_dt = np.concatenate([self.dt_off, self.dt_on])
        self.ax_t2.hist(all_dt, bins=30)
        self.ax_t2.set_xlabel("Δt (ms)"); self.ax_t2.set_ylabel("Count")
        self.fig_t.tight_layout(); self.canvas_t.draw()

        # --- 生スペクトルプロット (4段) ---
        ax_ref_int, ax_prb_int, ax_ref_abs, ax_prb_abs = self.axes_raw
        ax_ref_int.clear(); ax_prb_int.clear(); ax_ref_abs.clear(); ax_prb_abs.clear()

        ax_ref_int.plot(self.wl, self.i_ref_off, label="Pump OFF")
        ax_ref_int.plot(self.wl, self.i_ref_on, label="Pump ON")
        ax_ref_int.set_title("Ref Intensity"); ax_ref_int.set_ylabel("Intensity (counts)")
        ax_ref_int.legend(); ax_ref_int.grid(alpha=0.3)

        ax_prb_int.plot(self.wl, self.i_prb_off, label="Pump OFF")
        ax_prb_int.plot(self.wl, self.i_prb_on, label="Pump ON")
        ax_prb_int.set_title("Probe Intensity"); ax_prb_int.set_ylabel("Intensity (counts)")
        ax_prb_int.legend(); ax_prb_int.grid(alpha=0.3)

        with np.errstate(divide="ignore", invalid="ignore"):
            ref_abs_off = -np.log10(self.i_ref_off)
            ref_abs_on = -np.log10(self.i_ref_on)
            prb_abs_off = -np.log10(self.i_prb_off)
            prb_abs_on = -np.log10(self.i_prb_on)

        ax_ref_abs.plot(self.wl, ref_abs_off, label="Pump OFF")
        ax_ref_abs.plot(self.wl, ref_abs_on, label="Pump ON")
        ax_ref_abs.set_title("Ref Absorbance"); ax_ref_abs.set_ylabel("Absorbance (A.U.)")
        ax_ref_abs.legend(); ax_ref_abs.grid(alpha=0.3)

        ax_prb_abs.plot(self.wl, prb_abs_off, label="Pump OFF")
        ax_prb_abs.plot(self.wl, prb_abs_on, label="Pump ON")
        ax_prb_abs.set_title("Probe Absorbance"); ax_prb_abs.set_xlabel("Wavelength (nm)")
        ax_prb_abs.set_ylabel("Absorbance (A.U.)")
        ax_prb_abs.legend(); ax_prb_abs.grid(alpha=0.3)

        self.fig_raw.tight_layout(); self.canvas_raw.draw()
        self._log("すべてのグラフを更新しました。")

    def generate_report(self, run_dir: Path) -> str:
        """テキストレポートを生成する"""
        return textwrap.dedent(f"""
            ▼ Pump ON / OFF ΔA ＆ Timing レポート
            --------------------------------------------
            ショット数 (各状態) : {self.n_scans_spinbox.value()}
            積分時間            : {self.integ_ms_spinbox.value():.3f} ms
            Δt 平均 ±SD (OFF)  : {self.dt_off.mean():+.4f} ± {self.dt_off.std(ddof=1):.4f} ms
            Δt 平均 ±SD (ON)   : {self.dt_on.mean():+.4f} ± {self.dt_on.std(ddof=1):.4f} ms
            Spectrometer Ref   : {self.spectrometers['ref'].serial_number}
            Spectrometer Probe : {self.spectrometers['probe'].serial_number}
            保存先             : {run_dir.resolve()}
            """).strip()

    def reset_connection(self):
        """分光器の接続を解放し、UIを初期状態に戻す"""
        self._log("分光器の接続をリセットします...")
        if self.thread and self.thread.isRunning():
            self.worker.stop(); self.thread.quit(); self.thread.wait()
        for spec in self.spectrometers.values():
            try: spec.close()
            except Exception as e: print(f"分光器 {spec.serial_number} のクローズ中にエラー: {e}")
        self.spectrometers.clear()
        self._set_ui_state("initial")
        self.report_text.clear()
        all_axes = [self.ax_a, self.ax_t1, self.ax_t2, *self.axes_raw]
        for ax in all_axes: ax.clear()
        self.canvas_a.draw(); self.canvas_t.draw(); self.canvas_raw.draw()
        self._log("リセット完了。")

    def closeEvent(self, event):
        """アプリケーション終了時の処理"""
        self.reset_connection()
        event.accept()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_win = MainWindow()
    main_win.show()
    sys.exit(app.exec_())