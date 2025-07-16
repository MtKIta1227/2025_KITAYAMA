#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
pump_on_off_deltaA_with_timing_gui_v3.py
────────────────────────────────────────────────────────
元のスクリプトを PyQt5 ベースの GUI アプリケーションに変換（改訂版）。

[v3での改善点]
• 各グラフにズーム、パン、保存などが可能なナビゲーションツールバーを追加。
"""

from __future__ import annotations

import concurrent.futures as cf
import io
import json
import sys
import textwrap
import traceback
from datetime import datetime, timezone
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
from dateutil.parser import isoparse
# MatplotlibのPyQt5連携に必要なクラスをインポート
from matplotlib.backends.backend_qt5agg import (
    FigureCanvasQTAgg as FigureCanvas,
    NavigationToolbar2QT as NavigationToolbar,  # ツールバーのクラス
)
from PyQt5.QtCore import QObject, QThread, pyqtSignal
from PyQt5.QtWidgets import (
    QApplication,
    QDoubleSpinBox,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
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
    return datetime.now(timezone.utc).isoformat(timespec="microseconds")


class AcquisitionWorker(QObject):
    progress = pyqtSignal(int, str)
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)

    def __init__(self, specs: dict[str, Spectrometer], n_scans: int, label: str):
        super().__init__()
        self.spec_ref = specs["ref"]
        self.spec_probe = specs["probe"]
        self.n_scans = n_scans
        self.label = label
        self.is_running = True

    def _grab_once(self, spec: Spectrometer) -> tuple[str, np.ndarray]:
        ts = now_iso()
        return ts, spec.intensities().astype(np.float32)

    def run(self):
        try:
            wl = self.spec_ref.wavelengths()
            a_stack = np.empty((self.n_scans, wl.size), dtype=np.float32)
            ts_ref, ts_prb = [], []
            with cf.ThreadPoolExecutor(max_workers=2) as exe:
                for i in range(self.n_scans):
                    if not self.is_running: break
                    fut_ref = exe.submit(self._grab_once, self.spec_ref)
                    fut_prb = exe.submit(self._grab_once, self.spec_probe)
                    ts_r, i_ref = fut_ref.result()
                    ts_p, i_prb = fut_prb.result()
                    ts_ref.append(ts_r); ts_prb.append(ts_p)
                    with np.errstate(divide="ignore"):
                        a_stack[i] = -np.log10(i_prb / i_ref)
                    self.progress.emit(i + 1, self.label)
            if not self.is_running: return
            result = {"label": self.label, "wavelengths": wl, "A_mean": a_stack.mean(0),
                      "ts_ref_list": ts_ref, "ts_probe_list": ts_prb}
            self.finished.emit(result)
        except Exception:
            self.error.emit(f"測定中にエラーが発生しました:\n{traceback.format_exc()}")

    def stop(self):
        self.is_running = False


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Pump-Probe ΔA Measurement v3.0 (with Toolbar)")
        self.setGeometry(100, 100, 1200, 800)

        # ... (状態管理、データ保存用の変数は変更なし) ...
        self.spectrometers = {}
        self.worker = None
        self.thread = None
        self.wl = None
        self.a_off = self.a_on = None
        self.ts_off_ref = self.ts_off_prb = None
        self.ts_on_ref = self.ts_on_prb = None

        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QHBoxLayout(main_widget)

        # === 左パネル: 設定と制御 (変更なし) ===
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
        control_layout.addWidget(self.init_button)
        control_layout.addWidget(self.pump_off_button)
        control_layout.addWidget(self.pump_on_button)
        control_layout.addWidget(self.reset_button)
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
        
        # --- グラフタブの作成 (ツールバー追加) ---

        # ΔAプロットタブ
        plot_widget_a = QWidget()
        plot_layout_a = QVBoxLayout(plot_widget_a)
        self.fig_a, self.ax_a = plt.subplots()
        self.canvas_a = FigureCanvas(self.fig_a)
        toolbar_a = NavigationToolbar(self.canvas_a, self) # ツールバーを作成
        plot_layout_a.addWidget(toolbar_a)                  # ツールバーをレイアウトに追加
        plot_layout_a.addWidget(self.canvas_a)              # キャンバスをレイアウトに追加
        self.tabs.addTab(plot_widget_a, "ΔA スペクトル")

        # Δtプロットタブ
        plot_widget_t = QWidget()
        plot_layout_t = QVBoxLayout(plot_widget_t)
        self.fig_t, (self.ax_t1, self.ax_t2) = plt.subplots(2, 1, tight_layout=True)
        self.canvas_t = FigureCanvas(self.fig_t)
        toolbar_t = NavigationToolbar(self.canvas_t, self) # ツールバーを作成
        plot_layout_t.addWidget(toolbar_t)                  # ツールバーをレイアウトに追加
        plot_layout_t.addWidget(self.canvas_t)              # キャンバスをレイアウトに追加
        self.tabs.addTab(plot_widget_t, "タイミング (Δt)")

        # レポートタブ
        self.report_text = QTextEdit(); self.report_text.setReadOnly(True)
        self.report_text.setFontFamily("monospace")
        self.tabs.addTab(self.report_text, "レポート")
        
        right_panel.addWidget(self.tabs)
        
        # --- シグナルとスロットの接続 (変更なし) ---
        self.init_button.clicked.connect(self.initialize_spectrometers)
        self.pump_off_button.clicked.connect(self.start_off_measurement)
        self.pump_on_button.clicked.connect(self.start_on_measurement)
        self.reset_button.clicked.connect(self.reset_connection)

        self._set_ui_state("initial")
        self._log("アプリケーションを開始しました。")

    # === メソッド定義 (ここから下は変更なし) ===

    def _set_ui_state(self, state: str):
        if state == "initial":
            self.init_button.setEnabled(True)
            self.reset_button.setEnabled(False)
            self.pump_off_button.setEnabled(False)
            self.pump_on_button.setEnabled(False)
            self.n_scans_spinbox.setEnabled(True)
            self.integ_ms_spinbox.setEnabled(True)
            self.progress_bar.setValue(0)
        elif state == "ready_for_off":
            self.init_button.setEnabled(False)
            self.reset_button.setEnabled(True)
            self.pump_off_button.setEnabled(True)
            self.pump_on_button.setEnabled(False)
            self.n_scans_spinbox.setEnabled(True)
            self.integ_ms_spinbox.setEnabled(True)
        elif state == "measuring":
            self.init_button.setEnabled(False)
            self.reset_button.setEnabled(False)
            self.pump_off_button.setEnabled(False)
            self.pump_on_button.setEnabled(False)
            self.n_scans_spinbox.setEnabled(False)
            self.integ_ms_spinbox.setEnabled(False)
        elif state == "ready_for_on":
            self.init_button.setEnabled(False)
            self.reset_button.setEnabled(True)
            self.pump_off_button.setEnabled(False)
            self.pump_on_button.setEnabled(True)
            self.n_scans_spinbox.setEnabled(False)
            self.integ_ms_spinbox.setEnabled(False)
    
    def _log(self, message: str):
        self.log_text.append(f"[{datetime.now().strftime('%H:%M:%S')}] {message}")

    def initialize_spectrometers(self):
        try:
            self._log("分光器を検索中...")
            QApplication.processEvents()
            devs = list_devices()
            if len(devs) < 2:
                raise RuntimeError("分光器が2台検出されません。USB接続を確認してください。")
            self.spectrometers["ref"] = Spectrometer.from_serial_number(devs[0].serial_number)
            self.spectrometers["probe"] = Spectrometer.from_serial_number(devs[1].serial_number)
            self._log(f"参照用分光器: {devs[0].serial_number}")
            self._log(f"試料用分光器: {devs[1].serial_number}")
            self._set_ui_state("ready_for_off")
            self._log("初期化完了。ポンプOFF測定待機中...")
            self.tabs.setCurrentIndex(0)
            self.report_text.clear()
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"初期化に失敗しました:\n{e}")
            self._log(f"エラー: 初期化失敗 - {e}")
            self.reset_connection()

    def start_off_measurement(self):
        try:
            integ_us = round(self.integ_ms_spinbox.value() * 1_000)
            for s in self.spectrometers.values():
                s.integration_time_micros(integ_us)
            self._log(f"積分時間を {self.integ_ms_spinbox.value():.1f} ms に設定しました。")
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"積分時間の設定に失敗しました:\n{e}")
            return
        reply = QMessageBox.question(self, "確認", "ポンプ光が **OFF** になっていることを確認してください。\n測定を開始しますか？",
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            self._log("ポンプOFF測定を開始します。")
            self._start_acquisition("OFF")

    def start_on_measurement(self):
        reply = QMessageBox.question(self, "確認", "ポンプ光を **ON** に切り替えてください。\n測定を開始しますか？",
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            self._log("ポンプON測定を開始します。")
            self._start_acquisition("ON")

    def _start_acquisition(self, label: str):
        self._set_ui_state("measuring")
        n_scans = self.n_scans_spinbox.value()
        self.progress_bar.setMaximum(n_scans)
        self.thread = QThread()
        self.worker = AcquisitionWorker(self.spectrometers, n_scans, label)
        self.worker.moveToThread(self.thread)
        self.thread.started.connect(self.worker.run)
        self.worker.finished.connect(self.on_acquisition_finished)
        self.worker.error.connect(self.on_acquisition_error)
        self.worker.progress.connect(self.update_progress)
        self.thread.start()
        
    def update_progress(self, count, label):
        self.progress_bar.setValue(count)
        if count % 10 == 0 or count == self.n_scans_spinbox.value():
            self._log(f"[{label}] {count}/{self.n_scans_spinbox.value()} ショット完了")

    def on_acquisition_error(self, error_msg):
        QMessageBox.critical(self, "測定エラー", error_msg)
        self._log(error_msg)
        self.thread.quit()
        self.thread.wait()
        self._set_ui_state("ready_for_off")

    def on_acquisition_finished(self, result):
        label = result["label"]
        if label == "OFF":
            self.wl = result["wavelengths"]
            self.a_off = result["A_mean"]
            self.ts_off_ref = result["ts_ref_list"]
            self.ts_off_prb = result["ts_probe_list"]
            self._set_ui_state("ready_for_on")
            self._log("ポンプOFF測定完了。ポンプON測定待機中...")
        elif label == "ON":
            self.a_on = result["A_mean"]
            self.ts_on_ref = result["ts_ref_list"]
            self.ts_on_prb = result["ts_probe_list"]
            self._log("ポンプON測定完了。")
            self.process_and_display_results()
        self.thread.quit()
        self.thread.wait()

    def process_and_display_results(self):
        self._log("最終結果を計算・表示しています...")
        delta_a = self.a_on - self.a_off
        t0 = isoparse(self.ts_off_ref[0])
        def iso_list_to_ms(lst: list[str]) -> np.ndarray:
            return np.array([(isoparse(x) - t0).total_seconds() * 1_000 for x in lst])
        dt_off = iso_list_to_ms(self.ts_off_prb) - iso_list_to_ms(self.ts_off_ref)
        dt_on = iso_list_to_ms(self.ts_on_prb) - iso_list_to_ms(self.ts_on_ref)
        root = Path("raw_data")
        root.mkdir(exist_ok=True)
        run_dir = root / datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S_%fZ")
        run_dir.mkdir(exist_ok=False)
        self._log(f"結果を {run_dir.resolve()} に保存します。")
        self.plot_results(delta_a, dt_off, dt_on)
        self.save_files(run_dir, delta_a, dt_off, dt_on)
        report = self.generate_report(run_dir, dt_off, dt_on)
        self.report_text.setText(report)
        (run_dir / "deltaA_report.txt").write_text(report + "\n", encoding="utf-8")
        self._log("="*20 + "\n" + report + "\n" + "="*20)
        self._log("すべての処理が完了しました。パラメータを変更して次の測定を開始できます。")
        self._set_ui_state("ready_for_off")
        self.tabs.setCurrentIndex(2) 

    def save_files(self, run_dir, delta_a, dt_off, dt_on):
        (run_dir / "deltaA_mean.json").write_text(json.dumps({
                "wavelengths_nm": self.wl.tolist(), "deltaA": delta_a.tolist(),
                "A_on": self.a_on.tolist(), "A_off": self.a_off.tolist()}, indent=2), encoding="utf-8")
        np.savez(run_dir / "timing_and_absorbance.npz", wavelengths_nm=self.wl,
                 dt_on=dt_on, dt_off=dt_off, A_on=self.a_on, A_off=self.a_off)
        # ツールバーの保存機能と重複するが、自動保存のために残しておく
        self.fig_a.savefig(run_dir / "deltaA_plot.png", dpi=150)
        self.fig_t.savefig(run_dir / "timing_plot.png", dpi=150)
        self._log("JSON, NPZ, PNG ファイルを保存しました。")
        
    def plot_results(self, delta_a, dt_off, dt_on):
        self.ax_a.clear()
        self.ax_a.plot(self.wl, delta_a, label="ΔA = A_on − A_off")
        self.ax_a.set_xlabel("Wavelength (nm)"); self.ax_a.set_ylabel("ΔA")
        self.ax_a.set_title("Pump-Induced Absorbance Change")
        self.ax_a.grid(alpha=0.3); self.ax_a.legend(loc="best")
        self.canvas_a.draw()
        self.ax_t1.clear(); self.ax_t2.clear()
        self.ax_t1.plot(dt_off, ".-", label="OFF"); self.ax_t1.plot(dt_on, ".-", label="ON")
        self.ax_t1.set_ylabel("Δt (ms)\nProbe − Ref"); self.ax_t1.set_title("Start-time lag per shot")
        self.ax_t1.grid(alpha=0.3); self.ax_t1.legend()
        all_dt = np.concatenate([dt_off, dt_on])
        self.ax_t2.hist(all_dt, bins=30)
        self.ax_t2.set_xlabel("Δt (ms)"); self.ax_t2.set_ylabel("Count")
        self.ax_t2.set_title("Histogram of Δt (both states)"); self.ax_t2.grid(alpha=0.3)
        self.fig_t.tight_layout(); self.canvas_t.draw()
        self._log("グラフを更新しました。")

    def generate_report(self, run_dir, dt_off, dt_on):
        return textwrap.dedent(f"""
            ▼ Pump ON / OFF ΔA ＆ Timing レポート
            --------------------------------------------
            ショット数 (各状態) : {self.n_scans_spinbox.value()}
            積分時間            : {self.integ_ms_spinbox.value():.3f} ms
            Δt 平均 ±SD (OFF)  : {dt_off.mean():+.4f} ± {dt_off.std(ddof=1):.4f} ms
            Δt 平均 ±SD (ON)   : {dt_on.mean():+.4f} ± {dt_on.std(ddof=1):.4f} ms
            Spectrometer Ref   : {self.spectrometers['ref'].serial_number}
            Spectrometer Probe : {self.spectrometers['probe'].serial_number}
            保存先             : {run_dir.resolve()}
            """).strip()

    def reset_connection(self):
        self._log("分光器の接続をリセットします...")
        if self.thread and self.thread.isRunning():
             self.worker.stop(); self.thread.quit(); self.thread.wait()
        for spec in self.spectrometers.values():
            try: spec.close()
            except Exception as e: print(f"分光器 {spec.serial_number} のクローズ中にエラー: {e}")
        self.spectrometers.clear()
        self._set_ui_state("initial")
        self.report_text.clear()
        for ax in [self.ax_a, self.ax_t1, self.ax_t2]: ax.clear()
        self.canvas_a.draw(); self.canvas_t.draw()
        self._log("リセット完了。再度「分光器を初期化」から始めてください。")

    def closeEvent(self, event):
        self.reset_connection()
        event.accept()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_win = MainWindow()
    main_win.show()
    sys.exit(app.exec_())