import sys
import numpy as np
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QTextEdit,
    QDoubleSpinBox, QCheckBox
)
from PyQt5.QtCore import QTimer
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.backends.backend_qt5agg import NavigationToolbar2QT as NavigationToolbar
from matplotlib.figure import Figure

try:
    from seabreeze.spectrometers import Spectrometer
except ImportError:
    Spectrometer = None

class SimpleUSB4000GUI(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("USB4000リアルタイム測定ミニ")
        self.setGeometry(100, 100, 850, 560)
        self.spectrometer = None
        self.dark_spectrum = None
        self.timer = QTimer()
        self.timer.timeout.connect(self.auto_acquire)

        # --- GUIボタン ---
        self.connect_btn = QPushButton("接続")
        self.connect_btn.clicked.connect(self.connect_spectrometer)
        self.dark_btn = QPushButton("ダーク測定更新")
        self.dark_btn.clicked.connect(self.update_dark)
        self.start_btn = QPushButton("リアルタイム測定ON")
        self.start_btn.setCheckable(True)
        self.start_btn.clicked.connect(self.toggle_auto)
        self.reset_btn = QPushButton("グラフリセット")
        self.reset_btn.clicked.connect(self.reset_graph)
        self.status_label = QLabel("ステータス: 未接続")
        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setMaximumHeight(60)

        # --- X軸・Y軸コントロール ---
        self.xrange_chk = QCheckBox("波長範囲を固定")
        self.xmin_box = QDoubleSpinBox()
        self.xmin_box.setRange(0, 2000)
        self.xmin_box.setDecimals(1)
        self.xmin_box.setValue(400.0)
        self.xmax_box = QDoubleSpinBox()
        self.xmax_box.setRange(0, 2000)
        self.xmax_box.setDecimals(1)
        self.xmax_box.setValue(700.0)
        xrange_h = QHBoxLayout()
        xrange_h.addWidget(self.xrange_chk)
        xrange_h.addWidget(QLabel("最小:"))
        xrange_h.addWidget(self.xmin_box)
        xrange_h.addWidget(QLabel("最大:"))
        xrange_h.addWidget(self.xmax_box)

        self.yrange_chk = QCheckBox("強度範囲を固定")
        self.ymin_box = QDoubleSpinBox()
        self.ymin_box.setRange(-1e6, 1e6)
        self.ymin_box.setDecimals(1)
        self.ymin_box.setValue(0.0)
        self.ymax_box = QDoubleSpinBox()
        self.ymax_box.setRange(-1e6, 1e6)
        self.ymax_box.setDecimals(1)
        self.ymax_box.setValue(10000.0)
        yrange_h = QHBoxLayout()
        yrange_h.addWidget(self.yrange_chk)
        yrange_h.addWidget(QLabel("最小:"))
        yrange_h.addWidget(self.ymin_box)
        yrange_h.addWidget(QLabel("最大:"))
        yrange_h.addWidget(self.ymax_box)

        # --- グラフ ---
        self.figure = Figure()
        self.canvas = FigureCanvas(self.figure)
        self.ax = self.figure.add_subplot(111)
        self.ax.set_xlabel("Wavelength (nm)")
        self.ax.set_ylabel("Intensity")
        self.toolbar = NavigationToolbar(self.canvas, self)

        # --- レイアウト ---
        btnrow = QHBoxLayout()
        btnrow.addWidget(self.connect_btn)
        btnrow.addWidget(self.dark_btn)
        btnrow.addWidget(self.start_btn)
        btnrow.addWidget(self.reset_btn)
        vbox = QVBoxLayout()
        vbox.addLayout(btnrow)
        vbox.addLayout(xrange_h)
        vbox.addLayout(yrange_h)
        vbox.addWidget(self.toolbar)
        vbox.addWidget(self.canvas)
        vbox.addWidget(self.status_label)
        vbox.addWidget(self.log_box)
        self.setLayout(vbox)

    def connect_spectrometer(self):
        try:
            self.spectrometer = Spectrometer.from_first_available()
            self.status_label.setText("ステータス: 接続済み")
            self.log_box.append("分光器接続完了")
        except Exception as e:
            self.status_label.setText("ステータス: 接続失敗")
            self.log_box.append(f"接続エラー: {e}")

    def update_dark(self):
        if not self.spectrometer:
            self.log_box.append("分光器が接続されていません")
            return
        self.dark_spectrum = self.spectrometer.intensities()
        self.log_box.append("ダークスペクトルを更新しました")
        self.status_label.setText("ステータス: ダーク更新完了")

    def toggle_auto(self):
        if self.start_btn.isChecked():
            self.timer.start(50)  # 50msごと
            self.start_btn.setText("リアルタイム測定OFF")
            self.log_box.append("リアルタイム測定ON")
        else:
            self.timer.stop()
            self.start_btn.setText("リアルタイム測定ON")
            self.log_box.append("リアルタイム測定OFF")

    def auto_acquire(self):
        if not self.spectrometer or self.dark_spectrum is None:
            return
        intensities = self.spectrometer.intensities() - self.dark_spectrum
        wavelengths = self.spectrometer.wavelengths()
        self.ax.clear()
        self.ax.plot(wavelengths, intensities, label="Current (Dark-corrected)")
        self.ax.set_xlabel("Wavelength (nm)")
        self.ax.set_ylabel("Intensity")
        self.ax.legend()

        # X軸固定
        if self.xrange_chk.isChecked():
            self.ax.set_xlim(self.xmin_box.value(), self.xmax_box.value())
        else:
            self.ax.autoscale(enable=True, axis='x', tight=True)
        # Y軸固定
        if self.yrange_chk.isChecked():
            self.ax.set_ylim(self.ymin_box.value(), self.ymax_box.value())
        else:
            if len(intensities) > 0:
                y_min = np.nanmin(intensities)
                y_max = np.nanmax(intensities)
                margin = (y_max - y_min) * 0.05 if y_max > y_min else 1
                self.ax.set_ylim(y_min - margin, y_max + margin)
        self.canvas.draw()

    def reset_graph(self):
        self.ax.clear()
        self.ax.set_xlabel("Wavelength (nm)")
        self.ax.set_ylabel("Intensity")
        self.canvas.draw()
        self.log_box.append("グラフをリセットしました")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    gui = SimpleUSB4000GUI()
    gui.show()
    sys.exit(app.exec_())
