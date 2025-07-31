import sys
from oceandirect.OceanDirectAPI import OceanDirectAPI, OceanDirectError
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QComboBox, QStatusBar
)
from PyQt6.QtCore import Qt
import pyqtgraph as pg

class OceanDirectApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("OceanDirect 基本測定プログラム")
        self.setGeometry(100, 100, 800, 600)

        self.od = None
        self.spectrometer = None
        self.device_ids = {}
        self.dark_spectrum = None  # ダークスペクトル保存用

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.layout = QVBoxLayout(self.central_widget)

        connection_layout = QHBoxLayout()
        self.device_combo = QComboBox()
        self.device_combo.setPlaceholderText("利用可能な分光器ID...")
        self.connect_button = QPushButton("接続")
        connection_layout.addWidget(QLabel("デバイス:"))
        connection_layout.addWidget(self.device_combo, 1)
        connection_layout.addWidget(self.connect_button)

        controls_layout = QHBoxLayout()
        self.integration_time_input = QLineEdit("10000")
        self.integration_time_input.setToolTip("積分時間をマイクロ秒で入力")
        self.acquire_button = QPushButton("スペクトル取得")
        self.acquire_dark_button = QPushButton("ダーク測定")
        controls_layout.addWidget(QLabel("積分時間 (\u00b5s):"))
        controls_layout.addWidget(self.integration_time_input)
        controls_layout.addWidget(self.acquire_dark_button)
        controls_layout.addStretch()

        self.plot_widget = pg.PlotWidget()
        self.plot_widget.setBackground('w')
        self.plot_widget.setLabel('left', 'Intensity', units='counts')
        self.plot_widget.setLabel('bottom', 'Wavelength', units='nm')
        self.plot_widget.showGrid(x=True, y=True)
        self.spectrum_plot = self.plot_widget.plot(pen='b')

        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)

        self.layout.addLayout(connection_layout)
        self.layout.addLayout(controls_layout)
        self.layout.addWidget(self.acquire_button)
        self.layout.addWidget(self.plot_widget, 1)

        self.connect_button.clicked.connect(self.toggle_connection)
        self.acquire_button.clicked.connect(self.acquire_spectrum)
        self.acquire_dark_button.clicked.connect(self.acquire_dark_spectrum)

        self.initialize_api_and_devices()
        self.update_ui_state(connected=False)

    def initialize_api_and_devices(self):
        try:
            self.od = OceanDirectAPI()
            num_devices = self.od.find_usb_devices()
            if num_devices > 0:
                ids = self.od.get_device_ids()
                for dev_id in ids:
                    display_text = f"Device ID: {dev_id}"
                    self.device_ids[display_text] = dev_id
                    self.device_combo.addItem(display_text)
                self.status_bar.showMessage(f"{len(ids)}個のデバイスが見つかりました。", 5000)
            else:
                self.status_bar.showMessage("分光器が見つかりませんでした。")
        except OceanDirectError as e:
            self.status_bar.showMessage(f"API初期化エラー: {e}")
        except OSError as e:
            self.status_bar.showMessage(f"OSエラー（ドライバーを確認してください）: {e}")

    def toggle_connection(self):
        if self.spectrometer is None:
            if not self.device_ids:
                self.status_bar.showMessage("接続するデバイスがありません。")
                return

            selected_text = self.device_combo.currentText()
            device_id = self.device_ids.get(selected_text)

            try:
                self.spectrometer = self.od.open_device(device_id)
                sn = self.spectrometer.get_serial_number()
                model = self.spectrometer.get_model()

                self.wavelengths = self.spectrometer.get_wavelengths()
                min_integ = self.spectrometer.get_minimum_integration_time()
                max_integ = self.spectrometer.get_maximum_integration_time()
                self.integration_time_input.setToolTip(f"積分時間 (\u00b5s): {min_integ} - {max_integ}")

                self.update_ui_state(connected=True)
                self.status_bar.showMessage(f"接続完了: {model} ({sn})")

            except OceanDirectError as e:
                self.status_bar.showMessage(f"接続エラー: {e}")
                self.spectrometer = None
        else:
            try:
                self.spectrometer.close_device()
                self.status_bar.showMessage("切断しました。")
            except OceanDirectError as e:
                self.status_bar.showMessage(f"切断エラー: {e}")
            finally:
                self.spectrometer = None
                self.wavelengths = []
                self.dark_spectrum = None
                self.update_ui_state(connected=False)

    def acquire_spectrum(self):
        if self.spectrometer is None:
            self.status_bar.showMessage("分光器が接続されていません。")
            return

        try:
            integration_time_us = int(self.integration_time_input.text())
            self.spectrometer.set_integration_time(integration_time_us)
            spectrum_data = self.spectrometer.get_formatted_spectrum()

            if self.dark_spectrum:
                spectrum_data = [s - d for s, d in zip(spectrum_data, self.dark_spectrum)]

            if self.wavelengths and spectrum_data:
                self.spectrum_plot.setData(self.wavelengths, spectrum_data)
                self.status_bar.showMessage(f"スペクトル取得完了 (積分時間: {integration_time_us} \u00b5s)", 3000)

        except ValueError:
            self.status_bar.showMessage("積分時間には数値を入力してください。")
        except OceanDirectError as e:
            self.status_bar.showMessage(f"測定エラー: {e}")

    def acquire_dark_spectrum(self):
        if self.spectrometer is None:
            self.status_bar.showMessage("分光器が接続されていません。")
            return

        try:
            integration_time_us = int(self.integration_time_input.text())
            self.spectrometer.set_integration_time(integration_time_us)
            self.dark_spectrum = self.spectrometer.get_formatted_spectrum()
            self.status_bar.showMessage("ダークスペクトルを保存しました。", 3000)
        except ValueError:
            self.status_bar.showMessage("積分時間には数値を入力してください。")
        except OceanDirectError as e:
            self.status_bar.showMessage(f"ダーク測定エラー: {e}")

    def update_ui_state(self, connected: bool):
        self.device_combo.setEnabled(not connected)
        self.acquire_button.setEnabled(connected)
        self.acquire_dark_button.setEnabled(connected)
        self.integration_time_input.setEnabled(connected)
        if connected:
            self.connect_button.setText("切断")
        else:
            self.connect_button.setText("接続")
            self.spectrum_plot.clear()

    def closeEvent(self, event):
        if self.spectrometer:
            try:
                self.spectrometer.close_device()
            except OceanDirectError as e:
                print(f"デバイスクローズエラー: {e}")

        if self.od:
            try:
                self.od.shutdown()
            except OceanDirectError as e:
                print(f"APIシャットダウンエラー: {e}")

        event.accept()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = OceanDirectApp()
    window.show()
    sys.exit(app.exec())
