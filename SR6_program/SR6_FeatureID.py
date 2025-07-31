import sys
from oceandirect.OceanDirectAPI import OceanDirectAPI, OceanDirectError, FeatureID
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QComboBox, QStatusBar,
    QDialog, QTableWidget, QTableWidgetItem, QHeaderView
)
from PyQt6.QtCore import Qt
import pyqtgraph as pg

class OceanDirectApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("OceanDirect 基本測定プログラム")
        self.setGeometry(100, 100, 800, 600)

        # --- API / デバイス ---
        self.od = None
        self.spectrometer = None
        self.device_ids = {}

        # --- ウィジェット ---
        cw = QWidget()
        self.setCentralWidget(cw)
        layout = QVBoxLayout(cw)

        # 接続エリア
        conn_lay = QHBoxLayout()
        self.device_combo    = QComboBox()
        self.device_combo.setPlaceholderText("利用可能な分光器ID...")
        self.connect_button  = QPushButton("接続")
        self.features_button = QPushButton("機能確認")
        conn_lay.addWidget(QLabel("デバイス:"))
        conn_lay.addWidget(self.device_combo, 1)
        conn_lay.addWidget(self.connect_button)
        conn_lay.addWidget(self.features_button)

        # 測定エリア
        ctrl_lay = QHBoxLayout()
        self.integration_time_input = QLineEdit("10000")
        self.integration_time_input.setToolTip("積分時間をµsで入力")
        self.acquire_button = QPushButton("スペクトル取得")
        ctrl_lay.addWidget(QLabel("積分時間 (µs):"))
        ctrl_lay.addWidget(self.integration_time_input)
        ctrl_lay.addStretch()

        # グラフ
        self.plot_widget  = pg.PlotWidget()
        self.plot_widget.setBackground('w')
        self.plot_widget.setLabel('left', 'Intensity', units='counts')
        self.plot_widget.setLabel('bottom', 'Wavelength', units='nm')
        self.plot_widget.showGrid(x=True, y=True)
        self.spectrum_plot = self.plot_widget.plot(pen='b')

        # ステータスバー
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)

        # レイアウト組立
        layout.addLayout(conn_lay)
        layout.addLayout(ctrl_lay)
        layout.addWidget(self.acquire_button)
        layout.addWidget(self.plot_widget, 1)

        # シグナル ⇄ スロット
        self.connect_button.clicked.connect(self.toggle_connection)
        self.acquire_button.clicked.connect(self.acquire_spectrum)
        self.features_button.clicked.connect(self.show_supported_features)

        # 初期化
        self.initialize_api_and_devices()
        self.update_ui_state(connected=False)

    def initialize_api_and_devices(self):
        """OceanDirect APIを初期化し、利用可能なデバイスを検索"""
        try:
            self.od = OceanDirectAPI()
            self.od.find_devices()
            # 検出されたシリアルナンバーを取得する正しいメソッド名
            device_serials = self.od.get_found_serial_numbers()
            if not device_serials:
                self.status_bar.showMessage("分光器が見つかりません。")
                return

            for serial in device_serials:
                device_id = self.od.get_device_id(serial)
                self.device_ids[serial] = device_id
                self.device_combo.addItem(f"{self.od.get_model(device_id)} ({serial})", serial)
            self.status_bar.showMessage(f"{len(device_serials)}台の分光器を検出しました。")

        except OceanDirectError as e:
            self.status_bar.showMessage(f"APIエラー: {e}")
            self.od = None

    def toggle_connection(self):
        """分光器への接続・切断を切り替える"""
        if self.spectrometer is None:
            # 接続処理
            selected_serial = self.device_combo.currentData()
            if not selected_serial:
                self.status_bar.showMessage("接続するデバイスを選択してください。")
                return
            try:
                device_id = self.device_ids[selected_serial]
                self.spectrometer = self.od.open_device(device_id)
                self.status_bar.showMessage(f"{selected_serial} に接続しました。")
                self.update_ui_state(connected=True)
            except OceanDirectError as e:
                self.status_bar.showMessage(f"接続エラー: {e}")
                self.spectrometer = None
        else:
            # 切断処理
            try:
                serial_num = self.spectrometer.get_serial_number()
                self.od.close_device(self.spectrometer.device_id)
                self.status_bar.showMessage(f"{serial_num} から切断しました。")
            except OceanDirectError as e:
                 self.status_bar.showMessage(f"切断エラー: {e}")
            finally:
                self.spectrometer = None
                self.update_ui_state(connected=False)

    def acquire_spectrum(self):
        """スペクトルを取得してグラフを更新"""
        if not self.spectrometer:
            self.status_bar.showMessage("分光器が接続されていません。")
            return
        try:
            # 積分時間の設定
            integration_time = int(self.integration_time_input.text())
            self.spectrometer.set_integration_time(integration_time)

            # スペクトル取得とプロット
            wavelengths = self.spectrometer.get_wavelengths()
            intensities = self.spectrometer.get_formatted_spectrum()
            self.spectrum_plot.setData(wavelengths, intensities)
            self.status_bar.showMessage("スペクトルを取得しました。")

        except (ValueError, OceanDirectError) as e:
            self.status_bar.showMessage(f"測定エラー: {e}")


    def update_ui_state(self, connected: bool):
        """UIの状態を接続状態に応じて更新"""
        self.device_combo.setEnabled(not connected)
        self.integration_time_input.setEnabled(connected)
        self.acquire_button.setEnabled(connected)
        self.features_button.setEnabled(connected)
        if connected:
            self.connect_button.setText("切断")
        else:
            self.connect_button.setText("接続")
            self.spectrum_plot.clear() # 切断時にグラフをクリア

    def show_supported_features(self):
        """接続中の分光器がサポートする機能を一覧表示"""
        if not self.spectrometer:
            self.status_bar.showMessage("先に分光器を接続してください。")
            return

        dlg = QDialog(self)
        dlg.setWindowTitle("サポート機能一覧")
        lay = QVBoxLayout(dlg)

        features = list(FeatureID)
        table = QTableWidget(len(features), 2)
        table.setHorizontalHeaderLabels(["機能", "対応"])
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        table.verticalHeader().setVisible(False)

        for i, feat in enumerate(features):
            item_name = QTableWidgetItem(feat.name)
            item_name.setFlags(Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled)
            table.setItem(i, 0, item_name)

            try:
                # is_feature_id_enabled よりも is_feature_supported の方がより一般的
                ok = self.spectrometer.is_feature_supported(feat)
            except OceanDirectError:
                ok = False
            item_ok = QTableWidgetItem("○" if ok else "×")
            item_ok.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            item_ok.setFlags(Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled)
            table.setItem(i, 1, item_ok)

        lay.addWidget(table)
        dlg.resize(400, 600)
        dlg.exec()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = OceanDirectApp()
    win.show()
    sys.exit(app.exec())