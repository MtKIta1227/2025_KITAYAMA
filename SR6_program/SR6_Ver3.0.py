import sys
from oceandirect.OceanDirectAPI import OceanDirectAPI, OceanDirectError
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QComboBox, QStatusBar, QTreeWidget,
    QTreeWidgetItem, QInputDialog
)
from PyQt6.QtCore import Qt
import pyqtgraph as pg


class OceanDirectApp(QMainWindow):
    """OceanDirect を使った分光測定 GUI―グループ化対応版"""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("OceanDirect 基本測定プログラム")
        self.resize(1000, 600)

        # --- 測定関連の状態変数 ---
        self.od = None  # OceanDirectAPI インスタンス
        self.spectrometer = None  # Spectrometer オブジェクト
        self.device_ids: dict[str, int] = {}  # 表示テキスト→device_id
        self.dark_spectrum = None  # ダークスペクトル
        self.spectra_data: dict[str, list[tuple[str, list[float]]]] = {}  # グループ名→[(label, data)]
        self.spectrum_counter = 1
        self.current_group_item: QTreeWidgetItem | None = None
        self.wavelengths = []

        # ============================= UI =============================
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QHBoxLayout(central)

        # ---------- 左ペイン ----------
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        main_layout.addWidget(left_widget, 3)

        # デバイス接続
        connection_layout = QHBoxLayout()
        self.device_combo = QComboBox()
        self.device_combo.setPlaceholderText("利用可能な分光器ID…")
        self.connect_button = QPushButton("接続")
        connection_layout.addWidget(QLabel("デバイス:"))
        connection_layout.addWidget(self.device_combo, 1)
        connection_layout.addWidget(self.connect_button)
        left_layout.addLayout(connection_layout)

        # 測定コントロール
        controls_layout = QHBoxLayout()
        self.integration_time_input = QLineEdit("10000")
        self.integration_time_input.setToolTip("積分時間をマイクロ秒で入力")
        self.acquire_button = QPushButton("スペクトル取得")
        self.acquire_dark_button = QPushButton("ダーク測定")
        self.new_group_button = QPushButton("グループ作成")
        controls_layout.addWidget(QLabel("積分時間 (µs):"))
        controls_layout.addWidget(self.integration_time_input)
        controls_layout.addWidget(self.acquire_dark_button)
        controls_layout.addWidget(self.new_group_button)
        controls_layout.addStretch()
        left_layout.addLayout(controls_layout)
        left_layout.addWidget(self.acquire_button)

        # プロット
        self.plot_widget = pg.PlotWidget()
        self.plot_widget.setBackground('w')
        self.plot_widget.setLabel('left', 'Intensity', units='counts')
        self.plot_widget.setLabel('bottom', 'Wavelength', units='nm')
        self.plot_widget.showGrid(x=True, y=True)
        self.spectrum_plot = self.plot_widget.plot(pen='b')
        left_layout.addWidget(self.plot_widget, 1)

        # ---------- 右ペイン ----------
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        main_layout.addWidget(right_widget, 1)

        self.data_list = QTreeWidget()
        self.data_list.setHeaderLabels(["測定データ"])
        right_layout.addWidget(QLabel("仮保存スペクトルリスト"))
        right_layout.addWidget(self.data_list)

        # ステータスバー
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)

        # ---------- シグナル接続 ----------
        self.connect_button.clicked.connect(self.toggle_connection)
        self.acquire_button.clicked.connect(self.acquire_spectrum)
        self.acquire_dark_button.clicked.connect(self.acquire_dark_spectrum)
        self.new_group_button.clicked.connect(self.create_group)
        self.data_list.itemClicked.connect(self.on_item_clicked)

        # 初期化
        self.initialize_api_and_devices()
        self.update_ui_state(False)

    # ======================= API 初期化 =============================
    def initialize_api_and_devices(self):
        try:
            self.od = OceanDirectAPI()
            if (n := self.od.find_usb_devices()) > 0:
                for dev_id in self.od.get_device_ids():
                    text = f"Device ID: {dev_id}"
                    self.device_combo.addItem(text)
                    self.device_ids[text] = dev_id
                self.status_bar.showMessage(f"{n} 個のデバイスが見つかりました", 5000)
            else:
                self.status_bar.showMessage("分光器が見つかりませんでした")
        except (OceanDirectError, OSError) as e:
            self.status_bar.showMessage(f"API 初期化エラー: {e}")

    # ========================= 接続/切断 ============================
    def toggle_connection(self):
        if self.spectrometer is None:
            if not self.device_ids:
                self.status_bar.showMessage("接続するデバイスがありません")
                return
            device_id = self.device_ids.get(self.device_combo.currentText())
            try:
                self.spectrometer = self.od.open_device(device_id)
                self.wavelengths = self.spectrometer.get_wavelengths()
                # 積分時間範囲ヒント
                mi, ma = (self.spectrometer.get_minimum_integration_time(),
                          self.spectrometer.get_maximum_integration_time())
                self.integration_time_input.setToolTip(f"積分時間 (µs): {mi} - {ma}")
                self.status_bar.showMessage("接続完了", 3000)
                self.update_ui_state(True)
            except OceanDirectError as e:
                self.status_bar.showMessage(f"接続エラー: {e}")
        else:
            try:
                self.spectrometer.close_device()
            except OceanDirectError as e:
                self.status_bar.showMessage(f"切断エラー: {e}")
            self.spectrometer = None
            self.wavelengths = []
            self.dark_spectrum = None
            self.update_ui_state(False)

    # ========================= UI 状態更新 ==========================
    def update_ui_state(self, connected: bool):
        for w in [self.device_combo]:
            w.setEnabled(not connected)
        for w in [self.acquire_button, self.acquire_dark_button,
                  self.new_group_button, self.integration_time_input]:
            w.setEnabled(connected)
        self.connect_button.setText("切断" if connected else "接続")
        if not connected:
            self.spectrum_plot.clear()

    # ======================== グループ作成 =========================
    def create_group(self):
        name, ok = QInputDialog.getText(self, "新しいグループ", "グループ名を入力:")
        if ok and name:
            item = QTreeWidgetItem([name])
            self.data_list.addTopLevelItem(item)
            self.spectra_data[name] = []
            self.current_group_item = item
            self.data_list.setCurrentItem(item)
            self.status_bar.showMessage(f"グループ '{name}' を作成しました", 3000)

    # ======================== スペクトル取得 =======================
    def acquire_spectrum(self):
        if self.spectrometer is None:
            self.status_bar.showMessage("分光器が接続されていません")
            return
        if self.current_group_item is None:
            self.status_bar.showMessage("まずグループを選択/作成してください")
            return
        try:
            integ = int(self.integration_time_input.text())
            self.spectrometer.set_integration_time(integ)
            data = self.spectrometer.get_formatted_spectrum()
            if self.dark_spectrum:
                data = [s - d for s, d in zip(data, self.dark_spectrum)]
            if self.wavelengths:
                self.spectrum_plot.setData(self.wavelengths, data)
            label = f"spe_{self.spectrum_counter}"
            self.spectrum_counter += 1
            child = QTreeWidgetItem([label])
            self.current_group_item.addChild(child)
            group_name = self.current_group_item.text(0)
            self.spectra_data[group_name].append((label, data))
            self.status_bar.showMessage(f"{label} を {group_name} に追加", 3000)
        except ValueError:
            self.status_bar.showMessage("積分時間は数値で指定してください")
        except OceanDirectError as e:
            self.status_bar.showMessage(f"測定エラー: {e}")

    # ======================== ダーク測定 ===========================
    def acquire_dark_spectrum(self):
        if self.spectrometer is None:
            self.status_bar.showMessage("分光器が接続されていません")
            return
        try:
            integ = int(self.integration_time_input.text())
            self.spectrometer.set_integration_time(integ)
            self.dark_spectrum = self.spectrometer.get_formatted_spectrum()
            self.status_bar.showMessage("ダークスペクトルを保存しました", 3000)
        except ValueError:
            self.status_bar.showMessage("積分時間は数値で指定してください")
        except OceanDirectError as e:
            self.status_bar.showMessage(f"ダーク測定エラー: {e}")

    # ====================== ツリー選択処理 =========================
    def on_item_clicked(self, item: QTreeWidgetItem):
        parent = item.parent()
        # グループをクリック → グループ選択
        if parent is None:
            self.current_group_item = item
            self.status_bar.showMessage(f"グループ '{item.text(0)}' を選択", 2000)
            return
        # 子 (スペクトル) をクリック → プロット表示
        group_name = parent.text(0)
        label = item.text(0)
        for n, d in self.spectra_data.get(group_name, []):
            if n == label:
                self.spectrum_plot.setData(self.wavelengths, d)
                self.status_bar.showMessage(f"{label} を表示中", 3000)
                break

    # ===================== 終了処理 ================================
    def closeEvent(self, e):
        try:
            if self.spectrometer:
                self.spectrometer.close_device()
            if self.od:
                self.od.shutdown()
        except OceanDirectError as err:
            print(f"シャットダウンエラー: {err}")
        e.accept()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = OceanDirectApp()
    window.show()
    sys.exit(app.exec())
