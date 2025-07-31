import sys
import csv
from pathlib import Path
from oceandirect.OceanDirectAPI import OceanDirectAPI, OceanDirectError
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QComboBox, QStatusBar, QTreeWidget,
    QTreeWidgetItem, QFileDialog, QMenu, QAbstractItemView
)
from PyQt6.QtGui import QAction
from PyQt6.QtCore import Qt
import pyqtgraph as pg
import itertools

class OceanDirectApp(QMainWindow):
    """OceanDirect 分光測定 GUI（改良版グループ化・解除機能付き）"""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("OceanDirect 測定プログラム (ボタン操作版)")
        self.resize(1100, 650)
        self.od = None
        self.spectrometer = None
        self.device_ids: dict[str, int] = {}
        self.dark_spectrum: list[float] | None = None
        self.wavelengths: list[float] = []
        
        self.spectra_data: dict[str, tuple[str, list]] = {}
        
        self.spectrum_counter = 1
        self.group_counter = 1
        
        self.plot_colors = itertools.cycle([
            (29, 105, 222), (237, 102, 33), (60, 168, 56), (211, 47, 47),
            (142, 68, 173), (241, 196, 15), (46, 204, 113)
        ])

        self._build_ui()
        self.initialize_api_and_devices()
        self.update_ui_state(False)

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QHBoxLayout(central)

        # 左ペイン
        left = QWidget(); left_layout = QVBoxLayout(left)
        main_layout.addWidget(left, 3)

        conn_l = QHBoxLayout()
        self.device_combo = QComboBox(placeholderText="利用可能な分光器ID…")
        self.connect_button = QPushButton("接続")
        conn_l.addWidget(QLabel("デバイス:"))
        conn_l.addWidget(self.device_combo, 1)
        conn_l.addWidget(self.connect_button)
        left_layout.addLayout(conn_l)

        ctl_l = QHBoxLayout()
        self.integration_time_input = QLineEdit("10000")
        self.integration_time_input.setToolTip("積分時間をマイクロ秒で入力")
        self.acquire_button = QPushButton("スペクトル取得")
        self.acquire_dark_button = QPushButton("ダーク測定")
        ctl_l.addWidget(QLabel("積分時間 (µs):"))
        ctl_l.addWidget(self.integration_time_input)
        ctl_l.addWidget(self.acquire_dark_button)
        left_layout.addLayout(ctl_l)
        left_layout.addWidget(self.acquire_button)

        self.toggle_group_button = QPushButton("グループ化 / 解除")
        left_layout.addWidget(self.toggle_group_button)

        self.plot_widget = pg.PlotWidget(background='w')
        self.plot_widget.addLegend()
        self.plot_widget.setLabel('left', 'Intensity', units='counts')
        self.plot_widget.setLabel('bottom', 'Wavelength', units='nm')
        self.plot_widget.showGrid(x=True, y=True)
        left_layout.addWidget(self.plot_widget, 1)

        # 右ペイン
        right = QWidget(); right_layout = QVBoxLayout(right)
        main_layout.addWidget(right, 1)
        right_layout.addWidget(QLabel("測定スペクトルリスト"))
        self.data_list = QTreeWidget()
        self.data_list.setHeaderLabels(["測定データ"])
        self.data_list.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.data_list.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        right_layout.addWidget(self.data_list)

        self.status_bar = QStatusBar(); self.setStatusBar(self.status_bar)

        self.connect_button.clicked.connect(self.toggle_connection)
        self.acquire_button.clicked.connect(self.acquire_spectrum)
        self.acquire_dark_button.clicked.connect(self.acquire_dark_spectrum)
        self.data_list.itemClicked.connect(self.on_item_clicked)
        self.data_list.customContextMenuRequested.connect(self.show_context_menu)
        self.toggle_group_button.clicked.connect(self.toggle_group_action)

    def initialize_api_and_devices(self):
        try:
            self.od = OceanDirectAPI()
            n = self.od.find_usb_devices()
            if n > 0:
                for dev_id in self.od.get_device_ids():
                    text = f"Device ID: {dev_id}"
                    self.device_combo.addItem(text)
                    self.device_ids[text] = dev_id
                self.status_bar.showMessage(f"{n} 個のデバイスが見つかりました", 5000)
            else:
                self.status_bar.showMessage("分光器が見つかりませんでした")
        except (OceanDirectError, OSError) as e:
            self.status_bar.showMessage(f"API 初期化エラー: {e}")

    def toggle_connection(self):
        if self.spectrometer is None:
            if not self.device_ids:
                self.status_bar.showMessage("接続するデバイスがありません")
                return
            device_id = self.device_ids.get(self.device_combo.currentText())
            try:
                self.spectrometer = self.od.open_device(device_id)
                self.wavelengths = self.spectrometer.get_wavelengths()
                mi = self.spectrometer.get_minimum_integration_time()
                ma = self.spectrometer.get_maximum_integration_time()
                self.integration_time_input.setToolTip(f"積分時間 (µs): {mi} - {ma}")
                self.update_ui_state(True)
                self.status_bar.showMessage("接続完了", 3000)
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

    def update_ui_state(self, connected: bool):
        self.device_combo.setEnabled(not connected)
        for w in [self.acquire_button, self.acquire_dark_button, self.integration_time_input,
                  self.toggle_group_button]:
            w.setEnabled(connected)
        self.connect_button.setText("切断" if connected else "接続")
        if not connected:
            self.plot_widget.clear()

    def acquire_spectrum(self):
        if self.spectrometer is None:
            self.status_bar.showMessage("分光器が接続されていません")
            return
        try:
            integ = int(self.integration_time_input.text())
            self.spectrometer.set_integration_time(integ)
            data = self.spectrometer.get_formatted_spectrum()
            if self.dark_spectrum:
                data = [s - d for s, d in zip(data, self.dark_spectrum)]

            label = f"Spe_{self.spectrum_counter}"
            self.spectrum_counter += 1
            
            self.spectra_data[label] = ('spectrum', data)
            item = QTreeWidgetItem([label])
            self.data_list.addTopLevelItem(item)
            self.data_list.clearSelection()
            self.data_list.setCurrentItem(item)
            self.on_item_clicked(item)
            self.status_bar.showMessage(f"{label} を取得しました", 3000)

        except ValueError:
            self.status_bar.showMessage("積分時間は数値で指定してください")
        except OceanDirectError as e:
            self.status_bar.showMessage(f"測定エラー: {e}")

    def acquire_dark_spectrum(self):
        if self.spectrometer is None:
            self.status_bar.showMessage("分光器が接続されていません")
            return
        try:
            integ = int(self.integration_time_input.text())
            self.spectrometer.set_integration_time(integ)
            self.dark_spectrum = self.spectrometer.get_formatted_spectrum()
            self.status_bar.showMessage("ダークスペクトルを保存しました", 3000)
        except (ValueError, OceanDirectError) as e:
            self.status_bar.showMessage(f"ダーク取得エラー: {e}")

    def on_item_clicked(self, item: QTreeWidgetItem):
        # 複数選択時はグラフをクリアするだけ
        if len(self.data_list.selectedItems()) > 1:
            self.plot_widget.clear()
            self.status_bar.showMessage(f"{len(self.data_list.selectedItems())}個のアイテムを選択中", 2000)
            return
        
        self.plot_widget.clear()
        if not item: return

        if item.parent() is None and item.childCount() > 0:
            label = item.text(0)
            item_type, data = self.spectra_data.get(label, (None, None))
            if item_type == 'group':
                for spec_label, spec_data in data:
                    color = next(self.plot_colors)
                    self.plot_widget.plot(self.wavelengths, spec_data, pen=pg.mkPen(color=color), name=spec_label)
                self.status_bar.showMessage(f"グループ '{label}' を重ね書き表示中", 2000)
        
        else:
            label = item.text(0)
            data = None
            if item.parent():
                g_label = item.parent().text(0)
                _, group_data = self.spectra_data.get(g_label, (None, []))
                data = next((d for l, d in group_data if l == label), None)
            else:
                _, data = self.spectra_data.get(label, (None, None))
            
            if data:
                self.plot_widget.plot(self.wavelengths, data, pen='b', name=label)
                self.status_bar.showMessage(f"'{label}' を単独表示中", 2000)

    # --- 変更点: ボタンの動作を振り分けるロジックを更新 ---
    def toggle_group_action(self):
        """選択状態に応じてグループ化または部分的なグループ解除を行う"""
        selected_items = self.data_list.selectedItems()
        if not selected_items:
            self.status_bar.showMessage("リストからアイテムを選択してください", 3000)
            return

        # --- 解除のロジック ---
        # 選択アイテムが全て「同じ親を持つ子アイテム」かチェック
        parent = selected_items[0].parent()
        if parent is not None:
            if all(item.parent() == parent for item in selected_items):
                self.remove_items_from_group(selected_items)
                return

        # --- グループ化のロジック ---
        # 選択アイテムが全て「トップレベルの単独アイテム」かチェック
        are_all_spectra = all(item.parent() is None and item.childCount() == 0 for item in selected_items)
        if len(selected_items) > 1 and are_all_spectra:
            self.group_selected_spectra(selected_items)
            return

        self.status_bar.showMessage("無効な選択です。グループ化(単独アイテム複数選択)または解除(グループ内アイテム選択)を行ってください。", 4000)
    
    def group_selected_spectra(self, selected_items: list[QTreeWidgetItem]):
        """渡されたアイテムのリストを新しいグループにまとめる"""
        new_gname = f"Group{self.group_counter}"
        self.group_counter += 1
        
        new_group_data = []
        for item in selected_items:
            label = item.text(0)
            if label in self.spectra_data:
                _, data = self.spectra_data.pop(label)
                new_group_data.append((label, data))
        
        self.spectra_data[new_gname] = ('group', new_group_data)

        new_group_item = QTreeWidgetItem([new_gname])
        self.data_list.addTopLevelItem(new_group_item)
        
        for item in selected_items:
            (item.parent() or self.data_list.invisibleRootItem()).removeChild(item)
            new_group_item.addChild(item)

        new_group_item.setExpanded(True)
        self.data_list.clearSelection()
        self.data_list.setCurrentItem(new_group_item)
        self.on_item_clicked(new_group_item)
        self.status_bar.showMessage(f"'{new_gname}' を作成しました", 3000)

    # --- 変更点: 部分的なグループ解除を行うメソッドを新設 ---
    def remove_items_from_group(self, items_to_remove: list[QTreeWidgetItem]):
        """選択されたアイテムを現在のグループから解除する"""
        parent_item = items_to_remove[0].parent()
        if not parent_item: return
        
        g_label = parent_item.text(0)
        
        # データ構造を更新
        group_data_list = self.spectra_data[g_label][1]
        
        labels_to_remove = {item.text(0) for item in items_to_remove}
        new_group_list = []
        
        for spec_label, spec_data in group_data_list:
            if spec_label in labels_to_remove:
                # 解除するアイテムをトップレベルのデータとして復活させる
                self.spectra_data[spec_label] = ('spectrum', spec_data)
            else:
                new_group_list.append((spec_label, spec_data))
        
        # グループのデータを更新
        self.spectra_data[g_label] = ('group', new_group_list)
        
        # GUIを更新
        for item in items_to_remove:
            parent_item.removeChild(item)
            self.data_list.addTopLevelItem(item)
            
        self.data_list.clearSelection()
        
        # もしグループに残ったアイテムが1つ以下なら、グループを自動解散させる
        if parent_item.childCount() <= 1:
            self.dissolve_group(parent_item)
            self.status_bar.showMessage("アイテムをグループ解除し、残りのアイテムが少ないためグループを解散しました", 4000)
        else:
            self.status_bar.showMessage(f"{len(items_to_remove)}個のアイテムをグループ解除しました", 3000)
            
    def dissolve_group(self, group_item: QTreeWidgetItem):
        """指定されたグループを完全に解散し、中の全アイテムをトップレベルに戻す"""
        if not (group_item.parent() is None and group_item.childCount() > 0):
             return # グループでないなら何もしない

        g_label = group_item.text(0)
        _, group_data = self.spectra_data.pop(g_label, (None, []))

        for label, data in group_data:
            self.spectra_data[label] = ('spectrum', data)
        
        children_to_move = []
        while group_item.childCount() > 0:
            children_to_move.append(group_item.takeChild(0))
        
        for child in children_to_move:
            self.data_list.addTopLevelItem(child)

        idx = self.data_list.indexOfTopLevelItem(group_item)
        if idx != -1:
            self.data_list.takeTopLevelItem(idx)
        
        self.plot_widget.clear()

    def show_context_menu(self, pos):
        if len(self.data_list.selectedItems()) > 1: return
        
        item = self.data_list.itemAt(pos)
        if not item: return
        
        menu = QMenu(self)
        is_group = item.parent() is None and item.childCount() > 0
        
        delete_act = QAction("削除", self, triggered=lambda: self.delete_item(item))
        menu.addAction(delete_act)

        if is_group:
            save_act = QAction("グループをフォルダに保存", self, triggered=lambda: self.save_group(item))
            menu.addAction(save_act)
        elif item.parent() is None:
            save_act = QAction("スペクトルをCSV保存", self, triggered=lambda: self.save_spectrum(item))
            menu.addAction(save_act)
            
        menu.exec(self.data_list.viewport().mapToGlobal(pos))

    def delete_item(self, item: QTreeWidgetItem):
        if item.parent() is None and item.childCount() > 0:
            self.dissolve_group(item) # グループ削除時はまず解散させる
        elif item.parent():
            # 複数選択での部分解除に任せるため、単一の子アイテムの削除は実装しない
            self.status_bar.showMessage("グループ内アイテムを削除するには、グループ解除機能を使ってください", 4000)
            return
        else:
            label = item.text(0)
            self.spectra_data.pop(label, None)
            (item.parent() or self.data_list.invisibleRootItem()).removeChild(item)

        self.plot_widget.clear()
        self.status_bar.showMessage("削除しました", 3000)

    def save_group(self, item: QTreeWidgetItem):
        gname = item.text(0)
        _, data = self.spectra_data.get(gname, (None, None))
        if not data: return
        
        dir_path = QFileDialog.getExistingDirectory(self, "保存先フォルダを選択")
        if not dir_path: return
            
        for lbl, spec_data in data:
            self._write_csv(Path(dir_path) / f"{gname}_{lbl}.csv", spec_data)
        self.status_bar.showMessage(f"グループ '{gname}' を保存しました", 5000)

    def save_spectrum(self, item: QTreeWidgetItem):
        label = item.text(0)
        _, data = self.spectra_data.get(label, (None, None))
        if data is None: return
        
        path, _ = QFileDialog.getSaveFileName(self, "CSV 保存", f"{label}.csv", "CSV Files (*.csv)")
        if path:
            self._write_csv(Path(path), data)
            self.status_bar.showMessage(f"{label} を保存しました", 4000)

    def _write_csv(self, path: Path, data: list[float]):
        try:
            with path.open('w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(["wavelength[nm]", "intensity[counts]"])
                writer.writerows(zip(self.wavelengths, data))
        except OSError as e:
            self.status_bar.showMessage(f"保存エラー: {e}")

    def closeEvent(self, event):
        if self.spectrometer:
            try: self.spectrometer.close_device()
            except OceanDirectError: pass
        if self.od:
            try: self.od.shutdown()
            except OceanDirectError: pass
        event.accept()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = OceanDirectApp(); win.show()
    sys.exit(app.exec())