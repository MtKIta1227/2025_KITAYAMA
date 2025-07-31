import sys
import csv
import re
from pathlib import Path
from datetime import datetime
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QComboBox, QStatusBar, QTreeWidget,
    QTreeWidgetItem, QFileDialog, QMenu, QAbstractItemView, QTreeWidgetItemIterator
)
from PyQt6.QtGui import QAction
from PyQt6.QtCore import Qt, QTimer
import pyqtgraph as pg
import itertools

from .controllers.spectrometer_controller import SpectrometerController, OceanDirectError

class TimestampSortTreeWidgetItem(QTreeWidgetItem):
    def __lt__(self, other):
        ts1 = self.data(0, Qt.ItemDataRole.UserRole)
        ts2 = other.data(0, Qt.ItemDataRole.UserRole)
        if ts1 and ts2:
            return ts1 < ts2
        return super().__lt__(other)

class OceanDirectApp(QMainWindow):
    """GUIの表示とデータ管理を担当するクラス"""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("OceanDirect 測定プログラム")
        self.resize(1100, 650)
        
        self.spectrometer_controller = SpectrometerController()
        
        self.dark_spectrum: list[float] | None = None
        self.wavelengths: list[float] = []
        self.spectra_data: dict[str, tuple] = {}
        
        self.spectrum_counter = 1
        self.group_counter = 1
        
        self.plot_colors = itertools.cycle([
            (29, 105, 222), (237, 102, 33), (60, 168, 56), (211, 47, 47),
            (142, 68, 173), (241, 196, 15), (46, 204, 113)
        ])

        self._build_ui()
        self.find_and_populate_devices()
        self.update_ui_state(False)

    def _build_ui(self):
        menu_bar = self.menuBar()
        file_menu = menu_bar.addMenu("ファイル")
        load_action = QAction("CSVから読み込み", self)
        load_action.triggered.connect(self.load_data_from_csv)
        file_menu.addAction(load_action)
        save_action = QAction("名前を付けて保存", self)
        save_action.triggered.connect(self.save_data_as)
        file_menu.addAction(save_action)

        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QHBoxLayout(central)

        # Left Pane
        left_pane = QWidget()
        left_layout = QVBoxLayout(left_pane)
        main_layout.addWidget(left_pane, 3)

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
        self.plot_widget.setLabel('left', 'Intensity (a.u.)')
        self.plot_widget.setLabel('bottom', 'Wavelength / nm')
        self.plot_widget.showGrid(x=True, y=True)
        left_layout.addWidget(self.plot_widget, 1)

        # Right Pane
        right_pane = QWidget()
        right_layout = QVBoxLayout(right_pane)
        main_layout.addWidget(right_pane, 1)
        right_layout.addWidget(QLabel("測定スペクトルリスト"))
        self.data_list = QTreeWidget()
        self.data_list.setHeaderLabels(["測定データ"])
        self.data_list.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.data_list.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.data_list.setSortingEnabled(True)
        self.data_list.sortItems(0, Qt.SortOrder.AscendingOrder)
        right_layout.addWidget(self.data_list)

        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)

        # Connect signals
        self.connect_button.clicked.connect(self.toggle_connection)
        self.acquire_button.clicked.connect(self.acquire_spectrum)
        self.acquire_dark_button.clicked.connect(self.acquire_dark_spectrum)
        self.data_list.itemClicked.connect(self.on_item_clicked)
        self.data_list.customContextMenuRequested.connect(self.show_context_menu)
        self.toggle_group_button.clicked.connect(self.toggle_group_action)
        self.data_list.itemDoubleClicked.connect(self.edit_item_name)
        self.data_list.itemChanged.connect(self.update_item_name)

    def find_and_populate_devices(self):
        try:
            device_ids = self.spectrometer_controller.find_devices()
            self.device_combo.clear()
            if device_ids:
                for dev_id in device_ids:
                    self.device_combo.addItem(f"Device ID: {dev_id}", dev_id)
                self.status_bar.showMessage(f"{len(device_ids)} 個のデバイスが見つかりました", 5000)
            else:
                self.status_bar.showMessage("分光器が見つかりませんでした")
        except Exception as e:
            self.status_bar.showMessage(f"API 初期化エラー: {e}")

    def toggle_connection(self):
        if self.spectrometer_controller.is_connected:
            self.spectrometer_controller.disconnect()
            self.update_ui_state(False)
            self.status_bar.showMessage("切断しました")
        else:
            try:
                selected_id = self.device_combo.currentData()
                if selected_id is None:
                    self.status_bar.showMessage("接続するデバイスを選択してください", 3000)
                    return
                
                self.spectrometer_controller.connect(selected_id)
                self.wavelengths = self.spectrometer_controller.wavelengths
                limits = self.spectrometer_controller.integration_time_limits
                self.integration_time_input.setToolTip(f"積分時間 (µs): {limits[0]} - {limits[1]}")
                self.update_ui_state(True)
                self.status_bar.showMessage("接続完了")
            except OceanDirectError as e:
                self.status_bar.showMessage(f"接続エラー: {e}")

    def update_ui_state(self, connected: bool):
        self.device_combo.setEnabled(not connected)
        for w in [self.acquire_button, self.acquire_dark_button, self.integration_time_input]:
            w.setEnabled(connected)
        self.connect_button.setText("切断" if connected else "接続")

    def acquire_spectrum(self):
        try:
            integ = int(self.integration_time_input.text())
            self.spectrometer_controller.set_integration_time(integ)
            data = self.spectrometer_controller.acquire_spectrum()
            
            if self.dark_spectrum:
                data = [s - d for s, d in zip(data, self.dark_spectrum)]

            label = f"Spe_{self.spectrum_counter}"
            self.spectrum_counter += 1
            timestamp = datetime.now()
            self.spectra_data[label] = ('spectrum', data, timestamp)
            item = TimestampSortTreeWidgetItem([label])
            item.setData(0, Qt.ItemDataRole.UserRole, timestamp)
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsEditable)
            self.data_list.addTopLevelItem(item)
            self.data_list.sortItems(0, Qt.SortOrder.AscendingOrder)
            self.data_list.clearSelection()
            self.data_list.setCurrentItem(item)
            self.on_item_clicked(item)
            self.status_bar.showMessage(f"{label} を取得しました", 3000)
        except (ValueError, OceanDirectError) as e:
            self.status_bar.showMessage(f"測定エラー: {e}")

    def acquire_dark_spectrum(self):
        try:
            integ = int(self.integration_time_input.text())
            self.spectrometer_controller.set_integration_time(integ)
            self.dark_spectrum = self.spectrometer_controller.acquire_spectrum()
            self.status_bar.showMessage("ダークスペクトルを保存しました", 3000)
        except (ValueError, OceanDirectError) as e:
            self.status_bar.showMessage(f"ダーク取得エラー: {e}")

    def on_item_clicked(self, item: QTreeWidgetItem):
        if len(self.data_list.selectedItems()) > 1:
            self.plot_widget.clear()
            self.status_bar.showMessage(f"{len(self.data_list.selectedItems())}個のアイテムを選択中", 2000)
            return
        
        self.plot_widget.clear()
        if not item: return

        if item.parent() is None and item.childCount() > 0:
            label = item.text(0)
            _, data, _ = self.spectra_data.get(label, (None, None, None))
            if data:
                for spec_label, spec_data, _ in data:
                    color = next(self.plot_colors)
                    self.plot_widget.plot(self.wavelengths, spec_data, pen=pg.mkPen(color=color), name=spec_label)
                self.status_bar.showMessage(f"グループ '{label}' を重ね書き表示中", 2000)
        else:
            label = item.text(0)
            data = None
            if item.parent():
                g_label = item.parent().text(0)
                _, group_data, _ = self.spectra_data.get(g_label, (None, [], None))
                data = next((d for l, d, _ in group_data if l == label), None)
            else:
                _, data, _ = self.spectra_data.get(label, (None, None, None))
            if data:
                self.plot_widget.plot(self.wavelengths, data, pen='b', name=label)
                self.status_bar.showMessage(f"'{label}' を単独表示中", 2000)

    def toggle_group_action(self):
        selected_items = self.data_list.selectedItems()
        if not selected_items:
            self.status_bar.showMessage("リストからアイテムを選択してください", 3000)
            return

        parent = selected_items[0].parent()
        if parent is not None:
            if all(item.parent() == parent for item in selected_items):
                self.remove_items_from_group(selected_items)
            else:
                self.status_bar.showMessage("グループ解除は同じグループ内のアイテムでのみ行えます。", 3000)
            return

        if len(selected_items) > 1 and all(item.childCount() == 0 for item in selected_items):
            self.group_selected_spectra(selected_items)
        else:
            self.status_bar.showMessage("無効な選択です。グループ化(単独アイテム複数選択)または解除(グループ内アイテム選択)を行ってください。", 4000)

    def group_selected_spectra(self, selected_items: list[QTreeWidgetItem]):
        new_gname = f"Group{self.group_counter}"
        self.group_counter += 1
        new_group_data = []
        for item in selected_items:
            label = item.text(0)
            if label in self.spectra_data:
                _, data, timestamp = self.spectra_data.pop(label)
                new_group_data.append((label, data, timestamp))
        
        group_timestamp = datetime.now()
        self.spectra_data[new_gname] = ('group', new_group_data, group_timestamp)
        new_group_item = TimestampSortTreeWidgetItem([new_gname])
        new_group_item.setData(0, Qt.ItemDataRole.UserRole, group_timestamp)
        new_group_item.setFlags(new_group_item.flags() | Qt.ItemFlag.ItemIsEditable)
        self.data_list.addTopLevelItem(new_group_item)
        
        for item in selected_items:
            (item.parent() or self.data_list.invisibleRootItem()).removeChild(item)
            new_group_item.addChild(item)

        self.data_list.sortItems(0, Qt.SortOrder.AscendingOrder)
        new_group_item.setExpanded(True)
        self.data_list.clearSelection()
        self.data_list.setCurrentItem(new_group_item)
        self.on_item_clicked(new_group_item)
        self.status_bar.showMessage(f"'{new_gname}' を作成しました", 3000)

    def remove_items_from_group(self, items_to_remove: list[QTreeWidgetItem]):
        parent_item = items_to_remove[0].parent()
        if not parent_item: return
        
        g_label = parent_item.text(0)
        _, group_data_list, group_ts = self.spectra_data[g_label]
        labels_to_remove = {item.text(0) for item in items_to_remove}
        new_group_list = []
        
        for spec_label, spec_data, spec_ts in group_data_list:
            if spec_label in labels_to_remove:
                self.spectra_data[spec_label] = ('spectrum', spec_data, spec_ts)
            else:
                new_group_list.append((spec_label, spec_data, spec_ts))
        
        self.spectra_data[g_label] = ('group', new_group_list, group_ts)
        
        for item in items_to_remove:
            parent_item.removeChild(item)
            self.data_list.addTopLevelItem(item)
            
        self.data_list.clearSelection()
        self.data_list.sortItems(0, Qt.SortOrder.AscendingOrder)
        
        if parent_item.childCount() <= 1:
            self.dissolve_group(parent_item)
            self.status_bar.showMessage("アイテムをグループ解除し、残りのアイテムが少ないためグループを解散しました", 4000)
        else:
            self.status_bar.showMessage(f"{len(items_to_remove)}個のアイテムをグループ解除しました", 3000)

    def dissolve_group(self, group_item: QTreeWidgetItem):
        if group_item.parent() is not None: return
        g_label = group_item.text(0)
        if g_label in self.spectra_data:
            _, group_data, _ = self.spectra_data.pop(g_label)
            for label, data, timestamp in group_data:
                self.spectra_data[label] = ('spectrum', data, timestamp)
        
        children_to_move = []
        while group_item.childCount() > 0:
            children_to_move.append(group_item.takeChild(0))
        for child in children_to_move:
            self.data_list.addTopLevelItem(child)

        idx = self.data_list.indexOfTopLevelItem(group_item)
        if idx != -1:
            self.data_list.takeTopLevelItem(idx)
        
        self.data_list.sortItems(0, Qt.SortOrder.AscendingOrder)
        self.plot_widget.clear()

    def edit_item_name(self, item: QTreeWidgetItem, column: int):
        item.setData(0, Qt.ItemDataRole.UserRole + 1, item.text(0))
        QTimer.singleShot(0, lambda: self.data_list.editItem(item, column))

    def update_item_name(self, item: QTreeWidgetItem, column: int):
        old_name = item.data(0, Qt.ItemDataRole.UserRole + 1)
        new_name = item.text(0)

        if not new_name or old_name is None or old_name == new_name:
            item.setText(0, old_name or new_name)
            return
        if new_name in self.spectra_data:
            self.status_bar.showMessage(f"エラー: '{new_name}' は既に存在します", 3000)
            item.setText(0, old_name)
            return

        item_type, data, timestamp = self.spectra_data.pop(old_name)
        self.spectra_data[new_name] = (item_type, data, timestamp)
        
        if item.parent():
            parent = item.parent()
            g_label = parent.text(0)
            if g_label in self.spectra_data and self.spectra_data[g_label][0] == 'group':
                group_list = self.spectra_data[g_label][1]
                group_ts = self.spectra_data[g_label][2]
                for i, (lbl, spec_data, spec_ts) in enumerate(group_list):
                    if lbl == old_name:
                        group_list[i] = (new_name, spec_data, spec_ts)
                        self.spectra_data[g_label] = ('group', group_list, group_ts)
                        break
        
        self.status_bar.showMessage(f"'{old_name}' を '{new_name}' に変更しました", 3000)
        self.on_item_clicked(item)

    def show_context_menu(self, pos):
        if len(self.data_list.selectedItems()) > 1: return
        item = self.data_list.itemAt(pos)
        if not item: return
        menu = QMenu(self)
        delete_act = QAction("削除", self, triggered=lambda: self.delete_item(item))
        menu.addAction(delete_act)
        menu.exec(self.data_list.viewport().mapToGlobal(pos))

    def delete_item(self, item: QTreeWidgetItem):
        if item.parent() is None and item.childCount() > 0:
            g_label = item.text(0)
            if g_label in self.spectra_data:
                _, group_data, _ = self.spectra_data.pop(g_label)
                for spec_label, _, _ in group_data:
                    self.spectra_data.pop(spec_label, None)
        elif item.parent():
            parent_item = item.parent()
            g_label = parent_item.text(0)
            spec_label = item.text(0)
            if g_label in self.spectra_data and self.spectra_data[g_label][0] == 'group':
                group_data = self.spectra_data[g_label][1]
                group_ts = self.spectra_data[g_label][2]
                self.spectra_data[g_label] = ('group', [d for d in group_data if d[0] != spec_label], group_ts)
            self.spectra_data.pop(spec_label, None)
        else:
            label = item.text(0)
            self.spectra_data.pop(label, None)
        (item.parent() or self.data_list.invisibleRootItem()).removeChild(item)
        self.plot_widget.clear()
        self.status_bar.showMessage("削除しました", 3000)

    def save_data_as(self):
        selected_items = self.data_list.selectedItems()
        spectra_to_save = []
        
        if not selected_items:
            iterator = QTreeWidgetItemIterator(self.data_list)
            while iterator.value():
                item = iterator.value()
                if item.childCount() == 0:
                    label = item.text(0)
                    data = None
                    if item.parent():
                        parent_label = item.parent().text(0)
                        _, group_data, _ = self.spectra_data.get(parent_label, (None, [], None))
                        data = next((d for l, d, _ in group_data if l == label), None)
                    else:
                        _, data, _ = self.spectra_data.get(label, (None, None, None))
                    if data: spectra_to_save.append((label, data))
                iterator += 1
        else:
            for item in selected_items:
                if item.parent() is None and item.childCount() > 0:
                    g_label = item.text(0)
                    _, group_data, _ = self.spectra_data.get(g_label, (None, [], None))
                    for spec_label, spec_data, _ in group_data:
                        spectra_to_save.append((spec_label, spec_data))
                elif item.parent() is None:
                    label = item.text(0)
                    _, data, _ = self.spectra_data.get(label, (None, None, None))
                    if data: spectra_to_save.append((label, data))

        if not spectra_to_save:
            self.status_bar.showMessage("保存対象のスペクトルデータがありません", 3000)
            return

        path, _ = QFileDialog.getSaveFileName(self, "名前を付けて保存", "", "CSV Files (*.csv)")
        if not path: return
        self._write_spectra_to_csv(path, spectra_to_save)

    def load_data_from_csv(self):
        path, _ = QFileDialog.getOpenFileName(self, "CSVから読み込み", "", "CSV Files (*.csv)")
        if not path: return
            
        try:
            with open(path, 'r', encoding='utf-8') as f:
                reader = csv.reader(f)
                header = next(reader)
                data_rows = list(reader)

            if not header or len(header) < 2 or not data_rows:
                self.status_bar.showMessage("ファイルが空か、形式が不正です", 3000)
                return

            if not self.wavelengths:
                self.wavelengths = [float(row[0]) for row in data_rows]
            
            num_spectra_loaded = 0
            for i in range(1, len(header)):
                label = header[i]
                original_label = label
                count = 1
                while label in self.spectra_data:
                    label = f"{original_label}_imported_{count}"
                    count += 1
                
                data = [float(row[i]) for row in data_rows]
                timestamp = datetime.now()
                self.spectra_data[label] = ('spectrum', data, timestamp)
                item = TimestampSortTreeWidgetItem([label])
                item.setData(0, Qt.ItemDataRole.UserRole, timestamp)
                item.setFlags(item.flags() | Qt.ItemFlag.ItemIsEditable)
                self.data_list.addTopLevelItem(item)
                num_spectra_loaded += 1
            
            self.data_list.sortItems(0, Qt.SortOrder.AscendingOrder)
            self.status_bar.showMessage(f"{num_spectra_loaded}個のスペクトルを読み込みました", 4000)

        except (ValueError, IndexError, FileNotFoundError) as e:
            self.status_bar.showMessage(f"ファイルの読み込みエラー: {e}", 5000)

    def _write_spectra_to_csv(self, path, spectra_to_save):
        if not self.wavelengths: return
        try:
            with open(path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                header = [self.plot_widget.getAxis('bottom').labelText] + [label for label, _ in spectra_to_save]
                writer.writerow(header)

                for i, wl in enumerate(self.wavelengths):
                    row = [wl] + [data[i] for _, data in spectra_to_save]
                    writer.writerow(row)
            
            self.status_bar.showMessage(f"データを '{Path(path).name}' に保存しました", 4000)
        except OSError as e:
            self.status_bar.showMessage(f"保存エラー: {e}", 5000)

    def closeEvent(self, event):
        self.spectrometer_controller.shutdown()
        event.accept()