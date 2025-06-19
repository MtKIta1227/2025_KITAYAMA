import sys
import numpy as np
import matplotlib.pyplot as plt
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPlainTextEdit, QPushButton, QFileDialog, QSizePolicy, QToolBar, QListWidget, QMessageBox
from PyQt5 import QtGui
from PyQt5.QtGui import QPixmap
import json
import os
import pandas as pd
import re

class DataGraphApp(QMainWindow):
    def __init__(self):
        super().__init__()

        #global変数
        global filter_width
        filter_width = 5

        self.setWindowTitle("TAS_Grapher_ver3.6")
        # ウィンドウサイズをコンパクトにした
        self.resize(400, 500)
        # パルスを保存する辞書
        self.pulse_data = {}
        self.original_data = {}  # 元のデータを保存する辞書

        # メインウィジェットとレイアウト
        main_widget = QWidget()
        main_layout = QVBoxLayout()
        main_widget.setLayout(main_layout)

        # ツールバーの設定
        self.toolbar = QToolBar("ツールバー")
        self.addToolBar(self.toolbar)

        # ツールバーにボタンを追加
        ## メッセージ表示用のウィジェットを追加
        #main_layout.addWidget(QLabel("Status:"))
        ##self.message_display = QPlainTextEdit()
        ##self.message_display.setReadOnly(True)  # 読み取り専用にする
        ##self.message_display.setFixedSize(500, 20)
        ##self.message_display.setStyleSheet("background-color: lightgray")
        # メッセージが書かれている一番下の行にスクロールする
       ##self.message_display.moveCursor(QtGui.QTextCursor.End)
       ##self.message_display.ensureCursorVisible()


        #main_layout.addWidget(#self.message_display)  # メインレイアウトに追加

        load_all_button = QPushButton("Load")
        load_all_button.clicked.connect(self.load_all_data)
        self.toolbar.addWidget(load_all_button)

        save_all_button = QPushButton("Save")
        save_all_button.clicked.connect(self.save_all_data)
        self.toolbar.addWidget(save_all_button)

        plot_button = QPushButton("Plot ΔAbs")
        plot_button.clicked.connect(self.plot_graph)
        self.toolbar.addWidget(plot_button)
        plot_button.setStyleSheet("background-color: lightblue")

        overlay_selected_button = QPushButton("TAS")
        overlay_selected_button.clicked.connect(self.overlay_selected_pulses)
        self.toolbar.addWidget(overlay_selected_button)
        overlay_selected_button.setStyleSheet("background-color: lightgreen")

        #abs_plot_button = QPushButton("Plot ABS")  # 新しいボタン
        #abs_plot_button.clicked.connect(self.plot_abs)  # 新しいメソッド
        #self.toolbar.addWidget(abs_plot_button)

        #load_button = QPushButton("Load Data (Single)")
        #load_button.clicked.connect(self.load_data)
        #self.toolbar.addWidget(load_button)
#
        #save_button = QPushButton("Save Data (Single)")
        #save_button.clicked.connect(self.save_data)
        #self.toolbar.addWidget(save_button)

        save_excel_button = QPushButton("Output Excel")
        save_excel_button.clicked.connect(self.save_data_to_excel)
        self.toolbar.addWidget(save_excel_button)

        delete_button = QPushButton("Deleate")
        delete_button.clicked.connect(self.delete_selected_pulse)
        self.toolbar.addWidget(delete_button)
        #色を変更
        delete_button.setStyleSheet("background-color: red")

        # パルス入力用のレイアウト
        pulse_layout = QHBoxLayout()
        self.pulse_input = QPlainTextEdit()
        self.pulse_input.setFixedSize(100, 30)
        pulse_button = QPushButton("Save DataSet")
        pulse_button.clicked.connect(self.save_pulse_data)

        pulse_layout.addWidget(QLabel("DataSet Name:"))
        pulse_layout.addWidget(self.pulse_input)
        pulse_layout.addWidget(pulse_button)
        main_layout.addLayout(pulse_layout)

        # パルスリストの表示をQListWidgetに変更
        self.pulse_list = QListWidget()
        self.pulse_list.setSelectionMode(QListWidget.MultiSelection)
        self.pulse_list.itemClicked.connect(self.load_selected_pulse_data)  # 新しいシグナル接続
        main_layout.addWidget(QLabel("Saved Data list:"))
        main_layout.addWidget(self.pulse_list)

        # 左側のレイアウト (ダークデータ)
        dark_layout = QVBoxLayout()
        self.text_boxes = {}        
        self.graph_widgets = {}
        self.text_boxes['DARK_ref'], self.graph_widgets['DARK_ref'] = self.create_text_box_with_graph('DARK_ref', dark_layout)
        self.text_boxes['DARK_sig'], self.graph_widgets['DARK_sig'] = self.create_text_box_with_graph('DARK_sig', dark_layout)

        # 右側のレイアウト (残りのデータ)
        remaining_layout = QVBoxLayout()
        self.text_boxes['ref'], self.graph_widgets['ref'] = self.create_text_box_with_graph('ref', remaining_layout)
        self.text_boxes['sig'], self.graph_widgets['sig'] = self.create_text_box_with_graph('sig', remaining_layout)
        self.text_boxes['ref_p'], self.graph_widgets['ref_p'] = self.create_text_box_with_graph('ref_p', remaining_layout)
        self.text_boxes['sig_p'], self.graph_widgets['sig_p'] = self.create_text_box_with_graph('sig_p', remaining_layout)

        # 左側と右側のレイアウトをメインレイアウトに追加
        main_layout.addLayout(dark_layout)
        main_layout.addLayout(remaining_layout)

        self.setCentralWidget(main_widget)

        # ΔAbsのグラフを表示するためのQLabelを4列目に追加
        self.abs_graph_widget = QLabel()
        self.abs_graph_widget.setFixedSize(500, 100)
        main_layout.addWidget(self.abs_graph_widget)

        # テキストボックスの内容が変更されたときにグラフを更新する
        for text_box in self.text_boxes.values():
            text_box.textChanged.connect(self.update_graphs)
    #def display_message(self, message):
        #self.message_display.appendPlainText(message)

    def delete_selected_pulse(self):
            selected_items = self.pulse_list.selectedItems()

            if not selected_items:
                QMessageBox.warning(self, "警告", "削除する項目を選択してください。")
                return

            for item in selected_items:
                pulse_name = item.text()
                if pulse_name in self.pulse_data:
                    del self.pulse_data[pulse_name]  # データ辞書から削除
                    self.pulse_list.takeItem(self.pulse_list.row(item))  # リストウィジェットから削除

            QMessageBox.information(self, "Success", "選択したリストが削除されました。")            

    def create_text_box_with_graph(self, label, layout):
        h_layout = QHBoxLayout()
        
        lbl = QLabel(f'{label}:')
        lbl.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        lbl.setFixedWidth(50)
        h_layout.addWidget(lbl)
        
        text_box = QPlainTextEdit()
        text_box.setFixedSize(150, 80)
        text_box.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        h_layout.addWidget(text_box)

        graph_widget = QLabel()
        graph_widget.setFixedSize(90, 100)
        h_layout.addWidget(graph_widget)

        layout.addLayout(h_layout)
        return text_box, graph_widget

    def update_graphs(self):
        for label, text_box in self.text_boxes.items():
            data = text_box.toPlainText()
            self.update_graph(self.graph_widgets[label], data)  # グラフを更新  
    
            if data != self.original_data.get(label, ""):  # 内容が変更された場合
                text_box.setStyleSheet("color: red;")  # 変更されたことを示す色
            else:
                text_box.setStyleSheet("color: black;")  # 元の内容に戻った場合は黒に戻す

    def update_graph(self, graph_widget, data):
        plt.clf()
        if data.strip():
            x_values, y_values = self.parse_data(data)

            if len(x_values) != len(y_values):
                msg = QMessageBox()
                msg.setIcon(QMessageBox.Warning)
                msg.setText("警告: xとyの長さが一致しません。")
                msg.setInformativeText(f"xの長さ: {len(x_values)}, yの長さ: {len(y_values)}")
                msg.setWindowTitle("データエラー")
                msg.exec_()
                return

            plt.figure(figsize=(1.5, 1))
            plt.plot(x_values, y_values, color='k', alpha=0.7, linestyle='-', linewidth=1)
            plt.grid()
            plt.axis('off')
            plt.xlim(350, 800 if x_values else 1)
            plt.ylim(min(y_values), max(y_values) if y_values else 1)

            temp_file_path = 'temp.png'
            plt.savefig(temp_file_path, bbox_inches='tight', pad_inches=0)
            plt.close()

            graph_widget.setPixmap(QPixmap(temp_file_path))
            os.remove(temp_file_path)

    def parse_data(self, data):
        x_values, y_values = [], []
        for line in data.splitlines():
            line = line.strip()
            if not line:
                continue

            # 数値らしきものをすべて抽出する
            nums = re.findall(r"[-+]?\d*\.\d+|[-+]?\d+", line)
            if len(nums) >= 2:
                try:
                    x_values.append(float(nums[0]))
                    y_values.append(float(nums[1]))
                except ValueError:
                    # 変換できない場合は無視する
                    pass
            else:
                # 数値が2つ未満の場合は無視
                continue

        return x_values, y_values

    def plot_graph(self):
        selected_items = self.pulse_list.selectedItems()
        pulse_name = selected_items[0].text() if selected_items else "未選択"

        x_dark_ref, y_dark_ref = self.parse_data(self.text_boxes['DARK_ref'].toPlainText())
        _, y_dark_sig = self.parse_data(self.text_boxes['DARK_sig'].toPlainText())
        _, y_ref = self.parse_data(self.text_boxes['ref'].toPlainText())
        _, y_sig = self.parse_data(self.text_boxes['sig'].toPlainText())
        _, y_ref_p = self.parse_data(self.text_boxes['ref_p'].toPlainText())
        _, y_sig_p = self.parse_data(self.text_boxes['sig_p'].toPlainText())

        results = {
            'ref - DARK_ref': [ref - dark_ref for dark_ref, ref in zip(y_dark_ref, y_ref)],
            'sig - DARK_sig': [sig - dark_sig for dark_sig, sig in zip(y_dark_sig, y_sig)],
            'ref_p - DARK_ref': [ref_p - dark_ref for dark_ref, ref_p in zip(y_dark_ref, y_ref_p)],
            'sig_p - DARK_sig': [sig_p - dark_sig for dark_sig, sig_p in zip(y_dark_sig, y_sig_p)],
        }

        results['Difference_ref'] = [ref_p - ref for ref_p, ref in zip(results['ref_p - DARK_ref'], results['ref - DARK_ref'])]
        results['Difference_ref'] = [abs(value) for value in results['Difference_ref']]
        results['Difference_ref'] = [value / max(results['Difference_ref']) for value in results['Difference_ref']]

        results['Difference_sig'] = [sig_p - sig for sig_p, sig in zip(results['sig_p - DARK_sig'], results['sig - DARK_sig'])]
        results['Difference_sig'] = [abs(value) for value in results['Difference_sig']]
        results['Difference_sig'] = [value / max(results['Difference_sig']) for value in results['Difference_sig']]

        results['Difference'] = [ref - sig for ref, sig in zip(results['Difference_ref'], results['Difference_sig'])]
        results['Difference'] = [abs(value) for value in results['Difference']]
        results['Difference'] = [value / max(results['Difference']) for value in results['Difference']]

        log_values = []
        for ref_p_real, sig_real, sig_p_real, ref_real in zip(y_ref_p, y_sig, y_sig_p, y_ref):
            if ref_real != 0 and sig_p_real != 0:
                log_value = np.log((ref_p_real * sig_real) / (sig_p_real * ref_real))
                log_values.append(log_value)
            else:
                log_values.append(np.nan)
        #移動平均の窓幅をfilter_widthに指定
        log_values = np.convolve(log_values, np.ones(filter_width)/filter_width, mode='same')

        log_fig = plt.figure(figsize=(6, 4))
        plt.plot(x_dark_ref, log_values, color='black', alpha=0.7, linestyle='-',
                 label='LOG = log((ref_p * sig) / (sig_p * ref))')
        plt.axhspan(-0.01, 0.01, color='green', alpha=0.2)
        plt.title(f'Transient Absorption Spectrum -{pulse_name}-')
        plt.xlabel('Wavelength / nm')
        plt.ylabel('ΔAbs')
        plt.legend()
        plt.grid()
        plt.tight_layout()
        plt.get_current_fig_manager().window.setGeometry(600, 100, 800, 400)
        plt.show()

        plt.figure(figsize=(6, 10))  # 高さを調整
        plt.subplot(3, 1, 1)
        plt.plot(x_dark_ref, results['ref - DARK_ref'], color='black', alpha=0.7, linestyle='-', label='ref - DARK_ref')
        plt.plot(x_dark_ref, results['ref_p - DARK_ref'], color='black', linestyle='--', label='ref_p - DARK_ref')
        plt.xlabel('Wavelength / nm')
        plt.ylabel('Counts')
        plt.legend()
        plt.title(f'-{pulse_name}-')
        plt.tight_layout()
        plt.grid()

        plt.subplot(3, 1, 2)
        plt.plot(x_dark_ref, results['sig - DARK_sig'], color='black', alpha=0.7, linestyle='-', label='sig - DARK_sig')
        plt.plot(x_dark_ref, results['sig_p - DARK_sig'], color='black', linestyle='--', label='sig_p - DARK_sig')
        plt.xlabel('Wavelength / nm')
        plt.ylabel('Counts')
        plt.legend()
        plt.tight_layout()
        plt.grid()

        plt.subplot(3, 1, 3)
        plt.plot(x_dark_ref, results['Difference_ref'], color='r', alpha=0.6, linestyle='-', label='Difference_ref')
        plt.plot(x_dark_ref, results['Difference_sig'], color='b', alpha=0.6, linestyle='-', label='Difference_sig')
        plt.plot(x_dark_ref, results['Difference'], color='k', alpha=0.9, linestyle='-', label='Difference')

        plt.xlabel('Wavelength / nm')
        plt.ylabel('Difference')
        plt.ylim(0, 1)
        plt.legend()
        plt.grid()

        plt.tight_layout()
        plt.show()

    def plot_abs(self):
        x_dark_ref, y_dark_ref = self.parse_data(self.text_boxes['DARK_ref'].toPlainText())
        _, y_dark_sig = self.parse_data(self.text_boxes['DARK_sig'].toPlainText())  
        _, y_ref = self.parse_data(self.text_boxes['ref'].toPlainText())
        _, y_ref_p = self.parse_data(self.text_boxes['ref_p'].toPlainText())
        _, y_sig = self.parse_data(self.text_boxes['sig'].toPlainText())
        _, y_sig_p = self.parse_data(self.text_boxes['sig_p'].toPlainText())

        abs_ref = np.log((np.array(y_ref) - np.array(y_dark_ref)) / (np.array(y_ref_p) - np.array(y_dark_ref)))
        abs_sig = np.log((np.array(y_sig) - np.array(y_dark_sig)) / (np.array(y_sig_p) - np.array(y_dark_sig)))

        plt.figure(figsize=(8, 6))
        plt.plot(x_dark_ref, abs_ref, label='log((ref - DARK_ref) / (ref_p - DARK_ref))', color='blue')
        plt.plot(x_dark_ref, abs_sig, label='log((sig - DARK_sig) / (sig_p - DARK_sig))', color='orange')
        plt.title('Normal Absorbance')
        plt.xlabel('Wavelength / nm')
        plt.ylabel('Absorbance')
        plt.axhline(0, color='black', linewidth=0.5, linestyle='--')
        plt.legend()
        plt.grid()
        plt.tight_layout()
        plt.show()

    def overlay_selected_pulses(self):
        selected_items = self.pulse_list.selectedItems()
        plt.figure(figsize=(6, 4))
        
        x_dark_ref, y_dark_ref = self.parse_data(self.text_boxes['DARK_ref'].toPlainText())
        _, y_dark_sig = self.parse_data(self.text_boxes['DARK_sig'].toPlainText())
        _, y_ref = self.parse_data(self.text_boxes['ref'].toPlainText())
        _, y_sig = self.parse_data(self.text_boxes['sig'].toPlainText())
        _, y_ref_p = self.parse_data(self.text_boxes['ref_p'].toPlainText())
        _, y_sig_p = self.parse_data(self.text_boxes['sig_p'].toPlainText())

        plt.grid(True)
        for item in selected_items:
            pulse_value = item.text()
            if pulse_value in self.pulse_data:
                pulse_data = self.pulse_data[pulse_value]
                x_dark_ref_pulse, y_dark_ref = self.parse_data(pulse_data['DARK_ref'])
                _, y_dark_sig = self.parse_data(pulse_data['DARK_sig'])
                _, y_ref = self.parse_data(pulse_data['ref'])
                _, y_sig = self.parse_data(pulse_data['sig'])
                _, y_ref_p = self.parse_data(pulse_data['ref_p'])
                _, y_sig_p = self.parse_data(pulse_data['sig_p'])

                log_values = []
                for ref_p_real, sig_real, sig_p_real, ref_real in zip(y_ref_p, y_sig, y_sig_p, y_ref):
                    if ref_real != 0 and sig_p_real != 0:
                        log_value = np.log((ref_p_real * sig_real) / (sig_p_real * ref_real))
                        log_values.append(log_value)
                    else:
                        log_values.append(np.nan)
                #移動平均の窓幅をfilter_widthに指定
                #log_values = np.convolve(log_values, np.ones(filter_width)/filter_width, mode='same')
                #color_deltaの値は追加順に薄くしていく。色はRGBで指定
                color_deltaabs = (1-list(self.pulse_data.keys()).index(pulse_value) / len(self.pulse_data.keys()) * 0.9,0,0)
                plt.plot(x_dark_ref_pulse, log_values, label=pulse_value, color=color_deltaabs, linestyle='-', linewidth=1.5)

        plt.xlabel('Wavelength / nm')
        plt.ylabel('ΔAbs')
        plt.legend()
        plt.title('Transient Absorption Spectra')
        plt.tight_layout()
        plt.show()

    def save_data(self):
        data = {label: self.text_boxes[label].toPlainText() for label in self.text_boxes}
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getSaveFileName(self, "データを保存", "", "JSON Files (*.json);;All Files (*)", options=options)
        
        if file_path:
            with open(file_path, 'w') as f:
                json.dump(data, f)
            print("データが保存されました:", file_path)
            #self.display_message(f"データが保存されました: {file_path}")

    def save_all_data(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getSaveFileName(self, "All Dataを保存", "", "JSON Files (*.json);;All Files (*)", options=options)
        
        if file_path:
            with open(file_path, 'w') as f:
                json.dump(self.pulse_data, f)
            print("すべてのパルスデータが保存されました:", file_path)
            #self.display_message(f"すべてのパルスデータが保存されました: {file_path}")

    def load_data(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(self, "データを読み込む", "", "JSON Files (*.json);;All Files (*)", options=options)
        
        if file_path:
            with open(file_path, 'r') as f:
                data = json.load(f)
            for label, content in data.items():
                if label in self.text_boxes:
                    self.text_boxes[label].setPlainText(content)
            print("データが読み込まれました:", file_path)
            #self.display_message(f"データが読み込まれました: {file_path}")

    def load_all_data(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(self, "すべてのデータを読み込む", "", "JSON Files (*.json);;All Files (*)", options=options)
        
        if file_path:
            with open(file_path, 'r') as f:
                self.pulse_data = json.load(f)
            self.update_pulse_list()
            print("すべてのパルスデータが読み込まれました:", file_path)
            #self.display_message(f"すべてのパルスデータが読み込まれました: {file_path}")

    def save_data_to_excel(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getSaveFileName(self, "エクセルファイルを保存", "", "Excel Files (*.xlsx);;All Files (*)", options=options)
        #エクセルファイルを保存する
        if file_path:
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                matome_data = {'Wavelength': None}  # Matomeシート用のデータを収集する辞書
                for pulse, data in self.pulse_data.items():
                    x_dark_ref, y_dark_ref = self.parse_data(data['DARK_ref'])
                    _, y_dark_sig = self.parse_data(data['DARK_sig'])
                    _, y_ref = self.parse_data(data['ref'])
                    _, y_sig = self.parse_data(data['sig'])
                    _, y_ref_p = self.parse_data(data['ref_p'])
                    _, y_sig_p = self.parse_data(data['sig_p'])
                    #データフレームを作成
                    df = pd.DataFrame({
                        'Wavelength': x_dark_ref,
                        'DARK_ref': y_dark_ref,
                        'DARK_sig': y_dark_sig,
                        'ref': y_ref,
                        'sig': y_sig,
                        'ref_p': y_ref_p,
                        'sig_p': y_sig_p
                    })
                    #ΔAbsを計算
                    log_values = []
                    for ref_p_real, sig_real, sig_p_real, ref_real in zip(y_ref_p, y_sig, y_sig_p, y_ref):
                        if ref_real != 0 and sig_p_real != 0:
                            log_value = np.log((ref_p_real * sig_real) / (sig_p_real * ref_real))
                            log_values.append(log_value)
                        else:
                            log_values.append(np.nan)
                    df['ΔAbs'] = log_values  # ΔAbsをデータフレームに追加

                    # Matomeシート用のデータを収集
                    if matome_data['Wavelength'] is None:
                        matome_data['Wavelength'] = x_dark_ref
                    matome_data[pulse] = log_values

                    #データフレームをエクセルファイルに保存
                    df.to_excel(writer, sheet_name=pulse, index=False)

                # Matomeシートにデータをまとめて保存
                if matome_data:
                    matome_df = pd.DataFrame(matome_data)
                    matome_df.to_excel(writer, sheet_name='ΔAbs summary', index=False)

                                #with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
    def load_selected_pulse_data(self):
        selected_items = self.pulse_list.selectedItems()
        if len(selected_items) == 1:  # 選択されたアイテムが1つだけの場合
            pulse_value = selected_items[0].text()  # 最初の選択されたアイテムのテキストを取得
            if pulse_value in self.pulse_data:
                data = self.pulse_data[pulse_value]
                for label, content in data.items():
                    if label in self.text_boxes:
                        self.text_boxes[label].setPlainText(content)  # テキストボックスにデータを設定
                        self.text_boxes[label].setStyleSheet("color: black;")  # フォント色を黒に設定
                        self.original_data[label] = content  # 元のデータを保持
                print(f"Pulse {pulse_value}のデータが読み込まれました。")
                #self.display_message(f"Pulse {pulse_value}のデータが読み込まれました。")

    def update_pulse_list(self):
        self.pulse_list.clear()
        for pulse in self.pulse_data.keys():
            self.pulse_list.addItem(pulse)
            #self.display_message(f"Pulse {pulse}がリストに追加されました。")

    def save_pulse_data(self):
        pulse_value = self.pulse_input.toPlainText().strip()
        if pulse_value:
            self.pulse_data[pulse_value] = {
                'DARK_ref': self.text_boxes['DARK_ref'].toPlainText(),
                'DARK_sig': self.text_boxes['DARK_sig'].toPlainText(),
                'ref': self.text_boxes['ref'].toPlainText(),
                'sig': self.text_boxes['sig'].toPlainText(),
                'ref_p': self.text_boxes['ref_p'].toPlainText(),
                'sig_p': self.text_boxes['sig_p'].toPlainText()
            }
            self.update_pulse_list()
            self.plot_delta_abs()  # ΔAbsのグラフを更新
            print(f"Pulse {pulse_value}が保存されました。")
            #self.display_message(f"Pulse {pulse_value}が保存されました。")
        else:
            print("パルス名が空です。")
            #self.display_message("パルス名が空です。")

    def plot_delta_abs(self):
        # グラフをクリア
        plt.clf()
        
        # 保存されたすべてのパルスデータに対してΔAbsを計算
        #グラフサイズをウィジェットに合わせる
        plt.figure(figsize=(5, 2))
        for pulse_name, data in self.pulse_data.items():
            x_dark_ref, y_dark_ref = self.parse_data(data['DARK_ref'])
            _, y_dark_sig = self.parse_data(data['DARK_sig'])
            _, y_ref = self.parse_data(data['ref'])
            _, y_ref_p = self.parse_data(data['ref_p'])
            _, y_sig = self.parse_data(data['sig'])
            _, y_sig_p = self.parse_data(data['sig_p'])

            log_values = []
            for ref_p_real, sig_real, sig_p_real, ref_real in zip(y_ref_p, y_sig, y_sig_p, y_ref):
                if ref_real != 0 and sig_p_real != 0:
                    log_value = np.log((ref_p_real * sig_real) / (sig_p_real * ref_real))
                    log_values.append(log_value)
                else:
                    log_values.append(np.nan)
            #移動平均の窓幅をfilter_widthに指定
            #log_values = np.convolve(log_values, np.ones(filter_width)/filter_width, mode='same')

            # ΔAbsのプロット
            #color_deltaの値は追加順に薄くしていく。色はRGBで指定
            color_delta = (1 - list(self.pulse_data.keys()).index(pulse_name) / len(self.pulse_data.keys()) * 0.8, 0, 0)
            plt.plot(x_dark_ref, log_values, label=pulse_name, color=color_delta, alpha=0.7, linestyle='-', linewidth=1)
            plt.xlim(400,750)
            plt.ylim(min(log_values)-0.02, max(log_values)+0.02 if len(log_values) > 0 else 1)            #目盛りのフォントサイズを変更
            plt.tick_params(labelsize=8)
            #横軸の目盛りの位置を縦軸の0に合わせる
            plt.gca().spines['bottom'].set_position(('data', 0))

            plt.tight_layout()            

        plt.grid()
        
        # グラフをQLabelに表示
        temp_file_path = 'temp_abs.png'
        plt.savefig(temp_file_path, bbox_inches='tight', pad_inches=0)
        plt.close()
        self.abs_graph_widget.setPixmap(QPixmap(temp_file_path))
        os.remove(temp_file_path)

# アプリケーションの起動
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = DataGraphApp()
    window.show()
    sys.exit(app.exec_())
