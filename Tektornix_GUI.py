import sys
import matplotlib.pyplot as plt
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QLabel, QVBoxLayout, QWidget
)
from PyQt6.QtCore import Qt
from matplotlib.backends.backend_qtagg import (
    FigureCanvasQTAgg as FigureCanvas,
    NavigationToolbar2QT as NavigationToolbar
)

class DropLabel(QLabel):
    def __init__(self, main_window):
        super().__init__('Drop file here')
        self.main_window = main_window
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.setAcceptDrops(True)
        self.setStyleSheet('border: 2px dashed #aaa; padding: 20px; font-size: 16px;')

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        for url in event.mimeData().urls():
            self.main_window.load_and_plot(url.toLocalFile())

class PlotWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Tektronix 2ch')
        self.resize(1000, 700)
        self.label = DropLabel(self)
        self.figure = plt.figure()
        self.canvas = FigureCanvas(self.figure)
        self.toolbar = NavigationToolbar(self.canvas, self)

        layout = QVBoxLayout()
        layout.addWidget(self.label)
        layout.addWidget(self.toolbar)
        layout.addWidget(self.canvas)
        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def load_and_plot(self, filepath):
        try:
            t1, v1, t2, v2 = self.parse_tektronix_2ch_file(filepath)
            self.figure.clear()
            # 2段サブプロット
            ax1 = self.figure.add_subplot(2, 1, 1)
            ax2 = self.figure.add_subplot(2, 1, 2, sharex=ax1)
            ax1.plot(t1, v1, label='CH1')
            ax1.set_ylabel('CH1')
            ax1.grid(True)
            ax1.legend(loc='upper right')
            ax2.plot(t2, v2, label='CH2', color='orange')
            ax2.set_ylabel('CH2')
            ax2.set_xlabel('Time (s)')
            ax2.grid(True)
            ax2.legend(loc='upper right')
            self.figure.tight_layout()
            self.canvas.draw()
            self.label.setText(f'Loading Successful!: {filepath}')
        except Exception as e:
            self.label.setText(f'Error: {str(e)}')

    def parse_tektronix_2ch_file(self, filepath):
        t1, v1, t2, v2 = [], [], [], []
        with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
            for line in f:
                parts = line.strip().split()
                # 1行に4つ数字があれば両chとして扱う
                if len(parts) >= 4:
                    try:
                        t1.append(float(parts[0]))
                        v1.append(float(parts[1]))
                        t2.append(float(parts[2]))
                        v2.append(float(parts[3]))
                    except ValueError:
                        continue
        if not t1 or not t2:
            raise ValueError("No Data。")
        return t1, v1, t2, v2

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = PlotWindow()
    window.show()
    sys.exit(app.exec())