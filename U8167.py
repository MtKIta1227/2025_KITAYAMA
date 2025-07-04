# viewer_uint16.py  ──  ワンファイル完結の最小 GUI
import sys, os
import numpy as np
from PyQt5.QtWidgets import (
    QApplication, QWidget, QPushButton, QLabel, QFileDialog,
    QVBoxLayout, QMessageBox
)
from PyQt5.QtGui import QImage, QPixmap

# --- ここを環境に合わせて調整 -------------------------------------
HEIGHT, WIDTH = 639, 479               # 0–638ch × 0–478ch
DTYPE = np.uint16                      # 16bit
HEADER_BYTES = 11084                   # ファイル先頭のヘッダ長
# ------------------------------------------------------------------

def load_raw_image(path: str) -> np.ndarray:
    """raw ファイルを numpy 配列 (H, W) で返す"""
    file_size = os.path.getsize(path)
    expected = HEIGHT * WIDTH * np.dtype(DTYPE).itemsize
    if file_size < HEADER_BYTES + expected:
        raise ValueError("ファイルサイズが想定より小さい")
    with open(path, "rb") as f:
        f.seek(HEADER_BYTES)
        data = np.fromfile(f, dtype=DTYPE, count=HEIGHT * WIDTH)
    return data.reshape((HEIGHT, WIDTH))

def to_qimage(arr: np.ndarray) -> QImage:
    """16bit → 8bit に線形マッピングして QImage へ"""
    # 0–65535 → 0–255
    img8 = np.clip(arr / 256, 0, 255).astype(np.uint8)
    h, w = img8.shape
    return QImage(img8.data, w, h, w, QImage.Format_Grayscale8).copy()

class RawViewer(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Raw Image Viewer (uint16)")
        self.img_label = QLabel("画像なし")
        self.img_label.setAlignment(Qt.AlignCenter)

        load_btn = QPushButton("Load *.img")
        load_btn.clicked.connect(self.load_clicked)

        layout = QVBoxLayout(self)
        layout.addWidget(load_btn)
        layout.addWidget(self.img_label)

    def load_clicked(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Select .img file", "", "Raw Image (*.img);;All Files (*)")
        if not path:
            return
        try:
            arr = load_raw_image(path)
            qim = to_qimage(arr)
            self.img_label.setPixmap(QPixmap.fromImage(qim).scaled(
                self.img_label.width(), self.img_label.height(),
                aspectRatioMode=Qt.KeepAspectRatio))
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))

if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = RawViewer()
    w.resize(600, 800)
    w.show()
    sys.exit(app.exec_())
