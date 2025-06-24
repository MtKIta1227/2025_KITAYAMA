# Flame を External Level Trigger で 100 スキャン
from seabreeze.spectrometers import Spectrometer
import numpy as np

# 1) 初期化
spec = Spectrometer.from_first_available()   # 複数台なら .list_devices() で ID を確認
spec.integration_time_micros(10_000)        # 10 ms = 10,000 µs

# 2) Trigger モード設定（Level）
spec.trigger_mode(3)                         # モード番号は機種で要確認

print("準備完了。外部トリガ信号を送り込みます…")

# 3) 100 スキャン連続取得
wavelengths = spec.wavelengths()             # 波長軸は固定なので先に取得
spectra = np.empty((100, len(wavelengths)), dtype=float)

for i in range(100):
    intens = spec.intensities()              # 信号が来るまでブロッキング
    spectra[i] = intens
    if i % 10 == 0:
        print(f"{i+1} / 100 スキャン完了")

# 4) 後処理・保存など
np.savez("flame_level_100scans.npz",
         wavelengths=wavelengths,
         spectra=spectra)

spec.close()
print("測定終了。お疲れさまでした！")
