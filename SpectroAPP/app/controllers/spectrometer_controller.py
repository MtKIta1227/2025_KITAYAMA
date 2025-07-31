from oceandirect.OceanDirectAPI import OceanDirectAPI, OceanDirectError

class SpectrometerController:
    """分光器のハードウェア制御を専門に行うクラス"""

    def __init__(self):
        self.api = None
        self.spectrometer = None
        self.wavelengths = []
        self.integration_time_limits = (0, 0)

    def initialize_api(self):
        """OceanDirect APIを初期化する"""
        if self.api is None:
            self.api = OceanDirectAPI()

    def find_devices(self) -> list:
        """利用可能なUSBデバイスを検索し、そのIDリストを返す"""
        self.initialize_api()
        try:
            num_devices = self.api.find_usb_devices()
            if num_devices > 0:
                return self.api.get_device_ids()
            return []
        except OceanDirectError as e:
            print(f"デバイス検索エラー: {e}")
            return []

    def connect(self, device_id: int):
        """指定されたIDのデバイスに接続し、基本情報を取得する"""
        if self.is_connected:
            raise OceanDirectError("既に別のデバイスに接続されています。")
        
        self.initialize_api()
        self.spectrometer = self.api.open_device(device_id)
        self.wavelengths = self.spectrometer.get_wavelengths()
        min_integ = self.spectrometer.get_minimum_integration_time()
        max_integ = self.spectrometer.get_maximum_integration_time()
        self.integration_time_limits = (min_integ, max_integ)

    def disconnect(self):
        """デバイスとの接続を切断する"""
        if self.is_connected:
            try:
                self.spectrometer.close_device()
            except OceanDirectError as e:
                print(f"切断エラー: {e}")
            finally:
                self.spectrometer = None
                self.wavelengths = []
                self.integration_time_limits = (0, 0)

    @property
    def is_connected(self) -> bool:
        """接続状態を返すプロパティ"""
        return self.spectrometer is not None

    def set_integration_time(self, microseconds: int):
        """積分時間を設定する"""
        if not self.is_connected:
            raise OceanDirectError("分光器が接続されていません。")
        
        micros = max(self.integration_time_limits[0], microseconds)
        micros = min(self.integration_time_limits[1], micros)
        self.spectrometer.set_integration_time(micros)

    def acquire_spectrum(self) -> list[float]:
        """スペクトルデータを1フレーム取得して返す"""
        if not self.is_connected:
            raise OceanDirectError("分光器が接続されていません。")
        return self.spectrometer.get_formatted_spectrum()

    def shutdown(self):
        """APIをクリーンにシャットダウンする"""
        self.disconnect()
        if self.api:
            try:
                self.api.shutdown()
            except OceanDirectError as e:
                print(f"APIシャットダウンエラー: {e}")