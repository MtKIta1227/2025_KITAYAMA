"""
SR6 Spectrometer Feature Check

OceanDirect 3.x / Python 3.8+
動作確認モデル : SR-6（他モデルでも利用可）
"""

from oceandirect import OceanDirectAPI, FeatureID, OceanDirectError


def main() -> None:
    """
    1. USB もしくはネットワーク上の分光器をスキャン
    2. 最初に見つかったデバイスを開く
    3. FeatureID 全列挙 → is_feature_id_enabled() で可否を判定
    4. 結果をチェックマーク付きで表示
    5. 後始末
    """
    api = OceanDirectAPI()

    try:
        # --- デバイス検索 ---------------------------------------------------
        dev_count = api.get_number_devices()
        if dev_count == 0:
            print("分光器が見つかりません。USB 接続と電源を確認してください。")
            return

        dev_ids = api.get_device_ids()        # ID の一覧を取得
        spec = api.open_device(dev_ids[0])    # 先頭の 1 台だけ開く

        # --- 基本情報 -------------------------------------------------------
        model_name = spec.get_model()
        serial_no  = spec.get_serial_number()
        print(f"接続機種 : {model_name}  (S/N: {serial_no})\n")

        # --- Feature 一覧 ----------------------------------------------------
        print("=== Feature サポート一覧 ===")
        for fid in FeatureID:
            try:
                supported = spec.is_feature_id_enabled(fid)  # True / False
            except Exception as e:
                # 万一 SDK-enum と FW が不一致で例外が出た場合は未対応とみなす
                supported = False
                print(f"  ! {fid.name} 判定で例外: {e}")

            mark = "✔" if supported else "✖"
            print(f"{mark} {fid.name}")

        # --- 後始末 ----------------------------------------------------------
        api.close_device(spec.device_id)

    except OceanDirectError as exc:
        print(f"OceanDirectError: {exc}")

    finally:
        api.shutdown()


if __name__ == "__main__":
    main()
