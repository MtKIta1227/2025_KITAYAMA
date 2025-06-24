#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
check_spectrometers.py
────────────────────────────────────────────────────────
PCに接続されたOcean Insight分光器を自動検出して一覧表示。

表示項目：
- 検出数
- シリアル番号（serial_number）
- モデル名（model）
- ピクセル数（pixels）
- ファームウェアバージョン（firmware_version）
- USBプロトコル（usb_protocol）
"""

from seabreeze.spectrometers import list_devices, Spectrometer
import sys

def main():
    devices = list_devices()

    print(f"検出された分光器の数: {len(devices)} 台\n")

    if not devices:
        sys.exit("分光器が検出されませんでした。USB接続やドライバを確認してください。")

    for idx, dev in enumerate(devices, 1):
        try:
            spec = Spectrometer.from_serial_number(dev.serial_number)
            wavelengths = spec.wavelengths()
            print(f"【分光器 #{idx}】")
            print(f"  シリアル番号      : {spec.serial_number}")
            print(f"  モデル名          : {spec.model}")
            print(f"  ピクセル数        : {len(wavelengths)}")
            print(f"  ファームウェア    : {spec.firmware_version}")
            print(f"  USBプロトコル     : {spec.usb_protocol}")
            print("─────────────────────────────────")
            spec.close()
        except Exception as e:
            print(f"分光器 {dev.serial_number} への接続時エラー: {e}")

if __name__ == "__main__":
    main()
