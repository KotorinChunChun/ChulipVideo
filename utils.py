"""ユーティリティ関数モジュール.

時間変換、画像保存、アスペクト比計算などの汎用ヘルパー関数を提供。
"""
from __future__ import annotations

import os
import sys
from math import gcd

import cv2


def get_base_dir() -> str:
    """実行ファイルまたはスクリプトのベースディレクトリを返す.
    
    PyInstallerでバンドルした場合は実行ファイルの場所に、
    通常実行時はスクリプトの場所にする。
    """
    try:
        if getattr(sys, 'frozen', False):
            return os.path.dirname(sys.executable)
    except Exception:
        pass
    return os.path.dirname(os.path.abspath(__file__))


def sec_to_hhmmss(sec: float) -> str:
    """秒数をHHMMSS形式の文字列に変換する（ファイル名用）."""
    sec = int(sec)
    h = sec // 3600
    m = (sec % 3600) // 60
    s = sec % 60
    return f"{h:02d}{m:02d}{s:02d}"


def sec_to_display(sec: float) -> str:
    """秒数を表示用 HH:MM:SS.mmm 形式に変換する."""
    try:
        sec_f = float(sec)
    except Exception:
        sec_f = 0.0
    h = int(sec_f) // 3600
    m = (int(sec_f) % 3600) // 60
    s_int = int(sec_f % 60)
    ms = int((sec_f - int(sec_f)) * 1000)
    return f"{h:02d}:{m:02d}:{s_int:02d}.{ms:03d}"


def hhmmss_to_sec(time_str: str) -> float:
    """HH:MM:SS形式の文字列を秒数に変換する.
    
    対応する書式: HH:MM:SS.sss, MM:SS.sss, SS.sss, または単純な秒数（小数可）
    """
    try:
        if ":" in time_str:
            parts = time_str.split(":")
            parts = [p.strip() for p in parts if p.strip() != ""]
            if len(parts) == 3:
                h = int(parts[0])
                m = int(parts[1])
                s = float(parts[2])
                return h * 3600 + m * 60 + s
            elif len(parts) == 2:
                m = int(parts[0])
                s = float(parts[1])
                return m * 60 + s
            elif len(parts) == 1:
                return float(parts[0])
        else:
            return float(time_str)
    except Exception:
        pass
    return 0.0


def imwrite_jp(filename: str, img, params=None) -> bool:
    """日本語パス対応の画像保存関数."""
    try:
        ext = os.path.splitext(filename)[1]
        result, n = cv2.imencode(ext, img, params)

        if result:
            with open(filename, mode='w+b') as f:
                n.tofile(f)
            return True
        return False
    except Exception as e:
        print(f"Save Error: {e}")
        return False


def ratio_value_from_str(rstr: str) -> float | None:
    """アスペクト比文字列（例: '16:9'）を数値に変換する."""
    try:
        if isinstance(rstr, str) and ':' in rstr:
            a, b = rstr.split(':')
            return float(a) / float(b)
    except Exception:
        pass
    return None


def ratio_label_from_wh(w: int, h: int) -> str:
    """幅と高さからアスペクト比ラベルを生成する."""
    try:
        r = float(w) / float(h)
    except Exception:
        return '?:?'
    
    # 一般的なアスペクト比との照合
    candidates = [
        ('16:9', 16.0 / 9.0),
        ('9:16', 9.0 / 16.0),
        ('4:3', 4.0 / 3.0),
        ('3:4', 3.0 / 4.0),
        ('21:9', 21.0 / 9.0),
        ('1:1', 1.0),
        ('4:5', 4.0 / 5.0),
        ('5:4', 5.0 / 4.0),
    ]
    
    best = None
    best_diff = 1.0
    for label, val in candidates:
        diff = abs(r - val)
        if diff < best_diff:
            best_diff = diff
            best = label
    
    if best is not None and best_diff <= 0.03:
        return best
    
    try:
        g = gcd(int(w), int(h))
        return f"{int(w // g)}:{int(h // g)}"
    except Exception:
        return f"{w}:{h}"
