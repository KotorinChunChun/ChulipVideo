"""ユーティリティ関数モジュール.

時間変換、画像保存、アスペクト比計算などの汎用ヘルパー関数を提供。
"""
from __future__ import annotations

import os
import subprocess
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
    # srcフォルダの1つ上（プロジェクトルート）を返す
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def resource_path(relative_path: str) -> str:
    """リソースへの絶対パスを返す（PyInstallerの1ファイル化に対応）.
    
    PyInstallerで1ファイル化した際は一時フォルダ(_MEIPASS)を参照し、
    通常実行時は実行ディレクトリを参照する。
    """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = get_base_dir()

    return os.path.join(base_path, relative_path)


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


def open_folder_with_selection(path: str) -> None:
    """ファイルをハイライト（選択）した状態でフォルダを開く."""
    if not path or not os.path.exists(path):
        return

    path = os.path.abspath(path)
    try:
        if os.name == 'nt':
            # Windows: エクスプローラーでファイルを選択状態にする
            # フォルダの場合はそのフォルダ自体が選択される
            norm_path = os.path.normpath(path)
            subprocess.Popen(['explorer', f'/select,{norm_path}'])
        elif sys.platform == 'darwin':
            # macOS: Finderで表示（選択状態）
            subprocess.run(['open', '-R', path], check=True)
        else:
            # Linux等 (選択状態にする標準的な方法がないため、単に開く)
            parent = os.path.dirname(path) if os.path.isfile(path) else path
            subprocess.run(['xdg-open', parent], check=True)
    except Exception as e:
        print(f"Error opening folder: {e}")


import numpy as np

def get_safe_crop(img: np.ndarray, rect: list[int] | tuple[int, int, int, int], bg_color: tuple[int, int, int] = (0, 0, 0)) -> np.ndarray:
    """矩形範囲で画像をクロップする。画像範囲外は bg_color で埋める."""
    if img is None:
        return None
        
    x1, y1, x2, y2 = rect
    
    # 矩形の幅・高さ
    w = x2 - x1
    h = y2 - y1
    
    if w <= 0 or h <= 0:
        return np.zeros((1, 1, 3), dtype=np.uint8)
        
    img_h, img_w = img.shape[:2]
    
    # 画像の有効範囲との共通部分を計算
    ix1 = max(0, x1)
    iy1 = max(0, y1)
    ix2 = min(img_w, x2)
    iy2 = min(img_h, y2)
    
    # 出力画像（背景色で初期化）
    # OpenCV は BGR なので bg_color も BGR 想定だが、引数はタプルで渡される
    out_img = np.full((h, w, 3), bg_color, dtype=np.uint8)
    
    # 有効範囲が存在する場合のみコピー
    if ix1 < ix2 and iy1 < iy2:
        crop = img[iy1:iy2, ix1:ix2]
        
        # 出力画像上の配置位置
        ox1 = ix1 - x1
        oy1 = iy1 - y1
        ox2 = ox1 + (ix2 - ix1)
        oy2 = oy1 + (iy2 - iy1)
        
        out_img[oy1:oy2, ox1:ox2] = crop
        
    return out_img
