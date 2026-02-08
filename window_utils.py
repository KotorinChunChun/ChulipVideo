"""Windows API操作用ユーティリティモジュール."""
import ctypes
import ctypes.wintypes
import os
from typing import List, Tuple, Optional, Dict, Any

import cv2
import mss
import numpy as np


class BITMAPINFOHEADER(ctypes.Structure):
    _fields_ = [
        ('biSize', ctypes.c_uint32),
        ('biWidth', ctypes.c_int32),
        ('biHeight', ctypes.c_int32),
        ('biPlanes', ctypes.c_uint16),
        ('biBitCount', ctypes.c_uint16),
        ('biCompression', ctypes.c_uint32),
        ('biSizeImage', ctypes.c_uint32),
        ('biXPelsPerMeter', ctypes.c_int32),
        ('biYPelsPerMeter', ctypes.c_int32),
        ('biClrUsed', ctypes.c_uint32),
        ('biClrImportant', ctypes.c_uint32)
    ]


class WindowUtils:
    """Windows APIを使用した操作をまとめたクラス."""

    # 定数
    DWMWA_EXTENDED_FRAME_BOUNDS = 9
    PROCESS_QUERY_LIMITED_INFORMATION = 0x1000

    def __init__(self):
        self.sct = mss.mss()

    def get_monitor_info(self) -> List[Dict[str, Any]]:
        """モニター情報を取得する."""
        # sct.monitors[0] は全画面結合なので除外する
        return self.sct.monitors[1:]

    def enum_windows(self, filter_text: str = "") -> List[Tuple[Any, str]]:
        """可視ウィンドウの一覧を取得する.
        
        Args:
            filter_text: 検索フィルタ (小文字で比較)。タイトルまたはプロセス名に一致するものを返す。
            
        Returns:
            List of (hwnd, title, process_name, pid) tuples, sorted by process name then title.
        """
        windows = []
        filter_text = filter_text.lower() if filter_text else ""

        def enum_windows_proc(hwnd, lParam):
            if ctypes.windll.user32.IsWindowVisible(hwnd):
                length = ctypes.windll.user32.GetWindowTextLengthW(hwnd)
                if length > 0:
                    buff = ctypes.create_unicode_buffer(length + 1)
                    ctypes.windll.user32.GetWindowTextW(hwnd, buff, length + 1)
                    title = buff.value
                    if title and title != "録画ツール":
                        pname = self.get_process_name(hwnd)
                        pid = ctypes.c_ulong()
                        ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
                        
                        # 検索フィルタ適用 (タイトル または プロセス名)
                        if not filter_text or (filter_text in title.lower()) or (filter_text in pname.lower()):
                            windows.append((hwnd, title, pname, pid.value))
            return True

        ctypes.WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_int, ctypes.c_int)
        ctypes.windll.user32.EnumWindows(ctypes.WNDENUMPROC(enum_windows_proc), 0)
        
        # ソート: プロセス名 -> タイトル
        windows.sort(key=lambda x: (x[2].lower(), x[1].lower()))
        
        return windows

    def get_process_name(self, hwnd: Any) -> str:
        """hwndからプロセス名を取得する."""
        pid = ctypes.c_ulong()
        ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        
        h_process = ctypes.windll.kernel32.OpenProcess(self.PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
        if h_process:
            size = ctypes.c_uint32(260)
            buff = ctypes.create_unicode_buffer(size.value)
            if ctypes.windll.kernel32.QueryFullProcessImageNameW(h_process, 0, buff, ctypes.byref(size)):
                full_path = buff.value
                ctypes.windll.kernel32.CloseHandle(h_process)
                return os.path.basename(full_path)
            ctypes.windll.kernel32.CloseHandle(h_process)
        return "Unknown"

    def get_window_rect(self, hwnd: Any) -> Optional[Dict[str, int]]:
        """ウィンドウの正確な矩形を取得する (DWM使用)."""
        rect = ctypes.wintypes.RECT()
        res = ctypes.windll.dwmapi.DwmGetWindowAttribute(
            hwnd, self.DWMWA_EXTENDED_FRAME_BOUNDS, ctypes.byref(rect), ctypes.sizeof(rect)
        )
        if res == 0:
            w = rect.right - rect.left
            h = rect.bottom - rect.top
            return {'top': rect.top, 'left': rect.left, 'width': w, 'height': h}
        else:
            # フォールバック
            if ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(rect)):
                return {'top': rect.top, 'left': rect.left, 'width': rect.right - rect.left, 'height': rect.bottom - rect.top}
        return None

    def get_window_borders(self, hwnd: Any) -> Dict[str, int]:
        """ウィンドウの不可視枠（影など）のサイズを取得する."""
        rect_total = ctypes.wintypes.RECT()
        ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(rect_total))
        
        rect_visual = ctypes.wintypes.RECT()
        res = ctypes.windll.dwmapi.DwmGetWindowAttribute(
            hwnd, self.DWMWA_EXTENDED_FRAME_BOUNDS, ctypes.byref(rect_visual), ctypes.sizeof(rect_visual)
        )
        
        if res == 0:
            return {
                'left': rect_visual.left - rect_total.left,
                'top': rect_visual.top - rect_total.top,
                'right': rect_total.right - rect_visual.right,
                'bottom': rect_total.bottom - rect_visual.bottom,
                'width_diff': (rect_total.right - rect_total.left) - (rect_visual.right - rect_visual.left),
                'height_diff': (rect_total.bottom - rect_total.top) - (rect_visual.bottom - rect_visual.top)
            }
        return {'left': 0, 'top': 0, 'right': 0, 'bottom': 0, 'width_diff': 0, 'height_diff': 0}

    def set_window_position(self, hwnd: Any, x: int, y: int, width: int, height: int) -> bool:
        """ウィンドウの位置とサイズを変更する (視覚的な位置・サイズを指定)."""
        # 補正値を計算
        borders = self.get_window_borders(hwnd)
        
        # 指定された視覚的な座標・サイズから、設定すべき実際のウィンドウ座標・サイズ（枠含む）を計算
        # Target Visual X = Target Total X + Border Left
        # -> Target Total X = Target Visual X - Border Left
        real_x = x - borders['left']
        real_y = y - borders['top']
        real_w = width + borders['width_diff']
        real_h = height + borders['height_diff']
        
        # SWP_NOZORDER (0x0004) | SWP_NOACTIVATE (0x0010)
        uFlags = 0x0004 | 0x0010
        return ctypes.windll.user32.SetWindowPos(hwnd, 0, real_x, real_y, real_w, real_h, uFlags) != 0

    def set_window_display_affinity(self, hwnd: Any, exclude: bool = True) -> bool:
        """ウィンドウをキャプチャから除外するかどうかを設定する (Windows 10 2004+)."""
        # WDA_NONE = 0x00, WDA_MONITOR = 0x01, WDA_EXCLUDEFROMCAPTURE = 0x11
        WDA_EXCLUDEFROMCAPTURE = 0x11
        WDA_NONE = 0x00
        affinity = WDA_EXCLUDEFROMCAPTURE if exclude else WDA_NONE
        try:
            return ctypes.windll.user32.SetWindowDisplayAffinity(hwnd, affinity) != 0
        except Exception as e:
            print(f"SetWindowDisplayAffinity error: {e}")
            return False

    def capture_exclusive_window(self, hwnd: Any) -> Optional[np.ndarray]:
        """PrintWindow を使用して重なりを無視してウィンドウをキャプチャし、余白をクロップする."""
        try:
            # 視覚的な矩形 (DWM)
            rect_visual = self.get_window_rect(hwnd)
            if not rect_visual:
                return None
            
            # 全体の矩形 (影などを含む実際の設定座標)
            rect_total = ctypes.wintypes.RECT()
            if not ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(rect_total)):
                return None
            
            total_w = rect_total.right - rect_total.left
            total_h = rect_total.bottom - rect_total.top
            if total_w <= 0 or total_h <= 0: return None

            # 視覚的な矩形へのオフセットを計算 (PrintWindowの結果は全体の矩形基準)
            offset_x = rect_visual['left'] - rect_total.left
            offset_y = rect_visual['top'] - rect_total.top
            visual_w = rect_visual['width']
            visual_h = rect_visual['height']

            # DC作成
            hdc_screen = ctypes.windll.user32.GetDC(0)
            hdc_mem = ctypes.windll.gdi32.CreateCompatibleDC(hdc_screen)
            hbmp = ctypes.windll.gdi32.CreateCompatibleBitmap(hdc_screen, total_w, total_h)
            ctypes.windll.gdi32.SelectObject(hdc_mem, hbmp)

            # PW_RENDERFULLCONTENT (2) で描画
            ctypes.windll.user32.PrintWindow(hwnd, hdc_mem, 2)

            # Bitmap から numpy 配列 (BGR) への変換
            bi = BITMAPINFOHEADER()
            bi.biSize = ctypes.sizeof(BITMAPINFOHEADER)
            bi.biWidth = total_w
            bi.biHeight = -total_h  # Top-down
            bi.biPlanes = 1
            bi.biBitCount = 32
            bi.biCompression = 0 # BI_RGB

            buffer = ctypes.create_string_buffer(total_w * total_h * 4)
            ctypes.windll.gdi32.GetDIBits(hdc_mem, hbmp, 0, total_h, buffer, ctypes.byref(bi), 0)

            # クリーンアップ
            ctypes.windll.gdi32.DeleteObject(hbmp)
            ctypes.windll.gdi32.DeleteDC(hdc_mem)
            ctypes.windll.user32.ReleaseDC(0, hdc_screen)

            # buffer -> numpy
            raw_data = np.frombuffer(buffer, dtype=np.uint8)
            frame_bgra = raw_data.reshape((total_h, total_w, 4))
            frame_bgr = cv2.cvtColor(frame_bgra, cv2.COLOR_BGRA2BGR)
            
            # 視覚的な矩形に合わせてクロップ (ズレと白線の解消)
            # offset が負になることは通常ないが、クリップしておく
            y1 = max(0, offset_y)
            y2 = min(total_h, y1 + visual_h)
            x1 = max(0, offset_x)
            x2 = min(total_w, x1 + visual_w)
            
            cropped = frame_bgr[y1:y2, x1:x2]

            # 白線対策: クロップ後の最上部1pxを強制的に黒で塗りつぶす
            # PrintWindowの境界アーティファクト対策 (視覚的な上端に適用)
            if cropped.shape[0] > 0:
                cropped[0, :] = 0
            
            return cropped
        except Exception as e:
            print(f"Exclusive capture error: {e}")
            return None

    def check_single_instance(self, mutex_name: str, window_title: str) -> bool:
        """二重起動チェック。既に起動している場合はTrueを返す. (戻り値の意味を逆転させないように注意) -> Falseなら起動可"""
        # Mutex作成
        self.mutex = ctypes.windll.kernel32.CreateMutexW(None, False, mutex_name)
        last_error = ctypes.windll.kernel32.GetLastError()
        
        if last_error == 183: # ERROR_ALREADY_EXISTS
            # 既存ウィンドウを探す
            hwnd = ctypes.windll.user32.FindWindowW(None, window_title)
            if hwnd:
                # 最小化していたら元に戻す
                ctypes.windll.user32.ShowWindow(hwnd, 9) # SW_RESTORE
                # 前面に表示
                ctypes.windll.user32.SetForegroundWindow(hwnd)
            return False
        return True

