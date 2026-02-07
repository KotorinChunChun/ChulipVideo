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
            filter_text: 検索フィルタ (小文字で比較)
            
        Returns:
            List of (hwnd, title) tuples
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
                        # 検索フィルタ適用
                        if not filter_text or filter_text in title.lower():
                            windows.append((hwnd, title))
            return True

        ctypes.WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_int, ctypes.c_int)
        ctypes.windll.user32.EnumWindows(ctypes.WNDENUMPROC(enum_windows_proc), 0)
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

    def capture_exclusive_window(self, hwnd: Any) -> Optional[np.ndarray]:
        """PrintWindow を使用して重なりを無視してウィンドウをキャプチャする."""
        try:
            rect = self.get_window_rect(hwnd)
            if not rect:
                return None
                
            w = rect['width']
            h = rect['height']
            if w <= 0 or h <= 0: return None

            # DC作成
            hdc_screen = ctypes.windll.user32.GetDC(0)
            hdc_mem = ctypes.windll.gdi32.CreateCompatibleDC(hdc_screen)
            hbmp = ctypes.windll.gdi32.CreateCompatibleBitmap(hdc_screen, w, h)
            ctypes.windll.gdi32.SelectObject(hdc_mem, hbmp)

            # PW_RENDERFULLCONTENT (2) で描画
            ctypes.windll.user32.PrintWindow(hwnd, hdc_mem, 2)

            # Bitmap から numpy 配列 (BGR) への変換
            bi = BITMAPINFOHEADER()
            bi.biSize = ctypes.sizeof(BITMAPINFOHEADER)
            bi.biWidth = w
            bi.biHeight = -h  # Top-down
            bi.biPlanes = 1
            bi.biBitCount = 32
            bi.biCompression = 0 # BI_RGB

            buffer = ctypes.create_string_buffer(w * h * 4)
            ctypes.windll.gdi32.GetDIBits(hdc_mem, hbmp, 0, h, buffer, ctypes.byref(bi), 0)

            # クリーンアップ
            ctypes.windll.gdi32.DeleteObject(hbmp)
            ctypes.windll.gdi32.DeleteDC(hdc_mem)
            ctypes.windll.user32.ReleaseDC(0, hdc_screen)

            # numpy 配列化 (BGRA)
            frame = np.frombuffer(buffer, dtype=np.uint8).reshape((h, w, 4))
            return cv2.cvtColor(frame, cv2.COLOR_BGRA2BGR)
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
