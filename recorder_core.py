"""録画のコアロジックを扱うモジュール."""
import ctypes
import ctypes.wintypes
import datetime
import os
import threading
import time
from typing import Optional, Dict, List, Any, Callable

import cv2
import mss
import numpy as np

from window_utils import WindowUtils


class ScreenRecorderLogic:
    """録画処理の実行・管理を行うクラス."""

    def __init__(self, window_utils: WindowUtils):
        self.window_utils = window_utils
        self.is_recording = False
        self.recording_thread: Optional[threading.Thread] = None
        self.stop_event = threading.Event()
        self.trajectory_data: List[tuple] = []
        self.current_out_path = ""

    def start_recording(
        self,
        filepath: str,
        rect: Dict[str, int],
        fps: int,
        hwnd: Optional[Any] = None,
        record_cursor: bool = False,
        record_tsv: bool = True,
        exclusive_window: bool = False
    ):
        """録画を開始する."""
        if self.is_recording:
            return

        self.current_out_path = filepath
        self.stop_event.clear()
        self.is_recording = True
        self.trajectory_data = []

        self.recording_thread = threading.Thread(
            target=self._record_loop,
            args=(filepath, rect, fps, hwnd, record_cursor, record_tsv, exclusive_window)
        )
        self.recording_thread.start()

    def stop_recording(self):
        """録画を停止するリクエストを送る (非同期)."""
        self.stop_event.set()
        self.is_recording = False

    def wait_for_stop(self, timeout: Optional[float] = None):
        """録画スレッドの終了を待機する."""
        if self.recording_thread and self.recording_thread.is_alive():
            self.recording_thread.join(timeout)

    def save_trajectory_tsv(self):
        """記録されたマウス軌跡データをTSVに保存する."""
        if not self.trajectory_data or not self.current_out_path:
            return

        tsv_path = os.path.splitext(self.current_out_path)[0] + ".tsv"
        data_to_save = list(self.trajectory_data)
        
        # 呼び出し元でクリアする責務を持つか、ここでクリアするか。
        # ここではクリアせず、呼び出し元で管理しやすいようにするが、
        # メモリ節約のためクリアして渡すほうが安全かもしれない。
        # 今回はここでコピーを取っているので、呼び出し元でクリアしてもOK。

        def _save_worker(data):
            try:
                with open(tsv_path, "w", encoding="utf-8") as f:
                    f.write("timestamp\tframe\tx\ty\tclick\tkeys\n")
                    for row in data:
                        f.write(f"{row[0]}\t{row[1]}\t{row[2]}\t{row[3]}\t{row[4]}\t{row[5]}\n")
            except Exception as e:
                print(f"TSV保存エラー: {e}")

        threading.Thread(target=_save_worker, args=(data_to_save,)).start()

    def _record_loop(
        self,
        filepath: str,
        rect: Dict[str, int],
        fps: int,
        hwnd: Optional[Any],
        record_cursor: bool,
        record_tsv: bool,
        exclusive_window: bool
    ):
        """録画ループ本体."""
        # コーデック設定
        fourcc = cv2.VideoWriter_fourcc(*'mp4v')

        # 幅・高さは偶数である必要がある
        w = rect['width']
        h = rect['height']
        if w % 2 != 0: w -= 1
        if h % 2 != 0: h -= 1
        
        out = cv2.VideoWriter(filepath, fourcc, float(fps), (w, h))
        
        interval = 1.0 / fps
        next_time = time.time() + interval
        start_time = time.time()
        frame_idx = 0
        
        with mss.mss() as sct:
            while not self.stop_event.is_set():
                now = time.time()
                if now >= next_time:
                    # 追従時にDWMの正確な位置を取得
                    if hwnd:
                        try:
                            updated_rect = self.window_utils.get_window_rect(hwnd)
                            if updated_rect:
                                rect = updated_rect
                        except:
                            pass

                    try:
                        frame = None
                        capture_success = False

                        # ウィンドウ個別キャプチャ
                        if exclusive_window and hwnd:
                            frame = self.window_utils.capture_exclusive_window(hwnd)
                            if frame is not None:
                                capture_success = True
                        
                        # 通常のmssキャプチャ
                        if not capture_success:
                            img_sct = sct.grab(rect)
                            frame = np.array(img_sct)
                            frame = cv2.cvtColor(frame, cv2.COLOR_BGRA2BGR)
                            capture_success = True
                        
                        if capture_success and frame is not None:
                            # サイズ調整
                            if frame.shape[1] != w or frame.shape[0] != h:
                                frame = cv2.resize(frame, (w, h))

                            # マウス座標取得 (screen relative)
                            # ポインター位置取得のためにWinAPIを使う (Tkinter依存を避ける)
                            pt = ctypes.wintypes.POINT()
                            ctypes.windll.user32.GetCursorPos(ctypes.byref(pt))
                            cursor_x, cursor_y = pt.x, pt.y
                            
                            # 動画内相対座標の計算
                            rel_x = cursor_x - rect['left']
                            rel_y = cursor_y - rect['top']
                            
                            if record_tsv:
                                click_info, keys_info = self._get_input_state()
                                
                                self.trajectory_data.append((
                                    round(now - start_time, 3), 
                                    frame_idx, 
                                    rel_x, 
                                    rel_y, 
                                    click_info,
                                    ','.join(keys_info) if keys_info else "None"
                                ))

                            # 動画へのマウスポインタ直接描画
                            if record_cursor:
                                if 0 <= rel_x < w and 0 <= rel_y < h:
                                    cv2.circle(frame, (rel_x, rel_y), 5, (255, 255, 255), -1)
                                    cv2.circle(frame, (rel_x, rel_y), 5, (0, 0, 0), 1)
                            
                            out.write(frame)
                            frame_idx += 1
                    except Exception as e:
                        print(f"Record Error: {e}")
                    
                    next_time += interval
                    # スリープでCPU負荷調整
                    sleep_time = next_time - time.time()
                    if sleep_time > 0:
                        time.sleep(sleep_time)
                else:
                    time.sleep(0.001)
        
        out.release()

    def _get_input_state(self) -> tuple[str, List[str]]:
        """現在のマウス・キーボード入力状態を取得する."""
        # クリック状態の取得
        click_info = ""
        if ctypes.windll.user32.GetAsyncKeyState(0x01) & 0x8000: click_info += "L"
        if ctypes.windll.user32.GetAsyncKeyState(0x02) & 0x8000: click_info += "R"
        if ctypes.windll.user32.GetAsyncKeyState(0x04) & 0x8000: click_info += "M"
        
        if not click_info:
            click_info = "None"

        # キー状態の取得
        keys_info = []
        # 修飾キー
        if ctypes.windll.user32.GetAsyncKeyState(0x10) & 0x8000: keys_info.append("Shift")
        if ctypes.windll.user32.GetAsyncKeyState(0x11) & 0x8000: keys_info.append("Ctrl")
        if ctypes.windll.user32.GetAsyncKeyState(0x12) & 0x8000: keys_info.append("Alt")
        if (ctypes.windll.user32.GetAsyncKeyState(0x5B) & 0x8000) or (ctypes.windll.user32.GetAsyncKeyState(0x5C) & 0x8000):
            keys_info.append("Win")
        
        # 一般キー
        for vk, name in [(0x0D, "Enter"), (0x20, "Space"), (0x1B, "Esc"), (0x08, "BS"), (0x09, "Tab"), (0x2E, "Del")]:
            if ctypes.windll.user32.GetAsyncKeyState(vk) & 0x8000:
                keys_info.append(name)
        
        # 方向キー
        for vk, name in [(0x25, "Left"), (0x26, "Up"), (0x27, "Right"), (0x28, "Down")]:
            if ctypes.windll.user32.GetAsyncKeyState(vk) & 0x8000:
                keys_info.append(name)

        # A-Z
        for vk in range(0x41, 0x5B):
            if ctypes.windll.user32.GetAsyncKeyState(vk) & 0x8000:
                keys_info.append(chr(vk))
        
        # 0-9
        for vk in range(0x30, 0x3A):
            if ctypes.windll.user32.GetAsyncKeyState(vk) & 0x8000:
                keys_info.append(chr(vk))
        
        return click_info, keys_info
