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
from wgc_capture import WGCCapture


class ScreenRecorderLogic:
    """録画処理の実行・管理を行うクラス."""

    def __init__(self, window_utils: WindowUtils):
        self.window_utils = window_utils
        self.is_recording = False
        self.recording_thread: Optional[threading.Thread] = None
        self.stop_event = threading.Event()
        self.trajectory_data: List[tuple] = []
        self.current_out_path = ""
        self.shortcut_manager = None

    def start_recording(
        self,
        filepath: str,
        rect: Dict[str, int],
        fps: int,
        hwnd: Optional[Any] = None,
        record_tsv: bool = True,
        exclusive_window: bool = False,
        shortcut_manager: Optional[Any] = None
    ):
        """録画を開始する."""
        if self.is_recording:
            return

        self.current_out_path = filepath
        self.shortcut_manager = shortcut_manager
        self.stop_event.clear()
        self.is_recording = True
        self.trajectory_data = []

        self.recording_thread = threading.Thread(
            target=self._record_loop,
            args=(filepath, rect, fps, hwnd, record_tsv, exclusive_window)
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

    def get_filtered_input_state(
        self,
        rect: Dict[str, int],
        hwnd: Optional[Any] = None,
        wgc: Optional[Any] = None
    ) -> tuple[str, List[str], int, int]:
        """
        現在の入力状態を取得し、録画設定に基づいてフィルタリングを行った結果を返す.
        
        Args:
            rect: 録画対象の矩形 (left, top, width, height)
            hwnd: 録画対象のウィンドウハンドル (デスクトップ録画時はNone)
            wgc: WGCキャプチャインスタンス (任意)
            
        Returns:
            (click_info, keys_info, rel_x, rel_y)
            - click_info: "L", "R", "L,R", "None" など
            - keys_info: ["Ctrl", "A"] など
            - rel_x, rel_y: 録画領域内の相対座標
        """
        # マウス座標取得 (screen relative)
        pt = ctypes.wintypes.POINT()
        ctypes.windll.user32.GetCursorPos(ctypes.byref(pt))
        cursor_x, cursor_y = pt.x, pt.y
        
        # 動画内相対座標の計算
        rel_x = cursor_x - rect['left']
        rel_y = cursor_y - rect['top']
        w = rect['width']
        h = rect['height']

        # ベースとなる入力取得
        fg_hwnd = ctypes.windll.user32.GetForegroundWindow()
        
        # WGCかつウィンドウ録画時のみ、アクティブウィンドウチェックを行う
        # アクティブでなければ入力は一切記録しない ("None", [])
        if wgc and hwnd:
            if fg_hwnd != hwnd:
                return "None", [], rel_x, rel_y

        click_info, keys_info = self._get_input_state()
        is_desktop_recording = (hwnd is None)

        # --- デスクトップ録画時のフィルタリング ---
        if is_desktop_recording:
            # A. キー入力フィルタリング: アクティブウィンドウが録画領域内にあるか
            if keys_info:
                fg_rect = self.window_utils.get_window_rect(fg_hwnd)
                if fg_rect:
                    # 中心計算
                    cx = fg_rect['left'] + fg_rect['width'] / 2
                    cy = fg_rect['top'] + fg_rect['height'] / 2
                    
                    # 録画領域
                    r_left = rect['left']
                    r_top = rect['top']
                    r_right = r_left + rect['width']
                    r_bottom = r_top + rect['height']
                    
                    # 中心が領域外ならキー入力を無効化
                    if not (r_left <= cx < r_right and r_top <= cy < r_bottom):
                        keys_info = []
                else:
                    # 情報取得失敗時は念のため除外
                    keys_info = []

            # B. マウス入力フィルタリング: 座標が録画領域外なら除外
            # (Release "None" は除外しない)
            if click_info != "None":
                if not (0 <= rel_x < w and 0 <= rel_y < h):
                    click_info = "None"

        # --- 共通フィルタリング (Shortcut Manager) ---
        if self.shortcut_manager:
            # 組み合わせ文字列を生成
            modifiers = []
            others = []
            for k in keys_info:
                if k in ["Ctrl", "Shift", "Alt", "Win"]:
                    modifiers.append(k)
                else:
                    others.append(k)
            
            click_part = ""
            if "L" in click_info: click_part = "L-Click"
            elif "R" in click_info: click_part = "R-Click"
            elif "M" in click_info: click_part = "M-Click"
            
            allow = False
            
            # 何も入力がない場合は許可 (None)
            if click_info == "None" and not keys_info:
                allow = True
            else:
                # キーがある場合
                if keys_info:
                    key_str = "+".join(modifiers + others)
                    if self.shortcut_manager.is_allowed(key_str):
                        allow = True
                
                # クリックがある場合、修飾キー+クリックの判定
                if not allow and click_part:
                    if modifiers:
                        mod_click_str = "+".join(modifiers + [click_part])
                        if self.shortcut_manager.is_allowed(mod_click_str):
                            allow = True
                    else:
                        # 修飾キーなしクリック -> 常に許可方針
                        pass

        if not keys_info:
                # キー入力なしのクリックのみ -> 常に許可（動作履歴として重要）
                allow = True
        
        if not allow:
            # 記録しない
            click_info = "None"
            keys_info = []

        return click_info, keys_info, rel_x, rel_y

    def _record_loop(
        self,
        filepath: str,
        rect: Dict[str, int],
        fps: int,
        hwnd: Optional[Any],
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
        
        wgc: Optional[WGCCapture] = None
        if exclusive_window and hwnd:
            try:
                wgc = WGCCapture(hwnd)
                # WGC が正常に初期化されたか確認 (session が None なら失敗)
                if wgc.session is None:
                    # システム/ライブラリの状態により現代的なキャプチャ(WGC)が利用できない場合のフォールバック
                    print("Info: WGC session not initialized, falling back to standard window capture (PrintWindow)")
                    wgc.close()
                    wgc = None
            except Exception as e:
                print(f"WGC Startup Error: {e}")
                wgc = None  # PrintWindow フォールバックを使用

        try:
            with mss.mss() as sct:
                while not self.stop_event.is_set():
                    now = time.time()
                    if now >= next_time:
                        # 追従時にDWMの正確な位置を取得
                        if hwnd and not wgc: # WGC使用時はHWNDから直接取るので追従処理不要
                            try:
                                updated_rect = self.window_utils.get_window_rect(hwnd)
                                if updated_rect:
                                    rect = updated_rect
                            except:
                                pass

                        try:
                            frame = None
                            capture_success = False

                            # Windows Graphics Capture (WGC)
                            if wgc:
                                frame = wgc.get_latest_frame()
                                if frame is not None:
                                    capture_success = True

                            # WGC失敗時または WGC未使用時の独占モード -> PrintWindow
                            if not capture_success and exclusive_window and hwnd:
                                try:
                                    frame_bgr = self.window_utils.capture_exclusive_window(hwnd)
                                    if frame_bgr is not None:
                                        frame = frame_bgr
                                        capture_success = True
                                except Exception:
                                    pass
                            
                            # 通常のmssキャプチャ (独占モードでない場合のみ)
                            if not capture_success and not exclusive_window:
                                img_sct = sct.grab(rect)
                                frame = np.array(img_sct)
                                frame = cv2.cvtColor(frame, cv2.COLOR_BGRA2BGR)
                                capture_success = True
                        
                            if capture_success and frame is not None:
                                # サイズ調整
                                if frame.shape[1] != w or frame.shape[0] != h:
                                    frame = cv2.resize(frame, (w, h))

                                # 入力情報の取得 (共通化されたメソッドを使用)
                                if record_tsv:
                                    click_info, keys_info, rel_x, rel_y = self.get_filtered_input_state(rect, hwnd, wgc)
                                    
                                    self.trajectory_data.append((
                                        round(now - start_time, 3), 
                                        frame_idx, 
                                        rel_x, 
                                        rel_y, 
                                        click_info,
                                        ','.join(keys_info) if keys_info else "None"
                                    ))

                                # フレーム書き込み（TSV設定に関わらず実行）
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
        
        finally:
            if wgc:
                wgc.close()
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
        # 修飾キー (Ctrl, Shift, Alt, Win の順 - shortcut_manager.py と合わせる)
        if ctypes.windll.user32.GetAsyncKeyState(0x11) & 0x8000: keys_info.append("Ctrl")
        if ctypes.windll.user32.GetAsyncKeyState(0x10) & 0x8000: keys_info.append("Shift")
        if ctypes.windll.user32.GetAsyncKeyState(0x12) & 0x8000: keys_info.append("Alt")
        if (ctypes.windll.user32.GetAsyncKeyState(0x5B) & 0x8000) or (ctypes.windll.user32.GetAsyncKeyState(0x5C) & 0x8000):
            keys_info.append("Win")
        # 稀に 0x5B/0x5C で取れない環境があるための予備判定 (VK_LWIN/VK_RWIN は標準的なので基本は通るはず)
        
        # 一般キー
        for vk, name in [(0x0D, "Enter"), (0x20, "Space"), (0x1B, "Esc"), (0x08, "BS"), (0x09, "Tab"), (0x2E, "Del")]:
            if ctypes.windll.user32.GetAsyncKeyState(vk) & 0x8000:
                keys_info.append(name)
        
        # ナビゲーションキー
        for vk, name in [
            (0x21, "PageUp"), (0x22, "PageDown"), (0x23, "End"), (0x24, "Home"), (0x2D, "Insert")
        ]:
            if ctypes.windll.user32.GetAsyncKeyState(vk) & 0x8000:
                keys_info.append(name)

        # 方向キー
        for vk, name in [(0x25, "Left"), (0x26, "Up"), (0x27, "Right"), (0x28, "Down")]:
            if ctypes.windll.user32.GetAsyncKeyState(vk) & 0x8000:
                keys_info.append(name)

        # ファンクションキー (F1-F12)
        for vk in range(0x70, 0x7C):
            name = f"F{vk - 0x6F}"
            if ctypes.windll.user32.GetAsyncKeyState(vk) & 0x8000:
                keys_info.append(name)

        # その他特殊キー
        for vk, name in [(0x2C, "PrintScreen"), (0x13, "Pause"), (0x14, "CapsLock"), (0x91, "ScrollLock")]:
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
