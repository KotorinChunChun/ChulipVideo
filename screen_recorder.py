"""デスクトップ録画機能モジュール.

メインアプリケーションから呼び出される録画ツール。
"""
from __future__ import annotations

import ctypes
import datetime
import os
import threading
import time
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import filedialog, messagebox
from typing import TYPE_CHECKING, Optional

import cv2
import mss
import numpy as np
from PIL import Image, ImageTk

from config import get_base_dir

if TYPE_CHECKING:
    from video_frame_cropper import VideoCropperApp

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

class ScreenRecorderApp:
    """デスクトップ録画ツール.
    
    このクラスはスタンドアロンで起動することも、親アプリ
    (VideoCropperApp) から呼び出されることもできる。
    """
    
    # --- 定数 ---
    # 色定数
    COLOR_CANVAS_BG = "#444444"
    COLOR_BTN_RECORD_START = "#ffcccc"
    COLOR_BTN_RECORD_STOP = "#bbdefb"
    COLOR_BTN_RENAME = "#e3f2fd"
    COLOR_BTN_DELETE = "#ffebee"
    COLOR_BTN_UPDATE = "#e1f5fe"
    COLOR_BTN_EDIT = "#ccffcc"
    
    # リージョン設定
    REGION_THICKNESS = 5
    REGION_COLOR = "red"
    
    # Windows API 定数
    DWMWA_EXTENDED_FRAME_BOUNDS = 9
    PROCESS_QUERY_LIMITED_INFORMATION = 0x1000

    def __init__(self, root: Optional[tk.Tk] = None, parent_app: Optional[VideoCropperApp] = None):
        if root is None:
            self.root = tk.Tk()
            self.standalone = True
        else:
            self.root = tk.Toplevel(root)
            self.standalone = False
            
        self.root.title("録画ツール")
        self.root.geometry("900x600")
        self.parent_app = parent_app

        # 設定読み込み（親アプリから、あるいは個別設定）
        self.save_dir = ""
        if self.parent_app:
            # 親アプリの設定から取得するか、デフォルト値を使用
            self.save_dir = self._load_save_dir_from_config()
        
        if not self.save_dir:
            self.save_dir = os.path.join(os.path.expanduser("~"), "Videos", "AntigravityRecorder")
        
        if not os.path.exists(self.save_dir):
            os.makedirs(self.save_dir, exist_ok=True)

        # 状態変数
        self.is_recording = False
        self.recording_thread = None
        self.stop_event = threading.Event()
        self.preview_active = True
        
        # 再生状態変数
        self.cap = None
        self.is_playing = False
        self.playback_after_id = None
        self.video_total_frames = 0
        self.user_dragging_slider = False
        
        # UI変数
        self.source_var = tk.StringVar(value="desktop") # desktop / window
        self.target_var = tk.StringVar()
        self.fps_var = tk.IntVar(value=15)
        self.quality_var = tk.StringVar(value="中")
        self.save_path_var = tk.StringVar(value=self.save_dir)
        self.record_cursor_var = tk.BooleanVar(value=True) # マウスポインタ録画
        self.exclusive_window_var = tk.BooleanVar(value=False) # ウィンドウ単体キャプチャ
        
        # UIパーツ参照用
        self.widgets_to_lock = []
        self.region_window = None
        self.monitors = []
        self.windows = []
        self.sct = mss.mss()
        
        self._build_ui()
        self.update_source_list()
        self._start_preview()

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def _load_save_dir_from_config(self) -> str:
        # 簡易的に config.py や 親アプリのメソッドには頼らず、
        # 親の config 属性があればそこから読む、無ければ自前で管理
        # ここでは後で親アプリの config と連携するとして、一旦保留
        return ""

    def _build_ui(self):
        # 全面スプリットフレーム (PanedWindow)
        self.main_paned = tk.PanedWindow(self.root, orient=tk.HORIZONTAL, sashwidth=4, bg="#cccccc")
        self.main_paned.pack(fill=tk.BOTH, expand=True)

        left_container = tk.Frame(self.main_paned)
        right_container = tk.Frame(self.main_paned)
        self.main_paned.add(left_container, minsize=450) # +100 (350 -> 450)
        self.main_paned.add(right_container, minsize=450)

        # --- 左側: ファイル操作管理 (入れ替え) ---
        
        # 5. 保存先
        path_frame = tk.LabelFrame(left_container, text="保存設定")
        path_frame.pack(fill=tk.X, padx=10, pady=5)
        tk.Label(path_frame, text="保存先:").pack(side=tk.LEFT, padx=5, pady=10)
        tk.Entry(path_frame, textvariable=self.save_path_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        tk.Button(path_frame, text="...", command=self.browse_save_dir, width=3).pack(side=tk.LEFT, padx=5)

        # リスト
        list_label_frame = tk.LabelFrame(left_container, text="録画済みファイル")
        list_label_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        scrollbar = tk.Scrollbar(list_label_frame, orient=tk.VERTICAL)
        # ListView (Listbox) での実装 - selectmodeをEXTENDEDに変更
        self.file_listbox = tk.Listbox(list_label_frame, yscrollcommand=scrollbar.set, font=("Consolas", 10), selectmode=tk.EXTENDED)
        
        scrollbar.config(command=self.file_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=2, pady=2)
        
        self.file_listbox.bind("<Double-Button-1>", self.on_file_double_click)
        self.file_listbox.bind("<<ListboxSelect>>", self.on_file_select)
        
        # データ保持用
        self.file_items = [] # 一覧に表示されている実際のファイル名
        
        # リスト操作ボタン
        list_ctrl = tk.Frame(left_container)
        list_ctrl.pack(fill=tk.X, padx=10, pady=5)
        self.btn_rename = tk.Button(list_ctrl, text="名前変更", command=self.rename_file, state=tk.DISABLED, height=1, bg=self.COLOR_BTN_RENAME)
        self.btn_rename.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)
        self.btn_delete = tk.Button(list_ctrl, text="削除", command=self.delete_file, state=tk.DISABLED, height=1, bg=self.COLOR_BTN_DELETE)
        self.btn_delete.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)
        self.widgets_to_lock.extend([self.btn_rename, self.btn_delete])

        # キーバインド
        self.file_listbox.bind("<F2>", lambda e: self.rename_file())
        self.file_listbox.bind("<Delete>", lambda e: self.delete_file())
        self.file_listbox.bind("<Return>", lambda e: self.on_file_double_click(e)) # Enterでも再生
        self.file_listbox.bind("<Control-a>", lambda e: self._select_all_files())
        self.file_listbox.bind("<Control-A>", lambda e: self._select_all_files())

        # フッターアクション (左側下部)
        footer_frame = tk.Frame(left_container)
        footer_frame.pack(fill=tk.X, padx=10, pady=10)
        
        self.btn_close = tk.Button(footer_frame, text="閉じる", command=self.on_close, width=10)
        self.btn_close.pack(side=tk.RIGHT, padx=5)
        self.btn_edit = tk.Button(footer_frame, text="閉じて編集", command=self.close_and_edit, bg=self.COLOR_BTN_EDIT, width=12)
        self.btn_edit.pack(side=tk.RIGHT, padx=5)
        self.widgets_to_lock.extend([self.btn_close, self.btn_edit])


        # --- 右側: タブインターフェース (入れ替え) ---
        self.notebook = ttk.Notebook(right_container)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        self.tab_record = tk.Frame(self.notebook)
        self.tab_play = tk.Frame(self.notebook)
        self.notebook.add(self.tab_record, text="  録画  ")
        self.notebook.add(self.tab_play, text="  再生  ")

        # --- 録画タブの内容 ---
        # 1. ソース選択
        source_frame = tk.LabelFrame(self.tab_record, text="録画対象")
        source_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.radio_desktop = tk.Radiobutton(source_frame, text="デスクトップ", variable=self.source_var, 
                       value="desktop", command=self.update_source_list)
        self.radio_desktop.pack(side=tk.LEFT, padx=10)
        self.radio_window = tk.Radiobutton(source_frame, text="ウィンドウ", variable=self.source_var, 
                       value="window", command=self.update_source_list)
        self.radio_window.pack(side=tk.LEFT, padx=10)
        self.widgets_to_lock.extend([self.radio_desktop, self.radio_window])
        
        # 2. ターゲット選択
        target_frame = tk.Frame(self.tab_record)
        target_frame.pack(fill=tk.X, padx=10, pady=5)
        tk.Label(target_frame, text="対象:").pack(side=tk.LEFT)
        self.combo_target = ttk.Combobox(target_frame, textvariable=self.target_var, state="readonly")
        self.combo_target.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.combo_target.bind("<<ComboboxSelected>>", self.on_target_changed)
        self.btn_update = tk.Button(target_frame, text="更新", command=self.update_source_list, width=4, bg=self.COLOR_BTN_UPDATE)
        self.btn_update.pack(side=tk.LEFT)
        self.widgets_to_lock.append(self.btn_update)
        self.widgets_to_lock.append(self.combo_target)

        # 3. プレビュー
        preview_label_frame = tk.LabelFrame(self.tab_record, text="プレビュー")
        preview_label_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # 背景色をグレーに変更
        self.preview_canvas = tk.Canvas(preview_label_frame, bg=self.COLOR_CANVAS_BG, highlightthickness=0)
        self.preview_canvas.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.preview_image_id = None
        self.preview_canvas.bind("<Configure>", lambda e: self._on_canvas_resize("preview"))

        # 4. 録画設定 (FPS, Quality)
        settings_frame = tk.Frame(self.tab_record)
        settings_frame.pack(fill=tk.X, padx=10, pady=5)
        
        tk.Label(settings_frame, text="FPS:").pack(side=tk.LEFT)
        self.combo_fps = ttk.Combobox(settings_frame, textvariable=self.fps_var, values=[15, 30, 60], width=5, state="readonly")
        self.combo_fps.pack(side=tk.LEFT, padx=5)
        
        tk.Label(settings_frame, text="画質:").pack(side=tk.LEFT, padx=(10, 0))
        self.combo_quality = ttk.Combobox(settings_frame, textvariable=self.quality_var, values=["高", "中", "低"], width=5, state="readonly")
        self.combo_quality.pack(side=tk.LEFT, padx=5)
        self.widgets_to_lock.extend([self.combo_fps, self.combo_quality])

        # 6. オプション (マウスポインタ/重なり防止)
        options_frame = tk.Frame(self.tab_record)
        options_frame.pack(fill=tk.X, padx=10, pady=2)
        tk.Checkbutton(options_frame, text="マウスポインタを録画する", variable=self.record_cursor_var).pack(side=tk.LEFT)
        self.check_exclusive = tk.Checkbutton(options_frame, text="他のウィンドウの写り込みを防ぐ (Win10+)", variable=self.exclusive_window_var)
        self.check_exclusive.pack(side=tk.LEFT, padx=(10, 0))
        self.widgets_to_lock.append(self.check_exclusive)

        # 7. 録画ボタン
        self.btn_record = tk.Button(self.tab_record, text="● 録画開始", bg="#ffcccc", font=("Arial", 12, "bold"),
                                    command=self.toggle_recording)
        self.btn_record.pack(fill=tk.X, padx=15, pady=10)


        # --- 再生タブの内容 ---
        player_frame = tk.LabelFrame(self.tab_play, text="プレイヤー")
        player_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 背景色をグレーに変更
        self.player_canvas = tk.Canvas(player_frame, bg=self.COLOR_CANVAS_BG, highlightthickness=0)
        self.player_canvas.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.player_canvas.bind("<Configure>", lambda e: self._on_canvas_resize("player"))
        
        p_btns_container = tk.Frame(self.tab_play)
        p_btns_container.pack(fill=tk.X, padx=10, pady=(0, 10))

        p_btns = tk.Frame(p_btns_container)
        p_btns.pack(fill=tk.X, pady=2)
        
        self.btn_play = tk.Button(p_btns, text="▶", command=self.toggle_playback, state=tk.DISABLED, width=5)
        self.btn_play.pack(side=tk.LEFT, padx=5)
        
        # シークバー
        self.seek_var = tk.DoubleVar()
        self.slider = ttk.Scale(p_btns, from_=0, to=100, variable=self.seek_var, orient=tk.HORIZONTAL, command=self.on_slider_move)
        self.slider.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.slider.bind("<Button-1>", self.on_slider_press)
        self.slider.bind("<ButtonRelease-1>", self.on_slider_release)

        self.lbl_time = tk.Label(p_btns, text="00:00 / 00:00")
        self.lbl_time.pack(side=tk.LEFT, padx=5)

        self.refresh_file_list()

    def update_source_list(self):
        """録画対象のリストを更新"""
        self.combo_target['values'] = []
        mode = self.source_var.get()
        
        if mode == 'desktop':
            # モニター一覧取得
            self.monitors = self.sct.monitors[1:] # 0は全画面結合なので除外する場合が多いが、要件次第。一旦個別モニタのみ。
            if not self.monitors:
                self.monitors = [self.sct.monitors[0]] # フォールバック
            
            display_names = [f"Display {i+1}: {m['width']}x{m['height']}" for i, m in enumerate(self.monitors)]
            self.combo_target['values'] = display_names
            if display_names:
                self.combo_target.current(0)
                
        elif mode == 'window':
            # ウィンドウ一覧取得 (可視ウィンドウのみ)
            self.windows = []
            
            def enum_windows_proc(hwnd, lParam):
                if ctypes.windll.user32.IsWindowVisible(hwnd):
                    length = ctypes.windll.user32.GetWindowTextLengthW(hwnd)
                    if length > 0:
                        buff = ctypes.create_unicode_buffer(length + 1)
                        ctypes.windll.user32.GetWindowTextW(hwnd, buff, length + 1)
                        title = buff.value
                        if title:
                            self.windows.append((hwnd, title))
                return True
                
            WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_int, ctypes.c_int)
            ctypes.windll.user32.EnumWindows(WNDENUMPROC(enum_windows_proc), 0)
            
            display_names = []
            for h, t in self.windows:
                pname = self._get_process_name(h)
                display_names.append(f"[{pname}] {t} ({h})")
                
            self.combo_target['values'] = display_names
            if display_names:
                self.combo_target.current(0)
        
        self.on_target_changed(None)

    def on_target_changed(self, event):
        """ターゲット変更時の処理"""
        pass

    def _get_process_name(self, hwnd):
        """hwndからプロセス名を取得"""
        pid = ctypes.c_ulong()
        ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        
        # PROCESS_QUERY_LIMITED_INFORMATION (0x1000)
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

    def _get_start_monitor_idx(self):
        """mssのモニターインデックスが1始まりのため調整"""
        # self.monitors は sct.monitors[1:] なので、index + 1
        idx = self.combo_target.current()
        if idx < 0: return 1
        return idx + 1

    def _select_all_files(self):
        self.file_listbox.selection_set(0, tk.END)
        return "break"

    def _get_target_rect(self):
        """録画対象の矩形を取得 (DWMを使用して余白を除去)"""
        mode = self.source_var.get()
        if mode == 'desktop':
            idx = self.combo_target.current()
            if idx >= 0 and idx < len(self.monitors):
                return self.monitors[idx]
            return self.sct.monitors[1]
            
        elif mode == 'window':
            idx = self.combo_target.current()
            if idx >= 0 and idx < len(self.windows):
                hwnd = self.windows[idx][0]
                # DWMWA_EXTENDED_FRAME_BOUNDS = 9
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
                    ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(rect))
                    return {'top': rect.top, 'left': rect.left, 'width': rect.right - rect.left, 'height': rect.bottom - rect.top}
        
        return None

    def _capture_exclusive_window(self, hwnd):
        """PrintWindow を使用して重なりを無視してウィンドウをキャプチャ"""
        try:
            # ここも正確なサイズを使用
            rect = ctypes.wintypes.RECT()
            ctypes.windll.dwmapi.DwmGetWindowAttribute(hwnd, self.DWMWA_EXTENDED_FRAME_BOUNDS, ctypes.byref(rect), ctypes.sizeof(rect))
            w = rect.right - rect.left
            h = rect.bottom - rect.top
            if w <= 0 or h <= 0: return None

            # DC作成
            hdc_screen = ctypes.windll.user32.GetDC(0)
            hdc_mem = ctypes.windll.gdi32.CreateCompatibleDC(hdc_screen)
            hbmp = ctypes.windll.gdi32.CreateCompatibleBitmap(hdc_screen, w, h)
            ctypes.windll.gdi32.SelectObject(hdc_mem, hbmp)

            # PW_RENDERFULLCONTENT (2) で描画
            ctypes.windll.user32.PrintWindow(hwnd, hdc_mem, 2)

            # Bitmap から numpy 配列 (BGR) への変換
            # ここは ctypes だと非常に長いので、PIL Image を経由させる
            # または OpenCV ならもっと直接的な方法があるが、一旦 BITMAPINFO を作る
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

    def _start_preview(self):
        """プレビュー更新ループ"""
        if not self.preview_canvas.winfo_exists():
            return
            
        if self.preview_active:
            rect = self._get_target_rect()
            if rect:
                try:
                    capture_success = False
                    img = None

                    # ウィンドウ個別キャプチャ (重なり防止)
                    if self.source_var.get() == 'window' and self.exclusive_window_var.get():
                        idx = self.combo_target.current()
                        if idx >= 0 and idx < len(self.windows):
                            hwnd = self.windows[idx][0]
                            frame_bgr = self._capture_exclusive_window(hwnd)
                            if frame_bgr is not None:
                                frame_rgb = cv2.cvtColor(frame_bgr, cv2.COLOR_BGR2RGB)
                                img = Image.fromarray(frame_rgb)
                                capture_success = True
                    
                    # 通常の画面キャプチャ (mss)
                    if not capture_success:
                        img_sct = self.sct.grab(rect)
                        img = Image.frombytes("RGB", img_sct.size, img_sct.bgra, "raw", "BGRX")
                        capture_success = True
                    
                    if capture_success and img:
                        # キャンバスサイズに合わせてリサイズ
                        cw = self.preview_canvas.winfo_width()
                    ch = self.preview_canvas.winfo_height()
                    if cw > 1 and ch > 1:
                        img.thumbnail((cw, ch), Image.Resampling.NEAREST)
                        tk_img = ImageTk.PhotoImage(img)
                        
                        if self.preview_image_id:
                            self.preview_canvas.itemconfig(self.preview_image_id, image=tk_img)
                            self.preview_canvas.coords(self.preview_image_id, cw//2, ch//2)
                        else:
                            self.preview_image_id = self.preview_canvas.create_image(cw//2, ch//2, image=tk_img, anchor=tk.CENTER)
                        
                        # 参照保持
                        self.preview_canvas.image = tk_img
                    
                except Exception as e:
                    # キャプチャ失敗 (ウィンドウが最小化されている等)
                    pass

        # 録画中はプレビュー更新頻度を下げる、または止めるなどの調整も可
        interval = 100 if not self.is_recording else 500
        self.root.after(interval, self._start_preview)

    def toggle_recording(self):
        if self.is_recording:
            # 停止
            self.stop_recording()
        else:
            # 開始
            self.start_recording()

    def start_recording(self):
        rect = self._get_target_rect()
        if not rect:
            messagebox.showerror("Error", "録画対象が取得できません")
            return
            
        # 録画中のコントロールロック
        self._set_controls_state(tk.DISABLED)
        self._show_recording_region(rect)
        
        # ウィンドウ追従のためのhwnd取得
        hwnd = None
        if self.source_var.get() == 'window':
            idx = self.combo_target.current()
            if idx >= 0 and idx < len(self.windows):
                hwnd = self.windows[idx][0]

        # ファイル名
        now = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{now}.mp4"
        filepath = os.path.join(self.save_path_var.get(), filename)
        self.current_out_path = filepath
        
        fps = self.fps_var.get()
        quality_label = self.quality_var.get()
        # 画質設定: ビットレート直接指定はcv2 VideoWriterでは難しいので、
        # 必要ならリサイズ処理を入れるか、コーデックパラメータを探る
        # ここではFPS制御のみとする (画質は規定)
        
        self.stop_event.clear()
        self.is_recording = True
        self.btn_record.config(text="■ 録画停止", bg="#ff9999")
        
        self.recording_thread = threading.Thread(target=self._record_loop, args=(filepath, rect, fps, hwnd))
        self.recording_thread.start()
        
        # 赤枠追従ループ開始
        self._update_region_tracking()

    def _update_region_tracking(self):
        """録画中に赤枠を対象ウィンドウに追従させる"""
        if self.is_recording and self.region_window:
            rect = self._get_target_rect()
            if rect:
                thickness = self.REGION_THICKNESS
                x = rect['left'] - thickness
                y = rect['top'] - thickness
                # ジオメトリだけ更新
                if self.region_window.winfo_exists():
                    self.region_window.geometry(f"+{x}+{y}")
            
            # 滑らかにするために間隔を短縮 (100 -> 10)
            self.root.after(10, self._update_region_tracking)

    def stop_recording(self):
        self.stop_event.set()
        self.is_recording = False
        self.btn_record.config(text="処理中...", state=tk.DISABLED)
        
        # スレッド終了待機はUIを止めないようにチェックループへ
        self._wait_for_stop()

    def _wait_for_stop(self):
        if self.recording_thread and self.recording_thread.is_alive():
            self.root.after(100, self._wait_for_stop)
        else:
            self.btn_record.config(text="● 録画開始", bg="#ffcccc", state=tk.NORMAL)
            
            # ロック解除
            self._set_controls_state(tk.NORMAL)
            self._hide_recording_region()

            # リストを更新して新しく録画したファイルにフォーカス
            new_fname = os.path.basename(self.current_out_path)
            self.refresh_file_list(select_filename=new_fname)
            
            # 録画完了メッセージは不要とのことで削除

    def _set_controls_state(self, state):
        """録画中に各種コントロールをロック/アンロック"""
        self.file_listbox.config(state=state)
        for w in self.widgets_to_lock:
            try:
                target_state = state
                if state == tk.NORMAL and isinstance(w, ttk.Combobox):
                    target_state = "readonly"
                w.config(state=target_state)
            except tk.TclError:
                pass  # ウィジェットが既に破棄されている等

        if state == tk.DISABLED:
            self.btn_record.config(text="■ 録画停止", bg=self.COLOR_BTN_RECORD_STOP)
        else:
            self.btn_record.config(text="● 録画開始", bg=self.COLOR_BTN_RECORD_START)

    def _show_recording_region(self, rect):
        """デスクトップ上に赤枠のリージョンを表示"""
        if self.region_window: self._hide_recording_region()
        
        self.region_window = tk.Toplevel(self.root)
        self.region_window.overrideredirect(True)
        self.region_window.attributes("-topmost", True)
        self.region_window.attributes("-transparentcolor", "white")
        
        # 外側に表示
        thickness = self.REGION_THICKNESS
        x = rect['left'] - thickness
        y = rect['top'] - thickness
        # 幅・高さも調整 (mssの矩形に合わせて)
        w = rect['width'] + thickness * 2
        h = rect['height'] + thickness * 2
        
        self.region_window.geometry(f"{w}x{h}+{x}+{y}")
        canvas = tk.Canvas(self.region_window, width=w, height=h, bg="white", highlightthickness=0)
        canvas.pack()
        
        # 赤枠
        canvas.create_rectangle(thickness//2, thickness//2, w - thickness//2, h - thickness//2, outline=self.REGION_COLOR, width=thickness)

    def _hide_recording_region(self):
        """赤枠を非表示"""
        if self.region_window:
            self.region_window.destroy()
            self.region_window = None

    def _record_loop(self, filepath, rect, fps, hwnd=None):
        """録画ループ（別スレッド）"""
        fourcc = cv2.VideoWriter_fourcc(*'mp4v')
        
        # 幅・高さは偶数である必要がある
        w = rect['width']
        h = rect['height']
        if w % 2 != 0: w -= 1
        if h % 2 != 0: h -= 1
        
        out = cv2.VideoWriter(filepath, fourcc, float(fps), (w, h))
        
        interval = 1.0 / fps
        next_time = time.time() + interval
        
        # mss インスタンスはスレッドごとに持つ方が安全な場合がある
        with mss.mss() as sct:
            while not self.stop_event.is_set():
                now = time.time()
                if now >= next_time:
                    # 追従時にDWMの正確な位置を取得するように変更
                    if hwnd:
                        try:
                            w_rect = ctypes.wintypes.RECT()
                            ctypes.windll.dwmapi.DwmGetWindowAttribute(hwnd, self.DWMWA_EXTENDED_FRAME_BOUNDS, ctypes.byref(w_rect), ctypes.sizeof(w_rect))
                            rect = {
                                'top': w_rect.top,
                                'left': w_rect.left,
                                'width': w_rect.right - w_rect.left,
                                'height': w_rect.bottom - w_rect.top
                            }
                        except:
                            pass

                    try:
                        frame = None
                        capture_success = False

                        # ウィンドウ個別キャプチャ
                        if self.source_var.get() == 'window' and self.exclusive_window_var.get():
                            idx = self.combo_target.current()
                            if idx >= 0 and idx < len(self.windows):
                                hwnd = self.windows[idx][0]
                                frame = self._capture_exclusive_window(hwnd)
                                if frame is not None:
                                    capture_success = True
                        
                        # 通常のmssキャプチャ
                        if not capture_success:
                            img_sct = sct.grab(rect)
                            frame = np.array(img_sct)
                            frame = cv2.cvtColor(frame, cv2.COLOR_BGRA2BGR)
                            capture_success = True
                        
                        if capture_success and frame is not None:
                            # サイズ調整 (VideoWriterに渡すサイズと一致させる)
                            if frame.shape[1] != w or frame.shape[0] != h:
                                frame = cv2.resize(frame, (w, h))

                            # マウスポインタの描画
                            # ウィンドウキャプチャ時は座標系が異なるため、
                            # 正確に取得するにはウィンドウ内相対座標が必要
                            if self.record_cursor_var.get():
                                cursor_x, cursor_y = self.root.winfo_pointerxy()
                                if self.source_var.get() == 'window' and self.exclusive_window_var.get():
                                    # PrintWindow時はウィンドウ左上= (0,0)
                                    # GetWindowRect はタイトルバー等を含む
                                    # クライアント座標にするか等検討の予地あり
                                    rel_x = cursor_x - rect['left']
                                    rel_y = cursor_y - rect['top']
                                else:
                                    rel_x = cursor_x - rect['left']
                                    rel_y = cursor_y - rect['top']
                                
                                if 0 <= rel_x < w and 0 <= rel_y < h:
                                    cv2.circle(frame, (rel_x, rel_y), 5, (255, 255, 255), -1)
                                    cv2.circle(frame, (rel_x, rel_y), 5, (0, 0, 0), 1)
                            
                            out.write(frame)
                    except Exception as e:
                        print(f"Record Error: {e}")
                    
                    next_time += interval
                    # スリープでCPU負荷調整 (wait time is dynamic)
                    sleep_time = next_time - time.time()
                    if sleep_time > 0:
                        time.sleep(sleep_time)
                else:
                    time.sleep(0.001)
        
        out.release()

    def browse_save_dir(self):
        d = filedialog.askdirectory(initialdir=self.save_path_var.get())
        if d:
            self.save_path_var.set(d)
            self.save_dir = d
            self.refresh_file_list()

    def refresh_file_list(self, select_filename: Optional[str] = None):
        self.file_listbox.delete(0, tk.END)
        self.file_items = []

        d = self.save_path_var.get()
        if os.path.exists(d):
            files_info = []
            for f in os.listdir(d):
                if f.lower().endswith(".mp4"):
                    path = os.path.join(d, f)
                    try:
                        mtime = os.path.getmtime(path)
                        dt_str = datetime.datetime.fromtimestamp(mtime).strftime("%Y/%m/%d %H:%M:%S")
                        files_info.append((mtime, f, dt_str))
                    except:
                        pass
            
            # 更新日時順にソート (降順)
            files_info.sort(key=lambda x: x[0], reverse=True)

            for mtime, f, dt in files_info:
                # 等幅フォントで並ぶように調整
                display_str = f"{f:<30} | {dt}"
                self.file_listbox.insert(tk.END, display_str)
                self.file_items.append(f)
                
                if select_filename and f == select_filename:
                    idx = self.file_listbox.size() - 1
                    self.file_listbox.selection_set(idx)
                    self.file_listbox.see(idx)
                    self.file_listbox.activate(idx)

        self.on_file_select(None)

    def on_file_select(self, event):
        idx = self.file_listbox.curselection()
        if idx:
            self.btn_rename.config(state=tk.NORMAL)
            self.btn_delete.config(state=tk.NORMAL)
            self.btn_play.config(state=tk.NORMAL)
            # 再生を停止してリセット
            self.stop_playback()
            self.load_video_for_playback()
        else:
            self.btn_rename.config(state=tk.DISABLED)
            self.btn_delete.config(state=tk.DISABLED)
            self.btn_play.config(state=tk.DISABLED)
            self.stop_playback()
            self.clear_player_canvas()

    def on_file_double_click(self, event):
        """ダブルクリックまたはEnterで再生タブへ切り替えて再生"""
        idx = self.file_listbox.curselection()
        if idx:
            # 1. 再生タブを選択
            self.notebook.select(self.tab_play)
            # 2. 動画をロード（既に選択時にロードされているが念のため）
            self.load_video_for_playback()
            # 3. 再生開始
            if not self.is_playing:
                self.toggle_playback()

    def rename_file(self):
        indices = self.file_listbox.curselection()
        if not indices: return
        
        is_multiple = len(indices) > 1
        title = "一括名前変更 (接頭辞付与)" if is_multiple else "名前変更"
        label_text = "接頭辞を入力してください:" if is_multiple else "新しいファイル名:"
        
        # モーダルダイアログを自作
        dialog = tk.Toplevel(self.root)
        dialog.title(title)
        dialog.geometry("350x150")
        dialog.transient(self.root)
        dialog.grab_set()

        # 中央に配置
        self.root.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - 175
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - 75
        dialog.geometry(f"+{max(0, x)}+{max(0, y)}")

        tk.Label(dialog, text=label_text).pack(pady=5)
        
        initial_val = "" if is_multiple else os.path.splitext(self.file_items[indices[0]])[0]
        name_var = tk.StringVar(value=initial_val)
        entry = tk.Entry(dialog, textvariable=name_var, width=40)
        entry.pack(pady=5, padx=10)
        entry.focus_set()
        if not is_multiple:
            entry.select_range(0, tk.END)

        def on_ok(event=None):
            new_val = name_var.get().strip()
            if not new_val:
                dialog.destroy()
                return

            # ファイルロック解除
            self.stop_playback()
            if self.cap:
                self.cap.release()
                self.cap = None

            success_count = 0
            for i in indices:
                fname = self.file_items[i]
                base, ext = os.path.splitext(fname)
                old_path = os.path.join(self.save_dir, fname)
                
                if is_multiple:
                    # 先頭に付与
                    new_name = new_val + "_" + base + ext
                else:
                    # 名前を置換
                    new_name = new_val + ext
                
                new_path = os.path.join(self.save_dir, new_name)
                
                try:
                    os.rename(old_path, new_path)
                    success_count += 1
                except Exception as e:
                    print(f"Rename error for {fname}: {e}")

            if success_count > 0:
                self.refresh_file_list()
            
            if success_count < len(indices):
                messagebox.showwarning("Warning", "いくつかのファイルの名前変更に失敗しました。")
            
            dialog.destroy()
            self.file_listbox.focus_set()

        def on_cancel(event=None):
            dialog.destroy()
            self.file_listbox.focus_set()

        btn_frame = tk.Frame(dialog)
        btn_frame.pack(pady=10)
        tk.Button(btn_frame, text="OK", command=on_ok, width=10).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="キャンセル", command=on_cancel, width=10).pack(side=tk.LEFT, padx=5)

        entry.bind("<Return>", on_ok)
        entry.bind("<Escape>", on_cancel)

        self.root.wait_window(dialog)

    def delete_file(self):
        indices = self.file_listbox.curselection()
        if not indices: return
        
        count = len(indices)
        msg = f"{count} 個のアイテムを削除しますか？" if count > 1 else f"{self.file_items[indices[0]]} を削除しますか？"
        
        if messagebox.askyesno("確認", msg):
            self.stop_playback() # 再生中なら止める
            
            # 再生中でなくてもファイルハンドルを保持している可能性があるので明示的に解除
            if self.cap:
                self.cap.release()
                self.cap = None

            success_count = 0
            for i in sorted(indices, reverse=True):
                fname = self.file_items[i]
                path = os.path.join(self.save_dir, fname)
                try:
                    os.remove(path)
                    success_count += 1
                except Exception as e:
                    print(f"Delete error for {fname}: {e}")

            self.refresh_file_list()
            if success_count < count:
                messagebox.showerror("Error", "いくつかのファイルの削除に失敗しました。")

    def clear_player_canvas(self):
        self.player_canvas.delete("all")
        self.lbl_time.config(text="00:00 / 00:00")
        self.seek_var.set(0)

    def load_video_for_playback(self):
        idx = self.file_listbox.curselection()
        if not idx: return
        fname = self.file_items[idx[0]]
        path = os.path.join(self.save_dir, fname)
        
        if self.cap:
            self.cap.release()
            
        self.cap = cv2.VideoCapture(path)
        if not self.cap.isOpened():
            return
            
        self.video_total_frames = int(self.cap.get(cv2.CAP_PROP_FRAME_COUNT))
        self.video_fps = self.cap.get(cv2.CAP_PROP_FPS)
        self.slider.config(to=self.video_total_frames - 1)
        
        # 最初のフレームを表示
        self.show_frame(0)
        self.update_time_label(0)

    def toggle_playback(self):
        if self.is_playing:
            self.stop_playback()
        else:
            self.start_playback()

    def start_playback(self):
        if not self.cap or not self.cap.isOpened():
            return
        
        # 最後まで再生していたら最初に戻す
        curr = int(self.cap.get(cv2.CAP_PROP_POS_FRAMES))
        if curr >= self.video_total_frames - 1:
            self.cap.set(cv2.CAP_PROP_POS_FRAMES, 0)
            
        self.is_playing = True
        self.btn_play.config(text="Ⅱ")
        self.playback_loop()

    def stop_playback(self):
        self.is_playing = False
        self.btn_play.config(text="▶")
        if self.playback_after_id:
            self.root.after_cancel(self.playback_after_id)
            self.playback_after_id = None

    def playback_loop(self):
        if not self.is_playing or not self.cap:
            return
            
        ret, frame = self.cap.read()
        if ret:
            curr_frame = int(self.cap.get(cv2.CAP_PROP_POS_FRAMES))
            self.display_frame(frame)
            if not self.user_dragging_slider:
                self.seek_var.set(curr_frame)
            self.update_time_label(curr_frame)
            
            # FPSに合わせて待機 (15-30fps程度ならafterで実用範疇)
            delay = int(1000 / self.video_fps) if self.video_fps > 0 else 33
            self.playback_after_id = self.root.after(delay, self.playback_loop)
        else:
            self.stop_playback()

    def show_frame(self, frame_idx):
        if self.cap:
            self.cap.set(cv2.CAP_PROP_POS_FRAMES, frame_idx)
            ret, frame = self.cap.read()
            if ret:
                self.display_frame(frame)

    def display_frame(self, frame):
        # BGR -> RGB
        if frame is not None:
            self.last_player_frame = frame.copy()
            frame_to_disp = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            img = Image.fromarray(frame_to_disp)
        elif hasattr(self, 'last_player_frame'):
            frame_to_disp = cv2.cvtColor(self.last_player_frame, cv2.COLOR_BGR2RGB)
            img = Image.fromarray(frame_to_disp)
        else:
            return
        
        cw = self.player_canvas.winfo_width()
        ch = self.player_canvas.winfo_height()
        if cw > 1 and ch > 1:
            img.thumbnail((cw, ch), Image.Resampling.LANCZOS)
            tk_img = ImageTk.PhotoImage(img)
            self.player_canvas.delete("img")
            self.player_canvas.create_image(cw//2, ch//2, image=tk_img, anchor=tk.CENTER, tags="img")
            self.player_canvas.image = tk_img

    def _on_canvas_resize(self, mode):
        if mode == "player" and hasattr(self, 'last_player_frame'):
            # 再生中でない時のみ手動で再描画 (再生中はループが回るため)
            if not self.is_playing:
                self.display_frame(None)
        # preview はループが常時回っているので座標更新に任せる

    def update_time_label(self, curr_frame):
        if self.video_fps > 0:
            curr_sec = int(curr_frame / self.video_fps)
            total_sec = int(self.video_total_frames / self.video_fps)
            self.lbl_time.config(text=f"{self.format_time(curr_sec)} / {self.format_time(total_sec)}")

    def format_time(self, seconds):
        m, s = divmod(seconds, 60)
        return f"{m:02d}:{s:02d}"

    def on_slider_press(self, event):
        self.user_dragging_slider = True
        self.was_playing_before_drag = self.is_playing
        if self.is_playing:
            self.stop_playback()

    def on_slider_move(self, value):
        if self.user_dragging_slider:
            val = int(float(value))
            self.show_frame(val)
            self.update_time_label(val)

    def on_slider_release(self, event):
        self.user_dragging_slider = False
        val = int(self.seek_var.get())
        self.show_frame(val)
        if hasattr(self, 'was_playing_before_drag') and self.was_playing_before_drag:
            self.start_playback()

    def play_file(self):
        # このメソッドは当初外部プレイヤー起動用だったが、
        # アプリ内再生がメインになったため、もう直接は呼ばれないかもしれない。
        # 一応 toggle_playback に飛ばすか、残しておく。
        self.toggle_playback()

    def close_and_edit(self):
        idx = self.file_listbox.curselection()
        if idx:
            fname = self.file_items[idx[0]]
            path = os.path.join(self.save_dir, fname)
            
            # 親アプリにロードさせる
            if self.parent_app:
                # 親アプリがロードメソッドを持っている前提
                # load_video は引数を取らない作りになっているため、
                # 直接内部変数をセットしてロード処理を呼ぶか、引数対応が必要
                # ここでは簡易的に親のメソッドをハックする
                
                # 親の load_video を改修するのが筋だが、ここでは暫定的に
                # cap を直接作り直して UI 更新を呼ぶ
                
                # しかし video_frame_cropper.py のロジックを見ると
                # load_video() は filedialog を呼んでしまう。
                # 外部からパスを渡せる public method が必要。
                
                # 後続のタスクで video_frame_cropper.py に open_video(path) を追加する。
                if hasattr(self.parent_app, 'open_video_file'):
                     self.parent_app.open_video_file(path)
        
        self.on_close()

    def on_close(self):
        self.preview_active = False
        if self.is_recording:
            if messagebox.askyesno("確認", "録画中です。停止して閉じますか？"):
                self.stop_event.set()
                # 終了待ち
                if self.recording_thread:
                    self.recording_thread.join(timeout=2.0)
            else:
                return

        if self.cap:
            self.cap.release()
            
        self.root.destroy()
        if getattr(self, 'standalone', False):
            self.root.quit()

if __name__ == "__main__":
    app = ScreenRecorderApp()
    app.root.mainloop()
