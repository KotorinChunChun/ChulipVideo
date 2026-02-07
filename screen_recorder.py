"""デスクトップ録画機能モジュール.

メインアプリケーションから呼び出される録画ツール。
"""
from __future__ import annotations

import datetime
import os
import threading
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import filedialog, messagebox
from typing import TYPE_CHECKING, Optional, List, Tuple, Dict, Any

import cv2
import mss
import numpy as np
from PIL import Image, ImageTk

from config import get_base_dir
from ui_utils import add_tooltip
from window_utils import WindowUtils
from recorder_core import ScreenRecorderLogic

if TYPE_CHECKING:
    from video_frame_cropper import VideoCropperApp


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
    
    def __init__(self, root: Optional[tk.Tk] = None, parent_app: Optional[VideoCropperApp] = None):
        self.window_utils = WindowUtils()
        self.recorder_logic = ScreenRecorderLogic(self.window_utils)

        if root is None:
            self.root = tk.Tk()
            self.standalone = True
        else:
            self.root = tk.Toplevel(root)
            self.standalone = False
            
        self.root.title("録画ツール")
        # 二重起動チェック
        if not self._check_single_instance():
            self.root.destroy()
            return

        self.root.geometry("900x600")
        self.parent_app = parent_app

        # 設定読み込み
        self.save_dir = ""
        if self.parent_app:
            self.save_dir = self._load_save_dir_from_config()
        
        if not self.save_dir:
            self.save_dir = os.path.join(os.path.expanduser("~"), "Videos", "AntigravityRecorder")
        
        if not os.path.exists(self.save_dir):
            os.makedirs(self.save_dir, exist_ok=True)

        # 状態変数
        self.preview_active = True
        
        # 再生状態変数
        self.cap: Optional[cv2.VideoCapture] = None
        self.is_playing = False
        self.playback_after_id = None
        self.video_fps = 0.0
        self.video_total_frames = 0
        self.user_dragging_slider = False
        self.player_trajectory_data: List[tuple] = []
        self.last_player_frame: Optional[np.ndarray] = None
        
        # UI変数
        self.source_var = tk.StringVar(value="desktop") # desktop / window
        self.target_var = tk.StringVar()
        self.filter_var = tk.StringVar() # ウィンドウ検索用
        self.fps_var = tk.IntVar(value=15)
        self.quality_var = tk.StringVar(value="最高")
        self.save_path_var = tk.StringVar(value=self.save_dir)
        self.record_cursor_var = tk.BooleanVar(value=False) # マウスポインタ録画
        self.exclusive_window_var = tk.BooleanVar(value=False) # ウィンドウ単体キャプチャ
        self.record_tsv_var = tk.BooleanVar(value=True)
        self.seek_var = tk.DoubleVar()
        self.show_trajectory_var = tk.BooleanVar(value=True)
        
        # UIパーツ参照用
        self.widgets_to_lock: List[tk.Widget] = []
        self.region_window: Optional[tk.Toplevel] = None
        self.monitors: List[Dict[str, Any]] = []
        self.windows: List[Tuple[Any, str]] = []
        self.file_items: List[str] = [] # 一覧に表示されている実際のファイル名
        
        self._build_ui()
        self.update_source_list()
        self._start_preview()
        
        # ウィンドウ状態の復元
        self.load_window_geometry()

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def _check_single_instance(self) -> bool:
        """二重起動チェック。既に起動している場合は前面に出して終了。"""
        if not self.standalone:
            return True
        return self.window_utils.check_single_instance("AntigravityScreenRecorder_Mutex", "録画ツール")

    def load_window_geometry(self):
        """設定からウィンドウの状態を復元."""
        from config import load_global_config
        config = load_global_config()
        
        geom = config.get("recorder_geometry")
        if geom:
            try:
                self.root.geometry(geom)
            except:
                pass
        
        if config.get("recorder_maximized"):
            try:
                self.root.state('zoomed')
            except:
                pass
                
        sash_pos = config.get("recorder_sash_position")
        if sash_pos is not None:
            self.root.after(200, lambda: self.main_paned.sash_place(0, 0, sash_pos))

    def save_window_geometry(self):
        """設定にウィンドウの状態を保存."""
        from config import load_global_config, save_global_config
        config = load_global_config()
        
        config["recorder_maximized"] = (self.root.state() == 'zoomed')
        config["recorder_geometry"] = self.root.geometry()
        
        try:
            config["recorder_sash_position"] = self.main_paned.sash_coord(0)[1]
        except:
            pass
            
        save_global_config(config)

    def on_close(self):
        """終了時の処理."""
        self.save_window_geometry()
        self.preview_active = False
        
        if self.recorder_logic.is_recording:
            if messagebox.askyesno("確認", "録画中です。停止して閉じますか？"):
                self.recorder_logic.stop_recording()
                self.recorder_logic.wait_for_stop(timeout=2.0)
            else:
                return

        if self.cap:
            self.cap.release()
            
        self.root.destroy()
        if getattr(self, 'standalone', False):
            self.root.quit()

    def _load_save_dir_from_config(self) -> str:
        # 簡易実装
        return ""

    def _build_ui(self):
        # 全面スプリットフレーム (PanedWindow)
        self.main_paned = tk.PanedWindow(self.root, orient=tk.HORIZONTAL, sashwidth=4, bg="#cccccc")
        self.main_paned.pack(fill=tk.BOTH, expand=True)

        left_container = tk.Frame(self.main_paned)
        right_container = tk.Frame(self.main_paned)
        self.main_paned.add(left_container, minsize=450)
        self.main_paned.add(right_container, minsize=450)

        self._setup_left_panel(left_container)
        self._setup_right_panel(right_container)

    def _setup_left_panel(self, parent: tk.Widget):
        # 保存先設定
        path_frame = tk.LabelFrame(parent, text="保存設定")
        path_frame.pack(fill=tk.X, padx=10, pady=5)
        tk.Label(path_frame, text="保存先:").pack(side=tk.LEFT, padx=5, pady=10)
        tk.Entry(path_frame, textvariable=self.save_path_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        tk.Button(path_frame, text="...", command=self.browse_save_dir, width=3).pack(side=tk.LEFT, padx=5)

        # ファイルリスト
        list_label_frame = tk.LabelFrame(parent, text="録画済みファイル")
        list_label_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        scrollbar = tk.Scrollbar(list_label_frame, orient=tk.VERTICAL)
        self.file_listbox = tk.Listbox(list_label_frame, yscrollcommand=scrollbar.set, font=("Consolas", 10), selectmode=tk.EXTENDED)
        
        scrollbar.config(command=self.file_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=2, pady=2)
        
        self.file_listbox.bind("<Double-Button-1>", self.on_file_double_click)
        self.file_listbox.bind("<<ListboxSelect>>", self.on_file_select)
        
        # リスト操作ボタン
        list_ctrl = tk.Frame(parent)
        list_ctrl.pack(fill=tk.X, padx=10, pady=5)
        self.btn_rename = tk.Button(list_ctrl, text="名前変更", command=self.rename_file, state=tk.DISABLED, height=1, bg=self.COLOR_BTN_RENAME)
        self.btn_rename.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)
        self.btn_delete = tk.Button(list_ctrl, text="削除", command=self.delete_file, state=tk.DISABLED, height=1, bg=self.COLOR_BTN_DELETE)
        self.btn_delete.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)
        
        self.btn_open_tsv = tk.Button(list_ctrl, text="TSVを開く", command=self.open_tsv_file, state=tk.DISABLED, height=1)
        self.btn_open_tsv.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)
        
        self.widgets_to_lock.extend([self.btn_rename, self.btn_delete, self.btn_open_tsv])

        # キーバインド
        self.file_listbox.bind("<F2>", lambda e: self.rename_file())
        self.file_listbox.bind("<Delete>", lambda e: self.delete_file())
        self.file_listbox.bind("<Return>", lambda e: self.on_file_double_click(e))
        self.file_listbox.bind("<Control-a>", lambda e: self._select_all_files())
        self.file_listbox.bind("<Control-A>", lambda e: self._select_all_files())

        # フッターアクション
        footer_frame = tk.Frame(parent)
        footer_frame.pack(fill=tk.X, padx=10, pady=10)
        
        self.btn_close = tk.Button(footer_frame, text="閉じる", command=self.on_close, width=10)
        self.btn_close.pack(side=tk.RIGHT, padx=5)
        self.btn_edit = tk.Button(footer_frame, text="閉じて編集", command=self.close_and_edit, bg=self.COLOR_BTN_EDIT, width=12)
        self.btn_edit.pack(side=tk.RIGHT, padx=5)
        self.widgets_to_lock.extend([self.btn_close, self.btn_edit])

    def _setup_right_panel(self, parent: tk.Widget):
        self.notebook = ttk.Notebook(parent)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        self.tab_record = tk.Frame(self.notebook)
        self.tab_play = tk.Frame(self.notebook)
        self.notebook.add(self.tab_record, text="  録画  ")
        self.notebook.add(self.tab_play, text="  再生  ")

        self._setup_recording_tab()
        self._setup_playback_tab()

        self.refresh_file_list()

    def _setup_recording_tab(self):
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
        
        # 2. フィルター
        filter_frame = tk.Frame(self.tab_record)
        filter_frame.pack(fill=tk.X, padx=10, pady=0)
        tk.Label(filter_frame, text="検索:").pack(side=tk.LEFT)
        self.filter_var.trace_add("write", lambda *args: self.update_source_list())
        self.entry_filter = tk.Entry(filter_frame, textvariable=self.filter_var)
        self.entry_filter.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.widgets_to_lock.append(self.entry_filter)

        # 3. ターゲット選択
        target_frame = tk.Frame(self.tab_record)
        target_frame.pack(fill=tk.X, padx=10, pady=5)
        tk.Label(target_frame, text="対象:").pack(side=tk.LEFT)
        self.combo_target = ttk.Combobox(target_frame, textvariable=self.target_var, state="readonly")
        self.combo_target.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.combo_target.bind("<<ComboboxSelected>>", self.on_target_changed)
        self.btn_update = tk.Button(target_frame, text="更新", command=self.update_source_list, width=4, bg=self.COLOR_BTN_UPDATE)
        self.btn_update.pack(side=tk.LEFT)
        self.widgets_to_lock.extend([self.btn_update, self.combo_target])
        
        # 4. プレビュー
        preview_label_frame = tk.LabelFrame(self.tab_record, text="プレビュー")
        preview_label_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        self.preview_canvas = tk.Canvas(preview_label_frame, bg=self.COLOR_CANVAS_BG, highlightthickness=0)
        self.preview_canvas.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.preview_image_id = None
        self.preview_canvas.bind("<Configure>", lambda e: self._on_canvas_resize("preview"))

        # 5. 録画設定
        settings_frame = tk.Frame(self.tab_record)
        settings_frame.pack(fill=tk.X, padx=10, pady=5)
        
        tk.Label(settings_frame, text="FPS:").pack(side=tk.LEFT)
        self.combo_fps = ttk.Combobox(settings_frame, textvariable=self.fps_var, values=[15, 30, 60], width=5, state="readonly")
        self.combo_fps.pack(side=tk.LEFT, padx=5)
        
        tk.Label(settings_frame, text="画質:").pack(side=tk.LEFT, padx=(10, 0))
        self.combo_quality = ttk.Combobox(settings_frame, textvariable=self.quality_var, values=["最高", "高", "中", "低"], width=5, state="readonly")
        self.combo_quality.pack(side=tk.LEFT, padx=5)
        self.widgets_to_lock.extend([self.combo_fps, self.combo_quality])

        # 6. オプション
        options_frame = tk.Frame(self.tab_record)
        options_frame.pack(fill=tk.X, padx=10, pady=2)
        tk.Checkbutton(options_frame, text="マウスポインタを映像に含める", variable=self.record_cursor_var).pack(side=tk.LEFT)
        
        self.check_tsv = tk.Checkbutton(options_frame, text="操作履歴(TSV)", variable=self.record_tsv_var)
        self.check_tsv.pack(side=tk.LEFT, padx=5)
        self.widgets_to_lock.append(self.check_tsv)

        self.check_exclusive = tk.Checkbutton(options_frame, text="他窓除外(Win10+)", variable=self.exclusive_window_var)
        self.check_exclusive.pack(side=tk.LEFT, padx=5)
        self.widgets_to_lock.append(self.check_exclusive)

        # 7. 録画ボタン
        self.btn_record = tk.Button(self.tab_record, text="● 録画開始", bg="#ffcccc", font=("Arial", 12, "bold"),
                                    command=self.toggle_recording)
        self.btn_record.pack(fill=tk.X, padx=15, pady=10)

    def _setup_playback_tab(self):
        player_frame = tk.LabelFrame(self.tab_play, text="プレイヤー")
        player_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.player_canvas = tk.Canvas(player_frame, bg=self.COLOR_CANVAS_BG, highlightthickness=0)
        self.player_canvas.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.player_canvas.bind("<Configure>", lambda e: self._on_canvas_resize("player"))
        
        p_btns_container = tk.Frame(self.tab_play)
        p_btns_container.pack(fill=tk.X, padx=10, pady=(0, 10))

        p_btns = tk.Frame(p_btns_container)
        p_btns.pack(fill=tk.X, pady=2)
        
        self.btn_play = tk.Button(p_btns, text="▶", command=self.toggle_playback, state=tk.DISABLED, width=5)
        self.btn_play.pack(side=tk.LEFT, padx=5)
        
        self.slider = ttk.Scale(p_btns, from_=0, to=100, variable=self.seek_var, orient=tk.HORIZONTAL, command=self.on_slider_move)
        self.slider.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.slider.bind("<Button-1>", self.on_slider_press)
        self.slider.bind("<ButtonRelease-1>", self.on_slider_release)

        self.lbl_time = tk.Label(p_btns, text="00:00 / 00:00")
        self.lbl_time.pack(side=tk.LEFT, padx=5)

        self.chk_player_traj = tk.Checkbutton(p_btns_container, text="マウス軌跡を表示", variable=self.show_trajectory_var, command=self.refresh_player_canvas)
        self.chk_player_traj.pack(side=tk.LEFT, padx=5)

    def update_source_list(self):
        """録画対象のリストを更新"""
        self.combo_target['values'] = []
        mode = self.source_var.get()
        filter_text = self.filter_var.get().lower() if hasattr(self, 'filter_var') else ""
        
        if mode == 'desktop':
            # モニター一覧取得
            self.monitors = self.window_utils.get_monitor_info()
            
            display_names = [f"Display {i+1}: {m['width']}x{m['height']}" for i, m in enumerate(self.monitors)]
            self.combo_target['values'] = display_names
            if display_names:
                self.combo_target.current(0)
            self.entry_filter.config(state=tk.DISABLED)
                
        elif mode == 'window':
            # ウィンドウ一覧取得
            self.windows = self.window_utils.enum_windows(filter_text)
            
            display_names = []
            for h, t in self.windows:
                pname = self.window_utils.get_process_name(h)
                display_names.append(f"[{pname}] {t} ({h})")
                
            self.combo_target['values'] = display_names
            if display_names:
                # 現在の選択がまだリストにあれば維持、なければ最初
                curr = self.target_var.get()
                if curr in display_names:
                    self.combo_target.set(curr)
                else:
                    self.combo_target.current(0)
            else:
                self.target_var.set("")
            self.entry_filter.config(state=tk.NORMAL)
        
        self.on_target_changed(None)

    def on_target_changed(self, event):
        pass

    def _select_all_files(self):
        self.file_listbox.selection_set(0, tk.END)
        return "break"

    def _get_target_rect(self) -> Optional[Dict[str, int]]:
        """録画対象の矩形を取得"""
        mode = self.source_var.get()
        if mode == 'desktop':
            idx = self.combo_target.current()
            if idx >= 0 and idx < len(self.monitors):
                return self.monitors[idx]
            # fallback
            mons = self.window_utils.get_monitor_info()
            if mons: return mons[0]
            
        elif mode == 'window':
            idx = self.combo_target.current()
            if idx >= 0 and idx < len(self.windows):
                hwnd = self.windows[idx][0]
                return self.window_utils.get_window_rect(hwnd)
        
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

                    # ウィンドウ個別キャプチャ
                    if self.source_var.get() == 'window' and self.exclusive_window_var.get():
                        idx = self.combo_target.current()
                        if idx >= 0 and idx < len(self.windows):
                            hwnd = self.windows[idx][0]
                            frame_bgr = self.window_utils.capture_exclusive_window(hwnd)
                            if frame_bgr is not None:
                                frame_rgb = cv2.cvtColor(frame_bgr, cv2.COLOR_BGR2RGB)
                                img = Image.fromarray(frame_rgb)
                                capture_success = True
                    
                    # 通常キャプチャ
                    if not capture_success:
                        # mssのgrabはモニター座標系
                        # モニタ外の座標などを指定するとエラーになる場合があるため注意
                        # ここでは rect が正しいと仮定
                        img_sct = self.window_utils.sct.grab(rect)
                        img = Image.frombytes("RGB", img_sct.size, img_sct.bgra, "raw", "BGRX")
                        capture_success = True
                    
                    if capture_success and img:
                        # update canvas logic
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
                            
                            self.preview_canvas.image = tk_img
                except Exception as e:
                    pass

        # 録画ループ依存ではなくなったが、プレビュー更新頻度は調整
        interval = 100 if not self.recorder_logic.is_recording else 500
        self.root.after(interval, self._start_preview)

    def toggle_recording(self):
        if self.recorder_logic.is_recording:
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
        
        # UIロック
        self._set_controls_state(tk.DISABLED)
        self._show_recording_region(rect)
        
        # ウィンドウ追従のためのhwnd
        hwnd = None
        if self.source_var.get() == 'window':
            idx = self.combo_target.current()
            if idx >= 0 and idx < len(self.windows):
                hwnd = self.windows[idx][0]

        # ファイル名
        now = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{now}.mp4"
        filepath = os.path.join(self.save_path_var.get(), filename)
        
        fps = self.fps_var.get()

        self.btn_record.config(text="■ 録画停止", bg="#ff9999")
        
        self.recorder_logic.start_recording(
            filepath=filepath,
            rect=rect,
            fps=fps,
            hwnd=hwnd,
            record_cursor=self.record_cursor_var.get(),
            record_tsv=self.record_tsv_var.get(),
            exclusive_window=(self.source_var.get() == 'window' and self.exclusive_window_var.get())
        )
        
        # 赤枠追従ループ開始
        self._update_region_tracking()

    def _update_region_tracking(self):
        """録画中に赤枠を対象ウィンドウに追従させる"""
        if self.recorder_logic.is_recording and self.region_window:
            rect = self._get_target_rect()
            if rect:
                thickness = self.REGION_THICKNESS
                x = rect['left'] - thickness
                y = rect['top'] - thickness
                if self.region_window.winfo_exists():
                    self.region_window.geometry(f"+{x}+{y}")
            
            self.root.after(10, self._update_region_tracking)

    def stop_recording(self):
        self.recorder_logic.stop_recording()
        self.btn_record.config(text="処理中...", state=tk.DISABLED)
        self._wait_for_stop()

    def _wait_for_stop(self):
        """録画スレッドの終了を監視し、事後処理を行う"""
        if self.recorder_logic.recording_thread and self.recorder_logic.recording_thread.is_alive():
            self.root.after(100, self._wait_for_stop)
        else:
            self.btn_record.config(text="録画開始", state=tk.NORMAL, bg=self.COLOR_BTN_RECORD_START)
            self._set_controls_state(tk.NORMAL)
            self._hide_recording_region()
            
            # TSV保存呼び出し
            self.recorder_logic.save_trajectory_tsv()

            # ファイルリスト更新
            # ファイルが書き込まれるまで少しラグがあるかも？
            if self.recorder_logic.current_out_path:
                new_fname = os.path.basename(self.recorder_logic.current_out_path)
                self.refresh_file_list(select_filename=new_fname)

    def _set_controls_state(self, state):
        self.file_listbox.config(state=state)
        for w in self.widgets_to_lock:
            try:
                target_state = state
                if state == tk.NORMAL and isinstance(w, ttk.Combobox):
                    target_state = "readonly"
                w.config(state=target_state)
            except tk.TclError:
                pass

        if state == tk.DISABLED:
            self.btn_record.config(text="■ 録画停止", bg=self.COLOR_BTN_RECORD_STOP)
        else:
            self.btn_record.config(text="● 録画開始", bg=self.COLOR_BTN_RECORD_START)

    def _show_recording_region(self, rect):
        if self.region_window: self._hide_recording_region()
        
        self.region_window = tk.Toplevel(self.root)
        self.region_window.overrideredirect(True)
        self.region_window.attributes("-topmost", True)
        self.region_window.attributes("-transparentcolor", "white")
        
        thickness = self.REGION_THICKNESS
        x = rect['left'] - thickness
        y = rect['top'] - thickness
        w = rect['width'] + thickness * 2
        h = rect['height'] + thickness * 2
        
        self.region_window.geometry(f"{w}x{h}+{x}+{y}")
        canvas = tk.Canvas(self.region_window, width=w, height=h, bg="white", highlightthickness=0)
        canvas.pack()
        
        canvas.create_rectangle(thickness//2, thickness//2, w - thickness//2, h - thickness//2, outline=self.REGION_COLOR, width=thickness)

    def _hide_recording_region(self):
        if self.region_window:
            self.region_window.destroy()
            self.region_window = None

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
            
            files_info.sort(key=lambda x: x[0], reverse=True)

            for mtime, f, dt in files_info:
                display_str = f"{f:<30} | {dt}"
                self.file_listbox.insert(tk.END, display_str)
                self.file_items.append(f)
                
                if select_filename and f == select_filename:
                    idx = self.file_listbox.size() - 1
                    self.file_listbox.selection_set(idx)
                    self.file_listbox.see(idx)
                    self.file_listbox.activate(idx)

        self.on_file_select(None)

    def open_tsv_file(self):
        indices = self.file_listbox.curselection()
        if not indices: return
        
        fname = self.file_items[indices[0]]
        path = os.path.join(self.save_dir, fname)
        tsv_path = os.path.splitext(path)[0] + '.tsv'
        
        if os.path.exists(tsv_path):
            try:
                os.startfile(tsv_path)
            except Exception as e:
                messagebox.showerror("Error", f"TSVファイルを開けませんでした:\n{e}")
        else:
            messagebox.showwarning("Warning", "対応するTSVファイルが見つかりません。")

    def on_file_select(self, event):
        idx = self.file_listbox.curselection()
        if idx:
            self.btn_rename.config(state=tk.NORMAL)
            self.btn_delete.config(state=tk.NORMAL)
            self.btn_play.config(state=tk.NORMAL)
            
            fname = self.file_items[idx[0]]
            path = os.path.join(self.save_dir, fname)
            tsv_path = os.path.splitext(path)[0] + '.tsv'
            if os.path.exists(tsv_path):
                self.btn_open_tsv.config(state=tk.NORMAL)
            else:
                self.btn_open_tsv.config(state=tk.DISABLED)
                
            self.stop_playback()
            self.load_video_for_playback()
        else:
            self.btn_rename.config(state=tk.DISABLED)
            self.btn_delete.config(state=tk.DISABLED)
            self.btn_play.config(state=tk.DISABLED)
            self.btn_open_tsv.config(state=tk.DISABLED)
            self.stop_playback()
            self.clear_player_canvas()

    def on_file_double_click(self, event):
        idx = self.file_listbox.curselection()
        if idx:
            self.notebook.select(self.tab_play)
            self.load_video_for_playback()
            if not self.is_playing:
                self.toggle_playback()

    def rename_file(self):
        indices = self.file_listbox.curselection()
        if not indices: return
        
        is_multiple = len(indices) > 1
        title = "一括名前変更 (接頭辞付与)" if is_multiple else "名前変更"
        label_text = "接頭辞を入力してください:" if is_multiple else "新しいファイル名:"
        
        dialog = tk.Toplevel(self.root)
        dialog.title(title)
        dialog.geometry("350x150")
        dialog.transient(self.root)
        dialog.grab_set()

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
                    new_name = new_val + "_" + base + ext
                else:
                    new_name = new_val + ext
                
                new_path = os.path.join(self.save_dir, new_name)
                
                try:
                    os.rename(old_path, new_path)
                    
                    old_tsv = os.path.splitext(old_path)[0] + '.tsv'
                    if os.path.exists(old_tsv):
                        new_tsv = os.path.splitext(new_path)[0] + '.tsv'
                        try:
                            os.rename(old_tsv, new_tsv)
                        except Exception as te:
                            print(f"TSV rename error: {te}")
                            
                    success_count += 1
                except Exception as e:
                    print(f"Rename error for {fname}: {e}")

            if success_count > 0:
                self.refresh_file_list()
            
            dialog.destroy()
            self.file_listbox.focus_set()

        def on_cancel(event=None):
            dialog.destroy()
            self.file_listbox.focus_set()

        btn_frame = tk.Frame(dialog)
        btn_frame.pack(pady=10)
        tk.Button(btn_frame, text="OK", command=on_ok, width=10).pack(side=tk.LEFT, padx=5)
        self.root.bind("<Return>", on_ok) # entryだけにbindしてもいいが念のため

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
            self.stop_playback()
            if self.cap:
                self.cap.release()
                self.cap = None

            success_count = 0
            for i in sorted(indices, reverse=True):
                fname = self.file_items[i]
                path = os.path.join(self.save_dir, fname)
                try:
                    os.remove(path)
                    
                    tsv_path = os.path.splitext(path)[0] + '.tsv'
                    if os.path.exists(tsv_path):
                        try:
                            os.remove(tsv_path)
                        except Exception as te:
                            print(f"TSV delete error: {te}")
                            
                    success_count += 1
                except Exception as e:
                    print(f"Delete error for {fname}: {e}")

            self.refresh_file_list()

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
        
        self.load_video_trajectory(path)
        self.show_frame(0)
        self.update_time_label(0)

    def load_video_trajectory(self, video_path):
        self.player_trajectory_data = []
        tsv_path = os.path.splitext(video_path)[0] + ".tsv"
        if os.path.exists(tsv_path):
            try:
                with open(tsv_path, "r", encoding="utf-8") as f:
                    next(f, None)
                    for line in f:
                        parts = line.strip().split("\t")
                        if len(parts) >= 4:
                            # timestamp, frame, x, y, click, keys
                            click_info = parts[4] if len(parts) > 4 else "None"
                            keys_info = parts[5] if len(parts) > 5 else "None"
                            self.player_trajectory_data.append((
                                float(parts[0]), 
                                int(parts[1]), 
                                int(parts[2]), 
                                int(parts[3]),
                                click_info,
                                keys_info
                            ))
            except Exception as e:
                print(f"Player TSV load error: {e}")

    def refresh_player_canvas(self):
        if hasattr(self, 'last_player_frame') and self.last_player_frame is not None:
            self.display_frame(self.last_player_frame)

    def toggle_playback(self):
        if self.is_playing:
            self.stop_playback()
        else:
            self.start_playback()

    def start_playback(self):
        if not self.cap or not self.cap.isOpened():
            return
        
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

    def display_frame(self, frame: np.ndarray):
        if frame is not None:
            self.last_player_frame = frame.copy()
            frame_to_disp = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            img = Image.fromarray(frame_to_disp)
        elif self.last_player_frame is not None:
            frame_to_disp = cv2.cvtColor(self.last_player_frame, cv2.COLOR_BGR2RGB)
            img = Image.fromarray(frame_to_disp)
        else:
            return
        
        cw = self.player_canvas.winfo_width()
        ch = self.player_canvas.winfo_height()
        if cw > 1 and ch > 1:
            img.thumbnail((cw, ch), Image.Resampling.LANCZOS)
            tk_img = ImageTk.PhotoImage(img)
            self.player_canvas.delete("all")
            self.player_canvas.create_image(cw//2, ch//2, image=tk_img, anchor=tk.CENTER, tags="img")
            self.player_canvas.image = tk_img
            
            # 軌跡描画
            if self.show_trajectory_var.get() and self.player_trajectory_data:
                curr_pos = self.cap.get(cv2.CAP_PROP_POS_MSEC) / 1000.0 if self.cap else 0
                closest = None
                for row in self.player_trajectory_data:
                    if abs(row[0] - curr_pos) < 0.1:
                        closest = row
                        break
                
                if closest:
                    t, f_idx, vx, vy, click, keys = closest
                    img_w, img_h = img.size
                    orig_w = self.cap.get(cv2.CAP_PROP_FRAME_WIDTH)
                    orig_h = self.cap.get(cv2.CAP_PROP_FRAME_HEIGHT)
                    
                    if orig_w > 0 and orig_h > 0:
                        scale_x = img_w / orig_w
                        scale_y = img_h / orig_h
                        off_x = (cw - img_w) // 2
                        off_y = (ch - img_h) // 2
                        cx = vx * scale_x + off_x
                        cy = vy * scale_y + off_y
                        
                        color = "red"
                        outline_width = 2
                        if click != "None":
                            color = "yellow"
                            outline_width = 4
                            self.player_canvas.create_oval(cx-12, cy-12, cx+12, cy+12, outline="yellow", width=1, tags="overlay")
                        
                        r = 6
                        self.player_canvas.create_oval(cx-r, cy-r, cx+r, cy+r, outline=color, width=outline_width, tags="overlay")
                        self.player_canvas.create_oval(cx-2, cy-2, cx+2, cy+2, fill=color, tags="overlay")

                        if keys != "None":
                            self.player_canvas.create_text(
                                cw // 2, ch - 30, 
                                text=f"Keys: {keys}", 
                                fill="yellow", 
                                font=("Arial", 14, "bold"),
                                tags="overlay"
                            )

    def _on_canvas_resize(self, mode):
        if mode == "player" and self.last_player_frame is not None:
            if not self.is_playing:
                self.display_frame(None)

    def update_time_label(self, curr_frame):
        if hasattr(self, 'video_fps') and self.video_fps > 0:
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

    def close_and_edit(self):
        idx = self.file_listbox.curselection()
        if idx:
            fname = self.file_items[idx[0]]
            path = os.path.join(self.save_dir, fname)
            
            if self.parent_app and hasattr(self.parent_app, 'open_video_file'):
                self.parent_app.open_video_file(path)
        
        self.on_close()


if __name__ == "__main__":
    app = ScreenRecorderApp()
    app.root.mainloop()
