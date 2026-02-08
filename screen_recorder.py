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
from typing import TYPE_CHECKING, Optional, List, Tuple, Dict, Any

import cv2
import mss
import numpy as np
from PIL import Image, ImageTk

from config import get_base_dir, load_global_config, save_global_config, PROJECT_NAME
from utils import resource_path
from ui_utils import add_tooltip
from window_utils import WindowUtils
from recorder_core import ScreenRecorderLogic
import overlay_utils

if TYPE_CHECKING:
    from ChulipVideo import VideoCropperApp


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
            
        self.root.title(f"{PROJECT_NAME} 録画ツール")
        # アイコン設定
        try:
            self.icon_image = tk.PhotoImage(file=resource_path("ChulipVideo.png"))
            self.root.iconphoto(True, self.icon_image)
        except Exception as e:
            print(f"Recorder Icon Load Error: {e}")
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
            self.save_dir = os.path.join(os.path.expanduser("~"), "Videos", PROJECT_NAME)
        
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
        
        # 赤枠表示の安定性管理用
        self.last_target_rect = None
        self.last_rect_change_time = 0.0
        self.STABILITY_THRESHOLD = 0.1 # 秒
        
        # UI変数
        self.source_var = tk.StringVar(value="desktop") # desktop / window
        self.target_var = tk.StringVar()
        self.filter_var = tk.StringVar() # ウィンドウ検索用
        self.fps_var = tk.IntVar(value=15)
        self.quality_var = tk.StringVar(value="最高")
        self.save_path_var = tk.StringVar(value=self.save_dir)
        self.record_cursor_var = tk.BooleanVar(value=False) # 機能削除（変数は残すがUIは消す）
        self.show_region_var = tk.BooleanVar(value=True)    # 録画枠を表示するか
        self.exclusive_window_var = tk.BooleanVar(value=True) # ウィンドウ単体キャプチャ (デフォルトON)
        self.record_tsv_var = tk.BooleanVar(value=True)
        self.seek_var = tk.DoubleVar()
        self.show_trajectory_var = tk.BooleanVar(value=True)
        self.player_fit_var = tk.BooleanVar(value=True) # キャンバスに合わせる
        
        # プレイヤーのパン・ズーム用
        self.player_zoom = 1.0
        self.player_pan_x = 0
        self.player_pan_y = 0
        self._player_panning = False
        self._player_pan_start = (0, 0)

        # プレビューのパン・ズーム用
        self.preview_fit_var = tk.BooleanVar(value=True)
        self.preview_zoom = 1.0
        self.preview_pan_x = 0
        self.preview_pan_y = 0
        self._preview_panning = False
        self._preview_pan_start = (0, 0)

        # 共通設定 (テーマ) のロード
        self.global_config = load_global_config()
        self.theme = self.global_config.get("theme", {})
        
        # 設定から復元
        saved_source = self.global_config.get("recorder_source", "desktop")
        if saved_source in ["desktop", "window"]:
            self.source_var.set(saved_source)
        
        self.show_region_var.set(self.global_config.get("recorder_show_region", True))

        # UIパーツ参照用

        self.widgets_to_lock: List[tk.Widget] = []
        self.region_window: Optional[tk.Toplevel] = None
        self.monitors: List[Dict[str, Any]] = []
        self.windows: List[Tuple[Any, str, str, int]] = [] # (hwnd, title, process_name, pid)
        self.file_items: List[str] = [] # 一覧に表示されている実際のファイル名
        
        self._build_ui()
        self.update_source_list()
        self._start_preview()
        
        # ウィンドウ状態の復元
        self.load_window_geometry()

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        
        # タブ切り替えイベント
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_changed)
        
        # 初期表示の追従開始
        if self.notebook.select() == str(self.tab_record):
            self._update_region_tracking()

    def on_tab_changed(self, event):
        """タブ切り替え時の処理"""
        current_tab = self.notebook.select()
        if current_tab == str(self.tab_record):
            self._update_region_tracking()
        else:
            self._hide_recording_region()

    def _check_single_instance(self) -> bool:
        """二重起動チェック。既に起動している場合は前面に出して終了。"""
        if not self.standalone:
            return True
        return self.window_utils.check_single_instance(f"{PROJECT_NAME}ScreenRecorder_Mutex", "録画ツール")

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
        config["recorder_source"] = self.source_var.get()
        config["recorder_show_region"] = self.show_region_var.get()
        
        # ターゲット情報の保存
        self._save_recording_target_to_config(config)
        
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
            if messagebox.askyesno("確認", "録画中です。停止して閉じますか？", parent=self.root):
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

        # 3.5 座標・サイズ入力
        geo_frame = tk.Frame(self.tab_record)
        geo_frame.pack(fill=tk.X, padx=10, pady=2)
        
        self.geo_x = tk.IntVar()
        self.geo_y = tk.IntVar()
        self.geo_w = tk.IntVar()
        self.geo_h = tk.IntVar()
        
        tk.Label(geo_frame, text="座標 X:").pack(side=tk.LEFT)
        self.entry_x = tk.Entry(geo_frame, textvariable=self.geo_x, width=5)
        self.entry_x.pack(side=tk.LEFT, padx=2)
        
        tk.Label(geo_frame, text="Y:").pack(side=tk.LEFT)
        self.entry_y = tk.Entry(geo_frame, textvariable=self.geo_y, width=5)
        self.entry_y.pack(side=tk.LEFT, padx=2)
        
        tk.Label(geo_frame, text="W:").pack(side=tk.LEFT)
        self.entry_w = tk.Entry(geo_frame, textvariable=self.geo_w, width=5)
        self.entry_w.pack(side=tk.LEFT, padx=2)
        
        tk.Label(geo_frame, text="H:").pack(side=tk.LEFT)
        self.entry_h = tk.Entry(geo_frame, textvariable=self.geo_h, width=5)
        self.entry_h.pack(side=tk.LEFT, padx=2)
        
        self.btn_apply_geo = tk.Button(geo_frame, text="適用", command=self.apply_window_geometry, width=4, bg=self.COLOR_BTN_UPDATE)
        self.btn_apply_geo.pack(side=tk.LEFT, padx=5)
        
        for entry in [self.entry_x, self.entry_y, self.entry_w, self.entry_h]:
            entry.bind("<Return>", lambda e: self.apply_window_geometry())
            entry.bind("<Tab>", lambda e: self.apply_window_geometry())
            self.widgets_to_lock.append(entry)
        self.widgets_to_lock.append(self.btn_apply_geo)

        
        # 4. プレビュー
        preview_label_frame = tk.LabelFrame(self.tab_record, text="プレビュー")
        preview_label_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        self.preview_canvas = tk.Canvas(preview_label_frame, bg=self.COLOR_CANVAS_BG, highlightthickness=0)
        self.preview_canvas.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.preview_image_id = None
        self.preview_canvas.bind("<Configure>", lambda e: self._on_canvas_resize("preview"))
        self.preview_canvas.bind("<ButtonPress-2>", self._on_preview_middle_down)
        self.preview_canvas.bind("<B2-Motion>", self._on_preview_middle_drag)
        self.preview_canvas.bind("<ButtonRelease-2>", self._on_preview_middle_up)
        self.preview_canvas.bind("<Double-Button-2>", self._on_preview_middle_double_click)
        self.preview_canvas.bind("<MouseWheel>", self._on_preview_wheel)
        self.preview_canvas.bind("<Button-4>", self._on_preview_wheel)
        self.preview_canvas.bind("<Button-5>", self._on_preview_wheel)

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
        
        self.check_show_region = tk.Checkbutton(options_frame, text="録画枠を表示", variable=self.show_region_var)
        self.check_show_region.pack(side=tk.LEFT)
        self.widgets_to_lock.append(self.check_show_region)
        
        self.check_tsv = tk.Checkbutton(options_frame, text="操作履歴(TSV)", variable=self.record_tsv_var)
        self.check_tsv.pack(side=tk.LEFT, padx=5)
        self.widgets_to_lock.append(self.check_tsv)

        self.check_exclusive = tk.Checkbutton(options_frame, text="WGCで選択したウィンドウのみキャプチャする(Win10以降)", variable=self.exclusive_window_var)
        self.check_exclusive.pack(side=tk.LEFT, padx=5)
        self.widgets_to_lock.append(self.check_exclusive)

        self.check_preview_fit = tk.Checkbutton(options_frame, text="プレビューを拡大・縮小する", variable=self.preview_fit_var)
        self.check_preview_fit.pack(side=tk.LEFT)
        self.widgets_to_lock.append(self.check_preview_fit)

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

        self.chk_player_traj = tk.Checkbutton(p_btns_container, text="マウス・キー入力の軌跡を表示", variable=self.show_trajectory_var, command=self.refresh_player_canvas)
        self.chk_player_traj.pack(side=tk.LEFT, padx=5)

        self.chk_player_fit = tk.Checkbutton(p_btns_container, text="プレビューを拡大・縮小する", variable=self.player_fit_var, command=self.refresh_player_canvas)
        self.chk_player_fit.pack(side=tk.RIGHT, padx=5)

        # プレイヤーキャンバスへのバインド (パン・ズーム)
        self.player_canvas.bind("<ButtonPress-2>", self._on_player_middle_down)
        self.player_canvas.bind("<B2-Motion>", self._on_player_middle_drag)
        self.player_canvas.bind("<ButtonRelease-2>", self._on_player_middle_up)
        self.player_canvas.bind("<Double-Button-2>", self._on_player_middle_double_click)
        self.player_canvas.bind("<MouseWheel>", self._on_player_wheel)
        self.player_canvas.bind("<Button-4>", self._on_player_wheel)
        self.player_canvas.bind("<Button-5>", self._on_player_wheel)

    def _save_recording_target_to_config(self, config: Dict[str, Any]):
        """録画対象の情報を設定辞書に保存 (save_window_geometryから呼ばれる)"""
        mode = self.source_var.get()
        if mode == "desktop":
            config["recorder_desktop_index"] = self.combo_target.current()
        elif mode == "window":
            idx = self.combo_target.current()
            if idx >= 0 and idx < len(self.windows):
                # self.windows = [(hwnd, title, pname, pid), ...]
                hwnd, title, pname, pid = self.windows[idx]
                config["recorder_target_pid"] = pid
                config["recorder_target_process"] = pname
                config["recorder_target_title"] = title

    def _restore_recording_target_from_config(self, display_names: List[str]):
        """設定から録画対象を復元してコンボボックスを選択状態にする"""
        mode = self.source_var.get()
        
        # 1. Desktop Mode
        if mode == 'desktop':
            saved_idx = self.global_config.get("recorder_desktop_index")
            if saved_idx is not None and isinstance(saved_idx, int):
                if 0 <= saved_idx < len(display_names):
                    self.combo_target.current(saved_idx)
            return

        # 2. Window Mode
        if mode == 'window':
            # まず現在の選択が有効ならそれを維持 (リスト更新時など)
            curr = self.target_var.get()
            if curr in display_names:
                self.combo_target.set(curr)
                return

            # 無効(リストにない)または空なら、保存された設定からの復元を試みる
            target_pid = self.global_config.get("recorder_target_pid")
            target_proc = self.global_config.get("recorder_target_process")
            
            if target_pid or target_proc:
                best_idx = 0
                found = False
                
                # Priority 1: PID Match
                if target_pid:
                    for i, win in enumerate(self.windows):
                        # win = (hwnd, title, pname, pid)
                        if win[3] == target_pid:
                            best_idx = i
                            found = True
                            break
                
                # Priority 2: Process Name Match (if PID mismatch)
                if not found and target_proc:
                    for i, win in enumerate(self.windows):
                        if win[2] == target_proc:
                            best_idx = i
                            found = True
                            break
                
                self.combo_target.current(best_idx)
            else:
                # Default to first
                if display_names:
                    self.combo_target.current(0)

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
                self._restore_recording_target_from_config(display_names)
                if self.combo_target.current() == -1:
                    self.combo_target.current(0)
            else:
                self.target_var.set("")
            
            self.entry_filter.config(state=tk.DISABLED)
            for w in [self.entry_x, self.entry_y, self.entry_w, self.entry_h, self.btn_apply_geo]:
                w.config(state=tk.DISABLED)
                
        elif mode == 'window':
            # ウィンドウ一覧取得
            self.windows = self.window_utils.enum_windows(filter_text)
            display_names = [f"[{p}] {t} ({h})" for h, t, p, pid in self.windows]
            self.combo_target['values'] = display_names
            
            if display_names:
                self._restore_recording_target_from_config(display_names)
            else:
                self.target_var.set("")
            
            self.entry_filter.config(state=tk.NORMAL)
            for w in [self.entry_x, self.entry_y, self.entry_w, self.entry_h, self.btn_apply_geo]:
                w.config(state=tk.NORMAL)

        self.on_target_changed(None)

    def on_target_changed(self, event):
        rect = self._get_target_rect()
        if rect:
            self.geo_x.set(rect['left'])
            self.geo_y.set(rect['top'])
            self.geo_w.set(rect['width'])
            self.geo_h.set(rect['height'])
            # 赤枠の表示・更新は _update_region_tracking 内の安定性ロジックに任せる

    def apply_window_geometry(self):
        """入力ボックスの値でウィンドウを移動・リサイズ"""
        if self.source_var.get() == 'window':
            idx = self.combo_target.current()
            if idx >= 0 and idx < len(self.windows):
                hwnd = self.windows[idx][0]
                try:
                    x = self.geo_x.get()
                    y = self.geo_y.get()
                    w = self.geo_w.get()
                    h = self.geo_h.get()
                    self.window_utils.set_window_position(hwnd, x, y, w, h)
                    # 反映を確認するために少し待ってから更新
                    self.root.after(100, lambda: self._force_update_target_info())
                except Exception as e:
                    print(f"Geometry apply error: {e}")

    def _force_update_target_info(self):
        self.on_target_changed(None)


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
                            img_w, img_h = img.size
                            if self.preview_fit_var.get():
                                # 比例リサイズ
                                ratio = min(cw / img_w, ch / img_h)
                                new_w = int(img_w * ratio)
                                new_h = int(img_h * ratio)
                                if new_w > 0 and new_h > 0:
                                    img = img.resize((new_w, new_h), Image.Resampling.LANCZOS)
                                
                                # 中央配置
                                off_x = (cw - img.width) // 2
                                off_y = (ch - img.height) // 2
                            else:
                                # パン・ズーム
                                scale_view = self.preview_zoom
                                new_w = int(img_w * scale_view)
                                new_h = int(img_h * scale_view)
                                if new_w > 0 and new_h > 0:
                                    img = img.resize((new_w, new_h), Image.Resampling.LANCZOS)
                                
                                off_x = cw // 2 + self.preview_pan_x - img.width // 2
                                off_y = ch // 2 + self.preview_pan_y - img.height // 2

                            tk_img = ImageTk.PhotoImage(img)
                            
                            if self.preview_image_id:
                                self.preview_canvas.itemconfig(self.preview_image_id, image=tk_img)
                                # アンカーをNWに変更して座標指定
                                self.preview_canvas.coords(self.preview_image_id, off_x, off_y)
                                self.preview_canvas.itemconfig(self.preview_image_id, anchor=tk.NW)
                            else:
                                self.preview_image_id = self.preview_canvas.create_image(off_x, off_y, image=tk_img, anchor=tk.NW)
                            
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
            self.start_recording_sequence()

    def start_recording_sequence(self):
        rect = self._get_target_rect()
        if not rect:
            messagebox.showerror("Error", "録画対象が取得できません", parent=self.root)
            return
            
        # 録画開始ボタンを押した瞬間にUIをロック
        self._set_controls_state(tk.DISABLED)

        # カウントダウン表示 -> 終了後に start_recording 実行
        self._show_countdown(3, self.start_recording)

    def _show_countdown(self, count, callback):
        """画面中央にカウントダウンを表示"""
        if count > 0:
            # オーバーレイウィンドウ作成（透過）
            w = tk.Toplevel(self.root)
            w.overrideredirect(True)
            w.attributes("-topmost", True)
            w.attributes("-transparentcolor", "white")
            w.attributes("-alpha", 0.8)
            
            screen_w = self.root.winfo_screenwidth()
            screen_h = self.root.winfo_screenheight()
            
            # 対象領域の中心に表示したいが、簡易的に画面中央または対象領域中央
            rect = self._get_target_rect()
            if rect:
                cx = rect['left'] + rect['width'] // 2
                cy = rect['top'] + rect['height'] // 2
            else:
                cx = screen_w // 2
                cy = screen_h // 2
                
            size = 200
            w.geometry(f"{size}x{size}+{cx - size//2}+{cy - size//2}")
            
            lbl = tk.Label(w, text=str(count), font=("Arial", 100, "bold"), fg="red", bg="white")
            lbl.pack(fill=tk.BOTH, expand=True)
            
            # 1秒後に次へ
            self.root.after(1000, lambda: [w.destroy(), self._show_countdown(count - 1, callback)])
        else:
            callback()

    def start_recording(self):

        rect = self._get_target_rect()
        if not rect:
            messagebox.showerror("Error", "録画対象が取得できません", parent=self.root)
            return
        
        
        # UIロック (start_recording_sequence で実施済み)
        # self._set_controls_state(tk.DISABLED)
        # self._show_recording_region(rect) # 既に表示されているはずだが念のため
        
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
            record_tsv=self.record_tsv_var.get(),
            exclusive_window=(self.source_var.get() == 'window' and self.exclusive_window_var.get())
        )
        
        # 赤枠追従ループ開始
        self._update_region_tracking()

    def _update_region_tracking(self):
        """録画中または録画タブ表示中に赤枠を対象ウィンドウに追従させる"""
        # 条件: 録画中 OR 録画タブ表示中
        is_in_record_tab = (self.notebook.select() == str(self.tab_record))
        is_recording = self.recorder_logic.is_recording
        
        # 実際に枠を表示するかどうか
        # show_region_var が OFF の場合は即座に消す
        should_show = self.show_region_var.get() and (is_recording or is_in_record_tab)
        
        if not self.show_region_var.get():
            self._hide_recording_region()

        if should_show:
            rect = self._get_target_rect()
            
            # --- 赤枠表示の安定性ロジック ---
            is_stable = False
            if rect:
                if self.last_target_rect != rect:
                    # 矩形が変化した -> 非表示にしてタイマーリセット
                    self.last_target_rect = rect
                    self.last_rect_change_time = time.time()
                    self._hide_recording_region()
                else:
                    # 矩形が変化していない -> 指定時間経過したか確認
                    if time.time() - self.last_rect_change_time >= self.STABILITY_THRESHOLD:
                        is_stable = True
            
            if is_stable and rect:
                # 安定している場合のみ表示・更新
                if not self.region_window:
                    self._show_recording_region(rect)
                elif self.region_window.winfo_exists():
                    thickness = self.REGION_THICKNESS
                    x = rect['left'] - thickness
                    y = rect['top'] - thickness
                    w = rect['width'] + thickness * 2
                    h = rect['height'] + thickness * 2
                    self.region_window.geometry(f"{w}x{h}+{x}+{y}")
                    
                    canvas = self.region_window.winfo_children()[0]
                    if isinstance(canvas, tk.Canvas):
                        if canvas.winfo_width() != w or canvas.winfo_height() != h:
                            canvas.config(width=w, height=h)
                            canvas.delete("all")
                            canvas.create_rectangle(thickness//2, thickness//2, w - thickness//2, h - thickness//2, outline=self.REGION_COLOR, width=thickness)
            # 安定していない場合は何もしない（前回の _hide で消えているはず）
            # ------------------------------
            
            # 座標・サイズのUI自動更新
            # ユーザーが入力中でない場合のみ更新する
            try:
                focused_widget = self.root.focus_get()
            except KeyError:
                # Comboboxのドロップダウンなどがフォーカスを持っている場合 KeyError: 'popdown' が出ることがある
                focused_widget = None
            except Exception:
                focused_widget = None

            input_widgets = [self.entry_x, self.entry_y, self.entry_w, self.entry_h]
            if focused_widget not in input_widgets:
                if rect:
                    if self.geo_x.get() != rect['left']: self.geo_x.set(rect['left'])
                    if self.geo_y.get() != rect['top']: self.geo_y.set(rect['top'])
                    if self.geo_w.get() != rect['width']: self.geo_w.set(rect['width'])
                    if self.geo_h.get() != rect['height']: self.geo_h.set(rect['height'])

        # 録画中または録画タブならループを継続
        if is_recording or is_in_record_tab:
            self.root.after(50, self._update_region_tracking)
        else:
            self._hide_recording_region()

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
            # self._hide_recording_region() # 停止後も録画タブなら表示し続ける
            if self.notebook.select() != str(self.tab_record):
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
        
        try:
            import ctypes
            self.region_window.update_idletasks() # Ensure HWND is ready
            # winfo_id() returns the internal widget HWND, not the toplevel frame
            # We need to get the root ancestor (GA_ROOT = 2) for SetWindowDisplayAffinity
            internal_hwnd = self.region_window.winfo_id()
            GA_ROOT = 2
            hwnd = ctypes.windll.user32.GetAncestor(internal_hwnd, GA_ROOT)
            if hwnd == 0:
                hwnd = internal_hwnd  # Fallback to internal hwnd
            
            # print(f"Setting affinity for HWND: {hwnd} (internal: {internal_hwnd})")
            ret = self.window_utils.set_window_display_affinity(hwnd, True)
            if not ret:
                # GetLastError for debugging
                error_code = ctypes.windll.kernel32.GetLastError()
                # print(f"Failed to set display affinity for HWND: {hwnd}, Error: {error_code}")
            # else:
            #    print(f"Successfully set display affinity for HWND: {hwnd}")
        except Exception as e:
            pass

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
                messagebox.showerror("Error", f"TSVファイルを開けませんでした:\n{e}", parent=self.root)
        else:
            messagebox.showwarning("Warning", "対応するTSVファイルが見つかりません。", parent=self.root)

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
        
        if messagebox.askyesno("確認", msg, parent=self.root):
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
            img_w, img_h = img.size
            
            if self.player_fit_var.get():
                # キャンバスに適合 (比例リサイズ)
                ratio = min(cw / img_w, ch / img_h)
                new_w = int(img_w * ratio)
                new_h = int(img_h * ratio)
                if new_w > 0 and new_h > 0:
                    img = img.resize((new_w, new_h), Image.Resampling.LANCZOS)
                
                # 中央表示用のオフセット (オーバーレイ描画の座標計算に使用)
                off_x = (cw - img.width) // 2
                off_y = (ch - img.height) // 2
                scale_view = ratio
            else:
                # 自由変形 (ズーム・パン)
                scale_view = self.player_zoom
                new_w = int(img_w * scale_view)
                new_h = int(img_h * scale_view)
                if new_w > 0 and new_h > 0:
                    img = img.resize((new_w, new_h), Image.Resampling.LANCZOS)
                
                # キャンバス中央基準でパンを適用
                off_x = cw // 2 + self.player_pan_x - img.width // 2
                off_y = ch // 2 + self.player_pan_y - img.height // 2

            # 以降の計算で使用する img スケール (元動画 -> 表示画像)
            img_w_disp, img_h_disp = img.size
            orig_w = self.cap.get(cv2.CAP_PROP_FRAME_WIDTH)
            orig_h = self.cap.get(cv2.CAP_PROP_FRAME_HEIGHT)
            scale_x = img_w_disp / orig_w if orig_w > 0 else 1.0
            scale_y = img_h_disp / orig_h if orig_h > 0 else 1.0

            # 軌跡・オーバーレイ描画
            if self.show_trajectory_var.get() and self.player_trajectory_data:
                curr_pos = self.cap.get(cv2.CAP_PROP_POS_MSEC) / 1000.0 if self.cap else 0
                
                # 1. マウス軌跡
                mouse_data = None
                current_row_idx = -1
                for i, row in enumerate(self.player_trajectory_data):
                    t, f_idx, vx, vy, click, keys = row
                    if abs(t - curr_pos) < 0.1:
                        mouse_data = row
                        current_row_idx = i
                        break
                
                if mouse_data:
                    t_curr, f_idx, vx, vy, click, keys = mouse_data
                    # ripple logic / overlay_utils 呼び出し ... (省略せず元のロジックをスケールに合わせる)
                    ripple_age = 0.0
                    ripple_type = ""
                    lookback_sec = 0.5
                    if current_row_idx > 0:
                        for j in range(current_row_idx, 0, -1):
                            row_p = self.player_trajectory_data[j-1]
                            row_c = self.player_trajectory_data[j]
                            t_p, c_p = row_p[0], row_p[4]
                            t_c, c_c = row_c[0], row_c[4]
                            if t_curr - t_c > lookback_sec: break
                            for char, name in [("L", "left"), ("R", "right"), ("M", "middle")]:
                                if char in c_p and char not in c_c:
                                    ripple_type = name
                                    ripple_age = t_curr - t_c
                                    break
                            if ripple_type: break

                    overlay_utils.draw_mouse_overlay(
                        img, vx, vy, click, 
                        scale_x, scale_y, 
                        self.theme,
                        ripple_age=ripple_age,
                        ripple_type=ripple_type
                    )

                # 2. 入力履歴
                fade_duration = self.theme.get("input_overlay", {}).get("fade_duration", 2.0)
                history_inputs = []
                last_item = None
                
                start_idx = current_row_idx if current_row_idx >= 0 else 0
                for i in range(start_idx, -1, -1):
                    t, f_idx, vx, vy, click, keys = self.player_trajectory_data[i]
                    if curr_pos - t > fade_duration: break
                    item_text = overlay_utils.get_input_display_text(click, keys)
                    if not item_text: continue
                    if item_text != last_item:
                        history_inputs.append((item_text, curr_pos - t))
                        last_item = item_text
                
                if history_inputs:
                    overlay_utils.draw_input_overlay(img, history_inputs, scale_x, scale_y, self.theme)

            tk_img = ImageTk.PhotoImage(img)
            self.player_canvas.delete("all")
            # 指定されたオフセット (off_x, off_y) に合わせて描画 (anchor=NW)
            self.player_canvas.create_image(off_x, off_y, image=tk_img, anchor=tk.NW, tags="img")
            self.player_canvas.image = tk_img

    # プレイヤーのパンニング・ズームイベント
    def _on_player_middle_down(self, event):
        self._player_panning = True
        self._player_pan_start = (event.x, event.y)
        self.player_canvas.config(cursor="fleur")

    def _on_player_middle_drag(self, event):
        if self._player_panning:
            dx = event.x - self._player_pan_start[0]
            dy = event.y - self._player_pan_start[1]
            self.player_pan_x += dx
            self.player_pan_y += dy
            self._player_pan_start = (event.x, event.y)
            self.refresh_player_canvas()

    def _on_player_middle_up(self, event):
        self._player_panning = False
        self.player_canvas.config(cursor="")

    def _on_player_middle_double_click(self, event):
        self.player_pan_x = 0
        self.player_pan_y = 0
        self.player_zoom = 1.0
        self.refresh_player_canvas()

    def _on_player_wheel(self, event):
        if event.num == 4 or event.delta > 0:
            self.player_zoom *= 1.1
        elif event.num == 5 or event.delta < 0:
            self.player_zoom /= 1.1
        
        # 制限
        self.player_zoom = max(0.1, min(10.0, self.player_zoom))
        self.refresh_player_canvas()

    # プレビューのパンニング・ズームイベント
    def _on_preview_middle_down(self, event):
        self._preview_panning = True
        self._preview_pan_start = (event.x, event.y)
        self.preview_canvas.config(cursor="fleur")

    def _on_preview_middle_drag(self, event):
        if self._preview_panning:
            dx = event.x - self._preview_pan_start[0]
            dy = event.y - self._preview_pan_start[1]
            self.preview_pan_x += dx
            self.preview_pan_y += dy
            self._preview_pan_start = (event.x, event.y)
            # プレビューは _start_preview ループで更新されるが、即時反映させたい場合はここでも呼べる
            # ただしキャプチャ処理が重いので、変数の更新だけにして次のループに任せるか、
            # あるいは _start_preview を強制的に呼ぶか。
            # ここでは滑らかさを優先してループ任せにする (interval=100ms なので少しカクつくかも)

    def _on_preview_middle_up(self, event):
        self._preview_panning = False
        self.preview_canvas.config(cursor="")

    def _on_preview_middle_double_click(self, event):
        self.preview_pan_x = 0
        self.preview_pan_y = 0
        self.preview_zoom = 1.0
        # 即時リセット
        # self._start_preview() # may be too heavy

    def _on_preview_wheel(self, event):
        if event.num == 4 or event.delta > 0:
            self.preview_zoom *= 1.1
        elif event.num == 5 or event.delta < 0:
            self.preview_zoom /= 1.1
        
        # 制限
        self.preview_zoom = max(0.1, min(10.0, self.preview_zoom))

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
