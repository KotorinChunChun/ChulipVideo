"""画面録画機能モジュール

メインアプリケーションから呼び出される録画を行うためのウィンドウ。
"""
from __future__ import annotations

import ctypes
import ctypes.wintypes
import datetime
import os
import shutil
import subprocess
import threading
import time
import tkinter as tk
import tkinter.ttk as ttk
from utils import open_folder_with_selection
import glob
import urllib.request
import zipfile
from tkinter import filedialog, messagebox
from typing import TYPE_CHECKING, Any, Dict, List, Optional, Tuple

import cv2
import mss
import numpy as np
from PIL import Image, ImageDraw, ImageFont, ImageTk

import overlay_utils
from config import PROJECT_NAME, PROJECT_VERSION, get_base_dir, load_global_config, save_global_config
from recorder_core import ScreenRecorderLogic
from shortcut_manager import ShortcutManager
from shortcut_settings_dialog import ShortcutSettingsDialog
from ui_utils import add_tooltip
from utils import resource_path
from window_utils import WindowUtils
from multi_video_player import MultiVideoManager


# 仮想カメラライブラリのインポート試行
try:
    import pyvirtualcam
except ImportError:
    pyvirtualcam = None

class UnityCaptureInstaller:
    """Unity Captureドライバの自動インストールを行うヘルパークラス"""
    DOWNLOAD_URL = "https://github.com/schellingb/UnityCapture/archive/refs/heads/master.zip"
    
    @staticmethod
    def is_installed() -> bool:
        """ドライバがインストールされているか確認（pyvirtualcamで実際に開けるか試す）"""
        if pyvirtualcam is None:
            return False
        try:
            # 実際に unitycapture バックエンドで開けるかテスト
            # 極端に小さいサイズ (1x1など) を拒否するドライバがあるため、標準的なサイズでテスト
            with pyvirtualcam.Camera(width=640, height=480, fps=20, backend='unitycapture') as cam:
                return True
        except Exception:
            # ドライバ未導入、または他で使用中などの場合はここに来る
            return False

    @staticmethod
    def install(parent_window):
        """ダウンロードとインストール（要管理者権限）を実行"""
        try:
            # 1. ダウンロード
            target_dir = os.path.join(os.environ["TEMP"], "UnityCaptureInstall")
            if os.path.exists(target_dir):
                shutil.rmtree(target_dir)
            os.makedirs(target_dir)
            
            zip_path = os.path.join(target_dir, "UnityCapture.zip")
            
            # プログレス表示付きでダウンロード...は省略し、messageboxで通知
            messagebox.showinfo("インストールの準備", "Unity Captureドライバをダウンロードします。\nこれには数秒かかる場合があります。", parent=parent_window)
            
            urllib.request.urlretrieve(UnityCaptureInstaller.DOWNLOAD_URL, zip_path)
            
            # 2. 解凍
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(target_dir)
            
            # 3. インストールバッチの特定 (再帰的に探索)
            install_bat = None
            for root, dirs, files in os.walk(target_dir):
                if "Install.bat" in files:
                    install_bat = os.path.join(root, "Install.bat")
                    break
            
            if not install_bat or not os.path.exists(install_bat):
                raise FileNotFoundError("インストーラ(Install.bat)が見つかりませんでした。")
            
            # 4. 管理者権限で実行
            # ShellExecuteで 'runas' を指定
            messagebox.showinfo("確認", "次に表示されるユーザーアカウント制御(UAC)で「はい」を押して、\nドライバの登録を許可してください。", parent=parent_window)
            
            ret = ctypes.windll.shell32.ShellExecuteW(
                None, "runas", install_bat, None, os.path.dirname(install_bat), 1
            )
            
            if ret <= 32:
                raise RuntimeError(f"インストーラの起動に失敗しました。エラーコード: {ret}")
                
            messagebox.showinfo("インストール中", "インストーラの実行を開始しました。\n表示されたコマンドプロンプトが閉じ、処理が完了したことを確認してから OK を押してください。", parent=parent_window)
            
        except Exception as e:
            messagebox.showerror("インストールエラー", f"インストール中にエラーが発生しました:\n{e}", parent=parent_window)

if TYPE_CHECKING:
    from ChulipVideo import VideoCropperApp

class IndependentPreviewWindow(tk.Toplevel):
    """プレビュー専用の独立したフレームレスウィンドウ (OBS認識対応版)"""
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        # overrideredirect(True) を使うとOBSで認識されないため False。
        # 代わりに Windows API でスタイルを剥ぎ取る。
        self.overrideredirect(False)
        self.title(f"{PROJECT_NAME} - 録画プレビューツール")
        
        self.attributes("-topmost", True)
        self.canvas = tk.Canvas(self, bg="black", highlightthickness=0)
        self.canvas.pack(fill=tk.BOTH, expand=True)
        self.image_id = None
        
        # ドラッグ用
        self._drag_start_x = 0
        self._drag_start_y = 0
        self.canvas.bind("<Button-1>", self._on_drag_start)
        self.canvas.bind("<B1-Motion>", self._on_drag_motion)
        self.canvas.bind("<ButtonRelease-1>", self._on_drag_end)

        # ウィンドウを強制的に作成してからスタイルを剥ぎ取る
        self.update()
        self._remove_title_bar()
        
    def _remove_title_bar(self):
        """Windows APIを使用してタイトルバーや枠線を消す"""
        hwnd = ctypes.windll.user32.GetParent(self.winfo_id())
        if not hwnd: hwnd = self.winfo_id()
        
        # 基本スタイル (GWL_STYLE = -16)
        style = ctypes.windll.user32.GetWindowLongW(hwnd, -16)
        
        # WS_POPUP (0x80000000) を追加し、装飾スタイルを削除
        style &= ~0x00C00000 # WS_CAPTION
        style &= ~0x00040000 # WS_THICKFRAME
        style &= ~0x00020000 # WS_MINIMIZEBOX
        style &= ~0x00010000 # WS_MAXIMIZEBOX
        style &= ~0x00080000 # WS_SYSMENU
        style |= 0x80000000  # WS_POPUP
        
        ctypes.windll.user32.SetWindowLongW(hwnd, -16, style)
        
        # 拡張スタイル (GWL_EXSTYLE = -20) のクリーンアップ (余白・影対策)
        ex_style = ctypes.windll.user32.GetWindowLongW(hwnd, -20)
        # WS_EX_DLGMODALFRAME (0x00000001), WS_EX_WINDOWEDGE (0x00000100), 
        # WS_EX_CLIENTEDGE (0x00000200), WS_EX_STATICEDGE (0x00020000) を削除
        ex_style &= ~0x00000001
        ex_style &= ~0x00000100
        ex_style &= ~0x00000200
        ex_style &= ~0x00020000
        ctypes.windll.user32.SetWindowLongW(hwnd, -20, ex_style)
        
        # 変更を反映 (SWP_FRAMECHANGED = 0x0020, SWP_NOMOVE = 0x0002, SWP_NOSIZE = 0x0001)
        ctypes.windll.user32.SetWindowPos(hwnd, 0, 0, 0, 0, 0, 0x0004 | 0x0010 | 0x0002 | 0x0001 | 0x0020)
        
    def _on_drag_start(self, event):
        self._drag_start_x = event.x
        self._drag_start_y = event.y
        
    def _on_drag_motion(self, event):
        x = self.winfo_x() + (event.x - self._drag_start_x)
        y = self.winfo_y() + (event.y - self._drag_start_y)
        
        # geometry() を使うとOSの再描画と競合してチラつきや一瞬の黒枠が発生するため、
        # SetWindowPos API で位置のみを直接変更する。
        # SWP_NOSIZE(0x0001) | SWP_NOZORDER(0x0004) | SWP_NOACTIVATE(0x0010)
        hwnd = ctypes.windll.user32.GetParent(self.winfo_id())
        if not hwnd: hwnd = self.winfo_id()
        ctypes.windll.user32.SetWindowPos(hwnd, 0, x, y, 0, 0, 0x0001 | 0x0004 | 0x0010)

    def _on_drag_end(self, event):
        # 移動終了時に座標をアプリ側の変数に保存
        self.app.last_independent_preview_pos = f"+{self.winfo_x()}+{self.winfo_y()}"

    def update_image(self, img: Image.Image):
        """画像を1:1で表示。サイズが異なればウィンドウサイズを調整"""
        w, h = img.size
        curr_x = self.winfo_x()
        curr_y = self.winfo_y()
        
        if self.winfo_width() != w or self.winfo_height() != h:
            # 座標を維持したままサイズを更新
            # TkinterのgeometryではなくSetWindowPosで直接叩く（枠誤認回避）
            hwnd = ctypes.windll.user32.GetParent(self.winfo_id())
            if not hwnd: hwnd = self.winfo_id()
            # SWP_NOZORDER(4) | SWP_NOACTIVATE(10)
            ctypes.windll.user32.SetWindowPos(hwnd, 0, curr_x, curr_y, w, h, 0x0004 | 0x0010)
            # Tkinter側にも把握させる
            self.update_idletasks()
            
        self.tk_img = ImageTk.PhotoImage(img)
        if self.image_id is None:
            self.image_id = self.canvas.create_image(0, 0, anchor=tk.NW, image=self.tk_img)
        else:
            self.canvas.itemconfig(self.image_id, image=self.tk_img)
        
        # 確実に表示されるように前面へ
        self.lift()
        self.attributes("-topmost", True) # 念のため再度セット

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
    
    # 録画枠色 (外側 -> 中間 -> 内側) の3色グラデーション定義
    # 太さ比率は 1:3:1 (計5px)
    BORDER_COLORS = {
        "READY": ["#699483", "#97dc44", "#699483"], 
        "RECORDING": ["#aa0000", "#ff3030", "#699483"]
    }
    
    def __init__(self, root: Optional[tk.Tk] = None, parent_app: Optional[VideoCropperApp] = None):
        self.window_utils = WindowUtils()
        self.recorder_logic = ScreenRecorderLogic(self.window_utils)

        if root is None:
            self.root = tk.Tk()
            self.standalone = True
        else:
            self.root = tk.Toplevel(root)
            self.standalone = False
            
        self.root.title(f"{PROJECT_NAME} - 録画ツール")
        self.parent_app = parent_app
        
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

        # --- メンバ変数初期化 ---
        self._init_variables()
        
        # --- 設定ロード ---
        self._load_config()

        # --- UI構築 ---
        self._build_ui()
        
        # --- 初期化処理 ---
        self.update_source_list()
        self._check_vcam_driver_and_update_ui()
        self._start_preview()
        self.load_window_geometry()

        # --- イベントバインド ---
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_changed)
        
        # 初期表示の追従開始
        if self.notebook.select() == str(self.tab_record):
            self._update_region_tracking()

    def _init_variables(self):
        """メンバ変数の初期化"""
        self.window_utils = WindowUtils()
        self.recorder_logic = ScreenRecorderLogic(self.window_utils)

        # ショートカット管理
        self.shortcut_manager = ShortcutManager(os.path.join(get_base_dir(), "shortcuts.tsv"))

        # 状態変数
        self.preview_active = True
        self.is_counting_down = False
        self.countdown_after_id = None
        self.countdown_window = None
        
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
        self.fps_var = tk.IntVar(value=30)
        self.quality_var = tk.StringVar(value="最高")
        self.save_path_var = tk.StringVar()
        # self.record_cursor_var = tk.BooleanVar(value=False) # 機能削除
        self.show_region_var = tk.BooleanVar(value=True)
        self.exclusive_window_var = tk.BooleanVar(value=True)
        self.record_tsv_var = tk.BooleanVar(value=True)
        self.seek_var = tk.DoubleVar()
        self.show_trajectory_var = tk.BooleanVar(value=True)
        self.player_fit_var = tk.BooleanVar(value=True)
        self.preview_fit_var = tk.BooleanVar(value=True)
        self.show_preview_overlay_var = tk.BooleanVar(value=True)
        self.show_vcam_var = tk.BooleanVar(value=False)
        self.sync_geo_var = tk.BooleanVar(value=True) # 座標同期フラグ
        self.max_geo_var = tk.BooleanVar(value=False) # 最大化フラグ
        self.preview_independent_var = tk.BooleanVar(value=False) # 独立プレビューフラグ
        self.last_independent_preview_pos = "+0+0" # セッション内の表示座標
        
        # タイマー設定
        self.countdown_var = tk.IntVar(value=3)
        self.max_duration_var = tk.IntVar(value=0)
        self.recording_start_time = 0.0
        
        # プレイヤー/プレビューのパン・ズーム用
        self.player_zoom = 1.0
        self.player_pan_x = 0
        self.player_pan_y = 0
        self._player_panning = False
        self._player_pan_start = (0, 0)
        
        self.preview_zoom = 1.0
        self.preview_pan_x = 0
        self.preview_pan_y = 0
        self._preview_panning = False
        self._preview_pan_start = (0, 0)

        # プレビュー演出用
        self.preview_input_manager = overlay_utils.InputHistoryManager()
        self.last_mouse_state = "None"
        self.last_release_info = (0.0, "")

        # その他参照用
        self.vcam_instance = None
        self.widgets_to_lock: List[tk.Widget] = []
        self.region_window: Optional[tk.Toplevel] = None
        self.independent_preview_window: Optional[IndependentPreviewWindow] = None
        self.monitors: List[Dict[str, Any]] = []
        self.windows: List[Tuple[Any, str, str, int]] = []
        self.file_items: List[str] = []
        
        # マルチビデオ管理
        self.multi_video_manager: Optional[MultiVideoManager] = None
        self.is_multi_view = False


    def _load_config(self):
        """設定読み込み"""
        self.global_config = load_global_config()
        self.theme = self.global_config.get("theme", {})

        # 保存先ディレクトリ
        self.save_dir = ""
        if self.parent_app:
            self.save_dir = self._load_save_dir_from_config()
        if not self.save_dir:
            self.save_dir = os.path.join(os.path.expanduser("~"), "Videos", PROJECT_NAME)
        if not os.path.exists(self.save_dir):
            os.makedirs(self.save_dir, exist_ok=True)
        self.save_path_var.set(self.save_dir)

        # ショートカット管理
        self.shortcut_manager = ShortcutManager(os.path.join(get_base_dir(), "shortcuts.tsv"))

        # 設定値の復元
        saved_source = self.global_config.get("recorder_source", "desktop")
        if saved_source in ["desktop", "window"]:
            self.source_var.set(saved_source)
        
        self.fps_var.set(self.global_config.get("recorder_fps", 30))
        self.quality_var.set(self.global_config.get("recorder_quality", "最高"))
        self.show_region_var.set(self.global_config.get("recorder_show_region", True))
        self.show_preview_overlay_var.set(self.global_config.get("recorder_show_preview_overlay", True))
        self.preview_independent_var.set(self.global_config.get("recorder_preview_independent", False))

        # 設定変更時の自動保存トレース
        self.fps_var.trace_add("write", lambda *args: self.save_window_geometry())
        self.quality_var.trace_add("write", lambda *args: self.save_window_geometry())
        self.show_region_var.trace_add("write", lambda *args: self.save_window_geometry())
        self.show_preview_overlay_var.trace_add("write", lambda *args: self.save_window_geometry())
        self.preview_independent_var.trace_add("write", lambda *args: self.save_window_geometry())

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
        config["recorder_show_preview_overlay"] = self.show_preview_overlay_var.get()
        config["recorder_preview_independent"] = self.preview_independent_var.get()
        config["recorder_fps"] = self.fps_var.get()
        config["recorder_quality"] = self.quality_var.get()
        
        # ターゲット情報の保存
        self._save_recording_target_to_config(config)
        
        try:
            config["recorder_sash_position"] = self.main_paned.sash_coord(0)[1]
        except:
            pass
            
        save_global_config(config)

    def on_close(self):
        """終了時の処理."""
        if self.vcam_instance:
            try:
                self._send_vcam_no_signal()
                self.vcam_instance.close()
            except:
                pass
            self.vcam_instance = None
        
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
        self.btn_browse = tk.Button(path_frame, text="...", command=self.browse_save_dir, width=3, bg=self.theme.get("recorder_browse_bg"))
        self.btn_browse.pack(side=tk.LEFT, padx=5)

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

        # フォルダを開くボタン
        self.btn_open_folder = tk.Button(list_ctrl, text="保存フォルダを開く", command=lambda: open_folder_with_selection(self.save_path_var.get()))
        self.btn_open_folder.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)
        
        self.btn_open_tsv = tk.Button(list_ctrl, text="操作ログを開く", command=self.open_tsv_file, state=tk.DISABLED, height=1, bg=self.theme.get("recorder_open_tsv_bg"))
        self.btn_open_tsv.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)
        
        self.widgets_to_lock.extend([self.btn_rename, self.btn_delete, self.btn_open_folder, self.btn_open_tsv])
        self._fix_all_button_active_colors(path_frame)
        self._fix_all_button_active_colors(list_ctrl)

        # キーバインド
        self.file_listbox.bind("<F2>", lambda e: self.rename_file())
        self.file_listbox.bind("<Delete>", lambda e: self.delete_file())
        self.file_listbox.bind("<Return>", lambda e: self.on_file_double_click(e))
        self.file_listbox.bind("<Control-a>", lambda e: self._select_all_files())
        self.file_listbox.bind("<Control-A>", lambda e: self._select_all_files())

        # フッターアクション
        footer_frame = tk.Frame(parent)
        footer_frame.pack(fill=tk.X, padx=10, pady=10)
        
        self.btn_close = tk.Button(footer_frame, text="閉じる", command=self.on_close, width=10, bg=self.theme.get("recorder_close_bg"))
        self.btn_close.pack(side=tk.RIGHT, padx=5)
        self.btn_edit = tk.Button(footer_frame, text="閉じて編集", command=self.close_and_edit, bg=self.COLOR_BTN_EDIT, width=12)
        self.btn_edit.pack(side=tk.RIGHT, padx=5)
        self.btn_key_config = tk.Button(footer_frame, text="キー設定", command=self.open_shortcut_settings, height=1, bg=self.theme.get("recorder_key_config_bg"))
        self.btn_key_config.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)
        self.widgets_to_lock.extend([self.btn_key_config, self.btn_close, self.btn_edit])
        self._fix_all_button_active_colors(footer_frame)

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
        # 1. ソース選択 (録画対象フレーム)
        source_frame = tk.LabelFrame(self.tab_record, text="録画対象")
        source_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # 1-1. デスクトップ/ウィンドウ切り替え
        type_frame = tk.Frame(source_frame)
        type_frame.pack(fill=tk.X, padx=5, pady=2)
        
        self.radio_desktop = tk.Radiobutton(type_frame, text="デスクトップ", variable=self.source_var, 
                       value="desktop", command=self.update_source_list)
        self.radio_desktop.pack(side=tk.LEFT, padx=10)
        self.radio_window = tk.Radiobutton(type_frame, text="ウィンドウ", variable=self.source_var, 
                       value="window", command=self.update_source_list)
        self.radio_window.pack(side=tk.LEFT, padx=10)
        self.widgets_to_lock.extend([self.radio_desktop, self.radio_window])
        
        # 1-2. フィルター
        filter_frame = tk.Frame(source_frame)
        filter_frame.pack(fill=tk.X, padx=5, pady=2)
        tk.Label(filter_frame, text="検索:").pack(side=tk.LEFT)
        self.filter_var.trace_add("write", lambda *args: self.update_source_list())
        self.entry_filter = tk.Entry(filter_frame, textvariable=self.filter_var)
        self.entry_filter.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.widgets_to_lock.append(self.entry_filter)

        # 1-3. ターゲット選択
        target_frame = tk.Frame(source_frame)
        target_frame.pack(fill=tk.X, padx=5, pady=2)
        tk.Label(target_frame, text="対象:").pack(side=tk.LEFT)
        self.combo_target = ttk.Combobox(target_frame, textvariable=self.target_var, state="readonly")
        self.combo_target.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.combo_target.bind("<<ComboboxSelected>>", self.on_target_changed)
        self.btn_update = tk.Button(target_frame, text="更新", command=self.update_source_list, width=4, bg=self.COLOR_BTN_UPDATE)
        self.btn_update.pack(side=tk.LEFT)
        self.widgets_to_lock.extend([self.btn_update, self.combo_target])

        # 1-4. 座標・サイズ入力
        geo_frame = tk.Frame(source_frame)
        geo_frame.pack(fill=tk.X, padx=5, pady=2)
        
        self.geo_x = tk.IntVar()
        self.geo_y = tk.IntVar()
        self.geo_w = tk.IntVar()
        self.geo_h = tk.IntVar()
        
        # リアルタイム反映用のトレース設定
        for var in [self.geo_x, self.geo_y, self.geo_w, self.geo_h]:
            var.trace_add("write", self._on_geo_var_changed)
        
        # 最大化・同期チェックボックス
        self.check_max_geo = tk.Checkbutton(geo_frame, text="最大化", variable=self.max_geo_var, 
                                            command=self._on_max_changed, indicatoron=False, 
                                            selectcolor="", relief=tk.RAISED, overrelief=tk.RIDGE)
        self.check_max_geo.pack(side=tk.LEFT, padx=(0, 2), ipadx=5)
        self.widgets_to_lock.append(self.check_max_geo)

        self.check_sync_geo = tk.Checkbutton(geo_frame, text="同期", variable=self.sync_geo_var, command=self._on_geo_ctrl_changed)
        self.check_sync_geo.pack(side=tk.LEFT, padx=(0, 5))
        self.widgets_to_lock.append(self.check_sync_geo)

        tk.Label(geo_frame, text="座標 X:").pack(side=tk.LEFT)
        self.spin_x = tk.Spinbox(geo_frame, from_=-10000, to=10000, textvariable=self.geo_x, width=5)
        self.spin_x.pack(side=tk.LEFT, padx=2)
        
        tk.Label(geo_frame, text="Y:").pack(side=tk.LEFT)
        self.spin_y = tk.Spinbox(geo_frame, from_=-10000, to=10000, textvariable=self.geo_y, width=5)
        self.spin_y.pack(side=tk.LEFT, padx=2)
        
        tk.Label(geo_frame, text="W:").pack(side=tk.LEFT)
        self.spin_w = tk.Spinbox(geo_frame, from_=0, to=10000, textvariable=self.geo_w, width=5)
        self.spin_w.pack(side=tk.LEFT, padx=2)
        
        tk.Label(geo_frame, text="H:").pack(side=tk.LEFT)
        self.spin_h = tk.Spinbox(geo_frame, from_=0, to=10000, textvariable=self.geo_h, width=5)
        self.spin_h.pack(side=tk.LEFT, padx=2)
        
        self.btn_apply_geo = tk.Button(geo_frame, text="適用", command=self.apply_window_geometry, width=4, bg=self.COLOR_BTN_UPDATE)
        self.btn_apply_geo.pack(side=tk.LEFT, padx=5)
        
        for spin in [self.spin_x, self.spin_y, self.spin_w, self.spin_h]:
            spin.bind("<Return>", lambda e: self.apply_window_geometry())
            spin.bind("<MouseWheel>", self._on_geo_spin_wheel)
            self.widgets_to_lock.append(spin)
        self.widgets_to_lock.append(self.btn_apply_geo)
        
        # 初期状態反映
        self._on_sync_changed()

        
        # 4. プレビュー
        preview_label_frame = tk.LabelFrame(self.tab_record, text="プレビュー")
        preview_label_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        self.preview_canvas = tk.Canvas(preview_label_frame, bg=self.COLOR_CANVAS_BG, highlightthickness=0)
        self.preview_canvas = tk.Canvas(preview_label_frame, bg=self.COLOR_CANVAS_BG, highlightthickness=0)
        self.preview_canvas.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.preview_image_id = None
        self.preview_canvas.bind("<Configure>", lambda e: self._on_canvas_resize("preview"))
        # 共通のハンドラを使用（モード指定）
        self.preview_canvas.bind("<ButtonPress-2>", lambda e: self._on_canvas_middle_down(e, "preview"))
        self.preview_canvas.bind("<B2-Motion>", lambda e: self._on_canvas_middle_drag(e, "preview"))
        self.preview_canvas.bind("<ButtonRelease-2>", lambda e: self._on_canvas_middle_up(e, "preview"))
        self.preview_canvas.bind("<Double-Button-2>", lambda e: self._on_canvas_middle_double_click(e, "preview"))
        self.preview_canvas.bind("<MouseWheel>", lambda e: self._on_canvas_wheel(e, "preview"))
        self.preview_canvas.bind("<Button-4>", lambda e: self._on_canvas_wheel(e, "preview"))
        self.preview_canvas.bind("<Button-5>", lambda e: self._on_canvas_wheel(e, "preview"))

        # 5. 録画設定
        settings_frame = tk.Frame(self.tab_record)
        settings_frame.pack(fill=tk.X, padx=10, pady=5)
        
        tk.Label(settings_frame, text="フォーマット:").pack(side=tk.LEFT)
        self.format_var = tk.StringVar(value="mp4")
        self.combo_format = ttk.Combobox(settings_frame, textvariable=self.format_var, values=["mp4", "mkv"], width=5, state="readonly")
        self.combo_format.pack(side=tk.LEFT, padx=5)

        tk.Label(settings_frame, text="FPS:").pack(side=tk.LEFT, padx=(10, 0))
        self.combo_fps = ttk.Combobox(settings_frame, textvariable=self.fps_var, values=[15, 30, 60], width=5, state="readonly")
        self.combo_fps.pack(side=tk.LEFT, padx=5)
        
        tk.Label(settings_frame, text="画質:").pack(side=tk.LEFT, padx=(10, 0))
        self.combo_quality = ttk.Combobox(settings_frame, textvariable=self.quality_var, values=["最高", "高", "中", "低"], width=5, state="readonly")
        self.combo_quality.pack(side=tk.LEFT, padx=5)
        self.widgets_to_lock.extend([self.combo_format, self.combo_fps, self.combo_quality])

        # 5.5 タイマー設定
        timer_frame = tk.Frame(self.tab_record)
        timer_frame.pack(fill=tk.X, padx=10, pady=5)

        tk.Label(timer_frame, text="開始カウントダウン(秒):").pack(side=tk.LEFT)
        self.spin_countdown = tk.Spinbox(timer_frame, from_=0, to=60, textvariable=self.countdown_var, width=5)
        self.spin_countdown.pack(side=tk.LEFT, padx=5)
        
        tk.Label(timer_frame, text="最長録画時間(秒):").pack(side=tk.LEFT, padx=(10, 0))
        self.spin_max_duration = tk.Spinbox(timer_frame, from_=0, to=36000, textvariable=self.max_duration_var, width=6)
        self.spin_max_duration.pack(side=tk.LEFT, padx=5)
        tk.Label(timer_frame, text="(0で無制限)").pack(side=tk.LEFT)
        
        self.lbl_timer_info = tk.Label(timer_frame, text="", fg="blue")
        self.lbl_timer_info.pack(side=tk.LEFT, padx=10)
        
        self.widgets_to_lock.extend([self.spin_countdown, self.spin_max_duration])

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
        
        self.check_preview_overlay = tk.Checkbutton(options_frame, text="プレビューに演出を表示", variable=self.show_preview_overlay_var)
        self.check_preview_overlay.pack(side=tk.LEFT, padx=5)
        self.widgets_to_lock.append(self.check_preview_overlay)

        self.check_preview_independent = tk.Checkbutton(options_frame, text="プレビューウィンドウを独立させる", 
                                                        variable=self.preview_independent_var,
                                                        indicatoron=False, selectcolor="", 
                                                        relief=tk.RAISED, overrelief=tk.RIDGE)
        self.check_preview_independent.pack(side=tk.LEFT, padx=5, ipadx=5)
        self.widgets_to_lock.append(self.check_preview_independent)

        # 7. 録画ボタンエリア
        btn_frame = tk.Frame(self.tab_record)
        btn_frame.pack(fill=tk.X, padx=30, pady=10)
        
        self.btn_record = tk.Button(btn_frame, text="● 録画開始", bg="#ffcccc", font=("Arial", 12, "bold"),
                                    command=self.toggle_recording)
        self.btn_record.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.btn_vcam_toggle = tk.Button(btn_frame, text="● 仮想カメラ配信", bg="#ffcccc", font=("Arial", 12, "bold"),
                                         command=self._on_vcam_toggle_click)
        self.btn_vcam_toggle.pack(side=tk.LEFT, padx=(30, 0), fill=tk.Y)

        # 導入ボタン
        self.btn_install_vcam = tk.Button(btn_frame, text="導入", command=self._install_drive, font=("Arial", 10), bg=self.theme.get("recorder_install_vcam_bg"))
        self.btn_install_vcam.pack(side=tk.LEFT, padx=(10, 0), fill=tk.Y)

        # 全体の配色微調整
        self._fix_all_button_active_colors(self.tab_record)

    def _setup_playback_tab(self):
        # Controls at bottom (Fixed Height)
        p_btns_container = tk.Frame(self.tab_play, height=80) 
        p_btns_container.pack_propagate(False) # 固定高さ
        p_btns_container.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=(0, 10))

        # Player frame takes remaining space
        player_frame = tk.LabelFrame(self.tab_play, text="プレイヤー")
        player_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # プレイヤー表示エリア (Single / Multi 切り替え用)
        # Single View Canvas (Child of player_frame)
        self.player_canvas = tk.Canvas(player_frame, bg=self.COLOR_CANVAS_BG, highlightthickness=0)
        self.player_canvas.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.player_canvas.bind("<Configure>", lambda e: self._on_canvas_resize("player"))
        
        # Multi View Container (Child of player_frame, initially hidden)
        self.player_multi_container = tk.Frame(player_frame, bg=self.COLOR_CANVAS_BG)
        # Don't pack initially

        p_btns = tk.Frame(p_btns_container)
        p_btns.pack(fill=tk.X, pady=2)
        
        self.btn_play = tk.Button(p_btns, text="▶", command=self.toggle_playback, state=tk.DISABLED, width=5, bg=self.theme.get("recorder_play_bg"))
        self.btn_play.pack(side=tk.LEFT, padx=5)
        
        self.slider = ttk.Scale(p_btns, from_=0, to=100, variable=self.seek_var, orient=tk.HORIZONTAL, command=self.on_slider_move)
        self.slider.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.slider.bind("<Button-1>", self.on_slider_press)
        self.slider.bind("<ButtonRelease-1>", self.on_slider_release)

        self.lbl_time = tk.Label(p_btns, text="00:00 / 00:00")
        self.lbl_time.pack(side=tk.LEFT, padx=5)

        self.chk_player_traj = tk.Checkbutton(p_btns_container, text="マウス・キー入力の演出を表示", variable=self.show_trajectory_var, command=self.refresh_player_canvas)
        self.chk_player_traj.pack(side=tk.LEFT, padx=5)

        self.chk_player_fit = tk.Checkbutton(p_btns_container, text="プレビューを拡大・縮小する", variable=self.player_fit_var, command=self.refresh_player_canvas)
        self.chk_player_fit.pack(side=tk.RIGHT, padx=5)

        # 全体の配色微調整
        self._fix_all_button_active_colors(self.tab_play)

        # プレイヤーキャンバスへのバインド (パン・ズーム)
        self.player_canvas.bind("<ButtonPress-2>", lambda e: self._on_canvas_middle_down(e, "player"))
        self.player_canvas.bind("<B2-Motion>", lambda e: self._on_canvas_middle_drag(e, "player"))
        self.player_canvas.bind("<ButtonRelease-2>", lambda e: self._on_canvas_middle_up(e, "player"))
        self.player_canvas.bind("<Double-Button-2>", lambda e: self._on_canvas_middle_double_click(e, "player"))
        self.player_canvas.bind("<MouseWheel>", lambda e: self._on_canvas_wheel(e, "player"))
        self.player_canvas.bind("<Button-4>", lambda e: self._on_canvas_wheel(e, "player"))
        self.player_canvas.bind("<Button-5>", lambda e: self._on_canvas_wheel(e, "player"))

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
            self.check_max_geo.config(state=tk.DISABLED)
            self.max_geo_var.set(False)
            for w in [self.spin_x, self.spin_y, self.spin_w, self.spin_h, self.btn_apply_geo, self.check_sync_geo]:
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
            self.check_max_geo.config(state=tk.NORMAL)
            for w in [self.spin_x, self.spin_y, self.spin_w, self.spin_h, self.check_sync_geo]:
                w.config(state=tk.NORMAL)
            self._on_sync_changed()

        self.on_target_changed(None)

    def _on_max_changed(self):
        """最大化チェックボタンがクリックされた時の処理"""
        if self.source_var.get() == 'window':
            idx = self.combo_target.current()
            if idx >= 0 and idx < len(self.windows):
                hwnd = self.windows[idx][0]
                is_max = self.max_geo_var.get()
                if is_max:
                    self.window_utils.set_window_maximized(hwnd, True)
                else:
                    # OFFにしたとき、現在の(最大化時の)座標を取得して保持
                    rect = self.window_utils.get_window_rect(hwnd)
                    # 元に戻す
                    self.window_utils.set_window_maximized(hwnd, False)
                    # 保持していた座標を適用することでサイズを維持
                    if rect:
                        self.window_utils.set_window_position(hwnd, rect['left'], rect['top'], rect['width'], rect['height'])
        
        # 見た目の更新 (トグルボタンらしく凹凸を変える)
        if self.max_geo_var.get():
            self.check_max_geo.config(relief=tk.SUNKEN)
        else:
            self.check_max_geo.config(relief=tk.RAISED)
            
        self._on_geo_ctrl_changed()

    def _on_geo_ctrl_changed(self):
        """最大化または同期チェックボックスの状態変更時のUI更新処理"""
        is_maximized = self.max_geo_var.get()
        is_sync = self.sync_geo_var.get()
        
        # 最大化ONの時はすべて無効化
        if is_maximized:
            self.check_sync_geo.config(state=tk.DISABLED)
            for w in [self.spin_x, self.spin_y, self.spin_w, self.spin_h, self.btn_apply_geo]:
                w.config(state=tk.DISABLED)
            # 最大化ONになった瞬間に一度反映させる
            self.on_target_changed(None)
        else:
            self.check_sync_geo.config(state=tk.NORMAL)
            # 同期ONの時は数値を無効化せず、適用ボタンのみ無効化（ホイールでの即時反映は維持）
            # ただしユーザーの要望「最大化がONのときは...XYWHもすべて無効化して変更できなくする」
            # 同期ONのときは無効化するかどうか？ 前回の実装ではSpinboxは有効でした。
            # 要望に合わせて、同期ONのときもSpinboxは有効なまま（ホイール操作用）にします。
            for w in [self.spin_x, self.spin_y, self.spin_w, self.spin_h]:
                w.config(state=tk.NORMAL)
            
            if is_sync:
                self.btn_apply_geo.config(state=tk.DISABLED)
                self.on_target_changed(None)
            else:
                self.btn_apply_geo.config(state=tk.NORMAL)

    def _on_sync_changed(self):
        """(互換用・削除予定) 同期チェックボックスの状態変更時の処理"""
        self._on_geo_ctrl_changed()

    def _on_geo_spin_wheel(self, event):
        """Spinbox上でのマウスホイール操作"""
        delta = 1 if event.delta > 0 else -1
        # Linux/macOS対応 (念のため)
        if event.num == 4: delta = 1
        if event.num == 5: delta = -1
        
        widget: tk.Spinbox = event.widget
        try:
            curr = int(widget.get())
            widget.delete(0, tk.END)
            widget.insert(0, str(curr + delta))
            # self.geo_x 等の変数が更新されるため、traceにより apply_window_geometry が呼ばれる
        except:
            pass
        return "break"

    def _on_geo_var_changed(self, *args):
        """座標変数が変更された時の処理"""
        # 同期ONの時だけ即座に反映
        if self.sync_geo_var.get() and not self.max_geo_var.get():
            self.apply_window_geometry()

    def on_target_changed(self, event):
        if not self.sync_geo_var.get():
            return
            
        rect = self._get_target_rect()
        if rect:
            self.geo_x.set(rect['left'])
            self.geo_y.set(rect['top'])
            self.geo_w.set(rect['width'])
            self.geo_h.set(rect['height'])
            # 赤枠の表示・更新は _update_region_tracking 内の安定性ロジックに任せる

    def apply_window_geometry(self):
        """入力ボックスの値でウィンドウを移動・リサイズ"""
        if self.source_var.get() == 'window' and not self.max_geo_var.get():
            idx = self.combo_target.current()
            if idx >= 0 and idx < len(self.windows):
                hwnd = self.windows[idx][0]
                try:
                    x = self.geo_x.get()
                    y = self.geo_y.get()
                    w = self.geo_w.get()
                    h = self.geo_h.get()
                    self.window_utils.set_window_position(hwnd, x, y, w, h)
                except tk.TclError:
                    # 入力欄が空などの場合は無視
                    pass
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
                        # --- プレビュー演出 (オーバーレイ) の描画 (リサイズ前の元画像に行う) ---
                        if self.show_preview_overlay_var.get():
                            # ターゲット情報取得
                            target_hwnd = None
                            is_wgc = False
                            if self.source_var.get() == 'window':
                                idx = self.combo_target.current()
                                if idx >= 0 and idx < len(self.windows):
                                    target_hwnd = self.windows[idx][0]
                                if self.exclusive_window_var.get():
                                    is_wgc = True

                            # 現在の入力状態取得 (フィルタリング済み)
                            # is_wgc は True/False (get_filtered_input_state 内では Truthy チェックのみに使用されるため bool で可)
                            click_info, keys_info, rel_x, rel_y = self.recorder_logic.get_filtered_input_state(rect, target_hwnd, is_wgc)
                            keys_str = ",".join(keys_info) if keys_info else "None"
                            now_ts = time.time()
                            
                            # 1. 履歴更新
                            display_text = overlay_utils.get_input_display_text(click_info, keys_str)
                            self.preview_input_manager.update(now_ts, display_text)
                            
                            # 2. 波紋トリガー判定 (離された瞬間のみ)
                            ripple_type = ""
                            ripple_age = 0.0
                            for char, name in [("L", "left"), ("R", "right"), ("M", "middle")]:
                                if char in self.last_mouse_state and char not in click_info:
                                    # 離された
                                    self.last_release_info = (now_ts, name)
                                    break
                            
                            # 波紋描画 (離されてから一定時間内のみ)
                            if self.last_release_info[1]:
                                ripple_age = now_ts - self.last_release_info[0]
                                ripple_type = self.last_release_info[1]
                                # 波紋の持続時間を確認
                                ripple_duration = self.theme.get("mouse_overlay", {}).get(f"click_{ripple_type}", {}).get("ripple_duration", 0.5)
                                if ripple_age > ripple_duration:
                                    # 波紋終了
                                    self.last_release_info = (0.0, "")
                                    ripple_type = ""
                                    ripple_age = 0.0
                            
                            self.last_mouse_state = click_info

                            # --- 演出（オーバーレイ）の描画 ---
                            # WGCかつウィンドウ録画時、アクティブウィンドウが対象と異なれば完全に非表示（ホバーなども消す）
                            should_draw_overlay = True
                            if is_wgc and target_hwnd:
                                fg_hwnd = ctypes.windll.user32.GetForegroundWindow()
                                if fg_hwnd != target_hwnd:
                                    should_draw_overlay = False

                            if should_draw_overlay:
                                # マウスオーバーレイ描画
                                overlay_utils.draw_mouse_overlay(
                                    img, rel_x, rel_y, click_info,
                                    1.0, 1.0,
                                    self.theme,
                                    ripple_age=ripple_age,
                                    ripple_type=ripple_type
                                )
                                
                                # キー履歴オーバーレイ描画
                                active_inputs = self.preview_input_manager.get_active_inputs(
                                    now_ts, self.theme.get("input_overlay", {}).get("fade_duration", 1.0)
                                )
                                overlay_utils.draw_input_overlay(
                                    img, active_inputs,
                                    1.0, 1.0,
                                    self.theme
                                )

                        # --- 表示処理 (独立ウィンドウ or 本体Canvas) ---
                        try:
                            if self.preview_independent_var.get():
                                # ボタンの凹凸を制御
                                self.check_preview_independent.config(relief=tk.SUNKEN)

                                # 1. 独立ウィンドウの管理
                                if not self.independent_preview_window or not self.independent_preview_window.winfo_exists():
                                    self.independent_preview_window = IndependentPreviewWindow(self.root, self)
                                    # 前回の表示座標を使用
                                    self.independent_preview_window.geometry(self.last_independent_preview_pos)
                                
                                # 2. 独立ウィンドウを更新 (1:1描画)
                                self.independent_preview_window.update_image(img)
                                
                                # 3. 本体Canvasをクリア & 通知テキスト表示
                                self.preview_canvas.delete("all")
                                self.preview_canvas.create_text(
                                    self.preview_canvas.winfo_width() // 2,
                                    self.preview_canvas.winfo_height() // 2,
                                    text="別のウィンドウでプレビュー中です",
                                    fill="white", font=("Arial", 12)
                                )
                                self.preview_image_id = None
                            else:
                                # ボタンの凹凸を制御
                                self.check_preview_independent.config(relief=tk.RAISED)

                                # 独立ウィンドウが開いていれば閉じる
                                if self.independent_preview_window and self.independent_preview_window.winfo_exists():
                                    self.independent_preview_window.destroy()
                                    self.independent_preview_window = None

                                # --- Canvas への表示処理 (従来通り) ---
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
                                        
                                        # 中央配置
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
                        except Exception as disp_e:
                            # 表示処理のエラーはログに出して握りつぶし、配信を止めない
                            print(f"Preview Display Error: {disp_e}")
                            
                        # --- 仮想カメラへの送信 (分岐の外に出して常時実行) ---
                        try:
                            if self.show_vcam_var.get() and pyvirtualcam:
                                try:
                                    # 元画像(演出込み)をRGBに変換して送信
                                    # pyvirtualcam は numpy 配列 (H, W, 3) を期待
                                    frame = np.array(img.convert("RGB"))
                                    h, w = frame.shape[:2]
                                    
                                    # 初回またはサイズ変更時にインスタンス作成
                                    if self.vcam_instance is None or \
                                       self.vcam_instance.width != w or \
                                       self.vcam_instance.height != h:
                                        if self.vcam_instance:
                                            self.vcam_instance.close()
                                        
                                        # Unity Capture を優先的に使用
                                        self.vcam_instance = pyvirtualcam.Camera(width=w, height=h, fps=20, backend='unitycapture')
                                    
                                    self.vcam_instance.send(frame)
                                    self.vcam_instance.sleep_until_next_frame()
                                    
                                except Exception as v_e:
                                    # エラーが出たらOFFにする
                                    self.show_vcam_var.set(False)
                                    # ユーザーに通知（ループ内なので1回だけ出す制御が必要だが、ここでは簡易的にコンソール出力＆ボタン強調などに留める）
                                    # print(f"Virtual Camera Error: {v_e}")
                                    
                                    # エラー内容を見てドライバ未導入っぽければ案内を出す (UIスレッドで安全に)
                                    # "UnityCapture" という文字列が含まれていればバックエンドエラーの可能性大
                                    err_msg = str(v_e)
                                    self.root.after(0, lambda: self._handle_vcam_error(err_msg))

                            elif self.vcam_instance:
                                # OFFにされたら「No Signal」を送ってからクローズ
                                try:
                                    self._send_vcam_no_signal()
                                    self.vcam_instance.close()
                                except: pass
                                self.vcam_instance = None
                        except Exception as vcam_outer_e:
                            print(f"VCam Logic Error: {vcam_outer_e}")
                except Exception as e:
                    pass

        # プレビュー更新頻度の調整 (録画中は負荷軽減のため低頻度に)
        interval = 50 if not self.recorder_logic.is_recording else 200
        self.root.after(interval, self._start_preview)

    def toggle_recording(self):
        if self.is_counting_down:
            self.cancel_countdown()
            return
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
            
        count = self.countdown_var.get()
        if count > 0:
            self.is_counting_down = True
            # 録画開始ボタンを押した瞬間にUIをロック (ボタンはStopスタイルになる)
            self._set_controls_state(tk.DISABLED)
            
            # カウントダウン表示 -> 終了後に start_recording 実行
            self._show_countdown(count, self.start_recording)
        else:
            # カウントダウンなしで即開始
            self.start_recording()

    def _show_countdown(self, count, callback):
        """画面中央にカウントダウンを表示"""
        # ボタン表示を更新
        try:
            self.btn_record.config(text=f"中止 (残り{count}秒)")
        except:
            pass
            
        if count > 0:
            # オーバーレイウィンドウ作成（透過）
            w = tk.Toplevel(self.root)
            self.countdown_window = w
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
            def _next():
                if self.countdown_window:
                    try:
                        self.countdown_window.destroy()
                    except:
                        pass
                    self.countdown_window = None
                self._show_countdown(count - 1, callback)
                
            self.countdown_after_id = self.root.after(1000, _next)
        else:
            self.is_counting_down = False
            callback()

    def cancel_countdown(self):
        """カウントダウンを中止する"""
        self.is_counting_down = False
        if self.countdown_after_id:
            self.root.after_cancel(self.countdown_after_id)
            self.countdown_after_id = None
        if self.countdown_window:
            try:
                self.countdown_window.destroy()
            except:
                pass
            self.countdown_window = None
        
        # UIロックを解除
        self._set_controls_state(tk.NORMAL)

    def start_recording(self):
        self.is_counting_down = False
        self.recording_start_time = time.time()

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
        ext = self.format_var.get()
        filename = f"{now}.{ext}"
        filepath = os.path.join(self.save_path_var.get(), filename)
        
        fps = self.fps_var.get()

        self.btn_record.config(text="■ 録画停止", bg="#ff9999")
        
        self.recorder_logic.start_recording(
            filepath=filepath,
            rect=rect,
            fps=fps,
            hwnd=hwnd,
            record_tsv=self.record_tsv_var.get(),
            exclusive_window=(self.source_var.get() == 'window' and self.exclusive_window_var.get()),
            shortcut_manager=self.shortcut_manager
        )
        
        # 赤枠追従ループ開始
        self._update_region_tracking()

    def _update_region_tracking(self):
        """録画中または録画タブ表示中に赤枠を対象ウィンドウに追従させる"""
        # 条件: 録画中 OR 録画タブ表示中
        is_in_record_tab = (self.notebook.select() == str(self.tab_record))
        is_recording = self.recorder_logic.is_recording
        
        # show_region_var が OFF の場合は即座に消す
        should_show = self.show_region_var.get() and (is_recording or is_in_record_tab)
        
        # 自動停止チェック & 残り時間表示
        if is_recording:
            max_dur = self.max_duration_var.get()
            if max_dur > 0:
                elapsed = time.time() - self.recording_start_time
                remaining = max(0, max_dur - int(elapsed))
                
                # 終了予定時刻
                finish_time = datetime.datetime.fromtimestamp(self.recording_start_time + max_dur)
                finish_str = finish_time.strftime("%H:%M:%S")
                
                # MM:SS 形式 (必要なら HH:MM:SS)
                rem_m, rem_s = divmod(remaining, 60)
                rem_h, rem_m = divmod(rem_m, 60)
                if rem_h > 0:
                    rem_str = f"{rem_h}:{rem_m:02}:{rem_s:02}"
                else:
                    rem_str = f"{rem_m:02}:{rem_s:02}"
                
                self.lbl_timer_info.config(text=f"残り: {rem_str} (終了: {finish_str})")

                if elapsed >= max_dur:
                    self.stop_recording()
                    return
            else:
                self.lbl_timer_info.config(text="")
        else:
            self.lbl_timer_info.config(text="")
        
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
                    # 座標更新
                    self.region_window.geometry(f"{w}x{h}+{x}+{y}")
                    
                    # 録画状態に応じた色決定
                    state_key = "RECORDING" if is_recording else "READY"
                    colors = self.BORDER_COLORS[state_key]

                    canvas = self.region_window.winfo_children()[0]
                    if isinstance(canvas, tk.Canvas):
                        # サイズ変更または再描画定期的呼び出し
                        # (色が変わるタイミングがあるので常時描画更新する形にする、
                        #  ただし負荷軽減のためサイズ変更or色変更時のみがベストだが、
                        #  ここではシンプルに毎回描画呼出（Canvas要素は少ないので高速）)
                        if canvas.winfo_width() != w or canvas.winfo_height() != h:
                            canvas.config(width=w, height=h)
                        
                        self._draw_gradient_border(canvas, w, h, thickness, colors)
                    
                    # --- Zオーダー制御 ---
                    # ウィンドウキャプチャ時: ターゲットウィンドウの背面に配置
                    # デスクトップキャプチャ時: 最前面 (TopMost)
                    try:
                        self.region_window.update_idletasks()
                        # HWND取得
                        hwnd_region = self.region_window.winfo_id()
                        GA_ROOT = 2
                        hwnd_root = ctypes.windll.user32.GetAncestor(hwnd_region, GA_ROOT)
                        if hwnd_root == 0: hwnd_root = hwnd_region

                        if self.source_var.get() == 'window' and self.windows:
                            # ターゲットHWND特定
                            idx = self.combo_target.current()
                            if idx >= 0 and idx < len(self.windows):
                                target_hwnd = self.windows[idx][0]
                                
                                # TopMost解除 & ターゲットの後ろへ
                                SWP_NOSIZE = 0x0001
                                SWP_NOMOVE = 0x0002
                                SWP_NOACTIVATE = 0x0010
                                # HWND_NOTOPMOST = -2 (定数だがここではInsertAfterにターゲットを指定するので、
                                # そのターゲットがTopMostでなければ自然とTopMostでなくなるはず)
                                # ただし明示的にTopMostフラグを外す必要があるかは挙動次第。
                                # SetWindowPos(hwnd, hWndInsertAfter, x, y, cx, cy, uFlags)
                                
                                ctypes.windll.user32.SetWindowPos(
                                    hwnd_root, 
                                    target_hwnd, 
                                    0, 0, 0, 0, 
                                    SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE
                                )
                        else:
                            # 常に最前面
                            self.region_window.attributes("-topmost", True)
                            # 特定の環境ではattributesだけで効かない場合があるのでSetWindowPosも念のため呼ぶと確実だが、
                            # 既存コードで機能していたならattributesだけでOK。
                            # もしウィンドウモードから戻った場合にTopMostが外れていると困るので
                            # attributes("-topmost", True) は繰り返し呼んでも大丈夫。
                            pass

                    except Exception as e:
                        # print(f"Z-Order Logic Error: {e}")
                        pass

            # 安定していない場合は何もしない（前回の _hide で消えているはず）
            # ------------------------------
            
            # 座標・サイズのUI自動更新
            try:
                # 最大化状態の連動 (Syncに関わらず反映)
                if self.source_var.get() == 'window' and rect:
                    idx = self.combo_target.current()
                    if idx >= 0 and idx < len(self.windows):
                        hwnd = self.windows[idx][0]
                        is_maximized = self.window_utils.is_window_maximized(hwnd)
                        if self.max_geo_var.get() != is_maximized:
                            self.max_geo_var.set(is_maximized)
                            # 見た目の更新
                            if is_maximized:
                                self.check_max_geo.config(relief=tk.SUNKEN)
                            else:
                                self.check_max_geo.config(relief=tk.RAISED)
                            self._on_geo_ctrl_changed() # 状態が変わったらUI制御を更新

                # ユーザーが入力中でなく、かつ（同期ON または 最大化ON）の場合のみ数値を更新する
                is_max = self.max_geo_var.get()
                is_sync = self.sync_geo_var.get()
                
                try:
                    focused_widget = self.root.focus_get()
                except (KeyError, Exception):
                    focused_widget = None

                if is_max or is_sync:
                    input_widgets = [self.spin_x, self.spin_y, self.spin_w, self.spin_h]
                    if focused_widget not in input_widgets:
                        if rect:
                            if self.geo_x.get() != rect['left']: self.geo_x.set(rect['left'])
                            if self.geo_y.get() != rect['top']: self.geo_y.set(rect['top'])
                            if self.geo_w.get() != rect['width']: self.geo_w.set(rect['width'])
                            if self.geo_h.get() != rect['height']: self.geo_h.set(rect['height'])
            except Exception as e:
                # ウィンドウが閉じられた直後などのエラー無視
                pass

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

    def _on_vcam_toggle_click(self):
        """仮想カメラのトグルボタンがクリックされた時の処理"""
        # 現在の状態を反転
        self.show_vcam_var.set(not self.show_vcam_var.get())
        self._on_vcam_toggle()

    def _on_vcam_toggle(self):
        """仮想カメラのON/OFF状態が変更された時の共通処理"""
        self._update_vcam_button_ui()

    def _update_vcam_button_ui(self):
        """仮想カメラ配信ボタンの見た目を更新"""
        if not hasattr(self, "btn_vcam_toggle"): return
        
        if self.show_vcam_var.get():
            # 有効時: 録画停止ボタン（青系）に合わせる
            self.btn_vcam_toggle.config(text="仮想カメラ配信中", bg=self.COLOR_BTN_RECORD_STOP, relief=tk.SUNKEN)
        else:
            # 無効時: 録画開始ボタン（赤系）に合わせる
            self.btn_vcam_toggle.config(text="仮想カメラ配信", bg=self.COLOR_BTN_RECORD_START, relief=tk.RAISED)

    def _install_drive(self):
        """ドライバインストールボタン"""
        if messagebox.askyesno("ドライバのインストール", "Unity Captureドライバをダウンロードしてインストールしますか？\n(管理者権限が必要です)"):
            UnityCaptureInstaller.install(self.root)
            # インストール後に再チェックしてUIを更新
            self._check_vcam_driver_and_update_ui()

    def _check_vcam_driver_and_update_ui(self):
        """ドライバの状態を確認し、UIを更新する"""
        self.vcam_installed = UnityCaptureInstaller.is_installed()
        if self.vcam_installed:
            if hasattr(self, "btn_vcam_toggle"):
                self.btn_vcam_toggle.config(state=tk.NORMAL)
            if hasattr(self, "btn_install_vcam"):
                self.btn_install_vcam.config(state=tk.DISABLED) # インストール済みなら無効化
            self._update_vcam_button_ui()
        else:
            if hasattr(self, "btn_vcam_toggle"):
                self.btn_vcam_toggle.config(state=tk.DISABLED)
            if hasattr(self, "btn_install_vcam"):
                self.btn_install_vcam.config(state=tk.NORMAL)

    def _handle_vcam_error(self, err_msg: str):
        """仮想カメラエラー時のハンドリング"""
        if "backend" in err_msg.lower() or "unitycapture" in err_msg.lower():
            if messagebox.askyesno("仮想カメラエラー", 
                "仮想カメラの初期化に失敗しました。\nUnity Captureドライバがインストールされていない可能性があります。\n\n今すぐインストールしますか？"):
                UnityCaptureInstaller.install(self.root)
        else:
            messagebox.showerror("仮想カメラエラー", f"仮想カメラの起動に失敗しました:\n{err_msg}")

    def _send_vcam_no_signal(self):
        """仮想カメラに「No Signal」画面を送信する"""
        if not self.vcam_instance:
            return
            
        try:
            w, h = self.vcam_instance.width, self.vcam_instance.height
            # 黒背景の画像作成
            img = Image.new("RGB", (w, h), color="black")
            draw = ImageDraw.Draw(img)
            
            text = "No Signal"
            # フォント設定 (OS標準フォントを試行)
            font_size = max(20, min(w, h) // 10)
            try:
                # Windows標準のArialなどを使用
                font = ImageFont.truetype("arial.ttf", font_size)
            except:
                font = ImageFont.load_default()
            
            # テキストサイズ取得
            try:
                # Pillow 10.x+
                bbox = draw.textbbox((0, 0), text, font=font)
                tw = bbox[2] - bbox[0]
                th = bbox[3] - bbox[1]
            except:
                # 古い Pillow
                tw, th = draw.textsize(text, font=font)
                
            # 中央に描画
            draw.text(((w - tw) // 2, (h - th) // 2), text, font=font, fill="white")
            
            # 送信
            frame = np.array(img)
            self.vcam_instance.send(frame)
            # 少し待機（確実に届くように）
            time.sleep(0.1)
            
        except Exception as e:
            print(f"VCam No Signal Error: {e}")

    def _show_recording_region(self, rect):
        """録画枠ウィンドウを表示・更新."""
        # 枠作成 (初回のみ)
        if not self.region_window or not self.region_window.winfo_exists():
            self.region_window = tk.Toplevel(self.root)
            self.region_window.overrideredirect(True)
            self.region_window.attributes("-topmost", True)
            self.region_window.attributes("-alpha", 0.7)
            # 背景透明化 (Windowsのみ)
            if os.name == 'nt':
                self.region_window.attributes("-transparentcolor", "white")
                
            self.region_window.config(bg="white")
            
            # キャンバス
            canvas = tk.Canvas(self.region_window, bg="white", highlightthickness=0)
            canvas.pack(fill=tk.BOTH, expand=True)
        
        thickness = self.REGION_THICKNESS
        x = rect['left'] - thickness
        y = rect['top'] - thickness
        w = rect['width'] + thickness * 2
        h = rect['height'] + thickness * 2
        
        self.region_window.geometry(f"{w}x{h}+{x}+{y}")
        
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

        if self.recorder_logic.is_recording:
             colors = self.BORDER_COLORS["RECORDING"]
        else:
             colors = self.BORDER_COLORS["READY"]
             
        self._draw_gradient_border(canvas, w, h, thickness, colors)


    def _draw_gradient_border(self, canvas: tk.Canvas, w: int, h: int, thickness: int, colors: List[str]):
        """3色グラデーション枠を描画 (比率 1:3:1)"""
        canvas.delete("all")
        
        # thickness = 5 の場合: 1px, 3px, 1px
        # 実際には可変に対応するため比率で計算するが、今回は5px前提で整数のほうが綺麗に出るかも
        # colors = [Outer, Middle, Inner]
        
        t_outer = 1
        t_mid = 3
        t_inner = 1
        
        # もし thickness が変えられた場合用 (今回は 5固定前提)
        if thickness != 5:
            t_outer = max(1, int(thickness * 0.2))
            t_inner = max(1, int(thickness * 0.2))
            t_mid = thickness - t_outer - t_inner

        # 1. Outer (一番外側)
        # 描画順: Outer -> Middle -> Inner と重ねるか、あるいは
        # 大きい順に描画して中を塗りつぶさない（outlineのみ）方式だと
        # create_rectangle の outline は中心線基準で描画されるため、位置ずれに注意が必要。
        # ここでは「塗りつぶしなしの矩形を入れ子にする」のではなく、
        # 「枠線として面を描画する」アプローチをとるか、
        # シンプルに fill=None, outline=Color, width=T で描画するか。
        # width=T の線は、指定座標を中心に +/- T/2 の幅を持つ。
        
        # 座標計算 (矩形の中央線)
        # Outer: 幅 t_outer
        # 位置: 0 + t_outer/2
        current_offset = t_outer / 2.0
        canvas.create_rectangle(
            current_offset, current_offset, w - current_offset, h - current_offset,
            outline=colors[0], width=t_outer
        )
        
        # Middle: 幅 t_mid
        # 位置: t_outer + t_mid/2
        current_offset = t_outer + t_mid / 2.0
        canvas.create_rectangle(
            current_offset, current_offset, w - current_offset, h - current_offset,
            outline=colors[1], width=t_mid
        )
        
        # Inner: 幅 t_inner
        # 位置: t_outer + t_mid + t_inner/2
        current_offset = t_outer + t_mid + t_inner / 2.0
        canvas.create_rectangle(
            current_offset, current_offset, w - current_offset, h - current_offset,
            outline=colors[2], width=t_inner
        )

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
                if f.lower().endswith((".mp4", ".mkv")):
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
            
            # 複数選択チェック
            if len(idx) > 1:
                self.is_multi_view = True
                self._switch_to_multi_view(idx)
            else:
                self.is_multi_view = False
                self._switch_to_single_view()
                self.stop_playback()
                self.load_video_for_playback()
        else:
            self.btn_rename.config(state=tk.DISABLED)
            self.btn_delete.config(state=tk.DISABLED)
            self.btn_play.config(state=tk.DISABLED)
            self.btn_open_tsv.config(state=tk.DISABLED)
            self.stop_playback()
            self.is_multi_view = False
            self._switch_to_single_view()
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
        indices = [int(i) for i in indices]
        
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
            if self.multi_video_manager:
                self.multi_video_manager.release()
                # self.multi_video_manager = None # インスタンスは保持してもいいが、再ロード必要
                # 安全のため再生成させるフローにするか、releaseだけで十分か
                # MultiVideoManager.release() が cap.release() していればOK


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
        indices = [int(i) for i in indices]
        
        count = len(indices)
        msg = f"{count} 個のアイテムを削除しますか？" if count > 1 else f"{self.file_items[indices[0]]} を削除しますか？"
        
        if messagebox.askyesno("確認", msg, parent=self.root):
            self.stop_playback()
            if self.cap:
                self.cap.release()
                self.cap = None
            if self.multi_video_manager:
                self.multi_video_manager.release()


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

    def _switch_to_single_view(self):
        """シングルビューへの切り替え"""
        if self.multi_video_manager:
            self.multi_video_manager.release()
            self.multi_video_manager = None
            
        # Hide Multi
        self.player_multi_container.pack_forget()
        
        # Show Single
        if not self.player_canvas.winfo_ismapped():
            self.player_canvas.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
            
        # UI controls update
        self.chk_player_traj.config(state=tk.NORMAL)
        self.chk_player_fit.config(state=tk.NORMAL)

    def _switch_to_multi_view(self, indices):
        """マルチビューへの切り替え"""
        self.stop_playback()
        if self.cap:
            self.cap.release()
            self.cap = None
            
        # Hide Single
        self.player_canvas.pack_forget()
        
        # Show Multi
        self.player_multi_container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 選択されたファイルのパスリスト作成
        paths = []
        for i in indices:
            fname = self.file_items[i]
            paths.append(os.path.join(self.save_dir, fname))
            
        # MultiVideoManager初期化
        if self.multi_video_manager:
            self.multi_video_manager.release()
            
        self.multi_video_manager = MultiVideoManager(self.player_multi_container, paths)
        self.multi_video_manager.set_callback(self._on_multi_video_state_change)
        
        # グローバルコントロールの初期化
        
        # グローバルコントロールの初期化
        # Initial label
        self.lbl_time.config(text=self.multi_video_manager.get_current_time_str())
        self.seek_var.set(0)
        self.slider.config(to=100)
        
        # UI controls update
        self.chk_player_traj.config(state=tk.DISABLED) # Overlay not supported in multi yet
        self.chk_player_fit.config(state=tk.DISABLED)

    def load_video_for_playback(self):
        if self.is_multi_view: return

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
        if self.is_multi_view:
            if self.multi_video_manager:
                self.multi_video_manager.play_all()
                self.is_playing = True
                self.btn_play.config(text="Ⅱ")
                self.playback_loop()
            return

        if not self.cap or not self.cap.isOpened():
            return
        
        curr = int(self.cap.get(cv2.CAP_PROP_POS_FRAMES))
        if curr >= self.video_total_frames - 1:
            self.cap.set(cv2.CAP_PROP_POS_FRAMES, 0)
            
        self.is_playing = True
        self.btn_play.config(text="Ⅱ")
        self.playback_loop()

    def stop_playback(self):
        if self.is_multi_view:
            if self.multi_video_manager:
                self.multi_video_manager.stop_all()
        
        self.is_playing = False
        self.btn_play.config(text="▶")
        if self.playback_after_id:
            self.root.after_cancel(self.playback_after_id)
            self.playback_after_id = None

    def _on_multi_video_state_change(self, is_playing):
        """個別の動画が再生されたときにループを回す"""
        if is_playing and not self.playback_after_id:
            self.start_playback_loop_only()

    def start_playback_loop_only(self):
        """ループだけ回す (Global Play状態にはしない)"""
        if not self.playback_after_id:
            self.playback_loop()

    def playback_loop(self):
        # Multi-View Loop check first
        if self.is_multi_view and self.multi_video_manager:
             is_running = self.multi_video_manager.on_global_step()
             if is_running:
                 # Update slider only if Global Playing
                 if self.multi_video_manager.is_global_playing and not self.user_dragging_slider:
                     ratio = self.multi_video_manager.get_progress_ratio()
                     self.seek_var.set(ratio * 100)
                 
                 self.lbl_time.config(text=self.multi_video_manager.get_current_time_str())
                 self.playback_after_id = self.root.after(33, self.playback_loop)
             else:
                 self.playback_after_id = None
                 if self.multi_video_manager.is_global_playing:
                     self.stop_playback()
             return

        if not self.is_playing:
            return

        if not self.cap:
            return
            
        # 処理開始時間を記録 (ドリフト補正用)
        loop_start = time.time()
        
        ret, frame = self.cap.read()
        if ret:
            curr_frame = int(self.cap.get(cv2.CAP_PROP_POS_FRAMES))
            self.display_frame(frame)
            if not self.user_dragging_slider:
                self.seek_var.set(curr_frame)
            self.update_time_label(curr_frame)
            
            # 処理時間を考慮した待機時間を計算
            elapsed = time.time() - loop_start
            target_delay = 1.0 / self.video_fps if self.video_fps > 0 else 0.033
            wait_s = target_delay - elapsed
            if wait_s < 0:
                wait_s = 0
            
            self.playback_after_id = self.root.after(int(wait_s * 1000), self.playback_loop)
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
                    
                    # --- Playback Filtering Logic ---
                    # TSV saves keys as "Ctrl,C" etc. (comma separated)
                    # ShortcutManager expects "Ctrl+C" (plus separated)
                    
                    # 1. Keys
                    if keys and keys != "None":
                         k_list = keys.split(',')
                         # Reconstruct mod+key string for checking
                         # Assuming k_list is already in [Mods..., Key] order from recorder
                         combo = "+".join(k_list)
                         if not self.shortcut_manager.is_allowed_playback(combo):
                              keys = "None"

                    # 2. Mouse (ScreenRecorderLogic saves click separately)
                    # But ShortcutManager might treat "Ctrl+L-Click" as a combo.
                    # We need to check if the mouse click (with modifiers) is allowed.
                    if click and click != "None":
                         # Get modifiers from keys (if any)
                         k_list = keys.split(',') if keys and keys != "None" else []
                         parts = [k for k in k_list if k in ("Ctrl", "Shift", "Alt", "Win")]
                         parts.append(click)
                         combo = "+".join(parts)
                         
                         if not self.shortcut_manager.is_allowed_playback(combo):
                              click = "None"
                    
                    # If both filtered out, skip
                    if click == "None" and keys == "None":
                         continue
                    # --------------------------------

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

            self.player_canvas.image = tk_img

    # --- 共通イベントハンドラ (Player & Preview) ---

    def _on_canvas_middle_down(self, event, mode):
        """中ボタン押下: パンニング開始"""
        if mode == "preview":
            self._preview_panning = True
            self._preview_pan_start = (event.x, event.y)
            self.preview_canvas.config(cursor="fleur")
        elif mode == "player":
            self._player_panning = True
            self._player_pan_start = (event.x, event.y)
            self.player_canvas.config(cursor="fleur")

    def _on_canvas_middle_drag(self, event, mode):
        """中ボタンドラッグ: パンニング中"""
        if mode == "preview" and self._preview_panning:
            dx = event.x - self._preview_pan_start[0]
            dy = event.y - self._preview_pan_start[1]
            self.preview_pan_x += dx
            self.preview_pan_y += dy
            self._preview_pan_start = (event.x, event.y)
        elif mode == "player" and self._player_panning:
            dx = event.x - self._player_pan_start[0]
            dy = event.y - self._player_pan_start[1]
            self.player_pan_x += dx
            self.player_pan_y += dy
            self._player_pan_start = (event.x, event.y)
            self.refresh_player_canvas()

    def _on_canvas_middle_up(self, event, mode):
        """中ボタン離上: パンニング終了"""
        if mode == "preview":
            self._preview_panning = False
            self.preview_canvas.config(cursor="")
        elif mode == "player":
            self._player_panning = False
            self.player_canvas.config(cursor="")

    def _on_canvas_middle_double_click(self, event, mode):
        """中ボタンダブルクリック: ズーム・パンリセット"""
        if mode == "preview":
            self.preview_pan_x = 0
            self.preview_pan_y = 0
            self.preview_zoom = 1.0
        elif mode == "player":
            self.player_pan_x = 0
            self.player_pan_y = 0
            self.player_zoom = 1.0
            self.refresh_player_canvas()

    def _on_canvas_wheel(self, event, mode):
        """マウスホイール: ズームイン・アウト"""
        # 拡大率計算
        zoom_factor = 1.1
        if event.num == 5 or event.delta < 0:
            zoom_factor = 1.0 / 1.1

        if mode == "preview":
            self.preview_zoom *= zoom_factor
            self.preview_zoom = max(0.1, min(10.0, self.preview_zoom))
        elif mode == "player":
            self.player_zoom *= zoom_factor
            self.player_zoom = max(0.1, min(10.0, self.player_zoom))
            self.refresh_player_canvas()

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
            val = float(value)
            if self.is_multi_view and self.multi_video_manager:
                 ratio = val / 100.0
                 self.multi_video_manager.seek_all_by_ratio(ratio)
            else:
                 val = int(val)
                 self.show_frame(val)
                 self.update_time_label(val)

    def on_slider_release(self, event):
        self.user_dragging_slider = False
        val = float(self.seek_var.get())
        
        if self.is_multi_view and self.multi_video_manager:
            # Multi View Seek
            # Slider is 0-100 ratio
            ratio = val / 100.0
            self.multi_video_manager.seek_all_by_ratio(ratio)
        else:
            # Single View Seek
            val = int(val)
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


    def open_shortcut_settings(self):
        """ショートカットキー設定ダイアログを開く"""
        ShortcutSettingsDialog(self.root, self.shortcut_manager)

    def _fix_all_button_active_colors(self, container: tk.Widget):
        """コンテナ内の全ボタンの activebackground を背景色に合わせる."""
        from ui_utils import fix_button_active_colors
        fix_button_active_colors(container)
if __name__ == "__main__":
    app = ScreenRecorderApp()
    app.root.mainloop()
