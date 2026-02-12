"""アプリケーションメインモジュール

動画編集を行うためのウィンドウ。
"""
from __future__ import annotations

import json
import math
import os
import re
import time
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import filedialog, messagebox, simpledialog
from typing import TYPE_CHECKING, Any, Callable
import sys

import cv2
import CapchunScreen
import overlay_utils
from PIL import Image, ImageTk

# 分割したモジュールからインポート
from export import open_file
from utils import (
    get_base_dir,
    sec_to_hhmmss,
    sec_to_display,
    hhmmss_to_sec,
    imwrite_jp,
    ratio_value_from_str,
    ratio_label_from_wh,
    open_folder_with_selection,
    resource_path,
)
from config import (
    CONFIG_FILENAME,
    PROJECT_NAME,
    PROJECT_VERSION,
    load_global_config,
    save_global_config,
    load_video_settings,
    save_video_settings as save_video_settings_to_file,
    normalize_presets,
    get_default_presets_with_labels,
)
from clipboard import copy_image_to_clipboard
from seekbar import SeekbarMixin
from crop_handler import CropHandlerMixin
from export import ExportMixin
from ui_utils import add_tooltip as _add_tooltip, fix_button_active_colors


class VideoCropperApp(SeekbarMixin, CropHandlerMixin, ExportMixin):
    """動画のクロップと出力を行うGUIアプリケーション."""

    # ウィンドウの最小サイズ
    MIN_WINDOW_W: int = 800
    MIN_WINDOW_H: int = 600

    # キャンバスサイズ
    CANVAS_W: int = 640
    CANVAS_H: int = 360
    CANVAS_MIN_H: int = 60  # 映像表示領域の最小高さ

    # クロップ矩形の最小サイズ
    MIN_W: int = 20
    MIN_H: int = 20

    # シークバー設定
    SEEK_H: int = 100
    SEEK_MARGIN: int = 20

    # 矩形リサイズハンドルのサイズとエッジ判定マージン(デフォルト値、設定ファイルで上書き可能)
    HANDLE_SIZE: int = 8
    EDGE_MARGIN: int = 20

    # 下側パネルの固定高さ
    BOTTOM_PANEL_FIXED_HEIGHT: int = 450

    def __init__(self, root):
        self.root = root
        self.root.title(f"{PROJECT_NAME} {PROJECT_VERSION} - 編集ツール")
        
        # アイコン設定
        try:
            self.icon_image = tk.PhotoImage(file=resource_path("ChulipVideo.png"))
            self.root.iconphoto(True, self.icon_image)
        except Exception as e:
            print(f"Icon Load Error: {e}")
        
        # Load global config for theme
        self.global_config = load_global_config()
        self.theme = self.global_config.get("theme", {})
        
        # テーマ設定からハンドルサイズとエッジマージンを取得
        self.HANDLE_SIZE = self.theme.get("handle_size")
        self.EDGE_MARGIN = self.theme.get("edge_margin")

        # ウィンドウの最小サイズを設定
        self.root.minsize(self.MIN_WINDOW_W, self.MIN_WINDOW_H)

        # ビデオ / 再生状態
        self.cap = None
        self.frame = None
        self.fps = 30.0
        self.duration = 0
        self.playing = False
        self.current_time = 0
        self.speed = float(self.global_config.get("play_speed", 1.0))
        self.png_compression = int(self.global_config.get("png_compression", 3))
        self.video_filename = ""  # 動画ファイル名（拡張子除く）
        self.video_filepath = ""  # 動画ファイルのフルパス
        self.vid_w = 1920 # 初期値
        self.vid_h = 1080

        # トリム時間
        self.start_time = 0
        self.end_time = 0

        # クロップ矩形
        self.crop_rect = [100, 80, 300, 250]
        self.dragging_rect = False
        self.resizing_rect = False
        self.drag_offset = (0, 0)
        self.resize_edge = None
        self.orig_rect = None
        self.maintain_aspect_ratio = False  # Shift キー押下時のアスペクト比ロック
        self.orig_aspect_ratio = 1.0  # 元のアスペクト比
        # 矩形のフォーカス状態（左クリックでオレンジにする）
        self.rect_focused = False
        # 矩形のホバー状態
        self.rect_hovered = False

        # Seekbar dragging state
        self.drag_mode = None  # "current", "start", "end"

        # フレーム出力フィルタのチェックボックス状態
        self.check_prev_next = None  # 前後のフレームと異なるとき出力しない
        # Playback modifiers
        self.play_range_mode = False  # 区間再生モード
        self.check_duplicate = None  # 直前に出力したフレームと同一のとき出力しない

        # クロップの Undo スタック（メモリのみ、無制限）
        self.crop_history = []
        self.crop_redo = []

        # 解像度プリセット (name -> (w,h))
        self.resolution_presets = {}
        # アスペクト比オプション（UI で選べる固定リスト）
        self.aspect_options = ["未指定", "16:9", "4:3", "3:4", "9:16", "21:9", "1:1", "4:5", "5:4"]

        # キャンバスリサイズ用
        self.canvas_scale_x = 1.0  # キャンバスの横スケール比
        self.canvas_scale_y = 1.0  # キャンバスの縦スケール比
        self.canvas_offset_x = 0  # キャンバス内での画像のX オフセット
        self.canvas_offset_y = 0  # キャンバス内での画像のY オフセット
        # 画像のズームとパン（ホイールでズーム、中央クリックでパン）
        self.image_zoom = 1.0
        self.pan_offset_x = 0
        self.pan_offset_y = 0
        self._panning = False
        self._pan_start = (0, 0)

        self.lock_var = tk.BooleanVar(value=False)
        self.lock_move_var = tk.BooleanVar(value=False)

        # ツールチップ用のストレージ
        self._tooltips = {}

        # マウス軌跡のTSVデータ用
        self.trajectory_data = [] # list of (time, x, y)
        self.show_trajectory_var = tk.BooleanVar(value=self.global_config.get("show_trajectory", True))
        self.show_subtitle_var = tk.BooleanVar(value=self.global_config.get("show_subtitle", True))
        self.embed_overlay_var = tk.BooleanVar(value=self.global_config.get("embed_overlay", False))
        self.allow_oversize_var = tk.BooleanVar(value=False)

        # UIを先に構築
        self.build_ui()

        self._play_after_id = None
        # スペースキーをどのウィジェットにフォーカスがあっても捕まえる
        self.root.bind_all("<space>", lambda e: self._should_trigger_shortcut() and self.toggle_play())
        # Undo with Ctrl+Z
        self.root.bind_all("<Control-z>", lambda e: self._should_trigger_shortcut() and self.undo_crop())
        self.root.bind_all("<Control-Z>", lambda e: self._should_trigger_shortcut() and self.undo_crop())
        # Copy crop to clipboard with Ctrl+C
        self.root.bind_all("<Control-c>", lambda e: self._should_trigger_shortcut() and self.copy_crop_to_clipboard())
        self.root.bind_all("<Control-C>", lambda e: self._should_trigger_shortcut() and self.copy_crop_to_clipboard())
        # Redo with Ctrl+Y
        self.root.bind_all("<Control-y>", lambda e: self._should_trigger_shortcut() and self.redo_crop())
        self.root.bind_all("<Control-Y>", lambda e: self._should_trigger_shortcut() and self.redo_crop())
        # Arrow key repeat handlers (Left/Right)
        self._arrow_repeat_id = None
        self._arrow_dir = None
        self._arrow_start_time = None
        self.root.bind_all('<KeyPress-Left>', lambda e: self._should_trigger_shortcut() and self._on_arrow_press(e, -1))
        self.root.bind_all('<KeyRelease-Left>', lambda e: self._should_trigger_shortcut() and self._on_arrow_release(e))
        self.root.bind_all('<KeyPress-Right>', lambda e: self._should_trigger_shortcut() and self._on_arrow_press(e, 1))
        self.root.bind_all('<KeyRelease-Right>', lambda e: self._should_trigger_shortcut() and self._on_arrow_release(e))
        # Alt+矢印でクロップ矩形を移動（Ctrl併用で10px）
        self.root.bind_all('<Alt-Up>', lambda e: self._should_trigger_shortcut() and (self.move_crop_by(0, -10 if (e.state & 0x4) else -1) or "break"))
        self.root.bind_all('<Alt-Down>', lambda e: self._should_trigger_shortcut() and (self.move_crop_by(0, 10 if (e.state & 0x4) else 1) or "break"))
        self.root.bind_all('<Alt-Left>', lambda e: self._should_trigger_shortcut() and (self.move_crop_by(-10 if (e.state & 0x4) else -1, 0) or "break"))
        self.root.bind_all('<Alt-Right>', lambda e: self._should_trigger_shortcut() and (self.move_crop_by(10 if (e.state & 0x4) else 1, 0) or "break"))
        # Shift+矢印でクロップ矩形を拡大縮小（Ctrl併用で10px）
        self.root.bind_all('<Shift-Up>', lambda e: self._should_trigger_shortcut() and (self.expand_crop(0, 10 if (e.state & 0x4) else 1) or "break"))
        self.root.bind_all('<Shift-Down>', lambda e: self._should_trigger_shortcut() and (self.expand_crop(0, -10 if (e.state & 0x4) else -1) or "break"))
        self.root.bind_all('<Shift-Left>', lambda e: self._should_trigger_shortcut() and (self.expand_crop(-10 if (e.state & 0x4) else -1, 0) or "break"))
        self.root.bind_all('<Shift-Right>', lambda e: self._should_trigger_shortcut() and (self.expand_crop(10 if (e.state & 0x4) else 1, 0) or "break"))
        # Home/End bindings
        self.root.bind_all('<Home>', lambda e: self._should_trigger_shortcut() and self.set_current_time_direct(self.start_time))
        self.root.bind_all('<End>', lambda e: self._should_trigger_shortcut() and self.set_current_time_direct(self.end_time))
        self.root.bind_all('<Control-Home>', lambda e: self._should_trigger_shortcut() and self.set_current_time_direct(0))
        self.root.bind_all('<Control-End>', lambda e: self._should_trigger_shortcut() and self.set_current_time_direct(self.duration))
        # Ctrl+Sで現在のクロップ範囲をPNGとして保存
        self.root.bind_all('<Control-s>', lambda e: self._should_trigger_shortcut() and self.save_current_frame_as_png())
        self.root.bind_all('<Control-S>', lambda e: self._should_trigger_shortcut() and self.save_current_frame_as_png())
        
        # マウスの4, 5ボタン (戻る/進む) -> -1s, +1s
        # Windows の Tkinter では一般に Button-4, Button-5 がサイドボタンに割り当てられます
        self.root.bind_all("<Button-4>", lambda e: self._should_trigger_shortcut() and self.adjust_time(lambda: self.current_time, self.set_current_time_direct, -1))
        self.root.bind_all("<Button-5>", lambda e: self._should_trigger_shortcut() and self.adjust_time(lambda: self.current_time, self.set_current_time_direct, 1))

        # ウィンドウ終了時に設定を保存
        self.root.protocol("WM_DELETE_WINDOW", self.on_window_close)

        # 動画設定を読み込む
        self.load_config()

        # UI構築と設定読み込み完了後にウィンドウ位置・サイズを復元
        self.root.after(10, self.load_window_geometry)

        # 起動100ms後にレイアウト調整（シークバー等のリサイズ確実化）
        self.root.after(100, lambda: self.on_canvas_resize(None))

        # 入力履歴管理
        from overlay_utils import InputHistoryManager
        self.input_history_manager = InputHistoryManager()

        # ショートカット抑制フラグ (詳細設定でのテスト用)
        self.suppress_shortcuts = False

    def _should_trigger_shortcut(self) -> bool:
        """ショートカットキーイベントを処理すべきか判定する (別ウィンドウにフォーカスがある場合は無視)"""
        if self.suppress_shortcuts:
            return False
            
        # 現在フォーカスを持っているウィジェットを取得
        focused = self.root.focus_get()
        if focused:
            try:
                # そのウィジェットが属しているトップレベルウィンドウを取得
                top = focused.winfo_toplevel()
                # メインウィンドウ(root)でない場合はショートカットを無視
                if top != self.root:
                    return False
            except Exception:
                pass
        return True

    # ---------------- UI Construction ----------------
    def build_ui(self):
        # Top用フォント (カラー絵文字対応: Segoe UI Emoji)
        top_font = ("Segoe UI Emoji", 12)
        
        # Top: Load path and controls
        top_panel = tk.Frame(self.root)
        top_panel.pack(fill=tk.X, side=tk.TOP, padx=5, pady=6)

        # 録画ツール起動ボタン (赤系 -> Theme)
        btn_bg = self.theme.get("main_color")
        btn_rec = tk.Button(top_panel, text="🎥録画", command=self.open_CapchunScreen,
                  bg=btn_bg, width=6, font=top_font)
        btn_rec.pack(side=tk.LEFT, padx=5)
        self.add_tooltip(btn_rec, "録画ツールを起動する")

        btn_open = tk.Button(top_panel, text="🎞️選択", command=self.load_video,
                  width=6, bg=self.theme.get("button_normal_bg"), font=top_font)
        btn_open.pack(side=tk.LEFT, padx=5)
        self.add_tooltip(btn_open, "動画ファイルを選択して開く")

        self.entry_fullpath_var = tk.StringVar(value="")
        self.entry_fullpath = tk.Entry(top_panel, textvariable=self.entry_fullpath_var, font=top_font)
        self.entry_fullpath.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0,2))

        # 保存フォルダを開くボタン (フォルダアイコン 📂)
        btn_folder = tk.Button(top_panel, text="📂", 
                               command=lambda: open_folder_with_selection(self.video_filepath), 
                               width=3, bg=self.theme.get("button_normal_bg"), font=top_font)
        btn_folder.pack(side=tk.LEFT, padx=(0,5))
        self.add_tooltip(btn_folder, "保存フォルダを開く")

        # TSVデータを開くボタン (テキストアイコン 📄)
        btn_tsv = tk.Button(top_panel, text="📄", 
                            command=self.open_current_tsv, 
                            width=3, bg=self.theme.get("button_normal_bg"), font=top_font)
        btn_tsv.pack(side=tk.LEFT, padx=(0,5))
        self.add_tooltip(btn_tsv, "TSVファイルを開く")
        
        # 設定表示ボタン
        btn_check = tk.Button(top_panel, text="⚙️", command=self.open_video_settings, 
                               width=3, bg=self.theme.get("button_normal_bg"), font=top_font)
        btn_check.pack(side=tk.LEFT, padx=(0,5))
        self.add_tooltip(btn_check, "動画の【トリム/クロップ設定TSV】を開く")

        # 再読み込みボタン
        btn_reload = tk.Button(top_panel, text="🔃", command=self.load_config, 
                               width=3, bg=self.theme.get("button_reload_bg"), font=top_font)
        btn_reload.pack(side=tk.LEFT, padx=(0,5))
        self.add_tooltip(btn_reload, "【動画】・【トリム/クロップ設定JSON】・【入力履歴】を再読み込み")
        
        # 設定保存ボタン
        btn_save = tk.Button(top_panel, text="💾", command=self.save_video_settings, 
                             width=3, bg=self.theme.get("button_save_bg"), font=top_font)
        btn_save.pack(side=tk.LEFT, padx=(0,5))
        self.add_tooltip(btn_save, "動画の【トリム/クロップ設定JSON】を保存する")

        # 秒数表示ラベル（右隣）
        self.label_seconds = tk.Label(top_panel, text="(0s)", font=top_font)
        self.label_seconds.pack(side=tk.LEFT, padx=5)

        # 右上のヘルプボタン（ショートカット一覧）
        self.btn_help = tk.Button(top_panel, text="❓️", command=self.show_shortcuts, 
                                  width=3, bg=self.theme.get("button_help_bg"), font=top_font)
        self.btn_help.pack(side=tk.RIGHT, padx=4)
        self.add_tooltip(self.btn_help, "ショートカット一覧を表示する")

        # FPS表示ラベル (保存ボタンの左)
        self.label_fps = tk.Label(top_panel, text="--FPS", font=("Consolas", 14, "bold"), fg="#666666")
        self.label_fps.pack(side=tk.RIGHT, padx=10)
        self.add_tooltip(self.label_fps, "動画の1秒あたりのフレーム数(FPS)")

        # 表示倍率ラベル (FPSボタンの左)
        self.label_zoom = tk.Label(top_panel, text="100%", font=("Consolas", 14, "bold"), fg="#666666")
        self.label_zoom.pack(side=tk.RIGHT, padx=10)
        self.add_tooltip(self.label_zoom, "現在の動画表示倍率")

        # 1. Main Container (Replacing PanedWindow to lock the split)
        # 以前は PanedWindow を使用していましたが、下側パネルの高さを 450px に固定し、
        # ユーザーによる変更を不可にするため、通常の Frame パックに変更しました。
        
        # 2. Control Panel Frame (先に BOTTOM で pack して高さを固定)
        control_pane = tk.Frame(self.root, bg="#f5f5f5", height=self.BOTTOM_PANEL_FIXED_HEIGHT)
        control_pane.pack(side=tk.BOTTOM, fill=tk.X)
        control_pane.pack_propagate(False) # 子要素に寄らず高さを固定
        
        # 1.1 Video Canvas Pane (残りの領域をすべて埋める)
        canvas_pane = tk.Frame(self.root)
        canvas_pane.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        # キャンバスの背景色をテーマから取得
        canvas_bg = self.theme.get("canvas_bg")
        self.canvas = tk.Canvas(
            canvas_pane, width=self.CANVAS_W, height=self.CANVAS_H, bg=canvas_bg)
        self.canvas.pack(fill=tk.BOTH, expand=True)
        self.canvas_image = self.canvas.create_image(0, 0, anchor=tk.NW)
        linecolor = self.theme.get("crop_default_linecolor")
        width = self.theme.get("crop_width")
        linestyle = self.theme.get("crop_default_linestyle")
        # tkinter の dash 引数として適切な型(tuple)に変換
        if isinstance(linestyle, list):
            linestyle = tuple(linestyle)

        self.rect_id = self.canvas.create_rectangle(
            *self.crop_rect, outline=linecolor, width=width, dash=linestyle)

        self.canvas.bind("<ButtonPress-1>", self.on_mouse_down)
        self.canvas.bind("<B1-Motion>", self.on_mouse_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_mouse_up)
        # 右クリックで矩形移動を可能にする
        self.canvas.bind("<ButtonPress-3>", self.on_right_mouse_down)
        self.canvas.bind("<B3-Motion>", self.on_right_mouse_drag)
        self.canvas.bind("<ButtonRelease-3>", self.on_right_mouse_up)
        self.canvas.bind("<Configure>", self.on_canvas_resize)
        self.canvas.bind("<Motion>", self.on_canvas_motion)
        # 中央クリックで表示位置を移動（パン）
        self.canvas.bind("<ButtonPress-2>", self.on_middle_down)
        self.canvas.bind("<B2-Motion>", self.on_middle_drag)
        self.canvas.bind("<ButtonRelease-2>", self.on_middle_up)
        self.canvas.bind("<Double-Button-2>", self.on_middle_double_click)
        # マウスホイールでズーム（Windows と X11 両対応）
        self.canvas.bind("<MouseWheel>", self.on_canvas_wheel)
        # Note: Button-4/5 は Windows ではサイドボタンとして global bind したため、ここでは不要

        
        control_frame = tk.Frame(control_pane, bg="#f5f5f5")
        control_frame.pack(fill=tk.BOTH, expand=True)

        # 2. Main Controls (Play, Speed)
        main_ctrl = tk.Frame(control_frame)
        main_ctrl.pack(pady=2)

        self.btn_video_start = tk.Button(main_ctrl, text="◀◀先頭", command=self.go_to_video_start, width=8, bg=self.theme.get("button_normal_bg"))
        self.btn_video_start.pack(side=tk.LEFT, padx=4)
        self.add_tooltip(self.btn_video_start, "Ctrl+Home: 動画先頭へ")
        self.btn_trim_start = tk.Button(main_ctrl, text="◀開始位置", command=self.go_to_trim_start, width=10, bg=self.theme.get("button_trim_start_bg"))
        self.btn_trim_start.pack(side=tk.LEFT, padx=4)
        self.add_tooltip(self.btn_trim_start, "Home: 開始位置へ")

        self.btn_play = tk.Button(main_ctrl, text="▲再生", command=self.toggle_play, width=12, bg=self.theme.get("button_play_bg", "#B3E5FC"))
        # 区間再生はチェックボックス化（末尾ボタンの右）
        self.btn_play.pack(side=tk.LEFT, padx=4)
        self.add_tooltip(self.btn_play, "Space: 再生/停止")

        btn_end = tk.Button(main_ctrl, text="終了位置▶", command=self.go_to_trim_end, width=10, bg=self.theme.get("button_trim_end_bg"))
        btn_end.pack(side=tk.LEFT, padx=4)
        self.add_tooltip(btn_end, "End: 終了位置へ")
        btn_tail = tk.Button(main_ctrl, text="末尾▶▶", command=self.go_to_video_end, width=8, bg=self.theme.get("button_normal_bg"))
        btn_tail.pack(side=tk.LEFT, padx=4)
        self.add_tooltip(btn_tail, "Ctrl+End: 動画末尾へ")
        self.range_var = tk.BooleanVar(value=self.global_config.get("play_range", False))
        tk.Checkbutton(main_ctrl, text="区間再生", variable=self.range_var).pack(side=tk.LEFT, padx=(4,8))

        # ループ再生チェックボックス
        self.loop_var = tk.BooleanVar(value=self.global_config.get("play_loop", False))
        tk.Checkbutton(main_ctrl, text="ループ再生", variable=self.loop_var).pack(side=tk.LEFT, padx=(8,4))
        # 往復再生（ループがオンのときのみ有効）
        self.pingpong_var = tk.BooleanVar(value=self.global_config.get("play_pingpong", False))
        self.chk_pingpong = tk.Checkbutton(main_ctrl, text="往復再生", variable=self.pingpong_var, state=tk.NORMAL if self.loop_var.get() else tk.DISABLED)
        self.chk_pingpong.pack(side=tk.LEFT, padx=(4,8))
        self.add_tooltip(self.chk_pingpong, "ループ時のみ有効: 端で再生方向を反転")

        # ループ状態に応じて往復チェックの有効/無効を切替
        def _on_loop_change(*args):
            if self.loop_var.get():
                self.chk_pingpong.config(state=tk.NORMAL)
            else:
                self.chk_pingpong.config(state=tk.DISABLED)
                self.pingpong_var.set(False)
        self.loop_var.trace_add('write', _on_loop_change)

        # Speed input with Spinbox
        tk.Label(main_ctrl, text="再生速度:").pack(side=tk.LEFT, padx=(10, 2))
        self.speed_var = tk.StringVar(value=f"{self.speed:.1f}")
        self.speed_spinbox = tk.Spinbox(
            main_ctrl,
            from_=-9.9,
            to=9.9,
            increment=0.1,
            width=6,
            textvariable=self.speed_var,
            command=self.change_speed
        )
        self.speed_spinbox.pack(side=tk.LEFT, padx=5)
        self.add_tooltip(self.speed_spinbox, "再生速度: -9.9〜9.9 (Enterで確定)")

        # prevent space key from inserting into these controls (Space should toggle play)
        def _ignore_space(e):
            return "break"
        self.speed_spinbox.bind('<space>', _ignore_space)

        # Enterキーを押したときに更新
        self.speed_spinbox.bind("<Return>", self.change_speed)
        # 別の場所をクリックした（フォーカスが外れた）ときに更新
        self.speed_spinbox.bind("<FocusOut>", self.change_speed)
        # マウスホイールで速度を変更
        self.speed_spinbox.bind("<MouseWheel>", self.on_speed_wheel)
        self.speed_spinbox.bind("<Button-4>", self.on_speed_wheel)
        self.speed_spinbox.bind("<Button-5>", self.on_speed_wheel)

        

        # 3. Time Editor Panel (Start / Current / End) - will be shown above crop panel
        time_panel = tk.LabelFrame(control_frame, text="トリミング範囲", bd=1,
               relief=tk.SOLID, padx=5, pady=0)
        time_panel.pack(pady=1, fill=tk.X, padx=10)

        # 4. Seekbar Canvas (ウィンドウ幅に応じたサイズで初期化) -- placed above time controls
        # シークバーはウィンドウ幅いっぱいに広がるようにする
        self.seek_canvas = tk.Canvas(
            time_panel, height=self.SEEK_H, bg="#f0f0f0")
        self.seek_canvas.pack(pady=1, padx=0, fill=tk.X, expand=True)
        self.seek_canvas.bind("<Button-1>", self.seek_on_click)
        self.seek_canvas.bind("<B1-Motion>", self.seek_on_drag)
        self.seek_canvas.bind("<ButtonRelease-1>", self.seek_on_release)

        # Helper to create time controls
        def create_time_control(parent, label, color, var_getter, var_setter, add_move=False):
            frame = tk.Frame(parent)
            frame.pack(side=tk.LEFT, expand=True, pady=2)

            lbl = tk.Label(frame, text=label, fg=color, font=("Arial", 10, "bold"))
            lbl.pack(anchor=tk.CENTER)
            ctrl_f = tk.Frame(frame)
            # ラベルと入力欄の間に少し余白を入れる
            ctrl_f.pack(pady=3)

            btn_minus = tk.Button(ctrl_f, text="-1s", width=4, bg=color if add_move else self.theme.get("button_normal_bg"),
                      command=lambda: self.adjust_time(var_getter, var_setter, -1))
            btn_minus.pack(side=tk.LEFT)
            self.add_tooltip(btn_minus, "-1s: 1秒戻す")

            entry = tk.Entry(ctrl_f, width=14, font=(
                "Consolas", 12), justify="center")
            entry.insert(0, "00:00:00")
            entry.pack(side=tk.LEFT, padx=2)
            entry.bind("<Return>", lambda e: self.manual_entry_update(
                entry, var_setter))
            entry.bind("<FocusOut>", lambda e: self.manual_entry_update(
                entry, var_setter))

            btn_plus = tk.Button(ctrl_f, text="+1s", width=4, bg=color if add_move else self.theme.get("button_normal_bg"),
                      command=lambda: self.adjust_time(var_getter, var_setter, 1))
            btn_plus.pack(side=tk.LEFT)
            self.add_tooltip(btn_plus, "+1s: 1秒進める")

            return entry

        # Start Time Control
        self.entry_start = create_time_control(time_panel, "▼Start (開始)", self.theme.get("start_color_bg"),
                               lambda: self.start_time, self.set_start_time_direct, True)

        # Divider
        tk.Frame(time_panel, width=2, bg="#ccc", height=40).pack(side=tk.LEFT, padx=10)

        # Current Time Control
        self.entry_current = create_time_control(time_panel, "●Current (現在)", "#000000",
                             lambda: self.current_time, self.set_current_time_direct, False)
        # divider between controls
        tk.Frame(time_panel, width=2, bg="#ccc", height=40).pack(side=tk.LEFT, padx=10)

        # End Time Control
        self.entry_end = create_time_control(time_panel, "▲End (終了)", self.theme.get("end_color_bg"),
                             lambda: self.end_time, self.set_end_time_direct, True)

        # disable space input for time entry boxes
        self.entry_start.bind('<space>', _ignore_space)
        self.entry_current.bind('<space>', _ignore_space)
        self.entry_end.bind('<space>', _ignore_space)


        # 2.7. クロップサイズ入力パネル (フィードバックを反映した精密レイアウト)
        crop_panel = tk.LabelFrame(control_frame, text="クロップ範囲", bd=1, relief=tk.SOLID, padx=5, pady=5)
        crop_panel.pack(pady=5, fill=tk.X, padx=10)

        # 全要素を横に並べるメインコンテナ (確実に水平中央に配置しつつ垂直拡張)
        self.crop_hbox = tk.Frame(crop_panel)
        self.crop_hbox.pack(fill=tk.Y, expand=True, anchor=tk.CENTER)
        
        # 不要な外部へのサイド余白を削り、中央寄せを強調

        # --- セクション 1: 操作履歴 (Undo / Redo 横並び) ---
        sec1 = tk.Frame(self.crop_hbox)
        sec1.pack(side=tk.LEFT, padx=5, fill=tk.Y)
        self.btn_undo = tk.Button(sec1, text="↩️", command=self.undo_crop, width=5, font=("Segoe UI Emoji", 11), bg=self.theme.get("button_undo_bg"), relief=tk.RAISED)
        self.btn_undo.pack(side=tk.LEFT, padx=1, fill=tk.Y)
        self.btn_undo.config(state=tk.DISABLED)
        self.add_tooltip(self.btn_undo, "Ctrl+Z: 戻す")
        self.btn_redo = tk.Button(sec1, text="↪️", command=self.redo_crop, width=5, font=("Segoe UI Emoji", 11), bg=self.theme.get("button_redo_bg"), relief=tk.RAISED)
        self.btn_redo.pack(side=tk.LEFT, padx=1, fill=tk.Y)
        self.btn_redo.config(state=tk.DISABLED)
        self.add_tooltip(self.btn_redo, "Ctrl+Y: 進む")

        ttk.Separator(self.crop_hbox, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=10)

        # --- 枠外選択 ---
        sec_opt = tk.Frame(self.crop_hbox)
        sec_opt.pack(side=tk.LEFT, padx=5, fill=tk.Y)
        tk.Checkbutton(sec_opt, text="枠外\n選択", variable=self.allow_oversize_var).pack(side=tk.LEFT)
        
        ttk.Separator(self.crop_hbox, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=10)

        # --- セクション 2: 座標 ---
        sec2 = tk.Frame(self.crop_hbox)
        sec2.pack(side=tk.LEFT, padx=5, fill=tk.Y)
        tk.Label(sec2, text="座標:").pack(side=tk.LEFT, padx=(0,5))
        
        # 2段分の高さを持つ Lock ボタン (正方形)
        self.btn_lock_move = tk.Button(sec2, text="🔓", width=5, font=("Consolas", 12, "bold"), bg=self.theme.get("button_unlocked_bg"), command=self.toggle_move_lock)
        self.btn_lock_move.pack(side=tk.LEFT, fill=tk.Y, padx=(0,5), expand=True) # expand=True 追加
        self.add_tooltip(self.btn_lock_move, "クロップ位置(X, Y)をロック (リサイズは可能)")

        # X/Y のペア (上下2段)
        pair2 = tk.Frame(sec2)
        pair2.pack(side=tk.LEFT, fill=tk.Y, expand=True) # expand=True 追加
        # 上段: X
        f2a = tk.Frame(pair2)
        f2a.pack(side=tk.TOP, fill=tk.X, expand=True)
        tk.Label(f2a, text="X:", width=2).pack(side=tk.LEFT)
        self.entry_crop_x = tk.Entry(f2a, width=6, font=("Consolas", 11), justify="center")
        self.entry_crop_x.pack(side=tk.LEFT)
        self.entry_crop_x.bind("<Return>", lambda e: self.update_crop_from_entries())
        self.entry_crop_x.bind("<FocusOut>", lambda e: self.update_crop_from_entries())
        self.add_tooltip(self.entry_crop_x, "クロップ開始X座標 (Alt+左右で1px移動, Shift+左右で幅を拡大縮小)")
        # 下段: Y
        f2b = tk.Frame(pair2)
        f2b.pack(side=tk.TOP, fill=tk.X, expand=True, pady=(2,0))
        tk.Label(f2b, text="Y:", width=2).pack(side=tk.LEFT)
        self.entry_crop_y = tk.Entry(f2b, width=6, font=("Consolas", 11), justify="center")
        self.entry_crop_y.pack(side=tk.LEFT)
        self.entry_crop_y.bind("<Return>", lambda e: self.update_crop_from_entries())
        self.entry_crop_y.bind("<FocusOut>", lambda e: self.update_crop_from_entries())
        self.add_tooltip(self.entry_crop_y, "クロップ開始Y座標 (Alt+上下で1px移動, Shift+上下で高さを拡大縮小)")

        ttk.Separator(self.crop_hbox, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=10)

        # --- セクション 3: 解像度設定 ---
        sec3 = tk.Frame(self.crop_hbox)
        sec3.pack(side=tk.LEFT, padx=5, fill=tk.Y)
        tk.Label(sec3, text="解像度:").pack(side=tk.LEFT, padx=(0,5))
        self.ratio_var = tk.StringVar(value=self.aspect_options[0])
        # 2段分の高さを持つ Lock ボタン (正方形)
        self.btn_lock_res = tk.Button(sec3, text="🔓", width=5, font=("Consolas", 12, "bold"), bg=self.theme.get("button_unlocked_bg"), command=self.toggle_resolution_lock)
        self.btn_lock_res.pack(side=tk.LEFT, fill=tk.Y, padx=(0,5))
        self.add_tooltip(self.btn_lock_res, "解像度・アスペクト比をロック (移動は可能)")

        # 比率 / プリセット選択のペア
        p_aspect = tk.Frame(sec3)
        p_aspect.pack(side=tk.LEFT, padx=5)
        self.ratio_optionmenu = tk.OptionMenu(p_aspect, self.ratio_var, *self.aspect_options)
        self.ratio_optionmenu.config(width=10, font=("Meiryo UI", 9))
        self.ratio_optionmenu.pack(side=tk.TOP, fill=tk.X)
        self.add_tooltip(self.ratio_optionmenu, "特定のアスペクト比に固定・指定")
        self.resolution_var = tk.StringVar(value="カスタム")
        self.resolution_optionmenu = tk.OptionMenu(p_aspect, self.resolution_var, "カスタム")
        self.resolution_optionmenu.config(width=10, font=("Meiryo UI", 9))
        self.resolution_optionmenu.pack(side=tk.TOP, fill=tk.X, pady=(2,0))
        self.add_tooltip(self.resolution_optionmenu, "よく使う解像度プリセットを選択")
        self.ratio_var.trace_add('write', lambda *args: self.update_resolution_menu())

        # 幅 / 高 のペア
        p_size = tk.Frame(sec3)
        p_size.pack(side=tk.LEFT, fill=tk.Y, expand=True)

        # 上段: 幅
        f4a = tk.Frame(p_size)
        f4a.pack(side=tk.TOP, fill=tk.X, expand=True)
        tk.Label(f4a, text="幅:", width=2).pack(side=tk.LEFT)
        self.entry_crop_w = tk.Entry(f4a, width=6, font=("Consolas", 11), justify="center")
        self.entry_crop_w.pack(side=tk.LEFT)
        self.entry_crop_w.bind("<Return>", lambda e: self.update_crop_from_entries())
        self.entry_crop_w.bind("<FocusOut>", lambda e: self.update_crop_from_entries())
        self.add_tooltip(self.entry_crop_w, "出力される画像の幅 (ピクセル)")
        # 下段: 高
        f4b = tk.Frame(p_size)
        f4b.pack(side=tk.TOP, fill=tk.X, expand=True, pady=(2,0))
        tk.Label(f4b, text="高:", width=2).pack(side=tk.LEFT)
        self.entry_crop_h = tk.Entry(f4b, width=6, font=("Consolas", 11), justify="center")
        self.entry_crop_h.pack(side=tk.LEFT)
        self.entry_crop_h.bind("<Return>", lambda e: self.update_crop_from_entries())
        self.entry_crop_h.bind("<FocusOut>", lambda e: self.update_crop_from_entries())
        self.add_tooltip(self.entry_crop_h, "出力される画像の高さ (ピクセル)")

        # プリセット操作 (保存 / 削除) のペア
        p_presets = tk.Frame(sec3)
        p_presets.pack(side=tk.LEFT, padx=5)
        self.btn_save_preset = tk.Button(p_presets, text="プリセット保存", command=self.add_resolution_preset, bg=self.theme.get("button_reload_bg"), font=("Meiryo UI", 9))
        self.btn_save_preset.pack(side=tk.TOP, fill=tk.X)
        self.add_tooltip(self.btn_save_preset, "現在の解像度をプリセットに追加")
        self.btn_delete_preset = tk.Button(p_presets, text="プリセット削除", command=self.delete_resolution_preset, bg=self.theme.get("button_undo_bg"), font=("Meiryo UI", 9))
        self.btn_delete_preset.pack(side=tk.TOP, fill=tk.X, pady=(2,0))
        self.add_tooltip(self.btn_delete_preset, "選択中のプリセットを削除")

        

        # 5. 出力グループ（設定、PNG、動画/GIF）の3列構成 (1:2:1 の比率で配置)
        output_panel = tk.LabelFrame(control_frame, text="エクスポート", bd=1, relief=tk.SOLID, padx=5, pady=5)
        output_panel.pack(fill=tk.X, padx=10, pady=5)
        output_panel.columnconfigure(0, weight=1) # 操作の軌跡
        output_panel.columnconfigure(1, weight=0) # セパレータ
        output_panel.columnconfigure(2, weight=2) # 静止画出力
        output_panel.columnconfigure(3, weight=0) # セパレータ
        output_panel.columnconfigure(4, weight=1) # 動画出力
        output_panel.rowconfigure(0, weight=1)    # 縦方向の伸縮（はみ出し防止）

        # 列1: 操作の軌跡 (旧 設定)
        col_settings = tk.LabelFrame(output_panel, text="操作の軌跡", relief=tk.FLAT)
        col_settings.grid(row=0, column=0, sticky="nsew", padx=(2, 5))

        # セパレータ1
        ttk.Separator(output_panel, orient=tk.VERTICAL).grid(row=0, column=1, sticky="ns", padx=5)

        # 左側コンテナ（チェックボックス群）
        left_overlay = tk.Frame(col_settings, bg="#f5f5f5")
        left_overlay.pack(side=tk.LEFT, fill=tk.Y)

        self.chk_trajectory = tk.Checkbutton(left_overlay, text="マウスポインタを表示", variable=self.show_trajectory_var, command=self.update_canvas_image, bg="#f5f5f5")
        self.chk_trajectory.pack(anchor=tk.W)

        self.chk_sub = tk.Checkbutton(left_overlay, text="マウス・キー入力字幕を表示", variable=self.show_subtitle_var, command=self.update_canvas_image, bg="#f5f5f5")
        self.chk_sub.pack(anchor=tk.W)

        # 埋め込みチェックボックス
        self.chk_embed = tk.Checkbutton(left_overlay, text="字幕等を動画へ埋め込み", variable=self.embed_overlay_var, bg="#f5f5f5")
        self.chk_embed.pack(anchor=tk.W)

        # 右側: 詳細設定ボタン (高さは操作の軌跡いっぱいに)
        self.btn_overlay_settings = tk.Button(col_settings, text="⚙️ 詳細設定", command=self.open_overlay_settings, font=top_font, bg="#FFD700")
        self.btn_overlay_settings.pack(side=tk.RIGHT, padx=5, fill=tk.BOTH, expand=True, pady=2)


        # 列2: 静止画出力
        col_png = tk.LabelFrame(output_panel, text="静止画出力", relief=tk.FLAT)
        col_png.grid(row=0, column=2, sticky="nsew", padx=5)

        # セパレータ2
        ttk.Separator(output_panel, orient=tk.VERTICAL).grid(row=0, column=3, sticky="ns", padx=5)

        # 3分割用フレーム
        cp_split_frame = tk.Frame(col_png)
        cp_split_frame.pack(fill=tk.BOTH, expand=True)

        # 左端: クリップボードコピー
        left_f = tk.Frame(cp_split_frame)
        left_f.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.btn_copy_image = tk.Button(left_f, text="📷️ 表示内容をCOPY", command=self.copy_crop_to_clipboard, bg=self.theme.get("button_copy_bg"), cursor="hand2", font=top_font)
        self.btn_copy_image.pack(fill=tk.BOTH, expand=True, pady=2, padx=2)
        self.add_tooltip(self.btn_copy_image, "Ctrl+C: 現在のフレームをクリップボードにコピー")

        # 中央: 除外チェックボックス + 圧縮設定
        mid_f = tk.Frame(cp_split_frame)
        mid_f.pack(side=tk.LEFT, fill=tk.X, expand=False, padx=5)

        self.check_prev_next = tk.BooleanVar(value=self.global_config.get("check_prev_next", True))
        tk.Checkbutton(mid_f, text="前後不一致で除外", variable=self.check_prev_next).pack(anchor=tk.W)
        self.check_duplicate = tk.BooleanVar(value=self.global_config.get("check_duplicate", True))
        tk.Checkbutton(mid_f, text="直前重複で除外", variable=self.check_duplicate).pack(anchor=tk.W)

        # 圧縮設定
        compression_f = tk.Frame(mid_f)
        compression_f.pack(fill=tk.X, pady=(2, 0))
        tk.Label(compression_f, text="PNG圧縮:").pack(side=tk.LEFT)
        self.compression_var = tk.StringVar(value=str(self.png_compression))
        self.compression_spinbox = tk.Spinbox(compression_f, from_=0, to=9, increment=1, width=3, textvariable=self.compression_var, command=self.change_compression)
        self.compression_spinbox.pack(side=tk.LEFT, padx=3)

        # 右端: PNG保存ボタン (左端のコピーボタンと同じ比率で伸縮させる)
        right_f = tk.Frame(cp_split_frame)
        right_f.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.btn_export_png = tk.Button(right_f, text="🖼️ PNG連番保存", command=self.export_png, bg=self.theme.get("button_export_bg"), cursor="hand2", font=top_font)
        self.btn_export_png.pack(fill=tk.BOTH, expand=True, pady=2, padx=2)
        self.add_tooltip(self.btn_export_png, "指定範囲をPNG連番で保存")

        # 列3: 動画出力
        col_video = tk.LabelFrame(output_panel, text="動画出力", relief=tk.FLAT)
        col_video.grid(row=0, column=4, sticky="nsew", padx=(5, 2))
        
        # 保存ボタン容器 (高さを揃えるため fill=tk.Y。内部は expand=False にしてスリム化)
        video_btn_frame = tk.Frame(col_video)
        video_btn_frame.pack(fill=tk.BOTH, expand=True)

        self.btn_export_video = tk.Button(video_btn_frame, text="🎞️ MP4 動画保存", command=self.export_video, bg=self.theme.get("button_video_bg"), font=top_font)
        self.btn_export_video.pack(side=tk.TOP, fill=tk.X, expand=False, pady=1)
        self.add_tooltip(self.btn_export_video, "選択範囲をMP4動画としてエクスポート")
        
        self.btn_export_gif = tk.Button(video_btn_frame, text="🖼️ GIF アニメ保存", command=self.export_gif, bg=self.theme.get("button_gif_bg"), font=top_font)
        self.btn_export_gif.pack(side=tk.TOP, fill=tk.X, expand=False, pady=1)
        self.add_tooltip(self.btn_export_gif, "選択範囲をGIFアニメとしてエクスポート")

        # 全ボタンの activebackground を背景色に設定して色戻りを修正
        self._fix_all_button_active_colors(output_panel)
        self._fix_all_button_active_colors(main_ctrl)
        self._fix_all_button_active_colors(time_panel)
        self._fix_all_button_active_colors(crop_panel)
        self._fix_all_button_active_colors(top_panel)

        # 解像度プリセットをメニューに反映
        self.update_resolution_menu()

    # ------------------ ウィンドウ位置・サイズ管理 ------------------
    def _fix_all_button_active_colors(self, container):
        """コンテナ内の全ボタンの activebackground を背景色に合わせる."""
        fix_button_active_colors(container)

    def load_window_geometry(self):
        """設定ファイルからウィンドウの位置とサイズ、分割位置を読み込む"""
        config = load_global_config()
        
        # ウィンドウサイズと位置を個別に取得
        width = config.get("window_width", 1000)
        height = config.get("window_height", 700)
        x = config.get("window_x")
        y = config.get("window_y")
        
        # ジオメトリ文字列を構築（負の座標にも対応）
        if x is not None and y is not None:
            # 負の値の場合は自動的に-記号が付くので、+/-を適切に処理
            x_sign = '+' if x >= 0 else ''
            y_sign = '+' if y >= 0 else ''
            geometry = f"{width}x{height}{x_sign}{x}{y_sign}{y}"
        else:
            geometry = f"{width}x{height}"
        
        try:
            self.root.geometry(geometry)
            # ウィンドウの描画を確実にする
            self.root.update_idletasks()
        except Exception as e:
            print(f"ジオメトリ設定エラー: {e}, geometry={geometry}")
            self.root.geometry("1000x700")
        # 最大化状態
        if config.get("window_maximized", False):
            self.root.state('zoomed')
    
    def save_window_geometry(self):
        """ウィンドウの状態を保存"""
        config = load_global_config()

        # 最大化状態
        is_zoomed = (self.root.state() == 'zoomed')
        config["window_maximized"] = is_zoomed
        
        # ウィンドウの位置とサイズを個別に保存
        if not is_zoomed:
            # 最大化されていない場合のみ位置とサイズを保存
            config["window_x"] = self.root.winfo_x()
            config["window_y"] = self.root.winfo_y()
            config["window_width"] = self.root.winfo_width()
            config["window_height"] = self.root.winfo_height()
        
        # 下側パネルの高さはコード定数なので保存しない
        if "bottom_panel_height" in config:
            del config["bottom_panel_height"]

        # 分割位置は固定なので保存しない
        if "sash_position" in config:
            del config["sash_position"]

        save_global_config(config)

    def on_window_close(self):
        """ウィンドウ終了時の処理"""
        self.save_window_geometry()
        # メイン設定も保存
        self.save_config()
        self.root.destroy()

    # ------------------ 設定管理 ------------------
    def load_config(self):
        """設定ファイルから初期値を読み込む"""
        # アプリ共通設定の読み込み
        self.global_config = load_global_config()
        self.theme = self.global_config.get("theme", {})
        
        # プリセットの読み込み (config.pyから)
        self.resolution_presets = self.global_config.get("resolution_presets")
        if not self.resolution_presets:
            from config import get_default_presets_with_labels
            self.resolution_presets = get_default_presets_with_labels()

        # 各種変数の反映
        if hasattr(self, 'compression_var'):
            self.compression_var.set(str(self.global_config.get("png_compression", 3)))
            self.png_compression = int(self.compression_var.get())
        if hasattr(self, 'check_prev_next'):
            self.check_prev_next.set(self.global_config.get("check_prev_next", True))
        if hasattr(self, 'check_duplicate'):
            self.check_duplicate.set(self.global_config.get("check_duplicate", True))
        if hasattr(self, 'speed_var'):
            self.speed_var.set(f"{float(self.global_config.get('play_speed', 1.0)):.1f}")
            self.speed = float(self.speed_var.get())
        if hasattr(self, 'range_var'):
            self.range_var.set(self.global_config.get("play_range", False))
        if hasattr(self, 'loop_var'):
            self.loop_var.set(self.global_config.get("play_loop", False))
        if hasattr(self, 'pingpong_var'):
            self.pingpong_var.set(self.global_config.get("play_pingpong", False))
        if hasattr(self, 'show_trajectory_var'):
            self.show_trajectory_var.set(self.global_config.get("show_trajectory", True))
        if hasattr(self, 'show_subtitle_var'):
            self.show_subtitle_var.set(self.global_config.get("show_subtitle", True))
        if hasattr(self, 'embed_overlay_var'):
            self.embed_overlay_var.set(self.global_config.get("embed_overlay", False))
        
        # 現在開いている動画があれば再読み込み、なければ最後に開いた動画を読み込む
        if self.video_filepath and os.path.exists(self.video_filepath):
            # 再読み込みボタン用: 現在の動画を完全に再初期化
            self._load_video_internal(self.video_filepath)
        else:
            # 起動時: 最後に開いた動画があれば読み込む
            last_video = self.global_config.get("last_video_path", "")
            if last_video and os.path.exists(last_video):
                self._load_video_internal(last_video)
            else:
                self.update_resolution_menu()

    def _load_video_internal(self, video_file):
        """動画ファイルを読み込み、個別設定を反映させる"""
        if self.cap:
            self.cap.release()
        
        self.cap = cv2.VideoCapture(video_file)
        if not self.cap.isOpened():
            return False

        self.video_filename = os.path.splitext(os.path.basename(video_file))[0]
        self.video_filepath = os.path.abspath(video_file)
        self.entry_fullpath_var.set(self.video_filepath)
        
        self.fps = float(self.cap.get(cv2.CAP_PROP_FPS) or 30.0)
        frame_count = self.cap.get(cv2.CAP_PROP_FRAME_COUNT)
        self.duration = frame_count / self.fps if self.fps > 0 else 0

        # 動画情報を保存
        self.vid_w = int(self.cap.get(cv2.CAP_PROP_FRAME_WIDTH))
        self.vid_h = int(self.cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
        self.label_seconds.config(text=f"({self.duration:.3f}s, {self.vid_w}x{self.vid_h})")
        
        # FPSラベルを更新
        fps_display = f"{self.fps:.2f}".rstrip('0').rstrip('.')
        self.label_fps.config(text=f"{fps_display}FPS")

        # 動画個別設定の読み込み
        per_video_settings = load_video_settings(self.video_filepath)

        if per_video_settings:
            crop_rect_data = per_video_settings.get('crop_rect', {})
            self.start_time = float(per_video_settings.get('start_time', 0))
            self.end_time = float(per_video_settings.get('end_time', self.duration))
            self.end_time = float(per_video_settings.get('end_time', self.duration))
            # 現在の再生位置を復元（保存されていない場合は開始位置）
            self.current_time = float(per_video_settings.get('current_time', self.start_time))
            
            # 枠外選択設定の復元
            allow_oversize = per_video_settings.get('allow_oversize', False)
            if hasattr(self, 'allow_oversize_var'):
                self.allow_oversize_var.set(allow_oversize)
            
            if crop_rect_data:
                # 設定ファイルが古い（640x360基準）か新しい（ピクセル基準）かを簡易判定
                # 幅が640以下の場合は古い可能性があるが、ユーザーが小さいクロップを指定している場合と区別が難しい。
                # ここでは一度すべてピクセル基準として扱う（不整合が出る場合はユーザーに再設定してもらうのが安全）
                self.crop_rect = [
                    crop_rect_data.get("x1", self.vid_w // 4),
                    crop_rect_data.get("y1", self.vid_h // 4),
                    crop_rect_data.get("x2", self.vid_w * 3 // 4),
                    crop_rect_data.get("y2", self.vid_h * 3 // 4)
                ]
        else:
            # 個別設定がない場合はデフォルト値（中央付近）
            self.start_time = 0
            self.end_time = self.duration
            self.current_time = 0
            self.crop_rect = [self.vid_w//4, self.vid_h//4, self.vid_w*3//4, self.vid_h*3//4]
        self.update_resolution_menu()
        self.update_ui_texts()
        self.update_crop_entries()
        self.draw_seekbar()
        self.show_frame_at(self.current_time)
        self.load_trajectory()
        
        self.global_config["last_video_path"] = self.video_filepath
        save_global_config(self.global_config)
        return True

    def save_config(self):
        """アプリ共通設定を保存する (ウィンドウジオメトリ、最後に開いた動画、解像度プリセット等)"""
        config = load_global_config()
        
        # 最大化状態なら解除してジオメトリを取得（でないと前回の位置が取れない場合がある）
        is_zoomed = (self.root.state() == 'zoomed')
        config["window_maximized"] = is_zoomed
        
        # ウィンドウの位置とサイズを個別に保存
        if not is_zoomed:
            config["window_x"] = self.root.winfo_x()
            config["window_y"] = self.root.winfo_y()
            config["window_width"] = self.root.winfo_width()
            config["window_height"] = self.root.winfo_height()
        
        # 下側パネルの高さは定数化されたため保存不要。
        # 既存のキーがあれば削除しておく
        if "bottom_panel_height" in config:
            del config["bottom_panel_height"]
        if "sash_position" in config:
            del config["sash_position"]
            
        # 最後に開いた動画
        config["last_video_path"] = self.video_filepath if self.video_filepath else ""
        
        # 解像度プリセット
        config["resolution_presets"] = self.resolution_presets
        
        # 選択中の比率
        if hasattr(self, 'ratio_var'):
            config['selected_ratio'] = self.ratio_var.get()

        # 追加の設定項目
        config["png_compression"] = int(self.compression_var.get()) if hasattr(self, 'compression_var') else 3
        config["check_prev_next"] = self.check_prev_next.get() if hasattr(self, 'check_prev_next') else True
        config["check_duplicate"] = self.check_duplicate.get() if hasattr(self, 'check_duplicate') else True
        config["play_speed"] = float(self.speed_var.get()) if hasattr(self, 'speed_var') else 1.0
        config["play_range"] = self.range_var.get() if hasattr(self, 'range_var') else False
        config["play_loop"] = self.loop_var.get() if hasattr(self, 'loop_var') else False
        config["play_pingpong"] = self.pingpong_var.get() if hasattr(self, 'pingpong_var') else False
        config["show_trajectory"] = self.show_trajectory_var.get() if hasattr(self, 'show_trajectory_var') else True
        config["show_subtitle"] = self.show_subtitle_var.get() if hasattr(self, 'show_subtitle_var') else True
        config["embed_overlay"] = self.embed_overlay_var.get() if hasattr(self, 'embed_overlay_var') else False

        # アスペクト比オプション（UI で選べる固定リスト）
        config["aspect_options"] = self.aspect_options
        
        # 詳細設定ダイアログ等で変更された theme (mouse_overlay, input_overlay等) を保持
        if hasattr(self, 'global_config') and "theme" in self.global_config:
            config["theme"] = self.global_config["theme"]

        save_global_config(config)


    def save_video_settings(self):
        """当該動画ファイルに紐づく設定ファイルへ現在の設定（赤枠・時間等）を保存する"""
        if not self.video_filepath:
            messagebox.showwarning("Warning", "動画が読み込まれていません")
            return

        per_video_success = save_video_settings_to_file(
            self.video_filepath,
            self.crop_rect,
            self.start_time,
            self.end_time,
            self.current_time,
            allow_oversize=self.allow_oversize_var.get()
        )

        if per_video_success:
            messagebox.showinfo("Saved", f"動画個別の設定を保存しました:\n{per_video_success}")
        else:
            messagebox.showerror("Error", "個別設定の保存に失敗しました")

    def update_resolution_menu(self):
        """OptionMenu を現在の self.resolution_presets に合わせて更新する"""
        # (比率判定はクラスヘルパー `_ratio_value_from_str` / `_ratio_label_from_wh` を使用)
        menu = self.resolution_optionmenu["menu"]
        menu.delete(0, tk.END)
        # まずカスタム項目
        menu.add_command(label="カスタム", command=lambda v="カスタム": self.resolution_var.set(v))
        # 全画面項目
        menu.add_command(label="全画面", command=lambda v="全画面": (self.resolution_var.set(v), self.apply_resolution_preset(v)))

        # 比率フィルタが指定されていれば絞り込む
        selected_ratio = None
        try:
            selected_ratio = self.ratio_var.get()
        except Exception:
            selected_ratio = "未指定"

        target = self._ratio_value_from_str(selected_ratio) if selected_ratio and selected_ratio != '未指定' else None

        # build sortable list of (ratio_value, w, h, key_name)
        items = []
        for name, pair in self.resolution_presets.items():
            try:
                w, h = int(pair[0]), int(pair[1])
            except Exception:
                continue
            # determine ratio value: prefer prefix in name if present
            rval = None
            try:
                # name may start with e.g. '16:9 ' - try to parse
                first = name.split()[0]
                rval = self._ratio_value_from_str(first)
            except Exception:
                rval = None
            if rval is None:
                rval = float(w) / float(h) if h != 0 else 0.0
            # filter by selected ratio if needed
            if target is not None and abs(rval - target) > 0.03:
                continue
            items.append((rval, w, h, name))

        # sort by ratio, then width, then height
        items.sort(key=lambda t: (t[0], t[1], t[2]))
        for _, _, _, name in items:
            menu.add_command(label=name, command=lambda v=name: (self.resolution_var.set(v), self.apply_resolution_preset(v)))

    def add_resolution_preset(self):
        """現在の幅/高さをプリセットとして保存する（名前を入力）"""
        try:
            w = int(self.entry_crop_w.get())
            h = int(self.entry_crop_h.get())
        except Exception:
            messagebox.showerror("Err", "幅/高さが正しい整数ではありません")
            return

        name = simpledialog.askstring("プリセット名", "プリセット名を入力してください:")
        if not name:
            return
        # 名前の頭に比率ラベルを付与して保存
        try:
            lbl = self._ratio_label_from_wh(w, h)
        except Exception:
            lbl = f"{w}:{h}"
        full_name = f"{lbl} {name}"
        # avoid collision
        key = full_name
        idx = 1
        while key in self.resolution_presets:
            key = f"{full_name}#{idx}"
            idx += 1
        self.resolution_presets[key] = [int(w), int(h)]
        self.update_resolution_menu()
        self.save_config()

    def delete_resolution_preset(self):
        name = self.resolution_var.get()
        if name == "カスタム":
            messagebox.showinfo("Info", "削除できるプリセットを選択してください")
            return
        if name in self.resolution_presets:
            if messagebox.askyesno("確認", f"プリセット '{name}' を削除しますか？"):
                del self.resolution_presets[name]
                self.resolution_var.set("カスタム")
                self.update_resolution_menu()
                self.save_config()
        

    # ------------------ ヘルパー: 画像入出力 (utils モジュールに委譲) ------------------
    def imwrite_jp(self, filename, img, params=None):
        """日本語パス対応の画像保存関数（utils.imwrite_jp に委譲）"""
        return imwrite_jp(filename, img, params)


    def save_current_frame_as_png(self):
        """現在のフレームを指定された矩形で切り抜いて保存する (Ctrl+S用)"""
        if self.frame is None:
            return

        # クロップ矩形を動画の座標系に変換
        x1, y1, x2, y2 = self.crop_rect
        x1 = max(0, int(round(x1)))
        y1 = max(0, int(round(y1)))
        x2 = min(self.frame.shape[1], int(round(x2)))
        y2 = min(self.frame.shape[0], int(round(y2)))

        if x2 <= x1 or y2 <= y1:
            messagebox.showwarning("Warning", "有効なクロップ範囲が選択されていません。")
            return

        # 切り抜き
        cropped = self.frame[y1:y2, x1:x2]
        
        # 保存先決定 (動画と同じフォルダ)
        base_dir = os.path.dirname(self.video_filepath) if self.video_filepath else get_base_dir()
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        out_name = f"crop_{timestamp}.png"
        out_path = os.path.join(base_dir, out_name)

        # 保存
        try:
            if imwrite_jp(out_path, cropped):
                if messagebox.askyesno("Success", f"保存しました:\n{out_name}\n\n保存フォルダを開きますか？"):
                    open_folder_with_selection(out_path)
            else:
                messagebox.showerror("Error", "保存に失敗しました。")
        except Exception as e:
            messagebox.showerror("Error", f"例外が発生しました:\n{e}")

    def hhmmss_to_sec(self, time_str):
        result = hhmmss_to_sec(time_str)
        if result == 0.0 and time_str.strip() not in ("0", "0.0", "00:00:00", ""):
            # パースに失敗した場合のみエラーを表示
            try:
                float(time_str)
            except ValueError:
                messagebox.showerror("Err", f"時間指定が不正です: {time_str}")
        return result

    # ------------------ ロジック: 時間調整 ------------------
    def adjust_time(self, getter, setter, delta):
        if self.duration <= 0:
            return
        current_val = getter()
        # 秒数だけ増減する（フレーム内オフセットは保持）
        new_val = current_val + delta
        # clamp
        new_val = max(0, min(new_val, self.duration))
        setter(new_val)

    def manual_entry_update(self, entry_widget, setter):
        text = entry_widget.get()
        sec = self.hhmmss_to_sec(text)
        setter(sec)
        # フォーカスを外してUI更新を再開させる
        self.root.focus_set()

    def set_start_time_direct(self, val):
        val = max(0, min(val, self.duration))
        if val > self.end_time:
            val = self.end_time
        self.start_time = val
        self.current_time = val
        self.show_frame_at(self.current_time)
        self.update_ui_texts()
        try:
            # UIの更新を強制してから次の操作（再生）を受け付けやすくする
            # self.root.update_idletasks()
            pass
        except Exception:
            pass

    def set_end_time_direct(self, val):
        val = max(0, min(val, self.duration))
        if val < self.start_time:
            val = self.start_time
        self.end_time = val
        self.current_time = val
        self.show_frame_at(self.current_time)
        self.update_ui_texts()
        try:
            # UIの更新を強制してから次の操作（再生）を受け付けやすくする
            # self.root.update_idletasks()
            pass
        except Exception:
            pass

    def set_current_time_direct(self, val):
        val = max(0, min(val, self.duration))
        self.current_time = val
        self.show_frame_at(self.current_time)
        self.update_ui_texts()
        try:
            # UI の状態を即座に反映しておく
            # self.root.update_idletasks()
            pass
        except Exception:
            pass

    def go_to_video_start(self):
        # 動画の先頭へ移動
        self.set_current_time_direct(0)

    def go_to_video_end(self):
        # 動画の末尾へ移動
        self.set_current_time_direct(self.duration)

    def go_to_trim_start(self):
        # トリム開始位置へ移動
        self.set_current_time_direct(self.start_time)

    def go_to_trim_end(self):
        # トリム終了位置へ移動
        self.set_current_time_direct(self.end_time)

    def play_range(self):
        # 区間再生：開始位置へジャンプして再生開始、終了位置で停止
        if self.duration <= 0:
            return
        self.play_range_mode = True
        self.set_current_time_direct(self.start_time)
        if not self.playing:
            self.playing = True
            self._update_play_button_state()
            self.play_step()

    # ------------------ クロップの Undo/Redo ------------------
    def push_crop_history(self):
        # 現在の crop_rect を履歴に積む（重複は省く）
        cur = [int(v) for v in self.crop_rect]
        if not self.crop_history or self.crop_history[-1] != cur:
            self.crop_history.append(cur)
        # 新たな変更が入ったら redo 履歴はクリア
        try:
            self.crop_redo.clear()
        except Exception:
            self.crop_redo = []
        self.update_undo_button_state()

    def undo_crop(self, event=None):
        """Undo最後のクロップ矩形変更を元に戻す."""
        if not self.crop_history:
            return
        # 現在状態を redo に退避
        cur = [int(v) for v in self.crop_rect]
        self.crop_redo.append(cur)
        last = self.crop_history.pop()
        self.crop_rect = last
        # 矩形をキャンバスに反映し、角ハンドルも更新
        self._sync_crop_rect_ui()
        self.update_undo_button_state()

    def redo_crop(self) -> None:
        """Redo最後のクロップ矩形変更をやり直す."""
        if not getattr(self, 'crop_redo', None):
            return
        # 現在状態を undo 履歴に保存
        cur = [int(v) for v in self.crop_rect]
        self.crop_history.append(cur)
        nxt = self.crop_redo.pop()
        self.crop_rect = nxt
        # 矩形をキャンバスに反映し、角ハンドルも更新
        self._sync_crop_rect_ui()
        self.update_undo_button_state()

    def update_undo_button_state(self):
        try:
            if hasattr(self, 'btn_undo'):
                if self.crop_history:
                    self.btn_undo.config(state=tk.NORMAL)
                else:
                    self.btn_undo.config(state=tk.DISABLED)
            if hasattr(self, 'btn_redo'):
                if getattr(self, 'crop_redo', None):
                    self.btn_redo.config(state=tk.NORMAL)
                else:
                    self.btn_redo.config(state=tk.DISABLED)
        except Exception:
            pass

    # ------------------ ツールチップ ------------------
    def add_tooltip(self, widget, text):
        """フォーカスやホバー時にツールチップを表示する."""
        _add_tooltip(widget, text)

    def _scaled_rect_from_crop(self):
        # crop_rect (base coords) -> scaled coords on canvas
        x1, y1, x2, y2 = self.crop_rect
        return [int(x1 * self.canvas_scale_x) + self.canvas_offset_x,
                int(y1 * self.canvas_scale_y) + self.canvas_offset_y,
                int(x2 * self.canvas_scale_x) + self.canvas_offset_x,
                int(y2 * self.canvas_scale_y) + self.canvas_offset_y]

    def update_ui_texts(self):
        """テキストボックスの中身を現在の変数と同期する（フォーカスがない場合のみ）"""
        if self.duration <= 0:
            return

        # Start Entry
        if self.root.focus_get() != self.entry_start:
            self.entry_start.delete(0, tk.END)
            self.entry_start.insert(0, self.sec_to_display(self.start_time))

        # End Entry
        if self.root.focus_get() != self.entry_end:
            self.entry_end.delete(0, tk.END)
            self.entry_end.insert(0, self.sec_to_display(self.end_time))

        # Current Entry
        if self.root.focus_get() != self.entry_current:
            self.entry_current.delete(0, tk.END)
            self.entry_current.insert(0, self.sec_to_display(self.current_time))
        # シークバーを最新状態に更新
        try:
            self.draw_seekbar()
        except Exception:
            pass

    # ------------------ ヘルパー: 時間変換 (utils モジュールに委譲) ------------------
    def sec_to_display(self, sec):
        return sec_to_display(sec)

    def sec_to_hhmmss(self, sec):
        return sec_to_hhmmss(sec)

    def format_time(self, seconds: float) -> str:
        """秒数を時:分:秒の形式にフォーマットする."""
        return sec_to_display(seconds)

    # ------------------ 比率ヘルパー (utils モジュールに委譲) ------------------
    def _ratio_value_from_str(self, rstr):
        return ratio_value_from_str(rstr)

    def _ratio_label_from_wh(self, w, h):
        return ratio_label_from_wh(w, h)

    def update_crop_from_entries(self):
        """クロップサイズ入力フィールドから矩形を更新"""
        try:
            x1 = int(self.entry_crop_x.get())
            y1 = int(self.entry_crop_y.get())
            w = int(self.entry_crop_w.get())
            h = int(self.entry_crop_h.get())

            # 入力による変更は undo 履歴に保存
            try:
                self.push_crop_history()
            except Exception:
                pass

            x2 = x1 + w
            y2 = y1 + h
            self.crop_rect = self.clamp_rect_canvas([x1, y1, x2, y2])
            self.update_canvas_image()
        except ValueError:
            pass

    # ------------------ クリップボード (clipboard モジュールに委譲) ------------------
    def _copy_image_to_clipboard(self, pil_img):
        """PIL画像をクリップボードにコピーする（clipboard.copy_image_to_clipboard に委譲）"""
        return copy_image_to_clipboard(pil_img)

    # ------------------ 矢印キーのリピート処理 ------------------
    def _on_arrow_press(self, event, direction):
        # direction: -1 left, +1 right
        # avoid starting multiple
        if self._arrow_repeat_id is not None:
            return
        self._arrow_dir = direction
        self._arrow_start_time = time.time()
        # perform one immediate step
        self._arrow_step(event)
        # schedule repeat
        self._arrow_repeat_id = self.root.after(200, lambda: self._arrow_repeat(event))

    def _on_arrow_release(self, event):
        try:
            if self._arrow_repeat_id:
                self.root.after_cancel(self._arrow_repeat_id)
        except Exception:
            pass
        self._arrow_repeat_id = None
        self._arrow_dir = None
        self._arrow_start_time = None

    def _arrow_repeat(self, event):
        self._arrow_step(event)
        # accelerate: shorter delay over time
        elapsed = time.time() - (self._arrow_start_time or time.time())
        delay = 200
        if elapsed > 2.0:
            delay = 50
        elif elapsed > 1.0:
            delay = 100
        self._arrow_repeat_id = self.root.after(delay, lambda: self._arrow_repeat(event))

    def _arrow_step(self, event):
        # detect ctrl held
        ctrl = False
        try:
            ctrl = (event.state & 0x4) != 0
        except Exception:
            ctrl = False
        step = 1.0 if ctrl else 0.1
        # detect alt for crop movement when rectangle is focused
        try:
            # Alt is often 0x20000 on Windows or 0x8 on some systems
            alt_held = (event.state & (0x20000 | 0x8)) != 0
        except Exception:
            alt_held = False
        if alt_held and getattr(self, 'rect_focused', False) and self._arrow_dir in (-1, 1):
            # move crop horizontally by 1px or 10px per step
            base_dx = -1 if self._arrow_dir == -1 else 1
            dx = base_dx * 10 if ctrl else base_dx
            self.move_crop_by(dx, 0)
            return
        if self._arrow_dir == -1:
            self.set_current_time_direct(max(0, self.current_time - step))
        elif self._arrow_dir == 1:
            self.set_current_time_direct(min(self.duration, self.current_time + step))

    def copy_crop_to_clipboard(self, event=None):
        if not self.cap:
            return
        # ensure we have current frame
        if self.frame is None:
            self.show_frame_at(self.current_time)
        if self.frame is None:
            return
        # compute video pixel crop
        vx1, vy1, vx2, vy2 = [int(v) for v in self.crop_rect]
        vid_w = int(self.cap.get(cv2.CAP_PROP_FRAME_WIDTH))
        vid_h = int(self.cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
        
        # 安全のためクランプ
        vx1 = max(0, min(vid_w, vx1))
        vy1 = max(0, min(vid_h, vy1))
        vx2 = max(0, min(vid_w, vx2))
        vy2 = max(0, min(vid_h, vy2))
        ret, frm = True, self.frame
        try:
            crop = frm[vy1:vy2, vx1:vx2]
            if crop.size == 0:
                return
            rgb = cv2.cvtColor(crop, cv2.COLOR_BGR2RGB)
            pil = Image.fromarray(rgb)
            
            # 操作履歴のオーバーレイを描画
            history_manager = overlay_utils.InputHistoryManager()
            self._draw_overlay_on_image(pil, self.current_time, history_manager, 0, vx1, vy1)

            ok = self._copy_image_to_clipboard(pil)
            if ok:
                # 視覚的なフィードバック（フラッシュ＆メッセージ）
                self._show_copy_feedback()
            else:
                messagebox.showerror('Error', 'クリップボードへ画像をコピーできませんでした')
        except Exception as e:
            messagebox.showerror('Error', f'コピーに失敗しました:\n{e}')

    def _show_copy_feedback(self):
        """コピー時のフラッシュ効果とメッセージ表示を行う（改善版）."""
        # 1. クロップ範囲のキャンバス座標を計算
        x1, y1, x2, y2 = self.crop_rect
        cx1 = x1 * self.canvas_scale_x + self.canvas_offset_x
        cy1 = y1 * self.canvas_scale_y + self.canvas_offset_y
        cx2 = x2 * self.canvas_scale_x + self.canvas_offset_x
        cy2 = y2 * self.canvas_scale_y + self.canvas_offset_y

        # 2. フラッシュ（白い矩形）- 時間を 200ms に延長
        flash_id = self.canvas.create_rectangle(
            cx1, cy1, cx2, cy2, fill="white", outline="", stipple="gray50" if sys.platform != "win32" else ""
        )
        self.root.after(200, lambda: self.canvas.delete(flash_id))
        
        # 3. メッセージ表示
        mx = (cx1 + cx2) / 2
        my = (cy1 + cy2) / 2
        
        # モダンなフォント設定
        font_main = ("Segoe UI", 24, "bold")
        font_outline = ("Segoe UI", 24, "bold")
        
        # 縁取りテキスト (黒)
        shadows = []
        for dx, dy in [(-2,-2), (2,-2), (-2,2), (2,2), (0,-2), (0,2), (-2,0), (2,0)]:
            s_id = self.canvas.create_text(
                mx + dx, my + dy, text="コピーしたよ", fill="black", font=font_outline
            )
            shadows.append(s_id)
            
        # メインテキスト (白)
        text_id = self.canvas.create_text(
            mx, my, text="コピーしたよ", fill="white", font=font_main
        )

        # 4. 滑らかなフェードアウト (0.5秒)
        # Tkinterの色を段階的に透明（背景色）に近づける
        fade_steps = 10
        fade_duration = 500 # ms
        
        def get_fade_color(step: int, start_rgb: tuple, end_rgb: tuple) -> str:
            r = int(start_rgb[0] + (end_rgb[0] - start_rgb[0]) * (step / fade_steps))
            g = int(start_rgb[1] + (end_rgb[1] - start_rgb[1]) * (step / fade_steps))
            b = int(start_rgb[2] + (end_rgb[2] - start_rgb[2]) * (step / fade_steps))
            return f"#{r:02x}{g:02x}{b:02x}"

        # キャンバス背景色を取得（色の収束先）
        bg_hex = self.theme.get("canvas_bg", "#f5f5f5")
        bg_rgb = tuple(int(bg_hex.lstrip('#')[i:i+2], 16) for i in (0, 2, 4))
        
        white_rgb = (255, 255, 255)
        black_rgb = (0, 0, 0)

        def animate_fade(step=0):
            if step > fade_steps:
                self.canvas.delete(text_id)
                for s in shadows: self.canvas.delete(s)
                return
            
            # 文字の色を背景色に近づける
            new_text_color = get_fade_color(step, white_rgb, bg_rgb)
            new_shadow_color = get_fade_color(step, black_rgb, bg_rgb)
            
            self.canvas.itemconfig(text_id, fill=new_text_color)
            for s in shadows:
                self.canvas.itemconfig(s, fill=new_shadow_color)
            
            self.root.after(int(fade_duration / fade_steps), lambda: animate_fade(step + 1))

        # 少し待ってからフェード開始
        self.root.after(100, animate_fade)

    def show_shortcuts(self):
        # ショートカット一覧を表示（既に開いていれば再利用して前面に移動）
        try:
            if getattr(self, '_shortcuts_win', None):
                try:
                    self._shortcuts_win.deiconify()
                    self._shortcuts_win.lift()
                    return
                except Exception:
                    try:
                        self._shortcuts_win.destroy()
                    except Exception:
                        pass
                    self._shortcuts_win = None
        except Exception:
            pass

        # EXE化時のリソースパス解決用ヘルパー
        def resource_path(relative_path):
            import sys
            if hasattr(sys, '_MEIPASS'):
                return os.path.join(sys._MEIPASS, relative_path)
            return os.path.join(os.path.abspath("."), relative_path)

        # リソースパスからショートカット一覧を読み込む
        path = resource_path('README_shortcuts.md')
        text = ''
        if os.path.exists(path):
            try:
                with open(path, 'r', encoding='utf-8') as f:
                    text = f.read()
            except Exception:
                text = ''
        if not text:
            text = 'Ctrl+Z: 戻す\nCtrl+C: クロップをクリップボードへコピー\nSpace: 再生/停止\nHome/End: トリム範囲の先頭/末尾にジャンプ\nCtrl+Home/End: 動画の先頭/末尾にジャンプ\nAlt+矢印: 1px移動 / Alt+Ctrl+矢印: 10px移動\nShift+矢印: 1px拡縮 / Shift+Ctrl+矢印: 10px拡縮\nShift+左ドラッグ: 垂直/水平移動に限定'

        # show in simple Toplevel with scrollable Text
        top = tk.Toplevel(self.root)
        top.title('ショートカット一覧')
        self._shortcuts_win = top

        def _on_close():
            try:
                self._shortcuts_win.destroy()
            except Exception:
                pass
            self._shortcuts_win = None

        top.protocol('WM_DELETE_WINDOW', _on_close)
        txt = tk.Text(top, wrap='word', width=60, height=15)
        txt.insert('1.0', text)
        txt.config(state=tk.DISABLED)
        txt.pack(fill=tk.BOTH, expand=True)
        tk.Button(top, text='閉じる', command=_on_close).pack(pady=4)

        # 親ウィンドウの中央に表示
        top.update_idletasks()
        w = top.winfo_width()
        h = top.winfo_height()
        rx = self.root.winfo_rootx()
        ry = self.root.winfo_rooty()
        rw = self.root.winfo_width()
        rh = self.root.winfo_height()
        tx = rx + (rw // 2) - (w // 2)
        ty = ry + (rh // 2) - (h // 2)
        top.geometry(f"+{tx}+{ty}")

    def open_overlay_settings(self):
        """詳細設定ダイアログを開く"""
        try:
            from overlay_settings_dialog import InputOverlaySettingsDialog
            InputOverlaySettingsDialog(self, self.global_config, self.on_overlay_settings_update)
        except ImportError:
            messagebox.showinfo("Info", "詳細設定ダイアログの実装はまだです")

    def on_overlay_settings_update(self):
        """詳細設定ダイアログからの更新を受け取る"""
        # テーマ設定を更新
        self.theme = self.global_config.get("theme", {})
        
        # テーマ依存のパラメータを再設定
        self.HANDLE_SIZE = self.theme.get("handle_size", 8)
        self.EDGE_MARGIN = self.theme.get("edge_margin", 20)
        
        # キャンバス再描画
        self.update_canvas_image()

    def update_crop_entries(self):
        """クロップ矩形からサイズ入力フィールドを更新"""
        x1, y1, x2, y2 = self.crop_rect
        w = x2 - x1
        h = y2 - y1

        # フォーカスがないときのみ更新
        if self.root.focus_get() != self.entry_crop_x:
            self.entry_crop_x.delete(0, tk.END)
            self.entry_crop_x.insert(0, str(int(x1)))
        if self.root.focus_get() != self.entry_crop_y:
            self.entry_crop_y.delete(0, tk.END)
            self.entry_crop_y.insert(0, str(int(y1)))
        if self.root.focus_get() != self.entry_crop_w:
            self.entry_crop_w.delete(0, tk.END)
            self.entry_crop_w.insert(0, str(int(w)))
        if self.root.focus_get() != self.entry_crop_h:
            self.entry_crop_h.delete(0, tk.END)
            self.entry_crop_h.insert(0, str(int(h)))

    def toggle_resolution_lock(self):
        """解像度設定とマウスによるリサイズをロック/解除する"""
        is_locked = not self.lock_var.get()
        self.lock_var.set(is_locked)
        self.btn_lock_res.config(text="🔒" if is_locked else "🔓", 
                                 bg=self.theme.get("button_locked_bg") if is_locked else self.theme.get("button_unlocked_bg"))
        
        state = tk.DISABLED if is_locked else tk.NORMAL
        self.ratio_optionmenu.config(state=state)
        self.resolution_optionmenu.config(state=state)
        self.entry_crop_w.config(state=state)
        self.entry_crop_h.config(state=state)
        self.btn_undo.config(state=state if self.crop_history else tk.DISABLED)
        self.btn_redo.config(state=state if self.crop_redo else tk.DISABLED)
        self.btn_save_preset.config(state=state)
        self.btn_delete_preset.config(state=state)
        
        # ハンドルの表示・非表示を即座に反映
        self.update_canvas_image()

    def toggle_move_lock(self):
        """座標設定とマウスによる移動をロック/解除する"""
        is_locked = not self.lock_move_var.get()
        self.lock_move_var.set(is_locked)
        self.btn_lock_move.config(text="🔒" if is_locked else "🔓", 
                                  bg=self.theme.get("button_locked_bg") if is_locked else self.theme.get("button_unlocked_bg"))
        
        state = tk.DISABLED if is_locked else tk.NORMAL
        self.entry_crop_x.config(state=state)
        self.entry_crop_y.config(state=state)

    def move_crop_by(self, dx, dy):
        """クロップ矩形をdx,dyだけ移動（ピクセル単位）。Alt+矢印用。"""
        if self.lock_move_var.get():
            return
        try:
            # undo 履歴に現在の矩形を登録
            try:
                self.push_crop_history()
            except Exception:
                pass
            x1, y1, x2, y2 = self.crop_rect
            nx1 = x1 + dx
            ny1 = y1 + dy
            nx2 = x2 + dx
            ny2 = y2 + dy
            self.crop_rect = self.clamp_rect_canvas([nx1, ny1, nx2, ny2])
            self._sync_crop_rect_ui()
        except Exception:
            pass

    def expand_crop(self, dx, dy):
        """クロップ矩形を中心を維持して拡大・縮小（ピクセル単位）。Shift+矢印用。
        dx=1: 幅+2, dx=-1: 幅-2
        dy=1: 高+2, dy=-1: 高-2
        """
        if self.lock_var.get():
            return
        try:
            self.push_crop_history()
            x1, y1, x2, y2 = self.crop_rect
            nx1 = x1 - dx
            ny1 = y1 - dy
            nx2 = x2 + dx
            ny2 = y2 + dy

            # 最小サイズチェック
            if (nx2 - nx1) < self.MIN_W:
                nx1, nx2 = x1, x2
            if (ny2 - ny1) < self.MIN_H:
                ny1, ny2 = y1, y2

            self.crop_rect = self.clamp_rect_canvas([nx1, ny1, nx2, ny2])
            self._sync_crop_rect_ui()
        except Exception:
            pass

    def apply_resolution_preset(self, preset_name):
        """プリセット解像度を適用（self.resolution_presets を参照）"""
        if self.lock_var.get():
            return

        # 全画面対応
        if preset_name == "全画面":
            w = getattr(self, "vid_w", self.CANVAS_W)
            h = getattr(self, "vid_h", self.CANVAS_H)
            # 変更は undo 履歴に保存
            try:
                self.push_crop_history()
            except Exception:
                pass
            
            # クロップ矩形を全画面に設定
            self.crop_rect = [0, 0, w, h]
            
            # 入力欄も更新
            self.entry_crop_x.delete(0, tk.END)
            self.entry_crop_x.insert(0, "0")
            self.entry_crop_y.delete(0, tk.END)
            self.entry_crop_y.insert(0, "0")
            self.entry_crop_w.delete(0, tk.END)
            self.entry_crop_w.insert(0, str(w))
            self.entry_crop_h.delete(0, tk.END)
            self.entry_crop_h.insert(0, str(h))
            
            self._sync_crop_rect_ui()
            return

        presets = self.resolution_presets or {}
        if preset_name in presets:
            pair = presets[preset_name]
            if isinstance(pair, (list, tuple)) and len(pair) >= 2:
                w, h = pair[0], pair[1]
            elif isinstance(pair, int):
                w, h = pair, pair
            else:
                return
            # 変更は undo 履歴に保存
            try:
                self.push_crop_history()
            except Exception:
                pass
            self.entry_crop_w.delete(0, tk.END)
            self.entry_crop_w.insert(0, str(w))
            self.entry_crop_h.delete(0, tk.END)
            self.entry_crop_h.insert(0, str(h))
            self.update_crop_from_entries()

    def update_crop_from_entries(self):
        # ロック状態の取得
        size_locked = self.lock_var.get()
        move_locked = self.lock_move_var.get()
        
        if size_locked and move_locked:
            return
            
        # ドラッグ中・リサイズ中はEntryからの更新を無視する（ループ防止）
        if getattr(self, 'dragging_rect', False) or getattr(self, 'resizing_rect', False):
            return

        try:
            cur_x, cur_y, cur_x2, cur_y2 = self.crop_rect
            cur_w = cur_x2 - cur_x
            cur_h = cur_y2 - cur_y
            
            # 入力値を取得
            new_x = int(self.entry_crop_x.get()) if not move_locked else cur_x
            new_y = int(self.entry_crop_y.get()) if not move_locked else cur_y
            new_w = int(self.entry_crop_w.get()) if not size_locked else cur_w
            new_h = int(self.entry_crop_h.get()) if not size_locked else cur_h
            
            if (new_x == cur_x and new_y == cur_y and 
                new_w == cur_w and new_h == cur_h):
                return

            self.push_crop_history()
            self.crop_rect = self.clamp_rect_resize(new_x, new_y, new_x + new_w, new_y + new_h)
            self._sync_crop_rect_ui()
        except Exception:
            pass

    # ------------------ 動画再生ロジック ------------------
    def change_speed(self, event=None):
        try:
            v = self.speed_spinbox.get()
            val = float(v)
            # clamp range to -9.9 .. 9.9
            val = max(-9.9, min(9.9, val))
            self.speed = val
            self.speed_var.set(f"{self.speed:.1f}")
        except Exception as e:
            messagebox.showerror("speed Error", f"無効な速度の数値が入力されたので、1.0倍に戻しました。\n\n{str(e)}")
            self.speed = 1.0
            self.speed_var.set("1.0")  # UI上の表示も戻す
        finally:
            try:
                # Enterで確定したらフォーカスを外す
                self.root.focus_set()
            except Exception:
                pass

    def change_compression(self, event=None):
        try:
            val = int(self.compression_spinbox.get())
            # 0〜9の範囲に収める
            self.png_compression = max(0, min(9, val))
        except Exception as e:
            print(f"Compression Error: {e}")
            self.png_compression = 3
            self.compression_var.set("3")
        finally:
            try:
                self.root.focus_set()
            except Exception:
                pass

    def on_speed_wheel(self, event):
        """Spinbox をマウスホイールで増減させる。Windows は <MouseWheel>、X11 は Button-4/5。"""
        try:
            cur = float(self.speed_var.get())
        except Exception:
            cur = 1.0

        delta = 0.1
        # Windows: event.delta の符号で判定
        if hasattr(event, 'delta') and event.delta:
            if event.delta > 0:
                cur += delta
            else:
                cur -= delta
        else:
            # X11: Button-4=up, Button-5=down
            if str(event.num) == '4':
                cur += delta
            elif str(event.num) == '5':
                cur -= delta

        # Clamp range
        cur = max(-9.9, min(9.9, cur))
        self.speed_var.set(f"{cur:.1f}")
        try:
            self.change_speed()
        except Exception:
            pass

    def open_CapchunScreen(self):
        """録画ツールを起動"""
        app = CapchunScreen.CapchunScreenApp(self.root, parent_app=self)

    def open_video_file(self, result_path):
        """外部から指定された動画ファイルを開く"""
        if not os.path.exists(result_path):
            return
        self.load_video(target_path=result_path)

    def load_video(self, target_path=None):
        """動画ファイルを選択して読み込む"""
        file_path = target_path
        if not file_path:
            file_path = filedialog.askopenfilename(
                filetypes=[("Video files", "*.mp4 *.avi *.mov *.mkv *.wmv"), ("All files", "*.*")])
        if file_path:
            self._load_video_internal(file_path)
        if self.playing:
            self.toggle_play()


    def toggle_play(self):
        self.playing = not self.playing
        self._update_play_button_state()
        if self.playing:
            # 区間再生チェックが有効なら再生前に位置を範囲内に移動
            try:
                rng = getattr(self, 'range_var', None) is not None and self.range_var.get()
            except Exception:
                rng = False
            self.play_range_mode = bool(rng)
            if self.play_range_mode:
                # 現在位置が start..end の範囲内でなければ start に移動
                if not (self.start_time <= self.current_time <= self.end_time):
                    self.set_current_time_direct(self.start_time)
            try:
                # ensure UI state is flushed before starting playback
                # self.root.update_idletasks()
                pass
            except Exception:
                pass
            # If current position is at an unplayable end, restart from beginning
            try:
                loop_on = getattr(self, 'loop_var', None) is not None and self.loop_var.get()
            except Exception:
                loop_on = False
            try:
                pingpong_on = getattr(self, 'pingpong_var', None) is not None and self.pingpong_var.get()
            except Exception:
                pingpong_on = False
            # determine effective end position
            end_pos = self.end_time if self.play_range_mode else self.duration
            start_pos = self.start_time if self.play_range_mode else 0.0
            if self.current_time >= (end_pos - 1e-6) and not loop_on and not pingpong_on:
                self.set_current_time_direct(start_pos)
            self.play_step()
        else:
            # clear range mode if user stops playback manually
            self.play_range_mode = False
            if self._play_after_id:
                self.root.after_cancel(self._play_after_id)
                self._play_after_id = None

    def _update_play_button_state(self):
        """再生/停止ボタンの表示テキストと背景色を更新する"""
        if self.playing:
            text = "■ 停止"
            bg = self.theme.get("button_stop_bg", "#EF9A9A")
        else:
            text = "▲ 再生"
            bg = self.theme.get("button_play_bg", "#A5D6A7")
        self.btn_play.config(text=text, bg=bg)

    def play_step(self):
        if not (self.cap and self.playing):
            return

        self.cap.set(cv2.CAP_PROP_POS_MSEC, int(self.current_time * 1000))
        ret, frame = self.cap.read()
        if ret and frame is not None:
            self.frame = frame
            self.update_canvas_image()

        # time advance (supports negative speed for reverse)
        self.current_time += (1.0 / self.fps) * self.speed

        # Determine stop / loop behavior
        pingpong_on = getattr(self, 'pingpong_var', None) is not None and self.pingpong_var.get()
        if self.speed >= 0:
            # 区間再生モードなら end_time で停止
            if getattr(self, 'play_range_mode', False):
                # 区間再生かつループ有効なら start に戻して継続、そうでなければ停止
                loop_on = getattr(self, 'loop_var', None) is not None and self.loop_var.get()
                if self.current_time >= self.end_time:
                    if loop_on:
                        if pingpong_on:
                            # clamp and reverse direction
                            self.current_time = float(self.end_time)
                            self.speed = -abs(self.speed)
                            try:
                                self.speed_var.set(f"{self.speed:.1f}")
                            except Exception:
                                pass
                        else:
                            self.current_time = float(self.start_time)
                    else:
                        self.current_time = float(self.end_time)
                        self.playing = False
                        self.play_range_mode = False
                        self.btn_play.config(text="▲ 再生 (Space)")
                        self._update_play_button_state()
            else:
                # 通常再生: 末尾到達時はループ設定で先頭へ戻す
                loop_on = getattr(self, 'loop_var', None) is not None and self.loop_var.get()
                if self.current_time >= self.duration:
                    if loop_on:
                        if pingpong_on:
                            self.current_time = float(self.duration)
                            self.speed = -abs(self.speed)
                            try:
                                self.speed_var.set(f"{self.speed:.1f}")
                            except Exception:
                                pass
                        else:
                            self.current_time = 0.0
                    else:
                        self.current_time = float(self.duration)
                        self.playing = False
                        self.btn_play.config(text="▲ 再生 (Space)")
                        self._update_play_button_state()
        else:
            # 逆再生時
            if getattr(self, 'play_range_mode', False):
                loop_on = getattr(self, 'loop_var', None) is not None and self.loop_var.get()
                if self.current_time <= self.start_time:
                    if loop_on:
                        if pingpong_on:
                            self.current_time = float(self.start_time)
                            self.speed = abs(self.speed)
                            try:
                                self.speed_var.set(f"{self.speed:.1f}")
                            except Exception:
                                pass
                        else:
                            self.current_time = float(self.end_time)
                    else:
                        self.current_time = float(self.start_time)
                        self.playing = False
                        self.play_range_mode = False
                        self.btn_play.config(text="▲ 再生 (Space)")
                        self._update_play_button_state()
            else:
                loop_on = getattr(self, 'loop_var', None) is not None and self.loop_var.get()
                if self.current_time <= 0:
                    if loop_on:
                        if pingpong_on:
                            self.current_time = 0.0
                            self.speed = abs(self.speed)
                            try:
                                self.speed_var.set(f"{self.speed:.1f}")
                            except Exception:
                                pass
                        else:
                            self.current_time = float(self.duration)
                    else:
                        self.current_time = 0.0
                        self.playing = False
                        self.btn_play.config(text="▲ 再生 (Space)")
                        self._update_play_button_state()

        self.update_ui_texts()  # This updates texts and seekbar

        if self.playing:
            # use absolute speed for timer interval
            delay = max(5, int(1000.0 / (self.fps * max(0.01, abs(self.speed)))))
            self._play_after_id = self.root.after(delay, self.play_step)

    def show_frame_at(self, sec):
        if not self.cap:
            return
        sec = max(0, min(self.duration, sec))
        frame_no = int(sec * self.fps)
        self.cap.set(cv2.CAP_PROP_POS_FRAMES, frame_no)
        ret, frame = self.cap.read()
        if ret and frame is not None:
            self.frame = frame
            self.update_canvas_image()

    def update_canvas_image(self):
        if self.frame is None:
            return
        rgb = cv2.cvtColor(self.frame, cv2.COLOR_BGR2RGB)

        # キャンバスの現在のサイズを取得
        canvas_w = self.canvas.winfo_width()
        canvas_h = self.canvas.winfo_height()

        # キャンバスのサイズが確定していない場合はデフォルトサイズを使用
        if canvas_w <= 1:
            canvas_w = self.CANVAS_W
        if canvas_h <= 1:
            canvas_h = self.CANVAS_H
        # 最低キャンバス高さを保証（計算の安定用）
        canvas_h_effective = max(canvas_h, getattr(self, 'CANVAS_MIN_H', canvas_h))

        # 元のフレームのアスペクト比を取得
        frame_h, frame_w = rgb.shape[:2]
        frame_aspect = frame_w / frame_h
        canvas_aspect = canvas_w / canvas_h

        # アスペクト比を維持しながらリサイズサイズを計算
        if frame_aspect > canvas_aspect:
            # フレームが横長：幅に合わせる
            resized_w = canvas_w
            resized_h = int(canvas_w / frame_aspect)
        else:
            # フレームが縦長：高さに合わせる
            resized_h = canvas_h_effective
            resized_w = int(canvas_h_effective * frame_aspect)
        # ズームを考慮して最終的な表示サイズを決定
        zoom = getattr(self, 'image_zoom', 1.0)
        rw = max(1, int(resized_w * zoom))
        rh = max(1, int(resized_h * zoom))

        # リサイズ済みフレームを作成（ズーム後のサイズ）
        img = Image.fromarray(rgb).resize((rw, rh), Image.Resampling.LANCZOS)
        
        # マウス軌跡のオーバーレイ描画
        self.update_canvas_overlay(img)

        self.tk_img = ImageTk.PhotoImage(img)

        # パンオフセットを加味してキャンバス内に配置
        offset_x = (canvas_w - rw) // 2 + getattr(self, 'pan_offset_x', 0)
        offset_y = (canvas_h - rh) // 2 + getattr(self, 'pan_offset_y', 0)
        self.canvas_offset_x = offset_x
        self.canvas_offset_y = offset_y
        self.canvas.itemconfig(self.canvas_image, image=self.tk_img)
        self.canvas.coords(self.canvas_image, offset_x, offset_y)

        # スケール比を計算（元の動画座標系から表示上への変換）
        self.canvas_scale_x = rw / self.vid_w if self.vid_w > 0 else 1.0
        self.canvas_scale_y = rh / self.vid_h if self.vid_h > 0 else 1.0

        # 表示倍率（ズーム率）を更新
        if hasattr(self, 'label_zoom'):
            zoom_pct = int(round(self.canvas_scale_x * 100))
            self.label_zoom.config(text=f"{zoom_pct}%")

        # crop_rectをスケールして描画
        scaled_rect = [
            int(self.crop_rect[0] * self.canvas_scale_x) + offset_x,
            int(self.crop_rect[1] * self.canvas_scale_y) + offset_y,
            int(self.crop_rect[2] * self.canvas_scale_x) + offset_x,
            int(self.crop_rect[3] * self.canvas_scale_y) + offset_y
        ]
        
        # 矩形の外見(色、太さ、線種)を更新
        try:
            # テーマ設定の取得
            base_width = self.theme.get("crop_width")
            
            # 状態の判定
            is_active = getattr(self, 'rect_focused', False)
            is_dragging_or_resizing = self.dragging_rect or self.resizing_rect
            
            # 動画ピクセルからキャンバス表示ピクセルへのトータルスケールを計算
            total_scale = (rw / frame_w + rh / frame_h) / 2.0 if frame_w > 0 and frame_h > 0 else 1.0
            
            # 優先順位: 1.Focused (選択中/操作中) 2.Hover 3.Default
            if is_active or is_dragging_or_resizing:
                # 選択中または操作中 (既定: 赤色の実線)
                width = max(1, int(round((base_width + 1) * total_scale)))
                linecolor = self.theme.get("crop_focused_linecolor")
                linestyle = self.theme.get("crop_focused_linestyle")
            elif self.rect_hovered:
                # ホバー中（操作していない時、かつ未選択時。既定: 赤色の破線）
                width = max(1, int(round((base_width + 1) * total_scale)))
                linecolor = self.theme.get("crop_hover_linecolor")
                linestyle = self.theme.get("crop_hover_linestyle")
            else:
                # デフォルト（未選択・非ホバー。既定: 橙色の破線）
                width = max(1, int(round(base_width * total_scale)))
                linecolor = self.theme.get("crop_default_linecolor")
                linestyle = self.theme.get("crop_default_linestyle")

            # linestyle が空文字列やNoneの場合は実線、タプルやリストの場合は破線にする
            if not linestyle or linestyle == "":
                linestyle_arg = ""
            else:
                # tkinter の破線引数として適切な形式に変換
                linestyle_arg = tuple(linestyle) if isinstance(linestyle, (list, tuple)) else linestyle

            # 描画位置の調整（枠線を外側に広げる）
            # 太さが太くなっても内側（映像エリア）には被らないように、
            # 枠線の中心線を外側にずらす。
            # offset = 1 (1px gap) + width / 2 (center of the line)
            offset = 1.0 + width / 2.0
            draw_rect = [
                scaled_rect[0] - offset,
                scaled_rect[1] - offset,
                scaled_rect[2] + offset,
                scaled_rect[3] + offset
            ]

            self.canvas.coords(self.rect_id, *draw_rect)
            self.canvas.itemconfig(self.rect_id, width=width, outline=linecolor, dash=linestyle_arg)
        except Exception:
            pass

        # 角マーカー（ハンドル）を描画/更新する
        # ハンドルは論理的な角（scaled_rect）に合わせて配置する
        self._update_corner_handles(self._get_corner_coords(scaled_rect))

    # ------------------ キャンバスリサイズ処理 ------------------
    def on_canvas_resize(self, event=None):
        """キャンバスのリサイズイベントを処理"""
        cw = event.width if event else self.canvas.winfo_width()
        ch = event.height if event else self.canvas.winfo_height()
        
        if cw > 1 and ch > 1:
            # スケール比を計算
            h_eff = max(ch, getattr(self, 'CANVAS_MIN_H', ch))
            self.canvas_scale_x = cw / self.CANVAS_W
            self.canvas_scale_y = h_eff / self.CANVAS_H
            # フレームを再描画
            if self.frame is not None:
                self.update_canvas_image()
            # シークバーを再描画
            self.update_ui_texts()

    def _update_corner_handles(self, corners: list[tuple[int, int, int, int]]) -> None:
        """コーナーハンドルの矩形を更新(存在しなければ作成)."""
        if not hasattr(self, 'corner_ids'):
            self.corner_ids = [None, None, None, None]
        
        # リサイズロック時は非表示にする
        lock_res = self.lock_var.get()
        state = tk.HIDDEN if lock_res else tk.NORMAL
        
        # デフォルトのハンドルの色（通常時: 白）
        handle_color = self.theme.get("handle_color", "#FFFFFF")
        
        for i, rect in enumerate(corners):
            if self.corner_ids[i] is None:
                self.corner_ids[i] = self.canvas.create_rectangle(*rect, fill=handle_color, outline="black", width=1, state=state)
            else:
                self.canvas.coords(self.corner_ids[i], *rect)
                self.canvas.itemconfig(self.corner_ids[i], fill=handle_color, state=state)

    def _get_corner_coords(self, scaled_rect: list[int]) -> list[tuple[int, int, int, int]]:
        """スケール後の矩形座標から4隅のハンドル矩形座標を計算."""
        cx1, cy1, cx2, cy2 = scaled_rect
        size = self.theme.get("handle_size", 8)
        half = size // 2
        return [
            (cx1 - half, cy1 - half, cx1 + half, cy1 + half),
            (cx2 - half, cy1 - half, cx2 + half, cy1 + half),
            (cx1 - half, cy2 - half, cx1 + half, cy2 + half),
            (cx2 - half, cy2 - half, cx2 + half, cy2 + half),
        ]

    def _sync_crop_rect_ui(self) -> None:
        """クロップ矩形をキャンバスに反映し、ハンドルとエントリを更新."""
        if self.frame is not None:
            self.update_canvas_image()
        else:
            scaled_rect = self._scaled_rect_from_crop()
            
            # ここでも update_canvas_image と同様に枠線を外側に広げる計算を行う
            try:
                # 現在の幅を取得（floatで返ってくる可能性があるため安全に処理）
                current_width = float(self.canvas.itemcget(self.rect_id, "width"))
            except Exception:
                current_width = float(self.theme.get("crop_width", 1))

            offset = 1.0 + current_width / 2.0
            draw_rect = [
                scaled_rect[0] - offset,
                scaled_rect[1] - offset,
                scaled_rect[2] + offset,
                scaled_rect[3] + offset
            ]
            
            self.canvas.coords(self.rect_id, *draw_rect)
            self._update_corner_handles(self._get_corner_coords(scaled_rect))
        self.update_crop_entries()

    def load_trajectory(self):
        """動画と同名の .tsv からマウス軌跡データを読み込む."""
        self.trajectory_data = []
        if not self.video_filepath:
            return
        
        tsv_path = os.path.splitext(self.video_filepath)[0] + ".tsv"
        if os.path.exists(tsv_path):
            try:
                with open(tsv_path, "r", encoding="utf-8") as f:
                    # ヘッダーをスキップ
                    next(f, None)
                    for line in f:
                        parts = line.strip().split("\t")
                        if len(parts) >= 4:
                            # timestamp, frame, x, y, click, keys
                            try:
                                t = float(parts[0])
                                x = int(parts[2])
                                y = int(parts[3])
                                click = parts[4] if len(parts) > 4 else "None"
                                keys = parts[5] if len(parts) > 5 else "None"
                                self.trajectory_data.append((t, x, y, click, keys))
                            except:
                                pass
            except Exception as e:
                print(f"TSV読込エラー: {e}")

    def open_current_tsv(self):
        """動画と同名の .tsv ファイルを既定のエディタで開く."""
        if not self.video_filepath:
            messagebox.showinfo("Info", "動画が選択されていません")
            return
        
        tsv_path = os.path.splitext(self.video_filepath)[0] + ".tsv"
        if not os.path.exists(tsv_path):
            messagebox.showinfo("Info", f"TSVファイルが見つかりません:\n{tsv_path}")
            return
        
        open_file(tsv_path)

    def update_canvas_overlay(self, img):
        """マウス軌跡等のオーバーレイを表示中のフレームに合わせて描画."""
        if not self.show_trajectory_var.get() or not self.trajectory_data:
            return

        # --- 1. マウス軌跡のオーバーレイ (現在の時刻に最も近いデータ) ---
        mouse_data = None
        current_row_idx = -1
        for i, row in enumerate(self.trajectory_data):
            t, x, y, click, keys = row
            if abs(t - self.current_time) < (1.0 / self.fps):
                mouse_data = row
                current_row_idx = i
                break
        
        if mouse_data:
            t_curr, x, y, click, keys = mouse_data
            
            # --- 波紋ロジック ---
            ripple_age = 0.0
            ripple_type = ""
            lookback_sec = 0.5
            if current_row_idx > 0:
                for i in range(current_row_idx, 0, -1):
                    t_c, c_c = self.trajectory_data[i][0], self.trajectory_data[i][3]
                    t_p, c_p = self.trajectory_data[i-1][0], self.trajectory_data[i-1][3]
                    if t_curr - t_c > lookback_sec: break
                    for btn_char, btn_name in [("L", "left"), ("R", "right"), ("M", "middle")]:
                        if btn_char in c_p and btn_char not in c_c:
                            ripple_type = btn_name
                            ripple_age = t_curr - t_c
                            break
                    if ripple_type: break

            overlay_utils.draw_mouse_overlay(
                img, x, y, click, 
                self.canvas_scale_x, self.canvas_scale_y, 
                self.theme,
                ripple_age=ripple_age,
                ripple_type=ripple_type
            )

        # --- 2. 入力履歴オーバーレイ (ONの場合のみ) ---
        if self.show_subtitle_var.get():
            # 字幕用テキスト生成 (overlay_utilsの共通ロジックを使用)
            display_text = overlay_utils.get_input_display_text(click, keys)
            
            # 履歴マネージャを更新
            self.input_history_manager.update(self.current_time, display_text)
            
            # クロップ範囲のみに字幕を出すための描画範囲計算
            # キャンバス上の画像表示領域内におけるクロップ矩形の相対座標
            cx1, cy1, cx2, cy2 = self.crop_rect
            
            # 拡大率（現在の拡大表示を考慮）
            z = self.image_zoom
            # キャンバス座標系への変換
            # NOTE: self.crop_rect はビデオ本来の座標 [0, vid_w]
            # キャンバス上の画像は vid_w * self.canvas_scale_x * z に拡大されている
            sx = self.canvas_scale_x * z
            sy = self.canvas_scale_y * z
            
            # 描画対象の矩形範囲 (キャンバス画像内の相対座標)
            target_rect = [
                int(cx1 * sx),
                int(cy1 * sy),
                int(cx2 * sx),
                int(cy2 * sy)
            ]

            # 表示すべき入力を取得
            fade_duration = self.theme.get("input_overlay", {}).get("fade_duration", 1.0)
            active_inputs = self.input_history_manager.get_active_inputs(self.current_time, fade_duration)
            
            # 描画
            if active_inputs:
                overlay_utils.draw_input_overlay(
                    img, active_inputs,
                    sx, sy,
                    self.theme,
                    target_rect=target_rect
                )

    # ------------------ シークバーとマーカー ------------------
    def get_x(self, t):
        if self.duration <= 0:
            return self.SEEK_MARGIN
        # シークバーキャンバスの実際の幅を取得
        seek_w = self.seek_canvas.winfo_width()
        if seek_w <= 1:
            seek_w = self.CANVAS_W
        return self.SEEK_MARGIN + (t / self.duration) * (seek_w - 2*self.SEEK_MARGIN)

    def get_t(self, x):
        if self.duration <= 0:
            return 0
        # シークバーキャンバスの実際の幅を取得
        seek_w = self.seek_canvas.winfo_width()
        if seek_w <= 1:
            seek_w = self.CANVAS_W
        w = seek_w - 2*self.SEEK_MARGIN
        ratio = (x - self.SEEK_MARGIN) / w
        return max(0, min(1.0, ratio)) * self.duration

    def draw_seekbar(self):
        self.seek_canvas.delete("all")
        if self.duration <= 0:
            return

        # シークバーキャンバスの実際のサイズを取得
        w = self.seek_canvas.winfo_width()
        h = self.seek_canvas.winfo_height()
        if w <= 1:
            w = self.CANVAS_W
        if h <= 1:
            h = self.SEEK_H

        bar_y = h // 2

        # Base Bar
        self.seek_canvas.create_rectangle(
            self.SEEK_MARGIN, bar_y-4, w-self.SEEK_MARGIN, bar_y+4, fill="#ddd", outline="#aaa")

        # Play progress
        cx = self.get_x(self.current_time)
        self.seek_canvas.create_rectangle(
            self.SEEK_MARGIN, bar_y-4, cx, bar_y+4, fill="#4da6ff", outline="")

        # Start Marker (Top)
        sx = self.get_x(self.start_time)
        self.seek_canvas.create_polygon(
            sx-8, bar_y-8, sx+8, bar_y-8, sx, bar_y, fill=self.theme.get("start_color_bg"), outline="black")
        self.seek_canvas.create_text(
            sx, bar_y-20, text="Start", fill="#006600", font=("Arial", 8))

        # End Marker (Bottom)
        ex = self.get_x(self.end_time)
        self.seek_canvas.create_polygon(
            ex-8, bar_y+8, ex+8, bar_y+8, ex, bar_y, fill=self.theme.get("end_color_bg"), outline="black")
        self.seek_canvas.create_text(
            ex, bar_y+20, text="End", fill="#990000", font=("Arial", 8))

        # Current Thumb (Circle)
        self.seek_canvas.create_oval(
            cx-7, bar_y-7, cx+7, bar_y+7, fill="white", outline="#333", width=2)

    def seek_on_click(self, e):
        if self.duration <= 0:
            return
        # Determine closest target
        sx = self.get_x(self.start_time)
        ex = self.get_x(self.end_time)

        dist_s = abs(e.x - sx)
        dist_e = abs(e.x - ex)

        limit = 15
        if dist_s < limit and dist_s <= dist_e:
            self.drag_mode = "start"
        elif dist_e < limit:
            self.drag_mode = "end"
        else:
            self.drag_mode = "current"
            self.handle_seek_drag(e.x)

    def seek_on_drag(self, e):
        if self.duration <= 0 or not self.drag_mode:
            return
        self.handle_seek_drag(e.x)

    def seek_on_release(self, e):
        self.drag_mode = None

    def handle_seek_drag(self, mouse_x):
        # 1秒単位の丸めから、フレーム単位（FPSに基づく）の丸めに変更
        t_raw = self.get_t(mouse_x)
        fps = getattr(self, 'fps', 30.0) or 30.0
        t = round(t_raw * fps) / fps

        if self.drag_mode == "current":
            self.set_current_time_direct(t)
        elif self.drag_mode == "start":
            self.set_start_time_direct(t)
        elif self.drag_mode == "end":
            self.set_end_time_direct(t)

    # クロップ矩形のマウス操作は CropHandlerMixin のメソッドを使用します

    def on_mouse_down(self, e):
        edges = self.near_edge(e.x, e.y, m=self.theme.get("edge_margin", 20))
        if any(edges.values()):
            if self.lock_var.get():
                # ロック中はリサイズ不可だが、内側ならドラッグ開始（移動のみ許可）
                if self.inside_rect(e.x, e.y):
                    self.dragging_rect = True
                    self._start_dragging(e)
                return
            self.resizing_rect = True
            self.resize_edge = edges
            self.orig_rect = self.crop_rect.copy()
            try:
                self.push_crop_history()
            except Exception:
                pass
            self.maintain_aspect_ratio = (e.state & 0x1) != 0
            if self.maintain_aspect_ratio:
                w = self.orig_rect[2] - self.orig_rect[0]
                h = self.orig_rect[3] - self.orig_rect[1]
                self.orig_aspect_ratio = w / h if h > 0 else 1.0
            self.rect_focused = True
            try:
                self.update_canvas_image()
            except Exception:
                pass
        elif self.inside_rect(e.x, e.y):
            if self.lock_move_var.get():
                return
            self.dragging_rect = True
            self._start_dragging(e)
        else:
            # クリックが領域外ならフォーカス解除
            if getattr(self, 'rect_focused', False):
                self.rect_focused = False
                try:
                    self.update_canvas_image()
                except Exception:
                    pass

    def _start_dragging(self, e):
        try:
            self.push_crop_history()
        except Exception:
            pass
        img_x, img_y = self.canvas_mouse_to_image_coords(e.x, e.y)
        # ドラッグ開始時の左上座標を保存（軸固定移動用）
        self.drag_start_rect = self.crop_rect.copy()
        self.drag_offset = (img_x - self.crop_rect[0], img_y - self.crop_rect[1])
        self.rect_focused = True
        try:
            self.update_canvas_image()
        except Exception:
            pass


    def on_mouse_drag(self, e):
        if self.resizing_rect:
            # マウス座標を画像座標に変換
            ix, iy = self.canvas_mouse_to_image_coords(e.x, e.y)
            img_x, img_y = int(round(ix)), int(round(iy))

            # Ctrl キー (0x4), Shift キー (0x1), Alt キー (Windows: 0x20000 | 0x8)
            try:
                ctrl_held = (e.state & 0x4) != 0
                shift_held = (e.state & 0x1) != 0
                alt_held = (e.state & (0x20000 | 0x8)) != 0
            except Exception:
                ctrl_held = False
                shift_held = False
                alt_held = False

            x1, y1, x2, y2 = self.orig_rect
            
            if ctrl_held:
                # 対称リサイズ: 反対側も同じ分だけ動かす
                if self.resize_edge["l"] or self.resize_edge["r"]:
                    cx = (x1 + x2) / 2
                    if alt_held:
                        w = round(abs(img_x - cx) * 2 / 10) * 10
                    else:
                        w = abs(img_x - cx) * 2
                    x1 = cx - w/2
                    x2 = cx + w/2
                
                if self.resize_edge["t"] or self.resize_edge["b"]:
                    cy = (y1 + y2) / 2
                    if alt_held:
                        h = round(abs(img_y - cy) * 2 / 10) * 10
                    else:
                        h = abs(img_y - cy) * 2
                    y1 = cy - h/2
                    y2 = cy + h/2
            else:
                # 通常のリサイズ
                if self.resize_edge["l"]:
                    if alt_held:
                        x1 = x2 - round((x2 - img_x) / 10) * 10
                    else:
                        x1 = img_x
                if self.resize_edge["r"]:
                    if alt_held:
                        x2 = x1 + round((img_x - x1) / 10) * 10
                    else:
                        x2 = img_x
                if self.resize_edge["t"]:
                    if alt_held:
                        y1 = y2 - round((y2 - img_y) / 10) * 10
                    else:
                        y1 = img_y
                if self.resize_edge["b"]:
                    if alt_held:
                        y2 = y1 + round((img_y - y1) / 10) * 10
                    else:
                        y2 = img_y

            # Shift キー押下時のアスペクト比ロック
            self.maintain_aspect_ratio = shift_held
            if self.maintain_aspect_ratio:
                x1, y1, x2, y2 = self.maintain_aspect_ratio_resize(
                    x1, y1, x2, y2, ctrl_held=ctrl_held)

            self.crop_rect = self.clamp_rect_resize(x1, y1, x2, y2)

            # スケール後の座標で矩形を描画
            scaled_rect = [
                int(self.crop_rect[0] * self.canvas_scale_x) +
                self.canvas_offset_x,
                int(self.crop_rect[1] * self.canvas_scale_y) +
                self.canvas_offset_y,
                int(self.crop_rect[2] * self.canvas_scale_x) +
                self.canvas_offset_x,
                int(self.crop_rect[3] * self.canvas_scale_y) +
                self.canvas_offset_y
            ]
            self.canvas.coords(self.rect_id, *scaled_rect)
            # 矩形の外見をテーマに合わせて更新
            self.update_canvas_image()
            # update corner handles
            self._update_corner_handles(self._get_corner_coords(scaled_rect))
            self.update_crop_entries()
        elif self.dragging_rect:
            # マウス座標を画像座標に変換
            ix, iy = self.canvas_mouse_to_image_coords(e.x, e.y)
            img_x, img_y = int(round(ix)), int(round(iy))

            # 修飾キー判定 (Shift: 0x1, Alt: 0x20000 | 0x8)
            try:
                shift_held = (e.state & 0x1) != 0
                alt_held = (e.state & (0x20000 | 0x8)) != 0
            except Exception:
                shift_held = False
                alt_held = False

            w = self.crop_rect[2] - self.crop_rect[0]
            h = self.crop_rect[3] - self.crop_rect[1]
            
            # マウス位置に基づいた生の移動先候補
            raw_nx = img_x - self.drag_offset[0]
            raw_ny = img_y - self.drag_offset[1]

            # Shift押下時は軸固定（水平または垂直の移動量が大きい方を優先）
            if shift_held and hasattr(self, 'drag_start_rect'):
                sx, sy = self.drag_start_rect[0], self.drag_start_rect[1]
                dx = raw_nx - sx
                dy = raw_ny - sy
                if abs(dx) > abs(dy):
                    nx, ny = raw_nx, sy
                else:
                    nx, ny = sx, raw_ny
            else:
                nx, ny = raw_nx, raw_ny

            # Alt押下時は10px単位にスナップ
            if alt_held:
                nx = round(nx / 10) * 10
                ny = round(ny / 10) * 10
                
            self.crop_rect = self.clamp_rect_move(nx, ny, w, h)

            # スケール後の座標で矩形を描画
            scaled_rect = [
                int(self.crop_rect[0] * self.canvas_scale_x) +
                self.canvas_offset_x,
                int(self.crop_rect[1] * self.canvas_scale_y) +
                self.canvas_offset_y,
                int(self.crop_rect[2] * self.canvas_scale_x) +
                self.canvas_offset_x,
                int(self.crop_rect[3] * self.canvas_scale_y) +
                self.canvas_offset_y
            ]
            self.canvas.coords(self.rect_id, *scaled_rect)
            self._update_corner_handles(self._get_corner_coords(scaled_rect))
            self.update_crop_entries()
            # 矩形の外見をテーマに合わせて更新
            self.update_canvas_image()

    def on_mouse_up(self, e):
        self.dragging_rect = False
        self.resizing_rect = False
        # ドラッグ終了時にはアスペクト比ロックを解除（再度Shiftで有効）
        try:
            self.maintain_aspect_ratio = False
        except Exception:
            pass
        # 矩形の外見をテーマに合わせて更新
        self.update_canvas_image()

    def on_right_mouse_down(self, e):
        # 右クリックで矩形を移動開始（リサイズは行わない）
        if self.inside_rect(e.x, e.y):
            if self.lock_move_var.get():
                return
            self.dragging_rect = True
            try:
                self.push_crop_history()
            except Exception:
                pass
            img_x, img_y = self.canvas_mouse_to_image_coords(e.x, e.y)
            self.drag_offset = (
                img_x - self.crop_rect[0], img_y - self.crop_rect[1])
            # フォーカスを与える
            self.rect_focused = True
            try:
                self.update_canvas_image()
            except Exception:
                pass

    def on_right_mouse_drag(self, e):
        # 右ドラッグは矩形移動と同じ挙動
        if self.dragging_rect:
            # reuse existing dragging logic
            img_x, img_y = self.canvas_mouse_to_image_coords(e.x, e.y)
            w = self.crop_rect[2] - self.crop_rect[0]
            h = self.crop_rect[3] - self.crop_rect[1]
            nx = img_x - self.drag_offset[0]
            ny = img_y - self.drag_offset[1]
            self.crop_rect = self.clamp_rect_move(nx, ny, w, h)

            scaled_rect = self._scaled_rect_from_crop()
            self.canvas.coords(self.rect_id, *scaled_rect)
            self._update_corner_handles(self._get_corner_coords(scaled_rect))
            self.update_crop_entries()

    def on_right_mouse_up(self, e):
        self.dragging_rect = False

    # ------------------ 中央クリックでのパンとホイールズーム ------------------
    def on_middle_down(self, e):
        try:
            self._panning = True
            self._pan_start = (e.x, e.y)
        except Exception:
            self._panning = False

    def on_middle_drag(self, e):
        if not getattr(self, '_panning', False):
            return
        try:
            sx, sy = self._pan_start
            dx = e.x - sx
            dy = e.y - sy
            self.pan_offset_x = getattr(self, 'pan_offset_x', 0) + dx
            self.pan_offset_y = getattr(self, 'pan_offset_y', 0) + dy
            self._pan_start = (e.x, e.y)
            if self.frame is not None:
                self.update_canvas_image()
        except Exception:
            pass

    def on_middle_up(self, e):
        self._panning = False

    def on_middle_double_click(self, e):
        """中央ボタンのダブルクリックでズームとパンをリセット"""
        self.image_zoom = 1.0
        self.pan_offset_x = 0
        self.pan_offset_y = 0
        if self.frame is not None:
            self.update_canvas_image()

    def on_canvas_wheel(self, e):
        try:
            # Windows: e.delta (positive up), X11: Button-4/5
            if hasattr(e, 'delta') and e.delta:
                delta = e.delta
            else:
                # Button-4 == up, Button-5 == down
                if getattr(e, 'num', None) == 4:
                    delta = 120
                elif getattr(e, 'num', None) == 5:
                    delta = -120
                else:
                    delta = 0
            if delta == 0:
                return
                
            current_zoom = getattr(self, 'image_zoom', 1.0)
            # 安全のためゼロ以下除外
            if current_zoom <= 0.001:
                current_zoom = 0.1

            # 現在のズームレベルを計算 (base 1.1)
            try:
                current_level = round(math.log(current_zoom, 1.1))
            except Exception:
                current_level = 0
            
            if delta > 0:
                new_level = current_level + 1
            else:
                new_level = current_level - 1
            
            # レベル0は正確に1.0にする
            if new_level == 0:
                new_zoom = 1.0
            else:
                new_zoom = 1.1 ** new_level
            
            # キャンバス中央を起点としたズームのためにパンオフセットを調整
            # ズーム倍率の変化比率をパンオフセットに乗じる
            zoom_ratio = new_zoom / current_zoom
            self.pan_offset_x = getattr(self, 'pan_offset_x', 0) * zoom_ratio
            self.pan_offset_y = getattr(self, 'pan_offset_y', 0) * zoom_ratio

            self.image_zoom = max(0.1, min(10.0, new_zoom))

            if self.frame is not None:
                self.update_canvas_image()
        except Exception:
            pass

    def on_canvas_motion(self, e):
        # マウス移動時にカーソルを変更し、ハンドルをハイライトする
        edges = self.near_edge(e.x, e.y, m=self.theme.get("edge_margin", 20))  # ハイライト判定をon_mouse_downと統一
        inside = self.inside_rect(e.x, e.y)
        
        # ホバー状態の更新
        was_hovered = self.rect_hovered
        self.rect_hovered = inside or any(edges.values())
        if was_hovered != self.rect_hovered:
            self.update_canvas_image()

        res_lock = self.lock_var.get()
        move_lock = self.lock_move_var.get()
        cursor = ""
        
        # 角の判定
        if edges.get('l') and edges.get('t'):
            cursor = 'no' if res_lock else 'top_left_corner'
        elif edges.get('r') and edges.get('t'):
            cursor = 'no' if res_lock else 'top_right_corner'
        elif edges.get('l') and edges.get('b'):
            cursor = 'no' if res_lock else 'bottom_left_corner'
        elif edges.get('r') and edges.get('b'):
            cursor = 'no' if res_lock else 'bottom_right_corner'
        else:
            # 辺の判定
            if edges.get('l') or edges.get('r'):
                cursor = 'no' if res_lock else 'sb_h_double_arrow'
            elif edges.get('t') or edges.get('b'):
                cursor = 'no' if res_lock else 'sb_v_double_arrow'
            elif self.inside_rect(e.x, e.y):
                # 内側の判定（移動）
                cursor = 'no' if move_lock else 'fleur'
            else:
                cursor = ''

        try:
            self.canvas.config(cursor=cursor)
        except Exception:
            pass

            # ハンドルのハイライト
        if hasattr(self, 'corner_ids'):
            # クロップ範囲全体のホバー状態
            crop_hovered = self.rect_hovered
            
            # どの角がホバーされているか特定
            highlight_idx = None
            if edges.get('l') and edges.get('t'):
                highlight_idx = 0
            elif edges.get('r') and edges.get('t'):
                highlight_idx = 1
            elif edges.get('l') and edges.get('b'):
                highlight_idx = 2
            elif edges.get('r') and edges.get('b'):
                highlight_idx = 3

            for i, cid in enumerate(self.corner_ids):
                if cid is None:
                    continue
                
                # 色の決定
                # 1. 特定の角をホバー時 -> その角だけ赤
                # 2. クロップ範囲ホバー時 -> 全点黄色
                # 3. それ以外 -> 白
                if i == highlight_idx:
                    color = 'red'
                elif crop_hovered:
                    color = 'yellow'
                else:
                    color = 'white'
                
                try:
                    self.canvas.itemconfig(cid, fill=color)
                except Exception:
                    pass

def main():
    root = tk.Tk()
    VideoCropperApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
