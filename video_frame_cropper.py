from __future__ import annotations

import json
import os
import re
import time
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import filedialog, messagebox, simpledialog
from typing import TYPE_CHECKING, Any, Callable

import cv2
import screen_recorder
from PIL import Image, ImageTk

# 分割したモジュールからインポート
from utils import (
    get_base_dir,
    sec_to_hhmmss,
    sec_to_display,
    hhmmss_to_sec,
    imwrite_jp,
    ratio_value_from_str,
    ratio_label_from_wh,
)
from config import (
    CONFIG_FILENAME,
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


class VideoCropperApp(SeekbarMixin, CropHandlerMixin, ExportMixin):
    """動画のクロップと出力を行うGUIアプリケーション."""

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

    # 矩形リサイズハンドルのサイズとエッジ判定マージン
    HANDLE_SIZE: int = 8
    EDGE_MARGIN: int = 20

    def __init__(self, root):
        self.root = root
        self.root.title("動画クリップ取得ツール - Created By ことりちゅん - v0.2")

        # ウィンドウサイズと最小サイズの初期値（縦は小さくしてキャンバスを縮められるように）
        # 起動時にUI全体が見えるよう最小高さを少し小さくする
        self.root.minsize(800, 360)

        # ビデオ / 再生状態
        self.cap = None
        self.frame = None
        self.fps = 30.0
        self.duration = 0
        self.playing = False
        self.current_time = 0
        self.speed = 1.0
        self.png_compression = 3 
        self.video_filename = ""  # 動画ファイル名（拡張子除く）
        self.video_filepath = ""  # 動画ファイルのフルパス

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

        # 設定ファイルを読み込んでウィンドウ位置・サイズを復元
        self.load_window_geometry()

        self.build_ui()

        self._play_after_id = None
        # スペースキーをどのウィジェットにフォーカスがあっても捕まえる
        self.root.bind_all("<space>", lambda e: self.toggle_play())
        # Undo with Ctrl+Z
        self.root.bind_all("<Control-z>", lambda e: self.undo_crop())
        self.root.bind_all("<Control-Z>", lambda e: self.undo_crop())
        # Copy crop to clipboard with Ctrl+C
        self.root.bind_all("<Control-c>", lambda e: self.copy_crop_to_clipboard())
        self.root.bind_all("<Control-C>", lambda e: self.copy_crop_to_clipboard())
        # Redo with Ctrl+Y
        self.root.bind_all("<Control-y>", lambda e: self.redo_crop())
        self.root.bind_all("<Control-Y>", lambda e: self.redo_crop())
        # Arrow key repeat handlers (Left/Right)
        self._arrow_repeat_id = None
        self._arrow_dir = None
        self._arrow_start_time = None
        self.root.bind_all('<KeyPress-Left>', lambda e: self._on_arrow_press(e, -1))
        self.root.bind_all('<KeyRelease-Left>', lambda e: self._on_arrow_release(e))
        self.root.bind_all('<KeyPress-Right>', lambda e: self._on_arrow_press(e, 1))
        self.root.bind_all('<KeyRelease-Right>', lambda e: self._on_arrow_release(e))
        # Alt+矢印でクロップ矩形を1px移動
        # return "break" to stop further handling (avoid double-handling)
        self.root.bind_all('<Alt-Up>', lambda e: (self.move_crop_by(0, -1) or "break"))
        self.root.bind_all('<Alt-Down>', lambda e: (self.move_crop_by(0, 1) or "break"))
        self.root.bind_all('<Alt-Left>', lambda e: (self.move_crop_by(-1, 0) or "break"))
        self.root.bind_all('<Alt-Right>', lambda e: (self.move_crop_by(1, 0) or "break"))
        # Shift+矢印でクロップ矩形を拡大縮小（上下で高さ、左右で幅を2pxずつ）
        self.root.bind_all('<Shift-Up>', lambda e: (self.expand_crop(0, 1) or "break"))
        self.root.bind_all('<Shift-Down>', lambda e: (self.expand_crop(0, -1) or "break"))
        self.root.bind_all('<Shift-Left>', lambda e: (self.expand_crop(-1, 0) or "break"))
        self.root.bind_all('<Shift-Right>', lambda e: (self.expand_crop(1, 0) or "break"))
        # Home/End bindings
        self.root.bind_all('<Home>', lambda e: self.set_current_time_direct(self.start_time))
        self.root.bind_all('<End>', lambda e: self.set_current_time_direct(self.end_time))
        self.root.bind_all('<Control-Home>', lambda e: self.set_current_time_direct(0))
        self.root.bind_all('<Control-End>', lambda e: self.set_current_time_direct(self.duration))

        # ウィンドウ終了時に設定を保存
        self.root.protocol("WM_DELETE_WINDOW", self.on_window_close)

        # 動画設定を読み込む
        self.load_config()

    # ---------------- UI Construction ----------------
    def build_ui(self):
        # Top: Load path and controls
        top_panel = tk.Frame(self.root)
        top_panel.pack(fill=tk.X, side=tk.TOP, padx=5, pady=3)

        # 録画ツール起動ボタン (赤系)
        tk.Button(top_panel, text="録画ツール", command=self.open_screen_recorder,
                  bg="#ffcccc", width=10).pack(side=tk.LEFT, padx=5)

        tk.Button(top_panel, text="動画を開く", command=self.load_video,
                  width=10).pack(side=tk.LEFT, padx=5)

        self.entry_fullpath_var = tk.StringVar(value="")
        self.entry_fullpath = tk.Entry(top_panel, textvariable=self.entry_fullpath_var)
        self.entry_fullpath.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0,5))

        # 秒数表示ラベル（右隣）
        self.label_seconds = tk.Label(top_panel, text="(0s)")
        self.label_seconds.pack(side=tk.LEFT, padx=5)

        # 右上のヘルプボタン（ショートカット一覧）
        self.btn_help = tk.Button(top_panel, text="?", command=self.show_shortcuts, width=3)
        self.btn_help.pack(side=tk.RIGHT, padx=4)
        self.add_tooltip(self.btn_help, "ショートカット一覧を表示")

        # 1. Video Canvas (拡大縮小対応)
        self.canvas = tk.Canvas(
            self.root, width=self.CANVAS_W, height=self.CANVAS_H, bg="black")
        self.canvas.pack(fill=tk.BOTH, expand=True)
        self.canvas_image = self.canvas.create_image(0, 0, anchor=tk.NW)
        self.rect_id = self.canvas.create_rectangle(
            *self.crop_rect, outline="red", width=2)

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
        self.canvas.bind("<Button-4>", self.on_canvas_wheel)
        self.canvas.bind("<Button-5>", self.on_canvas_wheel)

        # 2. Control Panel Frame (固定サイズ、一つのフレームに格納)
        control_frame = tk.Frame(self.root, bg="#f5f5f5")
        control_frame.pack(fill=tk.X, side=tk.BOTTOM)

        # 2. Main Controls (Play, Speed)
        main_ctrl = tk.Frame(control_frame)
        main_ctrl.pack(pady=2)

        self.btn_video_start = tk.Button(main_ctrl, text="◀◀先頭", command=self.go_to_video_start, width=8)
        self.btn_video_start.pack(side=tk.LEFT, padx=4)
        self.add_tooltip(self.btn_video_start, "Ctrl+Home: 動画先頭へ")
        self.btn_trim_start = tk.Button(main_ctrl, text="◀開始位置", command=self.go_to_trim_start, width=10)
        self.btn_trim_start.pack(side=tk.LEFT, padx=4)
        self.add_tooltip(self.btn_trim_start, "Home: 開始位置へ")

        self.btn_play = tk.Button(main_ctrl, text="▲再生", command=self.toggle_play, width=12)
        # 区間再生はチェックボックス化（末尾ボタンの右）
        self.btn_play.pack(side=tk.LEFT, padx=4)
        self.add_tooltip(self.btn_play, "Space: 再生/停止")

        btn_end = tk.Button(main_ctrl, text="終了位置▶", command=self.go_to_trim_end, width=10)
        btn_end.pack(side=tk.LEFT, padx=4)
        self.add_tooltip(btn_end, "End: 終了位置へ")
        btn_tail = tk.Button(main_ctrl, text="末尾▶▶", command=self.go_to_video_end, width=8)
        btn_tail.pack(side=tk.LEFT, padx=4)
        self.add_tooltip(btn_tail, "Ctrl+End: 動画末尾へ")
        self.range_var = tk.BooleanVar(value=False)
        tk.Checkbutton(main_ctrl, text="区間再生", variable=self.range_var).pack(side=tk.LEFT, padx=(4,8))

        # ループ再生チェックボックス
        self.loop_var = tk.BooleanVar(value=False)
        tk.Checkbutton(main_ctrl, text="ループ再生", variable=self.loop_var).pack(side=tk.LEFT, padx=(8,4))
        # 往復再生（ループがオンのときのみ有効）
        self.pingpong_var = tk.BooleanVar(value=False)
        self.chk_pingpong = tk.Checkbutton(main_ctrl, text="往復再生", variable=self.pingpong_var, state=tk.DISABLED)
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
        tk.Label(main_ctrl, text="速度:").pack(side=tk.LEFT, padx=(10, 2))
        self.speed_var = tk.StringVar(value="1.0")
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
        self.add_tooltip(self.speed_spinbox, "速度: -9.9〜9.9 (変更後Enter)")

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

            btn_minus = tk.Button(ctrl_f, text="-1s", width=4,
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

            btn_plus = tk.Button(ctrl_f, text="+1s", width=4,
                      command=lambda: self.adjust_time(var_getter, var_setter, 1))
            btn_plus.pack(side=tk.LEFT)
            self.add_tooltip(btn_plus, "+1s: 1秒進める")

            return entry

        # Start Time Control
        self.entry_start = create_time_control(time_panel, "▼Start (開始)", "#00aa00",
                               lambda: self.start_time, self.set_start_time_direct, True)

        # Divider
        tk.Frame(time_panel, width=2, bg="#ccc", height=40).pack(side=tk.LEFT, padx=10)

        # Current Time Control
        self.entry_current = create_time_control(time_panel, "●Current (現在)", "#000000",
                             lambda: self.current_time, self.set_current_time_direct, False)
        # divider between controls
        tk.Frame(time_panel, width=2, bg="#ccc", height=40).pack(side=tk.LEFT, padx=10)

        # End Time Control
        self.entry_end = create_time_control(time_panel, "▲End (終了)", "#cc0000",
                             lambda: self.end_time, self.set_end_time_direct, True)

        # disable space input for time entry boxes
        self.entry_start.bind('<space>', _ignore_space)
        self.entry_current.bind('<space>', _ignore_space)
        self.entry_end.bind('<space>', _ignore_space)


        # 2.7. Crop Size Input Panel (placed below trimming range per user request)
        crop_panel = tk.LabelFrame(control_frame, text="クロップ範囲", bd=1,
                                  relief=tk.SOLID, padx=5, pady=5)
        crop_panel.pack(pady=5, fill=tk.X, padx=10)

        size_ctrl = tk.Frame(crop_panel)
        size_ctrl.pack()

        # 戻すボタン（Undo）
        self.btn_undo = tk.Button(size_ctrl, text="戻す", command=self.undo_crop, width=6)
        self.btn_undo.pack(side=tk.LEFT, padx=4)
        self.btn_undo.config(state=tk.DISABLED)
        self.add_tooltip(self.btn_undo, "Ctrl+Z: 戻す")
        # 進むボタン（Redo）
        self.btn_redo = tk.Button(size_ctrl, text="進む", command=self.redo_crop, width=6)
        self.btn_redo.pack(side=tk.LEFT, padx=4)
        self.btn_redo.config(state=tk.DISABLED)
        self.add_tooltip(self.btn_redo, "Ctrl+Y: 進む")

        # 座標パネルラベルとロックボタン
        ttk.Separator(size_ctrl, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=10)
        tk.Label(size_ctrl, text="座標:").pack(side=tk.LEFT, padx=(5,0))
        self.btn_lock_move = tk.Button(size_ctrl, text="🔓", width=3, font=("Consolas", 14, "bold"), bg="#ccffcc", command=self.toggle_move_lock)
        self.btn_lock_move.pack(side=tk.LEFT, padx=(2,5))
        self.add_tooltip(self.btn_lock_move, "クロップ位置(X, Y)をロック (リサイズは可能)")

        # X座標入力
        tk.Label(size_ctrl, text="X:").pack(side=tk.LEFT, padx=2)
        self.entry_crop_x = tk.Entry(
            size_ctrl, width=8, font=("Consolas", 10), justify="center")
        self.entry_crop_x.insert(0, "100")
        self.entry_crop_x.pack(side=tk.LEFT, padx=2)
        self.entry_crop_x.bind(
            "<Return>", lambda e: self.update_crop_from_entries())
        self.entry_crop_x.bind(
            "<FocusOut>", lambda e: self.update_crop_from_entries())
        self.add_tooltip(self.entry_crop_x, "クロップ開始X座標 (Alt+左右で1px移動, Shift+左右で幅を拡大縮小)")

        # Y座標入力
        tk.Label(size_ctrl, text="Y:").pack(side=tk.LEFT, padx=5)
        self.entry_crop_y = tk.Entry(
            size_ctrl, width=8, font=("Consolas", 10), justify="center")
        self.entry_crop_y.insert(0, "80")
        self.entry_crop_y.pack(side=tk.LEFT, padx=2)
        self.entry_crop_y.bind(
            "<Return>", lambda e: self.update_crop_from_entries())
        self.entry_crop_y.bind(
            "<FocusOut>", lambda e: self.update_crop_from_entries())
        self.add_tooltip(self.entry_crop_y, "クロップ開始Y座標 (Alt+上下で1px移動, Shift+上下で高さを拡大縮小)")

        # 幅と高さ選択
        ttk.Separator(size_ctrl, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=10)
        tk.Label(size_ctrl, text="解像度:").pack(side=tk.LEFT, padx=5)
        # 解像度ロックボタン
        self.btn_lock_res = tk.Button(size_ctrl, text="🔓", width=3, font=("Consolas", 14, "bold"), bg="#ccffcc", command=self.toggle_resolution_lock)
        self.btn_lock_res.pack(side=tk.LEFT, padx=(2,5))
        self.add_tooltip(self.btn_lock_res, "解像度・アスペクト比をロック (移動は可能)")

        # 比率選択
        self.ratio_var = tk.StringVar(value=self.aspect_options[0])
        self.ratio_optionmenu = tk.OptionMenu(size_ctrl, self.ratio_var, *self.aspect_options)
        self.ratio_optionmenu.pack(side=tk.LEFT, padx=(4,2))

        self.resolution_var = tk.StringVar(value="カスタム")
        self.resolution_optionmenu = tk.OptionMenu(size_ctrl, self.resolution_var, "カスタム")
        self.resolution_optionmenu.pack(side=tk.LEFT, padx=2)
        self.ratio_var.trace_add('write', lambda *args: self.update_resolution_menu())

        # 幅入力
        tk.Label(size_ctrl, text="幅:").pack(side=tk.LEFT, padx=5)
        self.entry_crop_w = tk.Entry(
            size_ctrl, width=8, font=("Consolas", 10), justify="center")
        self.entry_crop_w.insert(0, "200")
        self.entry_crop_w.pack(side=tk.LEFT, padx=2)
        self.entry_crop_w.bind(
            "<Return>", lambda e: self.update_crop_from_entries())
        self.entry_crop_w.bind(
            "<FocusOut>", lambda e: self.update_crop_from_entries())
        self.add_tooltip(self.entry_crop_w, "出力される画像の幅 (ピクセル)")

        # 高さ入力
        tk.Label(size_ctrl, text="高:").pack(side=tk.LEFT, padx=5)
        self.entry_crop_h = tk.Entry(
            size_ctrl, width=8, font=("Consolas", 10), justify="center")
        self.entry_crop_h.insert(0, "170")
        self.entry_crop_h.pack(side=tk.LEFT, padx=2)
        self.entry_crop_h.bind(
            "<Return>", lambda e: self.update_crop_from_entries())
        self.entry_crop_h.bind(
            "<FocusOut>", lambda e: self.update_crop_from_entries())
        self.add_tooltip(self.entry_crop_h, "出力される画像の高さ (ピクセル)")

        # プリセット保存/削除ボタン
        self.btn_save_preset = tk.Button(size_ctrl, text="プリセット保存", command=self.add_resolution_preset)
        self.btn_save_preset.pack(side=tk.LEFT, padx=4)
        self.btn_delete_preset = tk.Button(size_ctrl, text="プリセット削除", command=self.delete_resolution_preset)
        self.btn_delete_preset.pack(side=tk.LEFT, padx=4)

        

        # 5. 出力グループ（PNG出力・動画保存・フィルタ）
        # 出力パネルも他のパネルと同じ罫線デザインに揃える
        output_panel = tk.LabelFrame(control_frame, text="出力", bd=1, relief=tk.SOLID, padx=5, pady=5)
        output_panel.pack(fill=tk.X, padx=10, pady=5)

        # 左側: フィルタ + 保存ボタン
        left_out = tk.Frame(output_panel)
        left_out.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=6, pady=4)

        filter_panel = tk.Frame(left_out)
        filter_panel.pack(side=tk.TOP, fill=tk.X)
        self.check_prev_next = tk.BooleanVar(value=True)
        tk.Checkbutton(filter_panel, text="前後のフレームと異なるとき出力しない",
                       variable=self.check_prev_next).pack(side=tk.LEFT, padx=5)

        self.check_duplicate = tk.BooleanVar(value=True)
        tk.Checkbutton(filter_panel, text="直前に出力したフレームと同一のとき出力しない",
                       variable=self.check_duplicate).pack(side=tk.LEFT, padx=5)

        # 設定保存 / 設定確認 ボタン
        tk.Button(left_out, text="設定保存", command=self.save_video_settings).pack(side=tk.LEFT, padx=6, pady=4)
        tk.Button(left_out, text="設定確認", command=self.open_video_settings).pack(side=tk.LEFT, padx=6, pady=4)
        self.btn_reload_settings = tk.Button(left_out, text="設定再読み込み", command=self.load_config)
        self.btn_reload_settings.pack(side=tk.LEFT, padx=6, pady=4)
        self.add_tooltip(self.btn_reload_settings, "設定ファイルを再読み込み")

        # 右側: 出力操作（右下にまとめる）
        right_out = tk.Frame(output_panel)
        right_out.pack(side=tk.RIGHT, anchor=tk.S, padx=6, pady=4)
        # PNG圧縮設定（右寄せでPNG出力の近くへ）
        comp_frame_r = tk.Frame(right_out)
        comp_frame_r.pack(side=tk.TOP, anchor=tk.E)
        tk.Label(comp_frame_r, text="PNG圧縮:").pack(side=tk.LEFT, padx=(0,4))
        self.compression_var = tk.StringVar(value=str(self.png_compression))
        self.compression_spinbox = tk.Spinbox(
            comp_frame_r,
            from_=0,
            to=9,
            increment=1,
            width=3,
            textvariable=self.compression_var,
            command=self.change_compression
        )
        self.compression_spinbox.pack(side=tk.LEFT)
        self.compression_spinbox.bind("<Return>", self.change_compression)
        self.compression_spinbox.bind("<FocusOut>", self.change_compression)

        # Copy button + PNG output grouped
        png_btn_frame = tk.Frame(right_out)
        png_btn_frame.pack(side=tk.TOP, anchor=tk.E, pady=2)
        self.btn_copy_image = tk.Button(png_btn_frame, text="🖼️コピー", width=12, command=self.copy_crop_to_clipboard)
        self.btn_copy_image.pack(side=tk.LEFT, padx=(0,6))
        self.add_tooltip(self.btn_copy_image, "Ctrl+C: 現在のクロップをコピー")
        self.btn_export_png = tk.Button(png_btn_frame, text="PNG出力", width=12, command=self.export_png, bg="#ffdddd")
        self.btn_export_png.pack(side=tk.LEFT)
        self.add_tooltip(self.btn_export_png, "PNG出力（フォルダ選択）")

        tk.Button(right_out, text="動画保存", width=24, command=self.export_video, bg="#ddffdd").pack(side=tk.TOP, anchor=tk.E, pady=2)

        # 解像度プリセットをメニューに反映
        self.update_resolution_menu()

    # ------------------ ウィンドウ位置・サイズ管理 ------------------
    def load_window_geometry(self):
        """設定ファイルからウィンドウの位置とサイズを読み込む"""
        config_path = os.path.join(get_base_dir(), CONFIG_FILENAME)
        # デフォルト起動ジオメトリ（UIがウィンドウ内に収まるよう縦を抑える）
        default_geometry = "860x600"

        if os.path.exists(config_path):
            try:
                with open(config_path, "r", encoding="utf-8") as f:
                    config = json.load(f)

                window_geom = config.get("window_geometry", "")
                if window_geom:
                    # ジオメトリが有効かチェック (0でない値であることを確認)
                    try:
                        # ジオメトリをパースしてサイズをチェック
                        parts = window_geom.split('+')
                        size_part = parts[0].split('x')
                        width = int(size_part[0])
                        height = int(size_part[1])
                        if width > 0 and height > 0:
                            self.root.geometry(window_geom)
                            return
                    except:
                        pass
            except Exception as e:
                print(f"ウィンドウ設定の読み込みに失敗しました: {e}")

        # デフォルト値を使用
        self.root.geometry(default_geometry)

    def save_window_geometry(self):
        """ウィンドウの位置とサイズをvideo_frame_cropper_config.jsonに保存"""
        config_path = os.path.join(get_base_dir(), CONFIG_FILENAME)

        try:
            # 既存の設定を読み込む
            config = {}
            if os.path.exists(config_path):
                with open(config_path, "r", encoding="utf-8") as f:
                    config = json.load(f)

            # ウィンドウジオメトリを保存
            config["window_geometry"] = self.root.geometry()

            # 設定をファイルに保存
            with open(config_path, "w", encoding="utf-8") as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"ウィンドウ設定の保存に失敗しました: {e}")

    def on_window_close(self):
        """ウィンドウ終了時の処理"""
        self.save_window_geometry()
        self.root.destroy()

    # ------------------ 設定管理 ------------------
    def load_config(self):
        """設定ファイルから初期値を読み込む"""
        config_path = os.path.join(get_base_dir(), CONFIG_FILENAME)
        if os.path.exists(config_path):
            try:
                with open(config_path, "r", encoding="utf-8") as f:
                    config = json.load(f)

                # 動画ファイルが存在する場合は読み込む
                video_file = config.get("video_file", "")
                if video_file and os.path.exists(video_file):
                    self.cap = cv2.VideoCapture(video_file)
                    if self.cap.isOpened():
                        self.video_filename = os.path.splitext(
                            os.path.basename(video_file))[0]
                        self.video_filepath = os.path.abspath(video_file)
                        self.fps = float(self.cap.get(
                            cv2.CAP_PROP_FPS) or 30.0)
                        frames = int(self.cap.get(
                            cv2.CAP_PROP_FRAME_COUNT) or 0)
                        self.duration = max(
                            0, int(round(frames / self.fps))) if self.fps > 0 else 0

                        # まず、動画ごとの設定ファイル（<video>.settings.json）があれば優先して読み込む
                        per_video_settings = None
                        try:
                            settings_path = os.path.splitext(video_file)[0] + '.settings.json'
                            if os.path.exists(settings_path):
                                with open(settings_path, 'r', encoding='utf-8') as sf:
                                    per_video_settings = json.load(sf)
                        except Exception:
                            per_video_settings = None

                        if per_video_settings:
                            crop_rect = per_video_settings.get('crop_rect', {})
                            self.start_time = per_video_settings.get('start_time', 0)
                            self.end_time = per_video_settings.get('end_time', self.duration)
                        else:
                                        # グローバル設定(video_frame_cropper_config.json)から読み込む
                            crop_rect = config.get("crop_rect", {})
                            self.start_time = config.get("start_time", 0)
                            self.end_time = config.get("end_time", self.duration)

                        if crop_rect:
                            self.crop_rect = [
                                crop_rect.get("x1", 100),
                                crop_rect.get("y1", 80),
                                crop_rect.get("x2", 300),
                                crop_rect.get("y2", 250)
                            ]

                        self.current_time = self.start_time
                        # UIを更新 - スケールされた座標で矩形とハンドルを更新
                        try:
                            scaled = self._scaled_rect_from_crop()
                            self.canvas.coords(self.rect_id, *scaled)
                            self._update_corner_handles(self._get_corner_coords(scaled))
                        except Exception:
                            self.canvas.coords(self.rect_id, *self.crop_rect)
                        # フルパス表示と秒数表示を更新
                        try:
                            self.entry_fullpath_var.set(os.path.abspath(video_file))
                        except Exception:
                            pass
                        try:
                            self.label_seconds.config(text=f"({self.duration}s)")
                        except Exception:
                            pass
                        self.show_frame_at(self.current_time)
                        self.update_ui_texts()
                        self.update_crop_entries()
                        try:
                            self.draw_seekbar()
                        except Exception:
                            pass

                # 解像度プリセットがあれば読み込む、なければデフォルトを作成して保存
                presets = config.get("resolution_presets")
                default_presets = {
                    "160×120（カスタム）": [160, 120],
                    "320×240（QVGA）": [320, 240],
                    "480×320（HVGA）": [480, 320],
                    "640×480（VGA）": [640, 480],
                    "800×600（SVGA）": [800, 600],
                    "1024×768（XGA）": [1024, 768],
                    "1600×1200（UXGA）": [1600, 1200],
                    "426×240（SD 240p）": [426, 240],
                    "640×360（SD 360p）": [640, 360],
                    "854×480（SD 480p）": [854, 480],
                    "1280×720（HD 720p）": [1280, 720],
                    "1366×768（WXGA）": [1366, 768],
                    "1920×1080（FHD 1080p）": [1920, 1080],
                    "2560×1440（2K 1440p）": [2560, 1440],
                    "3840×2160（4K 2160p）": [3840, 2160],
                    "1080×1080（Instagram Feed）": [1080, 1080],
                    "1080×1920（Instagram Story）": [1080, 1920],
                    "1080×1920（TikTok）": [1080, 1920],
                    "1280×720（YouTube Thumbnail）": [1280, 720],
                    "1500×500（Twitter ヘッダー画像）": [1500, 500],
                    "400×400（Twitterプロフィール画像)": [400, 400],
                    "1200×675（Twitter 通常投稿・横長)": [1200, 675],
                    "1200×1200（Twitter 通常投稿・正方形)": [1200, 1200],
                    "1200×1500（Twitter 通常投稿・縦長)": [1200, 1500],
                    "1600×900（Twitter リンクカード大）": [1600, 900],
                    "800×800（Twitter リンクカード小）": [800, 800],
                    "1080×1080（Twitter 広告・正方形）": [1080, 1080],
                    "1920×1080（Twitter 広告・横長）": [1920, 1080],
                    "1200×628（Facebook Post）": [1200, 628],
                    "1080×1920（YouTube Short）": [1080, 1920],
                    "1080×1920（縦FHD 1080p）": [1080, 1920],
                    "32×32（1:1 アイコン）": [32, 32],
                    "1080×1080（1:1）": [1080, 1080],
                }
                if isinstance(presets, dict) and presets:
                    # 既存プリセットのキーが比率プレフィックスを含まない場合は付与して統一する
                    new_presets = {}
                    import re
                    for k, v in presets.items():
                        if re.match(r'^\d+:\d+\s', str(k)):
                            new_presets[k] = v
                        else:
                            try:
                                w, h = int(v[0]), int(v[1])
                                lbl = None
                                try:
                                    lbl = self._ratio_label_from_wh(w, h)
                                except Exception:
                                    lbl = f"{w}:{h}"
                                new_key = f"{lbl} {k}"
                                # avoid collision
                                if new_key in new_presets:
                                    # append suffix
                                    idx = 1
                                    while f"{new_key}#{idx}" in new_presets:
                                        idx += 1
                                    new_key = f"{new_key}#{idx}"
                                new_presets[new_key] = v
                            except Exception:
                                new_presets[k] = v
                    self.resolution_presets = new_presets
                else:
                    # 設定ファイルにプリセットがなければデフォルトを初期書き込み
                    # デフォルトのキーに比率ラベルを付与して保存
                    new_defaults = {}
                    for k, v in default_presets.items():
                        try:
                            w, h = int(v[0]), int(v[1])
                            lbl = self._ratio_label_from_wh(w, h)
                            new_key = f"{lbl} {k}"
                        except Exception:
                            new_key = k
                        new_defaults[new_key] = v
                    self.resolution_presets = new_defaults
                    try:
                        self.save_config()
                    except Exception:
                        pass
                # 読み込んだ設定から選択中の比率を復元（UI が構築済みであることが前提）
                try:
                    sel_ratio = config.get("selected_ratio", "未指定")
                    if hasattr(self, 'ratio_var'):
                        try:
                            self.ratio_var.set(sel_ratio)
                        except Exception:
                            pass
                except Exception:
                    pass

                try:
                    self.update_resolution_menu()
                except Exception:
                    pass

            except Exception as e:
                print(f"設定ファイルの読み込みに失敗しました: {e}")

    def save_config(self):
        """設定ファイルへ現在の重要設定を保存（マージして保存）"""
        config_path = os.path.join(get_base_dir(), "video_frame_cropper_config.json")
        config = {}
        if os.path.exists(config_path):
            try:
                with open(config_path, "r", encoding="utf-8") as f:
                    config = json.load(f)
            except Exception:
                config = {}

        # 更新する値
        config["video_file"] = self.video_filepath or config.get("video_file", "")
        config["crop_rect"] = {
            "x1": int(self.crop_rect[0]),
            "y1": int(self.crop_rect[1]),
            "x2": int(self.crop_rect[2]),
            "y2": int(self.crop_rect[3])
        }
        config["start_time"] = int(self.start_time)
        config["end_time"] = int(self.end_time)
        config["resolution_presets"] = self.resolution_presets
        # 選択中の比率を保存
        try:
            if hasattr(self, 'ratio_var'):
                config['selected_ratio'] = self.ratio_var.get()
            else:
                config['selected_ratio'] = '未指定'
        except Exception:
            config['selected_ratio'] = '未指定'

        try:
            with open(config_path, "w", encoding="utf-8") as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"設定ファイルの保存に失敗しました: {e}")

    def save_video_settings(self):
        """当該動画ファイルに紐づく設定ファイルへ現在の赤枠・開始・終了を保存する"""
        if not self.video_filepath:
            # 動画が選択されていない場合はファイルを選んで保存場所を決める
            path = filedialog.asksaveasfilename(defaultextension='.json', filetypes=[('JSON','*.json')], title='設定ファイルの保存先を選択')
            if not path:
                return
            save_path = path
        else:
            base = os.path.splitext(self.video_filepath)[0]
            save_path = base + '.settings.json'

        data = {
            'video_file': self.video_filepath,
            'crop_rect': {
                'x1': int(self.crop_rect[0]), 'y1': int(self.crop_rect[1]),
                'x2': int(self.crop_rect[2]), 'y2': int(self.crop_rect[3])
            },
            'start_time': int(self.start_time),
            'end_time': int(self.end_time)
        }
        try:
            with open(save_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            messagebox.showinfo('Saved', f'設定を保存しました:\n{save_path}')
        except Exception as e:
            messagebox.showerror('Error', f'設定の保存に失敗しました:\n{e}')

    def update_resolution_menu(self):
        """OptionMenu を現在の self.resolution_presets に合わせて更新する"""
        # (比率判定はクラスヘルパー `_ratio_value_from_str` / `_ratio_label_from_wh` を使用)
        menu = self.resolution_optionmenu["menu"]
        menu.delete(0, tk.END)
        # まずカスタム項目
        menu.add_command(label="カスタム", command=lambda v="カスタム": self.resolution_var.set(v))

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

    # ------------------ ヘルパー: 時間変換 (utils モジュールに委譲) ------------------
    def sec_to_hhmmss(self, sec):
        return sec_to_hhmmss(sec)

    def sec_to_display(self, sec):
        """表示用: HH:MM:SS.mmm (ミリ秒まで)"""
        return sec_to_display(sec)

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
            self.btn_play.config(text="停止")
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
        # フォーカスやホバー時にツールチップを表示する
        def show(e=None):
            try:
                if getattr(widget, '_tooltip_win', None):
                    return
                x = widget.winfo_rootx() + 20
                y = widget.winfo_rooty() + 20
                tw = tk.Toplevel(widget)
                tw.wm_overrideredirect(True)
                tw.wm_geometry(f"+{x}+{y}")
                lbl = tk.Label(tw, text=text, background="#ffffe0", relief='solid', borderwidth=1)
                lbl.pack()
                widget._tooltip_win = tw
            except Exception:
                pass

        def hide(e=None):
            try:
                if getattr(widget, '_tooltip_win', None):
                    widget._tooltip_win.destroy()
                    widget._tooltip_win = None
            except Exception:
                pass

        widget.bind('<FocusIn>', show)
        widget.bind('<FocusOut>', hide)
        widget.bind('<Enter>', show)
        widget.bind('<Leave>', hide)

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
            # move crop horizontally by 1px per step (no fractional)
            dx = -1 if self._arrow_dir == -1 else 1
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
        x1, y1, x2, y2 = self.crop_rect
        vid_w = int(self.cap.get(cv2.CAP_PROP_FRAME_WIDTH))
        vid_h = int(self.cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
        sx = vid_w / self.CANVAS_W
        sy = vid_h / self.CANVAS_H
        vx1, vy1 = int(x1 * sx), int(y1 * sy)
        vx2, vy2 = int(x2 * sx), int(y2 * sy)
        ret, frm = True, self.frame
        try:
            crop = frm[vy1:vy2, vx1:vx2]
            if crop.size == 0:
                return
            rgb = cv2.cvtColor(crop, cv2.COLOR_BGR2RGB)
            pil = Image.fromarray(rgb)
            ok = self._copy_image_to_clipboard(pil)
            if not ok:
                messagebox.showerror('Error', 'クリップボードへ画像をコピーできませんでした')
        except Exception as e:
            messagebox.showerror('Error', f'コピーに失敗しました:\n{e}')

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
            text = 'Ctrl+Z: 戻す\nCtrl+C: クロップをクリップボードへコピー\nSpace: 再生/停止\n区間再生: start->end を再生\nループ再生: 末尾到達で先頭に戻る'

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
        self.btn_lock_res.config(text="🔒" if is_locked else "🔓", bg="#ffcccc" if is_locked else "#ccffcc")
        
        state = tk.DISABLED if is_locked else tk.NORMAL
        self.ratio_optionmenu.config(state=state)
        self.resolution_optionmenu.config(state=state)
        self.entry_crop_w.config(state=state)
        self.entry_crop_h.config(state=state)
        self.btn_undo.config(state=state if self.crop_history else tk.DISABLED)
        self.btn_redo.config(state=state if self.crop_redo else tk.DISABLED)
        self.btn_save_preset.config(state=state)
        self.btn_delete_preset.config(state=state)

    def toggle_move_lock(self):
        """座標設定とマウスによる移動をロック/解除する"""
        is_locked = not self.lock_move_var.get()
        self.lock_move_var.set(is_locked)
        self.btn_lock_move.config(text="🔒" if is_locked else "🔓", bg="#ffcccc" if is_locked else "#ccffcc")
        
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

    def open_screen_recorder(self):
        """スクリーンレコーダーを起動"""
        app = screen_recorder.ScreenRecorderApp(self.root, parent_app=self)

    def open_video_file(self, result_path):
        """外部から指定された動画ファイルを開く"""
        if not os.path.exists(result_path):
            return
        self.load_video(target_path=result_path)

    def load_video(self, target_path=None):
        """Load video file.
        
        Args:
            target_path: 指定がある場合はそのパスを開く。Noneの場合はダイアログを表示。
        """
        if self.playing:
            self.toggle_play()

        if target_path:
            path = target_path
        else:
            path = filedialog.askopenfilename(filetypes=[
                ("動画ファイル", ("*.mp4", "*.mkv", "*.mov", "*.avi")),
                ("MP4", "*.mp4"),
                ("MKV", "*.mkv"),
                ("MOV", "*.mov"),
                ("AVI", "*.avi"),
                ("すべてのファイル", "*.*")
            ])
        
        if not path:
            return

        self.cap = cv2.VideoCapture(path)
        if not self.cap.isOpened():
            messagebox.showerror("Error", "動画を読み込めませんでした")
            return

        # ファイル名から拡張子を除いた名前とフルパスを保存
        self.video_filename = os.path.splitext(os.path.basename(path))[0]
        self.video_filepath = os.path.abspath(path)
        # ズーム/パン状態をリセット
        try:
            self.image_zoom = 1.0
            self.pan_offset_x = 0
            self.pan_offset_y = 0
        except Exception:
            pass

        # 動画を開いた時点で、最後に開いた動画のパスを設定に書き込む
        try:
            self.save_config()
        except Exception:
            pass

        self.fps = float(self.cap.get(cv2.CAP_PROP_FPS) or 30.0)
        frames = int(self.cap.get(cv2.CAP_PROP_FRAME_COUNT) or 0)
        self.duration = max(0, int(round(frames / self.fps))
                            ) if self.fps > 0 else 0

        # default times
        self.start_time = 0
        self.end_time = self.duration
        self.current_time = 0

        # Try to load per-video settings if present; otherwise fall back to global config
        try:
            settings_path = os.path.splitext(path)[0] + '.settings.json'
            per_video = None
            if os.path.exists(settings_path):
                with open(settings_path, 'r', encoding='utf-8') as sf:
                    per_video = json.load(sf)
            if per_video:
                crop_rect = per_video.get('crop_rect', {})
                self.start_time = per_video.get('start_time', self.start_time)
                self.end_time = per_video.get('end_time', self.end_time)
                if crop_rect:
                    self.crop_rect = [
                        crop_rect.get('x1', 100),
                        crop_rect.get('y1', 80),
                        crop_rect.get('x2', 300),
                        crop_rect.get('y2', 250),
                    ]
            else:
                # fallback to global config if present
                try:
                    cfg_path = os.path.join(get_base_dir(), CONFIG_FILENAME)
                    if os.path.exists(cfg_path):
                        with open(cfg_path, 'r', encoding='utf-8') as cf:
                            cfg = json.load(cf)
                            crop_rect = cfg.get('crop_rect', {})
                            self.start_time = cfg.get('start_time', self.start_time)
                            self.end_time = cfg.get('end_time', self.end_time)
                            if crop_rect:
                                self.crop_rect = [
                                    crop_rect.get('x1', 100),
                                    crop_rect.get('y1', 80),
                                    crop_rect.get('x2', 300),
                                    crop_rect.get('y2', 250),
                                ]
                except Exception:
                    # last-resort: center default rectangle
                    pass

        except Exception:
            # Prepare a centered default crop if any read fails
            cw, ch = self.CANVAS_W // 3, self.CANVAS_H // 3
            cx = (self.CANVAS_W - cw) // 2
            cy = (self.CANVAS_H - ch) // 2
            self.crop_rect = self.clamp_rect_canvas([cx, cy, cx+cw, cy+ch])
        else:
            # ensure crop rect is clamped to canvas
            try:
                self.crop_rect = self.clamp_rect_canvas(self.crop_rect)
            except Exception:
                cw, ch = self.CANVAS_W // 3, self.CANVAS_H // 3
                cx = (self.CANVAS_W - cw) // 2
                cy = (self.CANVAS_H - ch) // 2
                self.crop_rect = self.clamp_rect_canvas([cx, cy, cx+cw, cy+ch])
        try:
            scaled = self._scaled_rect_from_crop()
            self.canvas.coords(self.rect_id, *scaled)
            self._update_corner_handles(self._get_corner_coords(scaled))
        except Exception:
            self.canvas.coords(self.rect_id, *self.crop_rect)

        try:
            self.entry_fullpath_var.set(os.path.abspath(path))
        except Exception:
            pass
        try:
            self.label_seconds.config(text=f"({self.duration}s)")
        except Exception:
            pass

        self.show_frame_at(0)
        self.update_ui_texts()
        self.update_crop_entries()
        try:
            self.draw_seekbar()
        except Exception:
            pass

        # Clear undo/redo memory when switching video
        try:
            self.crop_history = []
            self.crop_redo = []
            try:
                self.btn_undo.config(state=tk.DISABLED)
            except Exception:
                pass
            try:
                self.btn_redo.config(state=tk.DISABLED)
            except Exception:
                pass
        except Exception:
            pass

    def toggle_play(self):
        self.playing = not self.playing
        self.btn_play.config(text="■停止" if self.playing else "▲再生")
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
                        self.btn_play.config(text="▲再生")
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
                        self.btn_play.config(text="▲再生")
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
                        self.btn_play.config(text="▲再生(Space)")
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
                        self.btn_play.config(text="▲再生(Space)")

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
        self.tk_img = ImageTk.PhotoImage(img)

        # パンオフセットを加味してキャンバス内に配置
        offset_x = (canvas_w - rw) // 2 + getattr(self, 'pan_offset_x', 0)
        offset_y = (canvas_h - rh) // 2 + getattr(self, 'pan_offset_y', 0)
        self.canvas_offset_x = offset_x
        self.canvas_offset_y = offset_y
        self.canvas.itemconfig(self.canvas_image, image=self.tk_img)
        self.canvas.coords(self.canvas_image, offset_x, offset_y)

        # スケール比を計算（元の座標系から表示上への変換）
        self.canvas_scale_x = rw / self.CANVAS_W
        self.canvas_scale_y = rh / self.CANVAS_H

        # crop_rectをスケールして描画
        scaled_rect = [
            int(self.crop_rect[0] * self.canvas_scale_x) + offset_x,
            int(self.crop_rect[1] * self.canvas_scale_y) + offset_y,
            int(self.crop_rect[2] * self.canvas_scale_x) + offset_x,
            int(self.crop_rect[3] * self.canvas_scale_y) + offset_y
        ]
        self.canvas.coords(self.rect_id, *scaled_rect)

        # ズーム/スケールに合わせて矩形枠線幅を調整（動画解像度における2ドット分の太さを再現）
        try:
            # 動画ピクセルからキャンバス表示ピクセルへのトータルスケールを計算
            total_scale_x = rw / frame_w if frame_w > 0 else 1.0
            total_scale_y = rh / frame_h if frame_h > 0 else 1.0
            avg_total_scale = (total_scale_x + total_scale_y) / 2.0
            # 元の動画の2ドット分を表示上の太さにする（最低1ドットは維持）
            line_w = max(1, int(round(2 * avg_total_scale)))
            self.canvas.itemconfig(self.rect_id, width=line_w)
        except Exception:
            pass

        # 角マーカー（ハンドル）を描画/更新する
        self._update_corner_handles(self._get_corner_coords(scaled_rect))

    # ------------------ キャンバスリサイズ処理 ------------------
    def on_canvas_resize(self, event):
        """キャンバスのリサイズイベントを処理"""
        if event.width > 0 and event.height > 0:
            # スケール比を計算（デフォルトサイズに対する比率）
            # 高さは最小値を保証して計算の安定を図る
            h_eff = max(event.height, getattr(self, 'CANVAS_MIN_H', event.height))
            self.canvas_scale_x = event.width / self.CANVAS_W
            self.canvas_scale_y = h_eff / self.CANVAS_H
            # フレームを再描画
            if self.frame is not None:
                self.update_canvas_image()
            # シークバーを再描画
            self.update_ui_texts()

    def _update_corner_handles(self, corners: list[tuple[int, int, int, int]]) -> None:
        """コーナーハンドルの矩形を更新（存在しなければ作成）."""
        if not hasattr(self, 'corner_ids'):
            self.corner_ids = [None, None, None, None]
        for i, rect in enumerate(corners):
            if self.corner_ids[i] is None:
                self.corner_ids[i] = self.canvas.create_rectangle(*rect, fill='red')
            else:
                self.canvas.coords(self.corner_ids[i], *rect)

    def _get_corner_coords(self, scaled_rect: list[int]) -> list[tuple[int, int, int, int]]:
        """スケール後の矩形座標から4隅のハンドル矩形座標を計算."""
        cx1, cy1, cx2, cy2 = scaled_rect
        size = self.HANDLE_SIZE
        half = size // 2
        return [
            (cx1 - half, cy1 - half, cx1 + half, cy1 + half),
            (cx2 - half, cy1 - half, cx2 + half, cy1 + half),
            (cx1 - half, cy2 - half, cx1 + half, cy2 + half),
            (cx2 - half, cy2 - half, cx2 + half, cy2 + half),
        ]

    def _sync_crop_rect_ui(self) -> None:
        """クロップ矩形をキャンバスに反映し、ハンドルとエントリを更新."""
        scaled_rect = self._scaled_rect_from_crop()
        self.canvas.coords(self.rect_id, *scaled_rect)
        self._update_corner_handles(self._get_corner_coords(scaled_rect))
        self.update_crop_entries()

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
            sx-8, bar_y-8, sx+8, bar_y-8, sx, bar_y, fill="#00aa00", outline="black")
        self.seek_canvas.create_text(
            sx, bar_y-20, text="Start", fill="#006600", font=("Arial", 8))

        # End Marker (Bottom)
        ex = self.get_x(self.end_time)
        self.seek_canvas.create_polygon(
            ex-8, bar_y+8, ex+8, bar_y+8, ex, bar_y, fill="#cc0000", outline="black")
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
        edges = self.near_edge(e.x, e.y)
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
                self.canvas.itemconfig(self.rect_id, outline='orange')
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
                    self.canvas.itemconfig(self.rect_id, outline='red')
                except Exception:
                    pass

    def _start_dragging(self, e):
        try:
            self.push_crop_history()
        except Exception:
            pass
        img_x, img_y = self.canvas_mouse_to_image_coords(e.x, e.y)
        self.drag_offset = (img_x - self.crop_rect[0], img_y - self.crop_rect[1])
        self.rect_focused = True
        try:
            self.canvas.itemconfig(self.rect_id, outline='orange')
        except Exception:
            pass


    def on_mouse_drag(self, e):
        if self.resizing_rect:
            # マウス座標を画像座標に変換
            img_x, img_y = self.canvas_mouse_to_image_coords(e.x, e.y)

            x1, y1, x2, y2 = self.orig_rect
            
            # Ctrl キー押下時の対称リサイズ判定
            try:
                ctrl_held = (e.state & 0x4) != 0
            except Exception:
                ctrl_held = False

            if ctrl_held:
                # 対称リサイズ: 反対側も同じ分だけ動かす
                if self.resize_edge["l"]:
                    dx = img_x - x1
                    x1 = img_x
                    x2 = x2 - dx
                elif self.resize_edge["r"]:
                    dx = img_x - x2
                    x2 = img_x
                    x1 = x1 - dx
                
                if self.resize_edge["t"]:
                    dy = img_y - y1
                    y1 = img_y
                    y2 = y2 - dy
                elif self.resize_edge["b"]:
                    dy = img_y - y2
                    y2 = img_y
                    y1 = y1 - dy
            else:
                # 通常のリサイズ
                if self.resize_edge["l"]:
                    x1 = img_x
                if self.resize_edge["r"]:
                    x2 = img_x
                if self.resize_edge["t"]:
                    y1 = img_y
                if self.resize_edge["b"]:
                    y2 = img_y

            # Shift キー押下時のアスペクト比ロック
            try:
                self.maintain_aspect_ratio = (e.state & 0x1) != 0
            except Exception:
                pass
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
            # 矩形のフォーカス色を維持
            try:
                self.canvas.itemconfig(self.rect_id, outline='orange' if self.rect_focused else 'red')
            except Exception:
                pass
            # update corner handles
            self._update_corner_handles(self._get_corner_coords(scaled_rect))
            self.update_crop_entries()
        elif self.dragging_rect:
            # マウス座標を画像座標に変換
            img_x, img_y = self.canvas_mouse_to_image_coords(e.x, e.y)

            w = self.crop_rect[2] - self.crop_rect[0]
            h = self.crop_rect[3] - self.crop_rect[1]
            nx = img_x - self.drag_offset[0]
            ny = img_y - self.drag_offset[1]
            self.crop_rect = self.clamp_rect_canvas([nx, ny, nx+w, ny+h])

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
            try:
                self.canvas.itemconfig(self.rect_id, outline='orange' if self.rect_focused else 'red')
            except Exception:
                pass
            try:
                self.canvas.itemconfig(self.rect_id, outline='orange' if self.rect_focused else 'red')
            except Exception:
                pass

    def on_mouse_up(self, e):
        self.dragging_rect = False
        self.resizing_rect = False
        # ドラッグ終了時にはアスペクト比ロックを解除（再度Shiftで有効）
        try:
            self.maintain_aspect_ratio = False
        except Exception:
            pass
        try:
            self.canvas.itemconfig(self.rect_id, outline='orange' if self.rect_focused else 'red')
        except Exception:
            pass

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
                self.canvas.itemconfig(self.rect_id, outline='orange')
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
            self.crop_rect = self.clamp_rect_canvas([nx, ny, nx+w, ny+h])

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
            step = 0.1
            if delta > 0:
                self.image_zoom = min(10.0, getattr(self, 'image_zoom', 1.0) * (1.0 + step))
            else:
                self.image_zoom = max(0.1, getattr(self, 'image_zoom', 1.0) * (1.0 - step))
            if self.frame is not None:
                self.update_canvas_image()
        except Exception:
            pass

    def on_canvas_motion(self, e):
        # マウス移動時にカーソルを変更し、ハンドルをハイライトする
        # 角優先で判定
        edges = self.near_edge(e.x, e.y, m=10)
        res_lock = self.lock_var.get()
        move_lock = self.lock_move_var.get()
        cursor = ""
        
        # corner detections
        if edges.get('l') and edges.get('t'):
            cursor = 'circle_slash' if res_lock else 'top_left_corner'
        elif edges.get('r') and edges.get('t'):
            cursor = 'circle_slash' if res_lock else 'top_right_corner'
        elif edges.get('l') and edges.get('b'):
            cursor = 'circle_slash' if res_lock else 'bottom_left_corner'
        elif edges.get('r') and edges.get('b'):
            cursor = 'circle_slash' if res_lock else 'bottom_right_corner'
        else:
            # edges only
            if edges.get('l') or edges.get('r'):
                cursor = 'circle_slash' if res_lock else 'sb_h_double_arrow'
            elif edges.get('t') or edges.get('b'):
                cursor = 'circle_slash' if res_lock else 'sb_v_double_arrow'
            elif self.inside_rect(e.x, e.y):
                cursor = 'circle_slash' if move_lock else 'fleur'
            else:
                cursor = ''

        try:
            self.canvas.config(cursor=cursor)
        except Exception:
            pass

        # ハンドルのハイライト
        if hasattr(self, 'corner_ids'):
            # determine which corner (if any)
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
                color = 'yellow' if i == highlight_idx else 'red'
                try:
                    self.canvas.itemconfig(cid, fill=color)
                except Exception:
                    pass

    # ------------------ 出力処理 ------------------
    def open_folder(self, path):
        """プラットフォーム依存でフォルダを開く"""
        try:
            if os.name == 'nt':
                os.startfile(path)
            elif sys.platform == 'darwin':
                subprocess.Popen(['open', path])
            else:
                # Linux 等
                subprocess.Popen(['xdg-open', path])
        except Exception as e:
            messagebox.showerror("Error", f"フォルダを開けませんでした:\n{e}")

    def open_video_settings(self):
        """当該動画に紐づく設定ファイルを既定のプログラムで開く"""
        if not self.video_filepath:
            messagebox.showinfo("Info", "設定ファイルを開く対象の動画が選択されていません")
            return
        settings_path = os.path.splitext(self.video_filepath)[0] + '.settings.json'
        if not os.path.exists(settings_path):
            messagebox.showinfo("Info", f"設定ファイルが見つかりません:\n{settings_path}")
            return
        try:
            if os.name == 'nt':
                os.startfile(settings_path)
            elif sys.platform == 'darwin':
                subprocess.Popen(['open', settings_path])
            else:
                subprocess.Popen(['xdg-open', settings_path])
        except Exception as e:
            messagebox.showerror('Error', f'設定ファイルを開けませんでした:\n{e}')

    def export_png(self):
        if not self.cap:
            messagebox.showerror("Error", "動画なし")
            return
        video_name = os.path.basename(self.video_filepath) if self.video_filepath else None
        video_dir = os.path.dirname(self.video_filepath) if self.video_filepath else None
        base_dir = filedialog.askdirectory(initialdir=video_dir)
        if not base_dir:
            return

        # 実行日時のフォルダを作成
        now = time.strftime("%Y%m%d_%H%M%S")
        save_dir = os.path.join(base_dir, f"{video_name}_crops_{now}")
        os.makedirs(save_dir, exist_ok=True)

        # PNG圧縮レベルを設定
        save_params = [int(cv2.IMWRITE_PNG_COMPRESSION), self.png_compression]

        # 座標変換
        x1, y1, x2, y2 = self.crop_rect
        vid_w = int(self.cap.get(cv2.CAP_PROP_FRAME_WIDTH))
        vid_h = int(self.cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
        sx = vid_w / self.CANVAS_W
        sy = vid_h / self.CANVAS_H
        vx1, vy1 = int(x1*sx), int(y1*sy)
        vx2, vy2 = int(x2*sx), int(y2*sy)

        was_playing = self.playing
        self.playing = False
        if self._play_after_id:
            self.root.after_cancel(self._play_after_id)

        # Create modal progress dialog
        try:
            progress_win = tk.Toplevel(self.root)
            progress_win.title("PNG 書き出し...")
            progress_win.transient(self.root)
            progress_win.attributes("-topmost", True)
            progress_win.grab_set()
            progress_win.resizable(False, False)
            
            # ウィンドウを一時的に隠してサイズ計算
            progress_win.withdraw()
            
            tk.Label(progress_win, text="PNG を出力中...").pack(padx=20, pady=(15, 5))
            pb = ttk.Progressbar(progress_win, orient=tk.HORIZONTAL, length=400, mode='determinate')
            pb.pack(padx=20, pady=(0, 10))
            prog_label = tk.Label(progress_win, text="0 / 0")
            prog_label.pack(padx=20, pady=(0, 15))
            
            # メインウィンドウの中央に配置
            progress_win.update_idletasks()
            win_w = progress_win.winfo_width()
            win_h = progress_win.winfo_height()
            root_x = self.root.winfo_x()
            root_y = self.root.winfo_y()
            root_w = self.root.winfo_width()
            root_h = self.root.winfo_height()
            pos_x = root_x + (root_w // 2) - (win_w // 2)
            pos_y = root_y + (root_h // 2) - (win_h // 2)
            progress_win.geometry(f"+{pos_x}+{pos_y}")
            progress_win.deiconify()
        except Exception:
            progress_win = None
            pb = None
            prog_label = None

        try:
            count = 0
            t = self.start_time
            limit = self.end_time
            frame_interval = 1.0 / self.fps
            # estimate total steps for progressbar
            try:
                total_steps = max(0, int((limit - t) / frame_interval) + 1)
            except Exception:
                total_steps = 0
            if pb is not None and total_steps > 0:
                pb['maximum'] = total_steps
                try:
                    prog_label.config(text=f"0 / {total_steps}")
                except Exception:
                    pass
            prev_crop = None
            next_crop = None
            last_saved_crop = None

            step_idx = 0
            while t <= limit:
                self.cap.set(cv2.CAP_PROP_POS_MSEC, t*1000)
                ret, frm = self.cap.read()
                if ret and frm is not None:
                    crop = frm[vy1:vy2, vx1:vx2]
                    if crop.size > 0:
                        # 次のフレームを取得
                        next_t = t + frame_interval
                        if next_t <= limit:
                            self.cap.set(cv2.CAP_PROP_POS_MSEC, next_t*1000)
                            ret_next, frm_next = self.cap.read()
                            if ret_next and frm_next is not None:
                                next_crop = frm_next[vy1:vy2, vx1:vx2]
                            else:
                                next_crop = None
                        else:
                            next_crop = None

                        # 前のフレーム、現在のフレーム、次のフレームが全て同じかチェック
                        is_matches_prev_next = False
                        if self.check_prev_next.get():
                            # チェック有効時：前後のフレーム比較を行う
                            if prev_crop is not None and next_crop is not None:
                                # 差分を計算してグレースケールに変換
                                diff1 = cv2.cvtColor(cv2.absdiff(
                                    prev_crop, crop), cv2.COLOR_BGR2GRAY)
                                diff2 = cv2.cvtColor(cv2.absdiff(
                                    crop, next_crop), cv2.COLOR_BGR2GRAY)
                                if (cv2.countNonZero(diff1) == 0 and
                                        cv2.countNonZero(diff2) == 0):
                                    is_matches_prev_next = True
                            elif prev_crop is None and next_crop is not None:
                                # 最初のフレーム：現在のフレームと次のフレームが同じかチェック
                                diff2 = cv2.cvtColor(cv2.absdiff(
                                    crop, next_crop), cv2.COLOR_BGR2GRAY)
                                if cv2.countNonZero(diff2) == 0:
                                    is_matches_prev_next = True
                            elif prev_crop is not None and next_crop is None:
                                # 最後のフレーム：前のフレームと現在のフレームが同じかチェック
                                diff1 = cv2.cvtColor(cv2.absdiff(
                                    prev_crop, crop), cv2.COLOR_BGR2GRAY)
                                if cv2.countNonZero(diff1) == 0:
                                    is_matches_prev_next = True
                        else:
                            # チェック無効時は常に真（フィルタリングしない）
                            is_matches_prev_next = True

                        # 直前に出力したフレームとも比較
                        is_same_as_last_saved = False
                        if self.check_duplicate.get() and last_saved_crop is not None:
                            diff_last = cv2.cvtColor(cv2.absdiff(
                                last_saved_crop, crop), cv2.COLOR_BGR2GRAY)
                            if cv2.countNonZero(diff_last) == 0:
                                is_same_as_last_saved = True

                        # チェックボックスの設定に応じて出力判定
                        if is_matches_prev_next and not is_same_as_last_saved:
                            time_str = self.sec_to_hhmmss(t)
                            # 実際の時間に基づいたフレーム番号を計算
                            frame_in_sec = int((t - int(t)) * self.fps)
                            filepath = os.path.join(save_dir, f"{self.video_filename}_{time_str}_{frame_in_sec:03d}.png")
                            # 日本語パス対応の画像保存関数を使用
                            self.imwrite_jp(filepath, crop, params=save_params)
                            last_saved_crop = crop.copy()
                            count += 1

                        prev_crop = crop.copy()
                t += frame_interval
                # update progress
                step_idx += 1
                if pb is not None:
                    try:
                        pb['value'] = step_idx
                        prog_label.config(text=f"{step_idx} / {total_steps}")
                        progress_win.update()
                    except Exception:
                        pass

            # 設定を保存（プリセットなども含めてマージ保存）
            try:
                self.save_config()
            except Exception:
                pass
            
            # 完了ダイアログの前にウィンドウを消す
            try:
                if progress_win is not None:
                    progress_win.grab_release()
                    progress_win.destroy()
                    progress_win = None
            except Exception:
                pass

            # 完了ダイアログとフォルダを開くかの確認
            open_now = messagebox.askyesno("完了", f"{count} images saved.\nフォルダを開きますか？")
            if open_now:
                self.open_folder(save_dir)
        # except Exception as e:
        #     messagebox.showerror("Err", str(e))
        finally:
            # Destroy progress dialog
            try:
                if progress_win is not None:
                    progress_win.grab_release()
                    progress_win.destroy()
            except Exception:
                pass
            self.playing = was_playing
            if self.playing:
                self.play_step()

    def export_video(self):
        """赤枠範囲をstart時間からend時間まで動画ファイルとして出力"""
        if not self.cap:
            messagebox.showerror("Error", "動画なし")
            return
        
        # 保存先とファイル名をユーザーに選ばせる（初期ファイル名を生成）
        start_h = self.sec_to_hhmmss(int(self.start_time))
        start_frame = int((self.start_time - int(self.start_time)) * max(1, self.fps))
        end_h = self.sec_to_hhmmss(int(self.end_time))
        end_frame = int((self.end_time - int(self.end_time)) * max(1, self.fps))
        default_name = f"{self.video_filename}_trim_{start_h}_{start_frame:03d}_{end_h}_{end_frame:03d}.mp4"
        video_dir = os.path.dirname(self.video_filepath) if self.video_filepath else None
        save_path = filedialog.asksaveasfilename(defaultextension='.mp4', initialfile=default_name,
                             initialdir=video_dir,
                             filetypes=[('MP4', '*.mp4')], title='保存先とファイル名を選択')
        if not save_path:
            return
        save_dir = os.path.dirname(save_path)

        # 座標変換（crop_rectは640x360基準、実際の動画サイズに変換）
        x1, y1, x2, y2 = self.crop_rect
        vid_w = int(self.cap.get(cv2.CAP_PROP_FRAME_WIDTH))
        vid_h = int(self.cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
        sx = vid_w / self.CANVAS_W
        sy = vid_h / self.CANVAS_H
        vx1, vy1 = int(x1*sx), int(y1*sy)
        vx2, vy2 = int(x2*sx), int(y2*sy)

        # クロップ後のサイズ
        crop_w = vx2 - vx1
        crop_h = vy2 - vy1

        if crop_w <= 0 or crop_h <= 0:
            messagebox.showerror("Error", "クロップ範囲が無効です")
            return

        # 再生を一時停止
        was_playing = self.playing
        self.playing = False
        if self._play_after_id:
            self.root.after_cancel(self._play_after_id)

        try:
            # VideoWriter の設定
            fourcc = cv2.VideoWriter_fourcc(*'mp4v')  # mp4v コーデック
            out = cv2.VideoWriter(save_path, fourcc, self.fps, (crop_w, crop_h))

            if not out.isOpened():
                messagebox.showerror("Error", "動画ファイルを作成できませんでした")
                return

            # start_time から end_time までのフレームを処理
            t = self.start_time
            limit = self.end_time
            frame_interval = 1.0 / self.fps
            frame_count = 0

            while t <= limit:
                self.cap.set(cv2.CAP_PROP_POS_MSEC, t*1000)
                ret, frm = self.cap.read()
                if ret and frm is not None:
                    # クロップ
                    crop = frm[vy1:vy2, vx1:vx2]
                    if crop.size > 0:
                        out.write(crop)
                        frame_count += 1
                t += frame_interval

            out.release()

            # 完了ダイアログとフォルダを開くかの確認
            open_now = messagebox.askyesno("完了", f"動画を保存しました。\n{frame_count} フレーム\nフォルダを開きますか？")
            if open_now:
                self.open_folder(save_dir)
        except Exception as e:
            messagebox.showerror("Error", f"動画保存中にエラーが発生しました:\n{e}")
        finally:
            self.playing = was_playing
            if self.playing:
                self.play_step()


if __name__ == "__main__":
    root = tk.Tk()
    VideoCropperApp(root)
    root.mainloop()
