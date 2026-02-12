import tkinter as tk
from tkinter import ttk, colorchooser, font as tkfont
from typing import Dict, Any, Callable, List, Tuple
from ui_utils import add_tooltip, fix_button_active_colors
import overlay_utils
from PIL import Image, ImageTk, ImageDraw
import copy
import time

class InputOverlaySettingsDialog:
    def __init__(self, app: Any, config: Dict[str, Any], on_update: Callable[[], None]):
        self.app = app
        self.parent = app.root
        self.original_config = config
        # 設定をディープコピーして編集用に保持（キャンセル時に破棄できる）
        self.config = copy.deepcopy(config)
        self.theme = self.config.get("theme", {})
        self.on_update = on_update

        # マウスオーバレイ設定を取得 (なければデフォルト作成)
        if "mouse_overlay" not in self.theme:
            self.theme["mouse_overlay"] = {}
        self.m_cfg = self.theme["mouse_overlay"]
        if "pointer" not in self.m_cfg:
            self.m_cfg["pointer"] = {}
        for btn in ["click_left", "click_right", "click_middle"]:
            if btn not in self.m_cfg:
                self.m_cfg[btn] = {}

        # 入力字幕設定を取得
        if "input_overlay" not in self.theme:
            self.theme["input_overlay"] = {}
        self.k_cfg = self.theme["input_overlay"]

        # プレビュー用の状態
        self.preview_images = {} # tab_id -> PhotoImage
        self.preview_canvases = {} # tab_id -> Canvas
        
        # マウステスト用状態
        self.test_mouse_pos = (150, 100)
        self.test_mouse_click = ""
        self.test_ripple_timer = 0
        self.test_ripple_type = ""

        # 入力字幕テスト用
        self.var_test_subtitle = tk.BooleanVar(value=False)
        self.test_input_manager = overlay_utils.InputHistoryManager()
        self.subtitle_settings_frames = [] # ロック対象のフレームリスト

        # テキスト縦余白
        self.var_sub_pad_y = tk.IntVar(value=self.k_cfg.get("padding_y", 5))

        # テーマプレビューのインタラクティブ状態
        self.test_crop_rect = [100, 60, 300, 180] # [x1, y1, x2, y2]
        self.test_drag_mode = None # "move", "resize_n", etc.
        self.test_drag_start = (0, 0)
        self.test_orig_rect = None

        # プレビュー更新タイマーID
        self._preview_timer_id = None

        self._create_dialog()
        self._refresh_previews()

    def _create_dialog(self):
        root = tk.Toplevel(self.parent)
        root.title("詳細設定")
        root.transient(self.parent)
        root.grab_set()

        # ウィンドウサイズと位置 (2カラム化に伴い少し幅広に)
        w, h = 800, 850
        root.geometry(f"{w}x{h}")
        root.resizable(False, False)
        # 親の中央に
        self.parent.update_idletasks() # 親ウィンドウの最新情報を取得
        root.update_idletasks()        # 自身の情報を確定
        
        rx = self.parent.winfo_rootx()
        ry = self.parent.winfo_rooty()
        rw = self.parent.winfo_width()
        rh = self.parent.winfo_height()
        
        bx = rx + (rw - w) // 2
        by = ry + (rh - h) // 2
        root.geometry(f"+{bx}+{by}")
        
        self.dialog_root = root

        # 大きめのフォント設定
        self.large_font = ("Meiryo UI", 10)
        self.btn_font = ("Meiryo UI", 11, "bold")

        # ノートブック（タブ）作成
        notebook = ttk.Notebook(root)
        notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # タブ1: マウスポインタ（波紋設定を統合）
        tab_mouse = ttk.Frame(notebook)
        notebook.add(tab_mouse, text="マウスポインタ")
        self._build_mouse_tab(tab_mouse)

        # タブ2: 入力字幕
        tab_subtitle = ttk.Frame(notebook)
        notebook.add(tab_subtitle, text="入力字幕")
        self._build_subtitle_tab(tab_subtitle)

        # タブ3: テーマ
        tab_theme = ttk.Frame(notebook)
        notebook.add(tab_theme, text="テーマ")
        self._build_theme_tab(tab_theme)

        # 下部ボタン
        frame_btns = tk.Frame(root)
        frame_btns.pack(fill=tk.X, padx=10, pady=10)

        btn_cancel = tk.Button(frame_btns, text="キャンセル", width=12, height=2, font=self.btn_font, command=self._on_cancel)
        btn_cancel.pack(side=tk.RIGHT, padx=5)

        btn_apply = tk.Button(frame_btns, text="保存", width=15, height=2, font=self.btn_font, bg="#d0f0d0", command=self._on_apply)
        btn_apply.pack(side=tk.RIGHT, padx=5)
        
        fix_button_active_colors(frame_btns)
        
        root.protocol("WM_DELETE_WINDOW", self._on_cancel)
        root.focus_set()
        self.parent.wait_window(root)

    def _create_preview_canvas(self, parent, tab_id: str):
        """タブ内にプレビュー用のキャンバスを作成する"""
        lbl = tk.Label(parent, text="プレビュー", font=self.large_font)
        lbl.pack(anchor=tk.W, padx=10, pady=(10, 0))
        
        # 16:9 の比率に設定
        cv_h = 225
        cv_w = 400
        if tab_id == "subtitle":
            cv_w, cv_h = 711, 400
            
        cv = tk.Canvas(parent, width=cv_w, height=cv_h, bg="#444444", highlightthickness=1, highlightbackground="gray")
        cv.pack(padx=10, pady=5)
        self.preview_canvases[tab_id] = cv
        
        # マウスタブ/波紋タブ用
        if tab_id in ["mouse", "ripple"]:
            cv.bind("<Motion>", lambda e: self._on_mouse_test(e, "move"))
            cv.bind("<ButtonPress-1>", lambda e: self._on_mouse_test(e, "L"))
            cv.bind("<ButtonPress-2>", lambda e: self._on_mouse_test(e, "M"))
            cv.bind("<ButtonPress-3>", lambda e: self._on_mouse_test(e, "R"))
            cv.bind("<ButtonRelease>", lambda e: self._on_mouse_test(e, "release"))
        # テーマタブ用 (インタラクティブ矩形)
        elif tab_id == "theme":
            cv.bind("<ButtonPress-1>", self._on_theme_mouse_down)
            cv.bind("<B1-Motion>", self._on_theme_mouse_drag)
            cv.bind("<ButtonRelease-1>", self._on_theme_mouse_up)
            cv.bind("<Motion>", self._on_theme_mouse_move) # カーソル形状変更用

    def _build_mouse_tab(self, parent):
        self._create_preview_canvas(parent, "mouse")
        
        # 2カラムレイアウト用フレーム
        two_col_frame = tk.Frame(parent)
        two_col_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # 左列: ポインタ本体
        left_col = tk.Frame(two_col_frame)
        left_col.pack(side=tk.LEFT, fill=tk.BOTH, expand=False, padx=(0, 5))
        
        # 右列: クリック波紋
        right_col = tk.Frame(two_col_frame)
        right_col.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(5, 0))

        # --- 左列: ポインタ本体 ---
        lbl_frame = tk.LabelFrame(left_col, text="ポインタ本体", padx=10, pady=10, font=self.large_font)
        lbl_frame.pack(fill=tk.X, pady=5)

        # 形状
        tk.Label(lbl_frame, text="形状:", font=self.large_font).grid(row=0, column=0, sticky=tk.W, pady=5)
        self.var_p_shape = tk.StringVar(value=self.m_cfg.get("pointer", {}).get("shape", "cursor"))
        cb_shape = ttk.Combobox(lbl_frame, values=["cursor", "circle", "square", "star", "rote_star"], textvariable=self.var_p_shape, state="readonly", width=12, font=self.large_font)
        cb_shape.grid(row=0, column=1, sticky=tk.W, padx=10)
        cb_shape.bind("<<ComboboxSelected>>", self._on_ui_change)

        # サイズ
        tk.Label(lbl_frame, text="サイズ:", font=self.large_font).grid(row=1, column=0, sticky=tk.W, pady=5)
        self.var_p_size = tk.IntVar(value=self.m_cfg.get("pointer", {}).get("radius", 6))
        sp = self._create_spinbox(lbl_frame, 1, 100, self.var_p_size, width=4)
        sp.grid(row=1, column=1, sticky=tk.W, padx=10)

        # 枠線幅
        tk.Label(lbl_frame, text="枠線幅:", font=self.large_font).grid(row=2, column=0, sticky=tk.W, pady=5)
        self.var_p_width = tk.IntVar(value=self.m_cfg.get("pointer", {}).get("width", 2))
        sp_w = self._create_spinbox(lbl_frame, 0, 20, self.var_p_width, width=4)
        sp_w.grid(row=2, column=1, sticky=tk.W, padx=10)

        # 色設定
        tk.Label(lbl_frame, text="枠線:", font=self.large_font).grid(row=3, column=0, sticky=tk.W, pady=5)
        f_col = tk.Frame(lbl_frame)
        f_col.grid(row=3, column=1, sticky=tk.W, padx=10)
        self.btn_p_color = tk.Button(f_col, text=" ", width=4, height=1, font=self.btn_font, command=lambda: self._pick_color("pointer", "color"))
        self.btn_p_color.pack(side=tk.LEFT)
        tk.Button(f_col, text="無", font=("Meiryo UI", 8), command=lambda: self._clear_color("pointer", "color")).pack(side=tk.LEFT, padx=5)
        self._update_color_btn(self.btn_p_color, self.m_cfg.get("pointer", {}).get("color", "#000000"))
        
        tk.Label(lbl_frame, text="塗り:", font=self.large_font).grid(row=4, column=0, sticky=tk.W, pady=5)
        f_fill = tk.Frame(lbl_frame)
        f_fill.grid(row=4, column=1, sticky=tk.W, padx=10)
        self.btn_p_fill = tk.Button(f_fill, text=" ", width=4, height=1, font=self.btn_font, command=lambda: self._pick_color("pointer", "fill"))
        self.btn_p_fill.pack(side=tk.LEFT)
        tk.Button(f_fill, text="無", font=("Meiryo UI", 8), command=lambda: self._clear_color("pointer", "fill")).pack(side=tk.LEFT, padx=5)
        self._update_color_btn(self.btn_p_fill, self.m_cfg.get("pointer", {}).get("fill", "#FFFFFF"))

        # --- 右列: クリック波紋 ---
        self.ripple_shape_vars = {}
        ripple_frame = tk.Frame(right_col)
        ripple_frame.pack(fill=tk.BOTH, expand=True)

        for btn_type, label in [("click_left", "左"), ("click_right", "右"), ("click_middle", "中")]:
            lf = tk.LabelFrame(ripple_frame, text=label, padx=10, pady=5, font=self.large_font)
            lf.pack(fill=tk.X, pady=2)

            # 1行に集約
            row = tk.Frame(lf)
            row.pack(fill=tk.X)
            
            # 色
            tk.Label(row, text="色:", font=self.large_font).pack(side=tk.LEFT)
            btn_col = tk.Button(row, text=" ", width=3, height=1, font=self.btn_font)
            btn_col.config(command=lambda t=btn_type, b=btn_col: self._pick_ripple_color(t, b))
            btn_col.pack(side=tk.LEFT, padx=2)
            default_col = "#FF0000" if "left" in btn_type else ("#0000FF" if "right" in btn_type else "#00FF00")
            current_col = self.m_cfg.get(btn_type, {}).get("color", default_col)
            self._update_color_btn(btn_col, current_col)

            # 形状
            tk.Label(row, text="形:", font=self.large_font).pack(side=tk.LEFT, padx=(5, 0))
            sv = tk.StringVar(value=self.m_cfg.get(btn_type, {}).get("shape", "circle"))
            self.ripple_shape_vars[btn_type] = sv
            cb_s = ttk.Combobox(row, values=["circle", "square", "star", "rote_star"], textvariable=sv, state="readonly", width=8, font=("Meiryo UI", 9))
            cb_s.pack(side=tk.LEFT, padx=2)
            cb_s.bind("<<ComboboxSelected>>", self._on_ui_change)

            # サイズ (旧: 大)
            tk.Label(row, text="サイズ:", font=self.large_font).pack(side=tk.LEFT, padx=(5, 0))
            var_size = tk.IntVar(value=self.m_cfg.get(btn_type, {}).get("ripple_range", 20))
            sp_size = self._create_spinbox(row, 0, 200, var_size, width=3)
            sp_size.config(command=lambda t=btn_type, v=var_size: self._update_ripple_val(t, "ripple_range", v))
            sp_size.pack(side=tk.LEFT, padx=2)
            sp_size.bind("<KeyRelease>", lambda e, t=btn_type, v=var_size: self._update_ripple_val(t, "ripple_range", v))

            # 秒
            tk.Label(row, text="秒:", font=self.large_font).pack(side=tk.LEFT, padx=(5, 0))
            var_dur = tk.DoubleVar(value=self.m_cfg.get(btn_type, {}).get("ripple_duration", 0.5))
            sp_dur = self._create_spinbox(row, 0.1, 10.0, var_dur, increment=0.1, width=3)
            sp_dur.config(command=lambda t=btn_type, v=var_dur: self._update_ripple_val(t, "ripple_duration", v))
            sp_dur.pack(side=tk.LEFT, padx=2)
            sp_dur.bind("<KeyRelease>", lambda e, t=btn_type, v=var_dur: self._update_ripple_val(t, "ripple_duration", v))

    def _build_theme_tab(self, parent):
        # テーマ設定（クロップ枠・ハンドル等）のタブ
        self._create_preview_canvas(parent, "theme")
        
        two_col_frame = tk.Frame(parent)
        two_col_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        left_col = tk.Frame(two_col_frame)
        left_col.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        
        right_col = tk.Frame(two_col_frame)
        right_col.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(5, 0))

        # --- 左列: クロップ枠 ---
        lf_crop = tk.LabelFrame(left_col, text="クロップ枠", padx=10, pady=10, font=self.large_font)
        lf_crop.pack(fill=tk.X, pady=5)

        crop_items = [
            ("通常時:", "crop_default_linecolor", "#FF8C00"),
            ("強調時:", "crop_focused_linecolor", "#FF0000"),
            ("ホバー時:", "crop_hover_linecolor", "#FF0000"),
        ]

        for i, (label, key, default) in enumerate(crop_items):
            tk.Label(lf_crop, text=label, font=self.large_font).grid(row=i, column=0, sticky=tk.W, pady=5)
            f_col = tk.Frame(lf_crop)
            f_col.grid(row=i, column=1, sticky=tk.W, padx=10)
            
            btn = tk.Button(f_col, text=" ", width=4, height=1, font=self.btn_font)
            btn.config(command=lambda k=key, b=btn: self._pick_theme_color(k, b))
            btn.pack(side=tk.LEFT)
            
            curr = self.theme.get(key, default)
            self._update_color_btn(btn, curr)

        # クロップ太さ
        tk.Label(lf_crop, text="枠線の太さ:", font=self.large_font).grid(row=len(crop_items), column=0, sticky=tk.W, pady=5)
        self.var_crop_width = tk.IntVar(value=self.theme.get("crop_linewidth", 2))
        sp_cw = self._create_spinbox(lf_crop, 1, 10, self.var_crop_width)
        sp_cw.grid(row=len(crop_items), column=1, sticky=tk.W, padx=10)
        sp_cw.config(command=lambda: self._update_theme_val("crop_linewidth", self.var_crop_width))

        # --- 右列: ハンドル・背景 ---
        lf_handle = tk.LabelFrame(right_col, text="ハンドル（四隅）・その他", padx=10, pady=10, font=self.large_font)
        lf_handle.pack(fill=tk.X, pady=5)

        # ハンドル色
        tk.Label(lf_handle, text="ハンドル色:", font=self.large_font).grid(row=0, column=0, sticky=tk.W, pady=5)
        f_hcol = tk.Frame(lf_handle)
        f_hcol.grid(row=0, column=1, sticky=tk.W, padx=10)
        btn_h = tk.Button(f_hcol, text=" ", width=4, height=1, font=self.btn_font)
        btn_h.config(command=lambda: self._pick_theme_color("handle_color", btn_h))
        btn_h.pack(side=tk.LEFT)
        self._update_color_btn(btn_h, self.theme.get("handle_color", "#FFFFFF"))

        # ハンドルサイズ
        tk.Label(lf_handle, text="ハンドル寸:", font=self.large_font).grid(row=1, column=0, sticky=tk.W, pady=5)
        self.var_h_size = tk.IntVar(value=self.theme.get("handle_size", 8))
        sp_hs = self._create_spinbox(lf_handle, 2, 30, self.var_h_size)
        sp_hs.grid(row=1, column=1, sticky=tk.W, padx=10)
        sp_hs.config(command=lambda: self._update_theme_val("handle_size", self.var_h_size))

        # キャンバス背景
        tk.Label(lf_handle, text="背景色:", font=self.large_font).grid(row=2, column=0, sticky=tk.W, pady=5)
        f_bgcol = tk.Frame(lf_handle)
        f_bgcol.grid(row=2, column=1, sticky=tk.W, padx=10)
        btn_bg = tk.Button(f_bgcol, text=" ", width=4, height=1, font=self.btn_font)
        btn_bg.config(command=lambda: self._pick_theme_color("canvas_bg", btn_bg))
        btn_bg.pack(side=tk.LEFT)
        self._update_color_btn(btn_bg, self.theme.get("canvas_bg", "#f5f5f5"))

    def _pick_theme_color(self, key, btn):
        curr = self.theme.get(key, "#FFFFFF")
        if not curr: curr = "#FFFFFF"
        new_col = colorchooser.askcolor(initialcolor=curr, title=f"色選択: {key}", parent=self.dialog_root)[1]
        if new_col:
            self.theme[key] = new_col
            self._update_color_btn(btn, new_col)
            self._on_ui_change()

    def _update_theme_val(self, key, var):
        try:
            self.theme[key] = var.get()
            self._on_ui_change()
        except: pass

    def _build_subtitle_tab(self, parent):
        self._create_preview_canvas(parent, "subtitle")
        
        # 左右2ライン化のためのメインフレーム
        two_col_frame = tk.Frame(parent)
        two_col_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        self.subtitle_settings_frames.append(two_col_frame)
        
        # 左列: 位置・レイアウト
        left_col = tk.Frame(two_col_frame)
        left_col.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        
        # 右列: フォント・色
        right_col = tk.Frame(two_col_frame)
        right_col.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(5, 0))

        # --- 左列 ---
        lf_disp = tk.LabelFrame(left_col, text="位置・レイアウト", padx=10, pady=10, font=self.large_font)
        lf_disp.pack(fill=tk.BOTH, expand=True)

        tk.Label(lf_disp, text="左右位置:", font=self.large_font).grid(row=0, column=0, sticky=tk.W, pady=5)
        self.var_sub_pos = tk.StringVar(value=self.k_cfg.get("position", "center"))
        cb_pos = ttk.Combobox(lf_disp, values=["left", "center", "right"], textvariable=self.var_sub_pos, state="readonly", width=12, font=self.large_font)
        cb_pos.grid(row=0, column=1, sticky=tk.W, padx=10)
        cb_pos.bind("<<ComboboxSelected>>", self._on_ui_change)

        tk.Label(lf_disp, text="上下位置:", font=self.large_font).grid(row=1, column=0, sticky=tk.W, pady=5)
        self.var_v_pos = tk.StringVar(value=self.k_cfg.get("v_position", "bottom"))
        cb_v_pos = ttk.Combobox(lf_disp, values=["top", "center", "bottom"], textvariable=self.var_v_pos, state="readonly", width=12, font=self.large_font)
        cb_v_pos.grid(row=1, column=1, sticky=tk.W, padx=10)
        cb_v_pos.bind("<<ComboboxSelected>>", self._on_ui_change)

        tk.Label(lf_disp, text="最大表示数:", font=self.large_font).grid(row=2, column=0, sticky=tk.W, pady=5)
        self.var_sub_stack = tk.IntVar(value=self.k_cfg.get("max_stack", 3))
        f_stack = tk.Frame(lf_disp)
        f_stack.grid(row=2, column=1, sticky=tk.W, padx=10)
        sp_stack = self._create_spinbox(f_stack, 1, 10, self.var_sub_stack)
        sp_stack.pack(side=tk.LEFT)
        
        # 検証(Test)ボタン
        self.btn_sub_test = tk.Checkbutton(f_stack, text="検証 (Test)", variable=self.var_test_subtitle, command=self._on_test_toggle, font=self.large_font, indicatoron=False, selectcolor="#ffefef", padx=10)
        self.btn_sub_test.pack(side=tk.LEFT, padx=(10, 0))
        add_tooltip(self.btn_sub_test, "ONにすると実際のキー入力やクリックを検知してプレビューに表示します（設定操作はロックされます）")

        tk.Label(lf_disp, text="表示時間(s):", font=self.large_font).grid(row=3, column=0, sticky=tk.W, pady=5)
        self.var_sub_dur = tk.DoubleVar(value=self.k_cfg.get("fade_duration", 1.0))
        sp_dur = self._create_spinbox(lf_disp, 0.1, 20.0, self.var_sub_dur, increment=0.1)
        sp_dur.grid(row=3, column=1, sticky=tk.W, padx=10)

        tk.Label(lf_disp, text="領域位置X:", font=self.large_font).grid(row=4, column=0, sticky=tk.W, pady=5)
        self.var_off_x = tk.IntVar(value=self.k_cfg.get("offset_x", 0))
        self._create_spinbox(lf_disp, -1000, 1000, self.var_off_x).grid(row=4, column=1, sticky=tk.W, padx=10)

        tk.Label(lf_disp, text="領域位置Y:", font=self.large_font).grid(row=5, column=0, sticky=tk.W, pady=5)
        self.var_off_y = tk.IntVar(value=self.k_cfg.get("offset_y", 0))
        self._create_spinbox(lf_disp, -1000, 1000, self.var_off_y).grid(row=5, column=1, sticky=tk.W, padx=10)

        tk.Label(lf_disp, text="文字中心Yオフセット:", font=self.large_font).grid(row=6, column=0, sticky=tk.W, pady=5)
        self.var_sub_y = tk.IntVar(value=self.k_cfg.get("text_offset_y", -2))
        self._create_spinbox(lf_disp, -100, 100, self.var_sub_y).grid(row=6, column=1, sticky=tk.W, padx=10)

        tk.Label(lf_disp, text="縦余白(px):", font=self.large_font).grid(row=7, column=0, sticky=tk.W, pady=5)
        self._create_spinbox(lf_disp, 0, 100, self.var_sub_pad_y).grid(row=7, column=1, sticky=tk.W, padx=10)

        # --- 右列 ---
        lf_font = tk.LabelFrame(right_col, text="フォント・色", padx=10, pady=10, font=self.large_font)
        lf_font.pack(fill=tk.BOTH, expand=True)

        tk.Label(lf_font, text="フォント:", font=self.large_font).grid(row=0, column=0, sticky=tk.W, pady=5)
        self.var_sub_font = tk.StringVar(value=self.k_cfg.get("font_family", "msgothic.ttc"))
        font_list = sorted(list(set(["msgothic.ttc", "meiryo.ttc", "arialbd.ttf", "BIZ-UDGothicR.ttc", self.var_sub_font.get()])))
        cb_font = ttk.Combobox(lf_font, values=font_list, textvariable=self.var_sub_font, width=15, font=self.large_font)
        cb_font.grid(row=0, column=1, sticky=tk.W, padx=10)
        cb_font.bind("<<ComboboxSelected>>", self._on_ui_change)
        cb_font.bind("<KeyRelease>", self._on_ui_change)

        tk.Label(lf_font, text="サイズ:", font=self.large_font).grid(row=1, column=0, sticky=tk.W, pady=5)
        self.var_f_size = tk.IntVar(value=self.k_cfg.get("font_size", 24))
        self._create_spinbox(lf_font, 8, 200, self.var_f_size).grid(row=1, column=1, sticky=tk.W, padx=10)

        row = 2
        for key, title, default in [
            ("font_color", "文字色:", "#000000"),
            ("bg_color", "背景色:", "#FFFFFF"),
            ("outline_color", "枠線色:", "")
        ]:
            tk.Label(lf_font, text=title, font=self.large_font).grid(row=row, column=0, sticky=tk.W, pady=5)
            f_col = tk.Frame(lf_font)
            f_col.grid(row=row, column=1, sticky=tk.W, padx=10)
            
            btn = tk.Button(f_col, text=" ", width=4, height=1, font=self.btn_font)
            btn.pack(side=tk.LEFT)
            tk.Button(f_col, text="なし", font=("Meiryo UI", 8), command=lambda k=key: self._clear_subtitle_color(k)).pack(side=tk.LEFT, padx=5)
            
            curr = self.k_cfg.get(key, default)
            self._update_color_btn(btn, curr)
            btn.config(command=lambda k=key, b=btn: self._pick_subtitle_color(k, b))
            
            if key == "font_color": self.btn_sub_font = btn
            elif key == "bg_color": self.btn_sub_bg = btn
            elif key == "outline_color": self.btn_sub_out = btn
            row += 1

    def _create_spinbox(self, parent, from_, to, var, increment=1, width=6):
        """強化版スピンボックス（マウスホイール対応）"""
        sp = tk.Spinbox(parent, from_=from_, to=to, increment=increment, textvariable=var, width=width, font=self.large_font, command=self._on_ui_change)
        
        def _on_wheel(event):
            delta = 1 if event.delta > 0 else -1
            if event.state & 0x1: # Shift押下
                delta *= 10
            try:
                cur = float(var.get())
                new_val = cur + delta * increment
                # 範囲内に制限 (クランプ)
                clamped_val = max(float(from_), min(float(to), new_val))
                
                # 型に合わせて設定 (重要: IntVarならintにしないと小数点が付く)
                if isinstance(var, tk.IntVar):
                    var.set(int(round(clamped_val)))
                else:
                    var.set(round(clamped_val, 2))
                
                self._on_ui_change()
            except: pass
            return "break"
        
        sp.bind("<MouseWheel>", _on_wheel)
        sp.bind("<KeyRelease>", self._on_ui_change)
        return sp

    def _refresh_previews(self):
        """全てのプレビューキャンバスを更新する"""
        if not self.dialog_root.winfo_exists(): return
        
        # 既存のタイマーがあればキャンセル（多重起動防止）
        if self._preview_timer_id is not None:
            self.dialog_root.after_cancel(self._preview_timer_id)
            self._preview_timer_id = None
        
        # キャンバスに描画
        for tab_id, cv in self.preview_canvases.items():
            cw = cv.winfo_width()
            ch = cv.winfo_height()
            # 初回などサイズ確定前はデフォルト値を使用
            if cw < 2:
                if tab_id == "subtitle": cw, ch = 711, 400
                else: cw, ch = 400, 225

            img = Image.new("RGBA", (cw, ch), (60, 60, 60, 255))
            
            if tab_id in ["mouse", "ripple"]:
                tx, ty = self.test_mouse_pos
                if self.test_ripple_timer > 0:
                    self.test_ripple_timer -= 0.02 # スローモーション (秒間 0.4 減算)
                else: self.test_ripple_type = ""
                
                overlay_utils.draw_mouse_overlay(
                    img, tx, ty, self.test_mouse_click,
                    1.0, 1.0, self.theme,
                    ripple_age=max(0, 0.5 - self.test_ripple_timer) if self.test_ripple_type else 0,
                    ripple_type=self.test_ripple_type
                )
            elif tab_id == "subtitle":
                if self.var_test_subtitle.get():
                    # テスト用入力マネージャから取得
                    fade_duration = self.var_sub_dur.get()
                    active_inputs = self.test_input_manager.get_active_inputs(time.time(), fade_duration)
                    dummy_theme = {"input_overlay": self.k_cfg, "mouse_overlay": self.m_cfg}
                    overlay_utils.draw_input_overlay(img, active_inputs, 1.0, 1.0, dummy_theme)
                else:
                    # 具体的なサンプル文字
                    sample_inputs = [
                        ("左クリック (最新)", 0.0),
                        ("Ctrl + C (1秒前)", 1.0),
                        ("Win + Shift + S (2秒前)", 2.0),
                        ("Sample (3秒前)", 3.0),
                    ]
                    dummy_theme = {"input_overlay": self.k_cfg, "mouse_overlay": self.m_cfg}
                    overlay_utils.draw_input_overlay(img, sample_inputs, 1.0, 1.0, dummy_theme)
            elif tab_id == "theme":
                # クロップ枠とハンドルのサンプル描画
                cw, ch = img.size
                x1, y1, x2, y2 = self.test_crop_rect
                
                draw = ImageDraw.Draw(img, "RGBA")
                # 背景色 (canvas_bg)
                bg_hex = self.theme.get("canvas_bg", "#444444")
                img.paste(overlay_utils._hex_to_rgba(bg_hex, 255), [0, 0, cw, ch])
                
                # 枠線 (focused 状態をサンプルとして描画)
                color = overlay_utils._hex_to_rgba(self.theme.get("crop_focused_linecolor", "#FF0000"), 255)
                width = self.theme.get("crop_linewidth", 2)
                draw.rectangle([x1, y1, x2, y2], outline=color, width=width)
                
                # ハンドル
                h_size = self.theme.get("handle_size", 8)
                h_color = overlay_utils._hex_to_rgba(self.theme.get("handle_color", "#FFFFFF"), 255)
                half = h_size // 2
                for hx, hy in [(x1, y1), (x2, y1), (x1, y2), (x2, y2)]:
                    draw.rectangle([hx - half, hy - half, hx + half, hy + half], fill=h_color)
                
                # 説明文字
                try:
                    from PIL import ImageFont
                    f = ImageFont.load_default()
                    draw.text((x1 + 10, y1 + 10), "Test Area", fill=(255,255,255,200), font=f)
                except: pass
            
            photo = ImageTk.PhotoImage(img)
            self.preview_images[tab_id] = photo
            cv.create_image(0, 0, anchor=tk.NW, image=photo)

        if self.test_ripple_timer > 0 or self.var_test_subtitle.get():
            self._preview_timer_id = self.dialog_root.after(50, self._refresh_previews)

    def _on_mouse_test(self, event, action):
        """プレビューキャンバス上でのマウス操作シミュレーション"""
        self.test_mouse_pos = (event.x, event.y)
        if action in ["L", "M", "R"]:
            self.test_mouse_click = action
            self.test_ripple_type = {"L":"left", "R":"right", "M":"middle"}.get(action, "")
            self.test_ripple_timer = 0.5
        elif action == "release":
            self.test_mouse_click = ""
        self._refresh_previews()

    # --- テーマプレビュー用インタラクティブ操作 ---
    def _on_theme_mouse_move(self, event):
        x1, y1, x2, y2 = self.test_crop_rect
        margin = self.theme.get("handle_size", 8) // 2 + 2
        mode = None
        if abs(event.x - x1) < margin and abs(event.y - y1) < margin: mode = "size_nw"
        elif abs(event.x - x2) < margin and abs(event.y - y1) < margin: mode = "size_ne"
        elif abs(event.x - x1) < margin and abs(event.y - y2) < margin: mode = "size_sw"
        elif abs(event.x - x2) < margin and abs(event.y - y2) < margin: mode = "size_se"
        elif x1 < event.x < x2 and y1 < event.y < y2: mode = "move"
        
        cursor_map = {"move": "fleur", "size_nw": "size_nw_se", "size_se": "size_nw_se", "size_ne": "size_ne_sw", "size_sw": "size_ne_sw"}
        cursor = cursor_map.get(mode, "")
        self.preview_canvases["theme"].config(cursor=cursor)

    def _on_theme_mouse_down(self, event):
        x1, y1, x2, y2 = self.test_crop_rect
        margin = self.theme.get("handle_size", 8) // 2 + 2
        self.test_drag_mode = None
        if abs(event.x - x1) < margin and abs(event.y - y1) < margin: self.test_drag_mode = "nw"
        elif abs(event.x - x2) < margin and abs(event.y - y1) < margin: self.test_drag_mode = "ne"
        elif abs(event.x - x1) < margin and abs(event.y - y2) < margin: self.test_drag_mode = "sw"
        elif abs(event.x - x2) < margin and abs(event.y - y2) < margin: self.test_drag_mode = "se"
        elif x1 < event.x < x2 and y1 < event.y < y2: self.test_drag_mode = "move"

        if self.test_drag_mode:
            self.test_drag_start = (event.x, event.y)
            self.test_orig_rect = list(self.test_crop_rect)

    def _on_theme_mouse_drag(self, event):
        if not self.test_drag_mode: return
        dx = event.x - self.test_drag_start[0]
        dy = event.y - self.test_drag_start[1]
        r = list(self.test_orig_rect)
        if self.test_drag_mode == "move":
            r[0]+=dx; r[1]+=dy; r[2]+=dx; r[3]+=dy
        elif self.test_drag_mode == "nw": r[0]+=dx; r[1]+=dy
        elif self.test_drag_mode == "ne": r[2]+=dx; r[1]+=dy
        elif self.test_drag_mode == "sw": r[0]+=dx; r[3]+=dy
        elif self.test_drag_mode == "se": r[2]+=dx; r[3]+=dy
        
        # Ensure valid rectangle (x1 < x2, y1 < y2)
        if r[0] > r[2]: r[0], r[2] = r[2], r[0]
        if r[1] > r[3]: r[1], r[3] = r[3], r[1]
            
        self.test_crop_rect = r
        self._refresh_previews()

    def _on_theme_mouse_up(self, event):
        self.test_drag_mode = None

    def _on_test_toggle(self):
        is_on = self.var_test_subtitle.get()
        state = tk.DISABLED if is_on else tk.NORMAL
        
        # 設定ウィジェットのロック/アンロック
        for frame in self.subtitle_settings_frames:
            self._set_widgets_state_recursive(frame, state)
        
        # 検証(Test)ボタン自体は有効にしておく
        self.btn_sub_test.configure(state=tk.NORMAL)

        if is_on:
            self.app.suppress_shortcuts = True
            # イベントバインド
            self.dialog_root.bind("<KeyPress>", self._on_test_any_input)
            self.dialog_root.bind("<ButtonPress>", self._on_test_any_input)
            # プレビュー表示のためにフォーカスを移す
            self.dialog_root.focus_set()
            # プレビュー更新ループを即座に開始
            self._refresh_previews()
        else:
            self.app.suppress_shortcuts = False
            # イベント解除
            self.dialog_root.unbind("<KeyPress>")
            self.dialog_root.unbind("<ButtonPress>")
            # 履歴をクリア
            self.test_input_manager.history.clear()
            self._on_ui_change()

    def _set_widgets_state_recursive(self, parent, state):
        """再帰的に子ウィジェットの状態を変更する"""
        for child in parent.winfo_children():
            # Testボタン自体は除外
            if child == self.btn_sub_test:
                continue
            try:
                # stateプロパティを持つウィジェットのみ変更
                child.configure(state=state)
            except:
                pass
            self._set_widgets_state_recursive(child, state)

    def _on_test_any_input(self, event):
        """テストモード時のキー入力/クリック検知"""
        if not self.var_test_subtitle.get():
            return
        
        # Testボタン自体への操作は、トグルを可能にするためにスルーする
        if event.widget == self.btn_sub_test:
            return

        click = "None"
        keys = "None"
        
        # 修飾キーの状態 (KeyPress/ButtonPress共通)
        # Windows環境では GetAsyncKeyState を使ったほうが確実
        import ctypes
        mods = []
        if ctypes.windll.user32.GetAsyncKeyState(0x11) & 0x8000: mods.append("Ctrl")
        if ctypes.windll.user32.GetAsyncKeyState(0x10) & 0x8000: mods.append("Shift")
        if ctypes.windll.user32.GetAsyncKeyState(0x12) & 0x8000: mods.append("Alt")
        if (ctypes.windll.user32.GetAsyncKeyState(0x5B) & 0x8000) or (ctypes.windll.user32.GetAsyncKeyState(0x5C) & 0x8000):
            mods.append("Win")
        
        # Tkinterのイベント状態もバックアップとして確認（念のため）
        if not mods:
            if event.state & 0x4: mods.append("Ctrl")
            if event.state & 0x1: mods.append("Shift")
            if event.state & (0x20000 | 0x8): mods.append("Alt")
        
        if event.type == tk.EventType.KeyPress:
            # TODO: モディファイアキー単体の場合は表示しないなどの判断は get_input_display_text 側に任せる
            keysym = event.keysym
            key_name = overlay_utils.convert_tk_event_to_key_name(keysym)
            
            # 修飾キー単体押下時もリストに含めるが、get_input_display_text がフィルタリングする
            keys = ",".join(mods + [key_name])
                
        elif event.type == tk.EventType.ButtonPress:
            if event.num == 1: click = "L"
            elif event.num == 2: click = "M"
            elif event.num == 3: click = "R"
            
            # マウスボタンの場合もキー修飾を渡す
            if mods:
                keys = ",".join(mods)
            
        text = overlay_utils.get_input_display_text(click, keys)
        if text:
            self.test_input_manager.update(time.time(), text)
            # 再描画は _refresh_previews のループで行われるが、即座に反映させるため呼ぶ
            self._refresh_previews()
        
        # 他のウィジェットにイベントを伝搬させない (ロック中なので)
        return "break"

    def _update_color_btn(self, btn, color_hex):
        if not color_hex:
            btn.config(bg="#f0f0f0", text="無", fg="black")
        else:
            try:
                r = int(color_hex[1:3], 16)
                g = int(color_hex[3:5], 16)
                b = int(color_hex[5:7], 16)
                lum = 0.299 * r + 0.587 * g + 0.114 * b
                fg = "white" if lum < 128 else "black"
                btn.config(bg=color_hex, text=" ", fg=fg)
            except:
                btn.config(bg="#FFFFFF", text="?", fg="black")

    def _clear_color(self, section, key):
        if section not in self.m_cfg: self.m_cfg[section] = {}
        self.m_cfg[section][key] = ""
        if key == "color": self._update_color_btn(self.btn_p_color, "")
        else: self._update_color_btn(self.btn_p_fill, "")
        self._on_ui_change()

    def _clear_subtitle_color(self, key):
        self.k_cfg[key] = ""
        if key == "font_color": self._update_color_btn(self.btn_sub_font, "")
        elif key == "bg_color": self._update_color_btn(self.btn_sub_bg, "")
        else: self._update_color_btn(self.btn_sub_out, "")
        self._on_ui_change()

    def _pick_color(self, section, key):
        current = self.m_cfg.get(section, {}).get(key, "#000000")
        if not current: current = "#FFFFFF"
        color = colorchooser.askcolor(current, title=f"色選択: {key}", parent=self.dialog_root)
        if color[1]:
            if section not in self.m_cfg: self.m_cfg[section] = {}
            self.m_cfg[section][key] = color[1]
            if key == "color": self._update_color_btn(self.btn_p_color, color[1])
            else: self._update_color_btn(self.btn_p_fill, color[1])
            self._on_ui_change()

    def _pick_ripple_color(self, btn_type, btn_widget):
        current = self.m_cfg.get(btn_type, {}).get("color", "#FF0000")
        color = colorchooser.askcolor(current, title=f"色選択: {btn_type}", parent=self.dialog_root)
        if color[1]:
            if btn_type not in self.m_cfg: self.m_cfg[btn_type] = {}
            self.m_cfg[btn_type]["color"] = color[1]
            self._update_color_btn(btn_widget, color[1])
            self._on_ui_change()

    def _update_ripple_val(self, btn_type, key, var):
        if btn_type not in self.m_cfg: self.m_cfg[btn_type] = {}
        try:
            val = float(var.get()) if isinstance(var, tk.DoubleVar) else int(var.get())
            self.m_cfg[btn_type][key] = val
            self._on_ui_change()
        except: pass

    def _pick_subtitle_color(self, key, btn_widget):
        current = self.k_cfg.get(key, "#000000")
        if not current: current = "#FFFFFF"
        color = colorchooser.askcolor(current, title=f"色選択: {key}", parent=self.dialog_root)
        if color[1]:
            self.k_cfg[key] = color[1]
            self._update_color_btn(btn_widget, color[1])
            self._on_ui_change()

    def _on_ui_change(self, event=None):
        """UI部品に変更があった際の内部状態更新（プレビュー反映のみ。即保存はしない。）」"""
        try:
            self.m_cfg["pointer"]["shape"] = self.var_p_shape.get()
            self.m_cfg["pointer"]["radius"] = int(self.var_p_size.get())
            self.m_cfg["pointer"]["width"] = int(self.var_p_width.get())
            
            for btn_type, var in self.ripple_shape_vars.items():
                if btn_type not in self.m_cfg: self.m_cfg[btn_type] = {}
                self.m_cfg[btn_type]["shape"] = var.get()

            self.k_cfg["position"] = self.var_sub_pos.get()
            self.k_cfg["v_position"] = self.var_v_pos.get()
            self.k_cfg["max_stack"] = self.var_sub_stack.get()
            self.k_cfg["fade_duration"] = self.var_sub_dur.get()
            self.k_cfg["offset_x"] = self.var_off_x.get()
            self.k_cfg["offset_y"] = self.var_off_y.get()
            self.k_cfg["text_offset_y"] = self.var_sub_y.get()
            self.k_cfg["padding_y"] = self.var_sub_pad_y.get()
            self.k_cfg["font_family"] = self.var_sub_font.get()
            self.k_cfg["font_size"] = self.var_f_size.get()
        except: pass

        self._refresh_previews()

    def _on_apply(self):
        """適用ボタン: 編集内容を元のオブジェクトに反映して閉じる"""
        self.app.suppress_shortcuts = False
        # 親の config を更新
        self.original_config["theme"] = self.config["theme"]
        # 反映（再描画など）を呼び出し
        if self.on_update: self.on_update()
        
        # ディスクに保存
        if hasattr(self.app, "save_config"):
            self.app.save_config()
            
        self.dialog_root.destroy()

    def _on_cancel(self):
        """キャンセルボタン: 何もせず閉じる"""
        self.app.suppress_shortcuts = False
        self.dialog_root.destroy()
