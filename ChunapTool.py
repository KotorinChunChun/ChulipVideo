"""ウィンドウ位置調整ツール.

対象ウィンドウの位置・サイズを調整するためのツールウィンドウ。
WGC または PrintWindow によるリアルタイムプレビュー、
座標の同期・画面外制限・最大化制御・Undo 機能を備える。
"""
import ctypes
import tkinter as tk
import tkinter.ttk as ttk
from typing import Any, Dict, List, Optional, Tuple

import cv2
from PIL import Image, ImageTk

from utils import resource_path
from wgc_capture import WGC_AVAILABLE, WGCCapture
from window_utils import WindowUtils

# ─── 型エイリアス ─────────────────────────────────────────────
# (hwnd, x, y, w, h, is_maximized)
GeoState = Tuple[Any, int, int, int, int, bool]

# ─── 定数 ─────────────────────────────────────────────────────
INITIAL_WIDTH = 640
INITIAL_HEIGHT = 240
MIN_WIDTH = 640
MIN_HEIGHT = 200
MIN_WINDOW_SIZE = 32          # ウィンドウ最小サイズ (px)
MAX_HISTORY = 100             # Undo 履歴の最大保持数
UI_POLL_INTERVAL_MS = 200     # UI 更新ループ間隔 (ms)
HISTORY_POLL_INTERVAL_MS = 500  # 履歴ポーリング間隔 (ms)
PREVIEW_WGC_INTERVAL_MS = 33  # WGC プレビュー間隔 (ms)
PREVIEW_FALLBACK_INTERVAL_MS = 200  # フォールバックプレビュー間隔 (ms)

# Win32 定数
SW_NORMAL = 1
SW_MAXIMIZE = 3
VK_LBUTTON = 0x01


class ChunapTool(tk.Tk):
    """ウィンドウ位置調整ツールのメインウィンドウ."""

    COLOR_BTN = "#e1f5fe"  # 水色ボタンの背景色

    # ──────────────────────────────────────────────────────────
    # 初期化
    # ──────────────────────────────────────────────────────────
    def __init__(self) -> None:
        super().__init__()
        self.title("ChunapTool - ウィンドウ位置調整ツール")

        # アイコン設定
        try:
            icon = tk.PhotoImage(file=resource_path("ChulipVideo.png"))
            self.iconphoto(False, icon)
        except Exception as e:
            print(f"Icon load error: {e}")

        # ウィンドウ配置（画面中央）
        self.update_idletasks()
        sx = (self.winfo_screenwidth() - INITIAL_WIDTH) // 2
        sy = (self.winfo_screenheight() - INITIAL_HEIGHT) // 2
        self.geometry(f"{INITIAL_WIDTH}x{INITIAL_HEIGHT}+{sx}+{sy}")
        self.minsize(MIN_WIDTH, MIN_HEIGHT)
        self.resizable(True, True)
        self.attributes("-topmost", True)
        self.attributes("-toolwindow", 1)

        # 外部ユーティリティ
        self.window_utils = WindowUtils()

        # Undo 履歴
        self.history: List[GeoState] = []
        self.last_state: Optional[GeoState] = None
        self._reading_rect = False  # UI 値セット中の再帰防止フラグ

        # ── Tkinter 変数 ──
        self.target_var = tk.StringVar()
        self.filter_var = tk.StringVar()
        self.geo_x = tk.IntVar()
        self.geo_y = tk.IntVar()
        self.geo_w = tk.IntVar()
        self.geo_h = tk.IntVar()
        self.max_geo_var = tk.BooleanVar(value=False)
        self.sync_geo_var = tk.BooleanVar(value=True)
        self.allow_out_of_bounds = tk.BooleanVar(value=False)

        # ── 内部状態 ──
        self.windows: List[Tuple[Any, str, str, int]] = []
        self.wgc_capture: Optional[WGCCapture] = None
        self.preview_image_id: Optional[int] = None

        # ── 構築 ──
        self._build_ui()
        self.update_source_list()

        # ── 定期ループ開始 ──
        self.after(UI_POLL_INTERVAL_MS, self._ui_update_loop)
        self.after(HISTORY_POLL_INTERVAL_MS, self._history_poll_loop)
        self.after(50, self._preview_loop)

        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _build_ui(self) -> None:
        """メインUIを構築する."""
        root = tk.Frame(self)
        root.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self._build_source_section(root)
        self._build_geometry_section(root)
        self._build_preview_section(root)

        # 初期状態を反映
        self._on_geo_ctrl_changed()

    def _build_source_section(self, parent: tk.Frame) -> None:
        """対象選択セクションを構築する."""
        group = tk.LabelFrame(parent, text="対象選択")
        group.pack(fill=tk.X, pady=5)

        # 検索フィルタ
        row_filter = tk.Frame(group)
        row_filter.pack(fill=tk.X, padx=5, pady=2)
        tk.Label(row_filter, text="検索:").pack(side=tk.LEFT)
        self.filter_var.trace_add("write", lambda *_: self.update_source_list())
        tk.Entry(row_filter, textvariable=self.filter_var).pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=5
        )

        # 対象コンボ
        row_target = tk.Frame(group)
        row_target.pack(fill=tk.X, padx=5, pady=2)
        tk.Label(row_target, text="対象:").pack(side=tk.LEFT)
        self.combo_target = ttk.Combobox(
            row_target, textvariable=self.target_var, state="readonly"
        )
        self.combo_target.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.combo_target.bind("<<ComboboxSelected>>", self._on_target_changed)
        tk.Button(
            row_target, text="更新", command=self.update_source_list, width=4, bg=self.COLOR_BTN
        ).pack(side=tk.LEFT)

    def _build_geometry_section(self, parent: tk.Frame) -> None:
        """位置・サイズ調整セクションを構築する."""
        group = tk.LabelFrame(parent, text="ウィンドウ位置・サイズ調整")
        group.pack(fill=tk.X, pady=5)

        row = tk.Frame(group)
        row.pack(fill=tk.X, padx=5, pady=5)

        # 最大化 / 同期 / 画面外可
        self.check_max_geo = tk.Checkbutton(
            row, text="最大化", variable=self.max_geo_var,
            command=self._on_max_changed, indicatoron=False,
            selectcolor=self.COLOR_BTN, relief=tk.RAISED, overrelief=tk.RIDGE,
            bg=self.COLOR_BTN,
        )
        self.check_max_geo.pack(side=tk.LEFT, padx=(0, 2), ipadx=5)

        self.check_sync_geo = tk.Checkbutton(
            row, text="同期", variable=self.sync_geo_var, command=self._on_geo_ctrl_changed
        )
        self.check_sync_geo.pack(side=tk.LEFT, padx=(0, 2))

        self.check_allow_out = tk.Checkbutton(
            row, text="画面外可", variable=self.allow_out_of_bounds,
            command=self.apply_window_geometry,
        )
        self.check_allow_out.pack(side=tk.LEFT, padx=(0, 5))

        # X / Y / W / H Spinbox
        spin_defs = [
            ("X:", self.geo_x, -10000, 10000),
            ("Y:", self.geo_y, -10000, 10000),
            ("W:", self.geo_w, 0, 10000),
            ("H:", self.geo_h, 0, 10000),
        ]
        self._spin_widgets: List[tk.Spinbox] = []
        for label, var, lo, hi in spin_defs:
            tk.Label(row, text=label).pack(side=tk.LEFT)
            spin = tk.Spinbox(
                row, from_=lo, to=hi, textvariable=var, width=7,
                command=self._on_geo_spin_cmd,
            )
            spin.pack(side=tk.LEFT, padx=2)
            self._spin_widgets.append(spin)

        # 適用 / 戻す ボタン
        self.btn_apply_geo = tk.Button(
            row, text="適用", command=self.apply_window_geometry, width=4, bg=self.COLOR_BTN
        )
        self.btn_apply_geo.pack(side=tk.LEFT, padx=5)

        self.btn_undo_geo = tk.Button(
            row, text="戻す", command=self._undo_geometry, width=4, bg=self.COLOR_BTN
        )
        self.btn_undo_geo.pack(side=tk.LEFT, padx=5)

        # Spinbox バインド（Enter / FocusOut / ホイール）
        for spin in self._spin_widgets:
            spin.bind("<Return>", lambda _: self.apply_window_geometry())
            spin.bind("<FocusOut>", lambda _: self.apply_window_geometry())
            spin.bind("<MouseWheel>", self._on_geo_spin_wheel)

    def _build_preview_section(self, parent: tk.Frame) -> None:
        """プレビューセクションを構築する."""
        group = tk.LabelFrame(parent, text="プレビュー")
        group.pack(fill=tk.BOTH, expand=True, pady=5)
        self.preview_canvas = tk.Canvas(group, bg="black")
        self.preview_canvas.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

    def update_source_list(self) -> None:
        """ウィンドウ一覧を更新し、最適な対象を選択する."""
        filter_text = self.filter_var.get().lower()
        self.windows = self.window_utils.enum_windows(filter_text)

        values = [
            f"[{w[2]}] {w[1]} ({w[3]})" for w in self.windows
        ]
        self.combo_target["values"] = values

        if not values:
            self.target_var.set("")
            return

        best_idx = self._find_best_match(values, filter_text)
        self.combo_target.current(best_idx)
        self._on_target_changed(None)

    @staticmethod
    def _find_best_match(values: List[str], filter_text: str) -> int:
        """フィルタに最も合致するインデックスを返す."""
        if not filter_text:
            return 0
        best = 0
        for i, v in enumerate(values):
            lower = v.lower()
            if f"[{filter_text}" in lower:
                return i
            if filter_text in lower and best == 0:
                best = i
        return best

    def _get_target_hwnd(self) -> Optional[Any]:
        """選択中のウィンドウハンドルを返す。未選択なら None."""
        idx = self.combo_target.current()
        if 0 <= idx < len(self.windows):
            return self.windows[idx][0]
        return None

    def _get_target_rect(self) -> Optional[Dict[str, int]]:
        """選択中ウィンドウの矩形情報を返す."""
        hwnd = self._get_target_hwnd()
        if hwnd is not None:
            return self.window_utils.get_window_rect(hwnd)
        return None

    def _on_target_changed(self, _event: Any) -> None:
        """対象ウィンドウが変わった時の処理."""
        rect = self._get_target_rect()
        if not rect:
            return

        # UI にウィンドウ座標を反映（再帰防止付き）
        self._set_geo_vars(rect)

        hwnd = self._get_target_hwnd()
        if hwnd is None:
            return

        # hwnd が変わった場合のみ履歴をリセット
        if self.last_state is None or self.last_state[0] != hwnd:
            self.history.clear()
            self.last_state = self._make_state(hwnd, rect)
            self.btn_undo_geo.config(state=tk.DISABLED)

        # WGC キャプチャの再初期化
        self._reset_wgc_capture(hwnd)

        # 最大化ボタンの有効・無効を制御
        if self.window_utils.can_maximize(hwnd):
            self.check_max_geo.config(state=tk.NORMAL)
        else:
            self.check_max_geo.config(state=tk.DISABLED)
            self.max_geo_var.set(False)

    def _reset_wgc_capture(self, hwnd: Any) -> None:
        """WGC キャプチャを再起動する."""
        if self.wgc_capture:
            self.wgc_capture.close()
            self.wgc_capture = None
        if WGC_AVAILABLE:
            try:
                self.wgc_capture = WGCCapture(hwnd)
            except Exception as e:
                print(f"Preview Start Error: {e}")

    def apply_window_geometry(self) -> None:
        """UI 上の座標値を対象ウィンドウに適用する."""
        if self.max_geo_var.get():
            return

        hwnd = self._get_target_hwnd()
        if hwnd is None:
            return

        try:
            tx, ty, tw, th = (
                self.geo_x.get(), self.geo_y.get(),
                self.geo_w.get(), self.geo_h.get(),
            )
        except (tk.TclError, ValueError):
            return

        curr_rect = self.window_utils.get_window_rect(hwnd)
        if not curr_rect:
            return

        # 境界制限の適用
        if not self.allow_out_of_bounds.get():
            bounds = self._get_work_bounds(curr_rect)
            if bounds:
                tx, ty, tw, th = self._clamp_to_bounds(tx, ty, tw, th, *bounds)

        # UIに調整結果を反映
        self._sync_geo_if_changed(tx, ty, tw, th)

        # ウィンドウに反映
        try:
            self.window_utils.set_window_position(hwnd, tx, ty, tw, th)
        except Exception as e:
            print(f"Geometry apply error: {e}")

    def _get_work_bounds(self, curr_rect: Dict[str, int]) -> Optional[Tuple[int, int, int, int]]:
        """ウィンドウ中心が属するモニターの作業領域 (left, top, right, bottom) を返す."""
        monitors = self.window_utils.get_monitor_info()
        if not monitors:
            return None

        # ウィンドウ中心で対象モニターを特定
        cx = curr_rect['left'] + curr_rect['width'] // 2
        cy = curr_rect['top'] + curr_rect['height'] // 2
        target = monitors[0]
        for m in monitors:
            if m['left'] <= cx < m['left'] + m['width'] and m['top'] <= cy < m['top'] + m['height']:
                target = m
                break

        m_left, m_top = target['left'], target['top']
        m_right = m_left + target['width']
        m_bottom = m_top + target['height']

        # 作業領域（タスクバー除外）があればそちらを優先
        work_areas = self.window_utils.get_workarea_info()
        if work_areas:
            target_work = work_areas[0]
            for wa in work_areas:
                if (wa['left'] < m_right and wa['left'] + wa['width'] > m_left
                        and wa['top'] < m_bottom and wa['top'] + wa['height'] > m_top):
                    target_work = wa
                    break
            m_left, m_top = target_work['left'], target_work['top']
            m_right = m_left + target_work['width']
            m_bottom = m_top + target_work['height']

        return m_left, m_top, m_right, m_bottom

    @staticmethod
    def _clamp_to_bounds(
        tx: int, ty: int, tw: int, th: int,
        m_left: int, m_top: int, m_right: int, m_bottom: int,
    ) -> Tuple[int, int, int, int]:
        """座標・サイズをモニター境界内に収める（サイズ維持優先）."""
        m_w, m_h = m_right - m_left, m_bottom - m_top

        # サイズをモニター内に収める
        tw = min(tw, m_w)
        th = min(th, m_h)

        # 位置を調整してモニター内に収める
        tx = max(tx, m_left)
        ty = max(ty, m_top)
        if tx + tw > m_right:
            tx = m_right - tw
        if ty + th > m_bottom:
            ty = m_bottom - th

        # 最小サイズの担保
        tw = max(tw, MIN_WINDOW_SIZE)
        th = max(th, MIN_WINDOW_SIZE)

        # 最小サイズ適用後の再位置調整
        if tx + tw > m_right:
            tx = m_right - tw
        if ty + th > m_bottom:
            ty = m_bottom - th
        tx = max(tx, m_left)
        ty = max(ty, m_top)

        return tx, ty, tw, th

    def _on_max_changed(self) -> None:
        """最大化チェック変更時の処理."""
        is_max = self.max_geo_var.get()
        state = tk.DISABLED if is_max else tk.NORMAL
        for w in self._spin_widgets + [self.check_sync_geo]:
            w.config(state=state)

        hwnd = self._get_target_hwnd()
        if hwnd is not None:
            if is_max:
                self.check_max_geo.config(relief=tk.SUNKEN)
                ctypes.windll.user32.ShowWindow(hwnd, SW_MAXIMIZE)
            else:
                self.check_max_geo.config(relief=tk.RAISED)
                ctypes.windll.user32.ShowWindow(hwnd, SW_NORMAL)
                self.apply_window_geometry()

        # 最大化・解除時は瞬時に履歴を確定させる
        self._record_history_step()

    def _on_geo_spin_cmd(self) -> None:
        """Spinbox の矢印ボタンで値が変わった時の反映."""
        if self.sync_geo_var.get() and not self.max_geo_var.get():
            self.apply_window_geometry()

    def _on_geo_spin_wheel(self, event: tk.Event) -> str:
        """Spinbox 上のホイール操作を処理する."""
        delta = 1 if event.delta > 0 else -1
        widget: tk.Spinbox = event.widget
        try:
            curr = int(widget.get())
            widget.delete(0, tk.END)
            widget.insert(0, str(curr + delta))
            if self.sync_geo_var.get() and not self.max_geo_var.get():
                self.apply_window_geometry()
        except (ValueError, tk.TclError):
            pass
        return "break"

    def _on_geo_ctrl_changed(self) -> None:
        """同期チェック変更時の処理."""
        if self.sync_geo_var.get():
            self.btn_apply_geo.config(state=tk.DISABLED)
            self._on_target_changed(None)
        else:
            self.btn_apply_geo.config(state=tk.NORMAL)

    def _history_poll_loop(self) -> None:
        """定期的に状態をチェックし、変化があれば履歴に保存する."""
        self._record_history_step()
        self.after(HISTORY_POLL_INTERVAL_MS, self._history_poll_loop)

    def _record_history_step(self) -> None:
        """現在のウィンドウ状態を確認し、変化があれば履歴に積む."""
        try:
            rect = self._get_target_rect()
            hwnd = self._get_target_hwnd()
            if not rect or hwnd is None:
                return

            is_max = self.window_utils.is_window_maximized(hwnd)
            new_state = self._make_state(hwnd, rect, is_max)

            # ドラッグ中は保存しない（ただし最大化フラグの変化は即座に保存）
            is_dragging = ctypes.windll.user32.GetAsyncKeyState(VK_LBUTTON) & 0x8000
            is_max_changed = self.last_state is not None and self.last_state[5] != is_max

            if not is_dragging or is_max_changed:
                if self.last_state is None:
                    self.last_state = new_state
                elif new_state != self.last_state:
                    self.history.append(self.last_state)
                    if len(self.history) > MAX_HISTORY:
                        self.history.pop(0)
                    self.last_state = new_state

            # 「戻す」ボタンの有効・無効を更新
            expected = tk.NORMAL if self.history else tk.DISABLED
            if str(self.btn_undo_geo["state"]) != str(expected):
                self.btn_undo_geo.config(state=expected)

        except Exception:
            pass

    def _undo_geometry(self) -> None:
        """直前の状態に1手戻す."""
        if not self.history:
            return

        target_state = self.history.pop()
        last_hwnd, tx, ty, tw, th, is_max = target_state
        self.last_state = target_state

        # 最大化状態を復元
        self.window_utils.set_window_maximized(last_hwnd, is_max)

        # 現在選択中のウィンドウと同じ場合、UIを更新
        if self._get_target_hwnd() == last_hwnd:
            self._reading_rect = True
            self.geo_x.set(tx)
            self.geo_y.set(ty)
            self.geo_w.set(tw)
            self.geo_h.set(th)
            self.max_geo_var.set(is_max)
            self.check_max_geo.config(relief=tk.SUNKEN if is_max else tk.RAISED)
            self._reading_rect = False

        if not is_max:
            self.window_utils.set_window_position(last_hwnd, tx, ty, tw, th)

    def _ui_update_loop(self) -> None:
        """UI 更新ループ."""
        try:
            rect = self._get_target_rect()
            hwnd = self._get_target_hwnd()

            if rect and hwnd is not None:
                self._sync_maximize_state(hwnd)
                self._record_history_step()

            self._sync_geometry_ui(rect, hwnd)
        except Exception:
            pass
        self.after(UI_POLL_INTERVAL_MS, self._ui_update_loop)

    def _sync_maximize_state(self, hwnd: Any) -> None:
        """ウィンドウの実際の最大化状態をUIに反映する."""
        actual_max = self.window_utils.is_window_maximized(hwnd)
        ui_max = self.max_geo_var.get()
        if actual_max != ui_max:
            self.max_geo_var.set(actual_max)
            self.check_max_geo.config(relief=tk.SUNKEN if actual_max else tk.RAISED)
            self._on_geo_ctrl_changed()

    def _sync_geometry_ui(self, rect: Optional[Dict[str, int]], hwnd: Optional[Any]) -> None:
        """座標情報を UI に反映する."""
        try:
            focused = self.focus_get()
        except (KeyError, tk.TclError):
            focused = None

        is_max = self.max_geo_var.get()
        is_sync = self.sync_geo_var.get()
        allow_out = self.allow_out_of_bounds.get()

        if focused in self._spin_widgets:
            return
        if not (is_sync or is_max or not allow_out):
            return
        if not rect:
            return

        tx, ty, tw, th = rect['left'], rect['top'], rect['width'], rect['height']

        if not allow_out and not is_max and hwnd is not None:
            bounds = self._get_work_bounds(rect)
            if bounds:
                new_tx, new_ty, new_tw, new_th = self._clamp_to_bounds(tx, ty, tw, th, *bounds)
                if (new_tx, new_ty, new_tw, new_th) != (tx, ty, tw, th):
                    tx, ty, tw, th = new_tx, new_ty, new_tw, new_th
                    self.window_utils.set_window_position(hwnd, tx, ty, tw, th)

        self._sync_geo_if_changed(tx, ty, tw, th)

    def _preview_loop(self) -> None:
        """プレビュー更新ループ."""
        try:
            frame = self._capture_preview_frame()
            if frame is not None:
                self._draw_preview(frame)
        except Exception:
            pass
        interval = (
            PREVIEW_WGC_INTERVAL_MS
            if self.wgc_capture and self.wgc_capture.session
            else PREVIEW_FALLBACK_INTERVAL_MS
        )
        self.after(interval, self._preview_loop)

    def _capture_preview_frame(self) -> Optional["cv2.Mat"]:
        """キャプチャフレームを取得."""
        if self.wgc_capture:
            frame = self.wgc_capture.get_latest_frame()
            if frame is not None:
                return frame
        hwnd = self._get_target_hwnd()
        if hwnd is not None:
            try:
                return self.window_utils.capture_exclusive_window(hwnd)
            except Exception:
                pass
        return None

    def _draw_preview(self, frame: "cv2.Mat") -> None:
        """キャンバスに描画."""
        cw, ch = self.preview_canvas.winfo_width(), self.preview_canvas.winfo_height()
        if cw <= 1 or ch <= 1:
            return
        fh, fw = frame.shape[:2]
        ratio = min(cw / fw, ch / fh)
        nw, nh = int(fw * ratio), int(fh * ratio)
        if nw <= 0 or nh <= 0:
            return
        resized = cv2.resize(frame, (nw, nh))
        tk_img = ImageTk.PhotoImage(Image.fromarray(cv2.cvtColor(resized, cv2.COLOR_BGR2RGB)))
        x, y = (cw - nw) // 2, (ch - nh) // 2
        if self.preview_image_id:
            self.preview_canvas.itemconfig(self.preview_image_id, image=tk_img)
            self.preview_canvas.coords(self.preview_image_id, x, y)
        else:
            self.preview_image_id = self.preview_canvas.create_image(x, y, image=tk_img, anchor=tk.NW)
        self.preview_canvas.image = tk_img

    def _set_geo_vars(self, rect: Dict[str, int]) -> None:
        """UI 変数を更新."""
        self._reading_rect = True
        self.geo_x.set(rect['left'])
        self.geo_y.set(rect['top'])
        self.geo_w.set(rect['width'])
        self.geo_h.set(rect['height'])
        self._reading_rect = False

    def _sync_geo_if_changed(self, tx: int, ty: int, tw: int, th: int) -> None:
        """差分がある時のみ UI 変数を更新."""
        if (tx != self.geo_x.get() or ty != self.geo_y.get() or tw != self.geo_w.get() or th != self.geo_h.get()):
            self._reading_rect = True
            self.geo_x.set(tx)
            self.geo_y.set(ty)
            self.geo_w.set(tw)
            self.geo_h.set(th)
            self._reading_rect = False

    @staticmethod
    def _make_state(hwnd: Any, rect: Dict[str, int], is_max: bool = False) -> GeoState:
        """GeoState 生成."""
        return (hwnd, rect['left'], rect['top'], rect['width'], rect['height'], is_max)

    def _on_close(self) -> None:
        """終了処理."""
        if self.wgc_capture:
            self.wgc_capture.close()
        self.destroy()

if __name__ == "__main__":
    app = ChunapTool()
    app.mainloop()
