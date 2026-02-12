"""ウィンドウ位置調整ツール

ウィンドウ位置調整を行うためのウィンドウ。
"""
import tkinter as tk
import tkinter.ttk as ttk
from typing import Optional, List, Dict, Any, Tuple
import ctypes
import os
import sys

import cv2
from PIL import Image, ImageTk

from window_utils import WindowUtils
from wgc_capture import WGCCapture, WGC_AVAILABLE
from config import PROJECT_NAME, load_global_config, save_global_config
from utils import resource_path

class ChunapTool(tk.Tk):
    COLOR_BTN_UPDATE = "#e1f5fe"

    def __init__(self):
        super().__init__()
        self.title("ChunapTool - ウィンドウ位置調整ツール")
        
        try:
            icon_path = resource_path("ChulipVideo.png")
            img = tk.PhotoImage(file=icon_path)
            self.iconphoto(False, img)
        except Exception as e:
            print(f"Icon load error: {e}")

        # ウィンドウを画面中央に配置
        self.update_idletasks()
        w = 640
        h = 240
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        x = (sw - w) // 2
        y = (sh - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")
        
        self.minsize(640, 200)
        self.resizable(True, True)
        self.attributes("-topmost", True)
        self.attributes("-toolwindow", 1)
        
        self.window_utils = WindowUtils()
        self.history: List[Tuple[Any, int, int, int, int, bool]] = [] # (hwnd, x, y, w, h, is_max)
        self.last_state: Optional[Tuple[Any, int, int, int, int, bool]] = None
        self._reading_rect = False
        
        # --- Variables ---
        self.source_var = tk.StringVar(value="window")
        self.target_var = tk.StringVar()
        self.filter_var = tk.StringVar()
        
        self.geo_x = tk.IntVar()
        self.geo_y = tk.IntVar()
        self.geo_w = tk.IntVar()
        self.geo_h = tk.IntVar()
        
        self.max_geo_var = tk.BooleanVar(value=False)
        self.sync_geo_var = tk.BooleanVar(value=True)
        self.allow_out_of_bounds = tk.BooleanVar(value=False)
        
        self.windows: List[Tuple[Any, str, str, int]] = []
        self.monitors: List[Dict[str, Any]] = []
        
        self.wgc_capture: Optional[WGCCapture] = None
        self.preview_image_id = None
        
        # --- UI Build ---
        self._build_ui()
        
        # --- Initial Load ---
        self._load_config()
        self.update_source_list()
        
        # --- Start Loop ---
        self.after(500, self._ui_update_loop)
        self.after(500, self._history_poll_loop)
        self.after(50, self._start_preview)
        
        # --- Bindings ---
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    def _build_ui(self):
        main_frame = tk.Frame(self)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 1. Source (Window only now)
        source_frame = tk.LabelFrame(main_frame, text="対象選択")
        source_frame.pack(fill=tk.X, pady=5)
        
        # Filter
        filter_frame = tk.Frame(source_frame)
        filter_frame.pack(fill=tk.X, padx=5, pady=2)
        tk.Label(filter_frame, text="検索:").pack(side=tk.LEFT)
        self.filter_var.trace_add("write", lambda *args: self.update_source_list())
        tk.Entry(filter_frame, textvariable=self.filter_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        # Target Combo
        target_frame = tk.Frame(source_frame)
        target_frame.pack(fill=tk.X, padx=5, pady=2)
        tk.Label(target_frame, text="対象:").pack(side=tk.LEFT)
        self.combo_target = ttk.Combobox(target_frame, textvariable=self.target_var, state="readonly")
        self.combo_target.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.combo_target.bind("<<ComboboxSelected>>", self.on_target_changed)
        tk.Button(target_frame, text="更新", command=self.update_source_list, width=4, bg=self.COLOR_BTN_UPDATE).pack(side=tk.LEFT)

        # 2. Geometry Settings
        geo_group = tk.LabelFrame(main_frame, text="ウィンドウ位置・サイズ調整")
        geo_group.pack(fill=tk.X, pady=5)
        
        geo_frame = tk.Frame(geo_group)
        geo_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Maximize / Sync
        self.check_max_geo = tk.Checkbutton(geo_frame, text="最大化", variable=self.max_geo_var, 
                                            command=self._on_max_changed, indicatoron=False, 
                                            selectcolor=self.COLOR_BTN_UPDATE, relief=tk.RAISED, overrelief=tk.RIDGE,
                                            bg=self.COLOR_BTN_UPDATE)
        self.check_max_geo.pack(side=tk.LEFT, padx=(0, 2), ipadx=5)
        
        self.check_sync_geo = tk.Checkbutton(geo_frame, text="同期", variable=self.sync_geo_var, command=self._on_geo_ctrl_changed)
        self.check_sync_geo.pack(side=tk.LEFT, padx=(0, 2))
        
        self.check_allow_out = tk.Checkbutton(geo_frame, text="画面外可", variable=self.allow_out_of_bounds, command=self.apply_window_geometry)
        self.check_allow_out.pack(side=tk.LEFT, padx=(0, 5))
        
        # XYWH Spinboxes
        tk.Label(geo_frame, text="X:").pack(side=tk.LEFT)
        self.spin_x = tk.Spinbox(geo_frame, from_=-10000, to=10000, textvariable=self.geo_x, width=7, command=self._on_geo_spin_cmd)
        self.spin_x.pack(side=tk.LEFT, padx=2)
        
        tk.Label(geo_frame, text="Y:").pack(side=tk.LEFT)
        self.spin_y = tk.Spinbox(geo_frame, from_=-10000, to=10000, textvariable=self.geo_y, width=7, command=self._on_geo_spin_cmd)
        self.spin_y.pack(side=tk.LEFT, padx=2)
        
        tk.Label(geo_frame, text="W:").pack(side=tk.LEFT)
        self.spin_w = tk.Spinbox(geo_frame, from_=0, to=10000, textvariable=self.geo_w, width=7, command=self._on_geo_spin_cmd)
        self.spin_w.pack(side=tk.LEFT, padx=2)
        
        tk.Label(geo_frame, text="H:").pack(side=tk.LEFT)
        self.spin_h = tk.Spinbox(geo_frame, from_=0, to=10000, textvariable=self.geo_h, width=7, command=self._on_geo_spin_cmd)
        self.spin_h.pack(side=tk.LEFT, padx=2)
        
        self.btn_apply_geo = tk.Button(geo_frame, text="適用", command=self.apply_window_geometry, width=4, bg=self.COLOR_BTN_UPDATE)
        self.btn_apply_geo.pack(side=tk.LEFT, padx=5)
        
        self.btn_undo_geo = tk.Button(geo_frame, text="戻す", command=self.undo_geometry, width=4, bg=self.COLOR_BTN_UPDATE)
        self.btn_undo_geo.pack(side=tk.LEFT, padx=5)
        
        # Bindings for spins
        for spin in [self.spin_x, self.spin_y, self.spin_w, self.spin_h]:
            spin.bind("<Return>", lambda e: self.apply_window_geometry())
            spin.bind("<FocusOut>", lambda e: self.apply_window_geometry())
            spin.bind("<MouseWheel>", self._on_geo_spin_wheel)

        # Preview Frame (Moved to bottom)
        preview_group = tk.LabelFrame(main_frame, text="プレビュー")
        preview_group.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.preview_canvas = tk.Canvas(preview_group, bg="black")
        self.preview_canvas.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self._on_sync_changed()

    def _load_config(self):
        config = load_global_config()
        saved_source = config.get("recorder_source", "window")
        self.source_var.set(saved_source)
        
    def update_source_list(self):
        filter_text = self.filter_var.get().lower()
        values = []
        self.windows = self.window_utils.enum_windows(filter_text)
        
        for w in self.windows:
            pname = w[2]
            title = w[1]
            pid = w[3]
            values.append(f"[{pname}] {title} ({pid})")
            
        self.monitors = []
        self.combo_target['values'] = values
        if values:
            best_idx = 0
            if filter_text:
                for i, v in enumerate(values):
                    if f"[{filter_text}" in v.lower():
                        best_idx = i
                        break
                    if filter_text in v.lower() and best_idx == 0:
                        best_idx = i
            
            self.combo_target.current(best_idx)
            self.on_target_changed(None)
        else:
            self.target_var.set("")

    def on_target_changed(self, event):
        mode = "window" 
        rect = self._get_target_rect()
        if rect:
            self._reading_rect = True
            self.geo_x.set(rect['left'])
            self.geo_y.set(rect['top'])
            self.geo_w.set(rect['width'])
            self.geo_h.set(rect['height'])
            self._reading_rect = False
            
            idx = self.combo_target.current()
            hwnd = None
            if mode == 'window' and 0 <= idx < len(self.windows):
                hwnd = self.windows[idx][0]

            # 対象変更時（hwndが変わった時）のみ履歴をリセット
            if self.last_state is None or (hwnd and self.last_state[0] != hwnd):
                self.history.clear()
                if hwnd:
                    self.last_state = (hwnd, rect['left'], rect['top'], rect['width'], rect['height'], self.window_utils.is_window_maximized(hwnd))
                else:
                    self.last_state = None
                self.btn_undo_geo.config(state=tk.DISABLED)
            
            if mode == 'window' and hwnd:
                # Clean up old
                if self.wgc_capture:
                    self.wgc_capture.close()
                    self.wgc_capture = None
                    
                if WGC_AVAILABLE:
                    try:
                        self.wgc_capture = WGCCapture(hwnd)
                    except Exception as e:
                        print(f"Preview Start Error: {e}")
                
                # Check if window can be maximized
                if self.window_utils.can_maximize(hwnd):
                    self.check_max_geo.config(state=tk.NORMAL)
                else:
                    self.check_max_geo.config(state=tk.DISABLED)
                    self.max_geo_var.set(False)

    def _get_target_rect(self) -> Optional[Dict[str, int]]:
        mode = "window" # Fixed
        idx = self.combo_target.current()
        if mode == 'window':
            if idx >= 0 and idx < len(self.windows):
                hwnd = self.windows[idx][0]
                return self.window_utils.get_window_rect(hwnd)
        return None

    def apply_window_geometry(self):
        if not self.max_geo_var.get():
            idx = self.combo_target.current()
            if idx >= 0 and idx < len(self.windows):
                hwnd = self.windows[idx][0]
                try:
                    tx = self.geo_x.get()
                    ty = self.geo_y.get()
                    tw = self.geo_w.get()
                    th = self.geo_h.get()

                    curr_rect = self.window_utils.get_window_rect(hwnd)
                    if not curr_rect:
                        return

                    monitors = self.window_utils.get_monitor_info()
                    if not monitors:
                        self.window_utils.set_window_position(hwnd, tx, ty, tw, th)
                        return

                    cx = curr_rect['left'] + curr_rect['width'] // 2
                    cy = curr_rect['top'] + curr_rect['height'] // 2
                    
                    target_monitor = monitors[0]
                    for m in monitors:
                        if m['left'] <= cx < m['left'] + m['width'] and m['top'] <= cy < m['top'] + m['height']:
                            target_monitor = m
                            break
                    
                    m_left = target_monitor['left']
                    m_top = target_monitor['top']
                    m_right = m_left + target_monitor['width']
                    m_bottom = m_top + target_monitor['height']

                    # 境界チェックと自動調整
                    if not self.allow_out_of_bounds.get():
                        work_areas = self.window_utils.get_workarea_info()
                        if work_areas:
                            target_work = work_areas[0]
                            for wa in work_areas:
                                if (wa['left'] < m_right and wa['left'] + wa['width'] > m_left and 
                                    wa['top'] < m_bottom and wa['top'] + wa['height'] > m_top):
                                    target_work = wa
                                    break
                            m_left, m_top = target_work['left'], target_work['top']
                            m_right, m_bottom = m_left + target_work['width'], m_top + target_work['height']

                        # サイズをモニター内に収める
                        m_w = m_right - m_left
                        m_h = m_bottom - m_top
                        if tw > m_w: tw = m_w
                        if th > m_h: th = m_h

                        # 位置を調整してモニター内に収める (サイズ維持優先)
                        if tx < m_left: tx = m_left
                        if ty < m_top: ty = m_top
                        if tx + tw > m_right: tx = m_right - tw
                        if ty + th > m_bottom: ty = m_bottom - th

                        # 最小サイズの担保
                        if tw < 32: tw = 32
                        if th < 32: th = 32
                        
                        # サイズ変更後の再位置調整
                        if tx + tw > m_right: tx = m_right - tw
                        if ty + th > m_bottom: ty = m_bottom - th
                        if tx < m_left: tx = m_left
                        if ty < m_top: ty = m_top

                    # UIに調整結果を反映
                    if tx != self.geo_x.get() or ty != self.geo_y.get() or tw != self.geo_w.get() or th != self.geo_h.get():
                        self._reading_rect = True
                        self.geo_x.set(tx)
                        self.geo_y.set(ty)
                        self.geo_w.set(tw)
                        self.geo_h.set(th)
                        self._reading_rect = False

                    # 適用
                    self.window_utils.set_window_position(hwnd, tx, ty, tw, th)

                except tk.TclError:
                    pass
                except Exception as e:
                    print(f"Geometry apply error: {e}")

    def _on_max_changed(self):
        is_max = self.max_geo_var.get()
        state = tk.DISABLED if is_max else tk.NORMAL
        for w in [self.spin_x, self.spin_y, self.spin_w, self.spin_h, self.check_sync_geo]:
            w.config(state=state)
            
        idx = self.combo_target.current()
        if idx >= 0 and idx < len(self.windows):
            hwnd = self.windows[idx][0]
            if is_max:
                self.check_max_geo.config(relief=tk.SUNKEN)
                ctypes.windll.user32.ShowWindow(hwnd, 3) # SW_MAXIMIZE
            else:
                self.check_max_geo.config(relief=tk.RAISED)
                ctypes.windll.user32.ShowWindow(hwnd, 1) # SW_NORMAL
                self.apply_window_geometry() # Restore size
        
        # 最大化・解除時は瞬時に履歴を確定させる
        self._record_history_step()

    def _on_geo_spin_cmd(self):
        """Spinboxの矢印等で値が変わった時の反映"""
        if self.sync_geo_var.get() and not self.max_geo_var.get():
            self.apply_window_geometry()

    def _history_poll_loop(self):
        """0.5秒おきに状態をチェックして履歴に保存する"""
        self._record_history_step()
        self.after(500, self._history_poll_loop)

    def _record_history_step(self):
        """現在の状態をチェックし、変化があれば履歴に保存する"""
        try:
            rect = self._get_target_rect()
            if rect:
                idx = self.combo_target.current()
                if idx >= 0 and idx < len(self.windows):
                    hwnd = self.windows[idx][0]
                    is_max = self.window_utils.is_window_maximized(hwnd)
                    new_state = (hwnd, rect['left'], rect['top'], rect['width'], rect['height'], is_max)
                    
                    # 左クリック状態を確認 (0x01)
                    is_dragging = ctypes.windll.user32.GetAsyncKeyState(0x01) & 0x8000
                    
                    # 座標変化の保存は「クリック中でない」とき。
                    # ただし「最大化フラグのみの変化」はクリック中（ボタン押し中）でも即座に保存して良い
                    is_max_changed = (self.last_state and self.last_state[5] != is_max)
                    
                    if not is_dragging or is_max_changed:
                        if self.last_state is None:
                            self.last_state = new_state
                        elif new_state != self.last_state:
                            # 状態が変わっていたら履歴に積む
                            self.history.append(self.last_state)
                            if len(self.history) > 100:
                                self.history.pop(0)
                            self.last_state = new_state
            
            # 戻すボタンの有効・無効切り替え
            state = tk.NORMAL if self.history else tk.DISABLED
            if self.btn_undo_geo['state'] != state:
                self.btn_undo_geo.config(state=state)
                
        except Exception:
            pass

    def undo_geometry(self):
        """1手戻す"""
        if not self.history:
            return
        
        target_state = self.history.pop()
        last_hwnd, tx, ty, tw, th, is_max = target_state
        self.last_state = target_state
        
        self.window_utils.set_window_maximized(last_hwnd, is_max)
        
        idx = self.combo_target.current()
        if idx >= 0 and idx < len(self.windows) and self.windows[idx][0] == last_hwnd:
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

    def _on_geo_ctrl_changed(self):
        is_sync = self.sync_geo_var.get()
        if is_sync:
            self.btn_apply_geo.config(state=tk.DISABLED)
            self.on_target_changed(None)
        else:
            self.btn_apply_geo.config(state=tk.NORMAL)

    def _ui_update_loop(self):
        """Poll window position and update UI if Sync/Maximize is enabled"""
        try:
            rect = self._get_target_rect()
            
            # 1. Update Maximize State
            if rect:
                idx = self.combo_target.current()
                if idx >= 0 and idx < len(self.windows):
                    hwnd = self.windows[idx][0]
                    if self.window_utils.is_window_maximized(hwnd):
                        if not self.max_geo_var.get():
                            self.max_geo_var.set(True)
                            self.check_max_geo.config(relief=tk.SUNKEN)
                            self._on_geo_ctrl_changed()
                    else:
                        if self.max_geo_var.get():
                            self.max_geo_var.set(False)
                            self.check_max_geo.config(relief=tk.RAISED)
                            self._on_geo_ctrl_changed()
                    
                    self._record_history_step()

            # 2. Update Geometry UI
            try:
                focused_widget = self.focus_get()
            except:
                focused_widget = None

            is_max = self.max_geo_var.get()
            is_sync = self.sync_geo_var.get()
            allow_out = self.allow_out_of_bounds.get()
            
            input_widgets = [self.spin_x, self.spin_y, self.spin_w, self.spin_h]
            if (is_sync or is_max or not allow_out) and (focused_widget not in input_widgets):
                if rect:
                    tx, ty, tw, th = rect['left'], rect['top'], rect['width'], rect['height']
                    
                    if not allow_out and not is_max:
                        work_areas = self.window_utils.get_workarea_info()
                        if work_areas:
                            cx = tx + tw // 2
                            cy = ty + th // 2
                            target_work = work_areas[0]
                            for wa in work_areas:
                                if wa['left'] <= cx < wa['left'] + wa['width'] and wa['top'] <= cy < wa['top'] + wa['height']:
                                    target_work = wa
                                    break
                            m_left, m_top = target_work['left'], target_work['top']
                            m_right, m_bottom = m_left + target_work['width'], m_top + target_work['height']
                            
                            adjusted = False
                            if tx < m_left:
                                diff = m_left - tx
                                tw -= diff
                                tx = m_left
                                adjusted = True
                            if ty < m_top:
                                diff = m_top - ty
                                th -= diff
                                ty = m_top
                                adjusted = True
                            if tx + tw > m_right: tw = m_right - tx; adjusted = True
                            if ty + th > m_bottom: th = m_bottom - ty; adjusted = True
                            if tw < 32: tw = 32; adjusted = True
                            if th < 32: th = 32; adjusted = True
                            
                            if adjusted:
                                if tx + tw > m_right: tx = m_right - tw
                                if ty + th > m_bottom: ty = m_bottom - th
                                if tx < m_left: tx = m_left
                                if ty < m_top: ty = m_top
                                idx = self.combo_target.current()
                                if idx >= 0 and idx < len(self.windows):
                                    hwnd = self.windows[idx][0]
                                    self.window_utils.set_window_position(hwnd, tx, ty, tw, th)

                    try:
                        self._reading_rect = True
                        if self.geo_x.get() != tx: self.geo_x.set(tx)
                        if self.geo_y.get() != ty: self.geo_y.set(ty)
                        if self.geo_w.get() != tw: self.geo_w.set(tw)
                        if self.geo_h.get() != th: self.geo_h.set(th)
                    except tk.TclError:
                        pass
                    finally:
                        self._reading_rect = False
                        
        except Exception:
            pass
        self.after(200, self._ui_update_loop)

    def _start_preview(self):
        try:
            has_wgc_frame = False
            if self.wgc_capture:
                frame = self.wgc_capture.get_latest_frame()
                if frame is not None:
                    self._update_preview_canvas(frame)
                    has_wgc_frame = True
            if not has_wgc_frame:
                self._try_fallback_preview()
        except Exception:
            pass
        interval = 33 if (self.wgc_capture and self.wgc_capture.session) else 200
        self.after(interval, self._start_preview)

    def _try_fallback_preview(self):
        idx = self.combo_target.current()
        if idx >= 0 and idx < len(self.windows):
            hwnd = self.windows[idx][0]
            try:
                frame = self.window_utils.capture_exclusive_window(hwnd)
                if frame is not None:
                    self._update_preview_canvas(frame)
            except Exception:
                pass

    def _update_preview_canvas(self, frame):
        try:
            cw = self.preview_canvas.winfo_width()
            ch = self.preview_canvas.winfo_height()
            if cw > 1 and ch > 1:
                h, w = frame.shape[:2]
                ratio = min(cw / w, ch / h)
                new_w, new_h = int(w * ratio), int(h * ratio)
                if new_w > 0 and new_h > 0:
                    frame = cv2.resize(frame, (new_w, new_h))
                    frame_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                    tk_img = ImageTk.PhotoImage(Image.fromarray(frame_rgb))
                    x, y = (cw - new_w) // 2, (ch - new_h) // 2
                    if self.preview_image_id:
                        self.preview_canvas.itemconfig(self.preview_image_id, image=tk_img)
                        self.preview_canvas.coords(self.preview_image_id, x, y)
                    else:
                        self.preview_image_id = self.preview_canvas.create_image(x, y, image=tk_img, anchor=tk.NW)
                    self.preview_canvas.image = tk_img
        except Exception:
            pass

    def _on_sync_changed(self):
        self._on_geo_ctrl_changed()

    def _on_geo_spin_wheel(self, event):
        delta = 1 if event.delta > 0 else -1
        widget: tk.Spinbox = event.widget
        try:
            curr = int(widget.get())
            widget.delete(0, tk.END)
            widget.insert(0, str(curr + delta))
            if self.sync_geo_var.get() and not self.max_geo_var.get():
                self.apply_window_geometry()
        except:
            pass
        return "break"

    def on_close(self):
        if self.wgc_capture:
            self.wgc_capture.close()
        self.destroy()

if __name__ == "__main__":
    app = ChunapTool()
    app.mainloop()
