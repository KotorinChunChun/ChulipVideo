"""シークバー関連モジュール.

シークバーの描画とマウス操作のロジックを提供。
"""
from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    import tkinter as tk


class SeekbarMixin:
    """シークバー関連のメソッドを提供するMixinクラス.
    
    VideoCropperAppクラスにミックスインして使用する。
    """
    
    # 以下の属性は VideoCropperApp から継承される想定
    # seek_canvas: tk.Canvas
    # duration: float
    # current_time: float
    # start_time: float
    # end_time: float
    # SEEK_MARGIN: int
    # SEEK_H: int
    # CANVAS_W: int
    # drag_mode: str | None
    
    def get_x(self, t: float) -> float:
        """時間からシークバー上のX座標を計算する."""
        if self.duration <= 0:
            return self.SEEK_MARGIN
        # シークバーキャンバスの実際の幅を取得
        seek_w = self.seek_canvas.winfo_width()
        if seek_w <= 1:
            seek_w = self.CANVAS_W
        return self.SEEK_MARGIN + (t / self.duration) * (seek_w - 2 * self.SEEK_MARGIN)

    def get_t(self, x: float) -> float:
        """シークバー上のX座標から時間を計算する."""
        if self.duration <= 0:
            return 0
        # シークバーキャンバスの実際の幅を取得
        seek_w = self.seek_canvas.winfo_width()
        if seek_w <= 1:
            seek_w = self.CANVAS_W
        w = seek_w - 2 * self.SEEK_MARGIN
        ratio = (x - self.SEEK_MARGIN) / w
        return max(0, min(1.0, ratio)) * self.duration

    def draw_seekbar(self) -> None:
        """シークバーを描画する."""
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
            self.SEEK_MARGIN, bar_y - 4, w - self.SEEK_MARGIN, bar_y + 4,
            fill="#ddd", outline="#aaa"
        )

        # Play progress
        cx = self.get_x(self.current_time)
        self.seek_canvas.create_rectangle(
            self.SEEK_MARGIN, bar_y - 4, cx, bar_y + 4,
            fill="#4da6ff", outline=""
        )

        # Start Marker (Top)
        sx = self.get_x(self.start_time)
        self.seek_canvas.create_polygon(
            sx - 8, bar_y - 8, sx + 8, bar_y - 8, sx, bar_y,
            fill="#00aa00", outline="black"
        )
        self.seek_canvas.create_text(
            sx, bar_y - 20, text="Start", fill="#006600", font=("Arial", 8)
        )

        # End Marker (Bottom)
        ex = self.get_x(self.end_time)
        self.seek_canvas.create_polygon(
            ex - 8, bar_y + 8, ex + 8, bar_y + 8, ex, bar_y,
            fill="#cc0000", outline="black"
        )
        self.seek_canvas.create_text(
            ex, bar_y + 20, text="End", fill="#990000", font=("Arial", 8)
        )

        # Current Thumb (Circle)
        self.seek_canvas.create_oval(
            cx - 7, bar_y - 7, cx + 7, bar_y + 7,
            fill="white", outline="#333", width=2
        )

    def seek_on_click(self, e) -> None:
        """シークバーのクリックイベントを処理する."""
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

    def seek_on_drag(self, e) -> None:
        """シークバーのドラッグイベントを処理する."""
        if self.duration <= 0 or not self.drag_mode:
            return
        self.handle_seek_drag(e.x)

    def seek_on_release(self, e) -> None:
        """シークバーのリリースイベントを処理する."""
        self.drag_mode = None

    def handle_seek_drag(self, mouse_x: float) -> None:
        """シークバーのドラッグ処理を行う."""
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
