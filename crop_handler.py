"""クロップ矩形操作モジュール.

クロップ矩形のドラッグ、リサイズ、Undo/Redo などのロジックを提供。
"""
from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    pass


class CropHandlerMixin:
    """クロップ矩形操作のメソッドを提供するMixinクラス.
    
    VideoCropperAppクラスにミックスインして使用する。
    """
    
    # 以下の属性は VideoCropperApp から継承される想定
    # crop_rect: list[int]
    # crop_history: list
    # crop_redo: list
    # MIN_W, MIN_H: int
    # CANVAS_W, CANVAS_H: int
    # canvas_scale_x, canvas_scale_y: float
    # canvas_offset_x, canvas_offset_y: int
    # resize_edge: dict
    # orig_rect: list[int]
    # orig_aspect_ratio: float
    
    def push_crop_history(self) -> None:
        """現在のクロップ矩形を履歴に追加する."""
        current = self.crop_rect.copy()
        if self.crop_history and self.crop_history[-1] == current:
            return
        self.crop_history.append(current)
        # Redo スタックはクリア（新しい操作が入ったため）
        self.crop_redo.clear()
        self.update_undo_button_state()

    def undo_crop(self, event=None) -> None:
        """Undo: 最後のクロップ矩形変更を元に戻す."""
        if not self.crop_history:
            return
        # 現在の状態を Redo スタックに保存
        self.crop_redo.append(self.crop_rect.copy())
        self.crop_rect = self.crop_history.pop()
        self._sync_crop_rect_ui()
        self.update_undo_button_state()

    def redo_crop(self) -> None:
        """Redo: 最後のクロップ矩形変更をやり直す."""
        if not self.crop_redo:
            return
        # 現在の状態を履歴に保存
        self.crop_history.append(self.crop_rect.copy())
        self.crop_rect = self.crop_redo.pop()
        self._sync_crop_rect_ui()
        self.update_undo_button_state()

    def update_undo_button_state(self) -> None:
        """Undo/Redoボタンの有効/無効を更新する."""
        try:
            import tkinter as tk
            if hasattr(self, 'btn_undo'):
                if self.crop_history:
                    self.btn_undo.config(state=tk.NORMAL)
                else:
                    self.btn_undo.config(state=tk.DISABLED)
            if hasattr(self, 'btn_redo'):
                if self.crop_redo:
                    self.btn_redo.config(state=tk.NORMAL)
                else:
                    self.btn_redo.config(state=tk.DISABLED)
        except Exception:
            pass

    def clamp_rect_canvas(self, r: list[int]) -> list[int]:
        """矩形をキャンバス範囲内に制約する."""
        x1, y1, x2, y2 = r
        x1, x2 = sorted([x1, x2])
        y1, y2 = sorted([y1, y2])
        w = max(self.MIN_W, x2 - x1)
        h = max(self.MIN_H, y2 - y1)
        
        # クランプ対象の最大値を動画解像度に変更（フォールバックでキャンバス）
        vw = getattr(self, "vid_w", self.CANVAS_W)
        vh = getattr(self, "vid_h", self.CANVAS_H)

        if getattr(self, 'allow_oversize_var', None) and self.allow_oversize_var.get():
             # 枠外選択許可時は動画解像度によるクランプを行わない
             # ただし、あまりに遠くに行き過ぎないように緩和された制限を設ける（例えば解像度の3倍程度）
             # ここではシンプルに、width/height の最小値だけ保証して座標制限は事実上外す
             # (UI操作で戻ってこられる範囲ならOK)
             limit_margin = 5000 # 画面外許容範囲
             if x1 < -limit_margin: x1 = -limit_margin
             if y1 < -limit_margin: y1 = -limit_margin
             if x1 > vw + limit_margin: x1 = vw + limit_margin
             if y1 > vh + limit_margin: y1 = vh + limit_margin
             # x2, y2 は w, h から再計算されるので x1, y1 だけ見ればよいが念のため
        else:
            if x1 < 0:
                x1 = 0
            if y1 < 0:
                y1 = 0
            if x1 + w > vw:
                x1 = vw - w
            if y1 + h > vh:
                y1 = vh - h
                
        return [int(x1), int(y1), int(x1 + w), int(y1 + h)]

    def clamp_rect_move(self, x1: int, y1: int, w: int, h: int) -> list[int]:
        """移動専用のクランプ処理（サイズ固定）.
        
        指定されたw, hを必ず維持し、位置のみを調整する。
        """
        # クランプ対象の最大値を動画解像度に変更（フォールバックでキャンバス）
        vw = getattr(self, "vid_w", self.CANVAS_W)
        vh = getattr(self, "vid_h", self.CANVAS_H)

        allow_oversize = getattr(self, 'allow_oversize_var', None) and self.allow_oversize_var.get()

        if allow_oversize:
             limit_margin = 5000 # 画面外許容範囲
             if x1 < -limit_margin: x1 = -limit_margin
             if y1 < -limit_margin: y1 = -limit_margin
             # x2 = x1 + w が vw + limit_margin を超えないようにする
             if x1 + w > vw + limit_margin: x1 = (vw + limit_margin) - w
             if y1 + h > vh + limit_margin: y1 = (vh + limit_margin) - h
        else:
            # キャンバス（動画）内に収める
            if x1 < 0: x1 = 0
            if y1 < 0: y1 = 0
            if x1 + w > vw: x1 = vw - w
            if y1 + h > vh: y1 = vh - h
            
        return [int(x1), int(y1), int(x1 + w), int(y1 + h)]

    def maintain_aspect_ratio_resize(
        self, x1: int, y1: int, x2: int, y2: int, ctrl_held: bool = False
    ) -> tuple[int, int, int, int]:
        """アスペクト比を維持しながらリサイズする."""
        # どの角がドラッグされているかを判定
        edges = self.resize_edge
        
        if ctrl_held:
            # 中心固定アスペクト比リサイズ
            # 元の中心を計算
            ox1, oy1, ox2, oy2 = self.orig_rect
            cx, cy = (ox1 + ox2) / 2.0, (oy1 + oy2) / 2.0
            
            new_w = abs(x2 - x1)
            new_h = abs(y2 - y1)
            
            # 高さの変更を優先（上下辺または四角をドラッグ中）
            if edges["t"] or edges["b"]:
                new_w = new_h * self.orig_aspect_ratio
            # 幅の変更を優先（左右辺のみをドラッグ中）
            elif edges["l"] or edges["r"]:
                new_h = new_w / self.orig_aspect_ratio
                
            x1 = int(cx - new_w / 2.0)
            x2 = int(cx + new_w / 2.0)
            y1 = int(cy - new_h / 2.0)
            y2 = int(cy + new_h / 2.0)
        else:
            # 通常のアスペクト比リサイズ（対角固定）
            new_w = x2 - x1
            new_h = y2 - y1

            # 高さの変更に基づいて幅を調整
            if edges["t"] or edges["b"]:
                new_w = int(new_h * self.orig_aspect_ratio)
                if edges["l"]:
                    x1 = x2 - new_w
                else:
                    x2 = x1 + new_w
            # 幅の変更に基づいて高さを調整
            elif edges["l"] or edges["r"]:
                new_h = int(new_w / self.orig_aspect_ratio)
                if edges["t"]:
                    y1 = y2 - new_h
                else:
                    y2 = y1 + new_h

        return x1, y1, x2, y2

    def clamp_rect_resize(
        self, x1: int, y1: int, x2: int, y2: int
    ) -> list[int]:
        """リサイズ時に矩形をクランプする（各辺を独立に制約）."""
        # 動画の実際の解像度を取得（フォールバックとしてCANVAS値）
        vw = getattr(self, "vid_w", self.CANVAS_W)
        vh = getattr(self, "vid_h", self.CANVAS_H)
        
        edges = self.resize_edge or {}
        allow_oversize = getattr(self, 'allow_oversize_var', None) and self.allow_oversize_var.get()

        # 最小サイズを保証
        if x2 - x1 < self.MIN_W:
            if edges.get("r", False):
                x2 = x1 + self.MIN_W
            else:
                x1 = x2 - self.MIN_W
        if y2 - y1 < self.MIN_H:
            if edges.get("b", False):
                y2 = y1 + self.MIN_H
            else:
                y1 = y2 - self.MIN_H

        if not allow_oversize:
            # 動画解像度の範囲内に制約
            if x1 < 0:
                x1 = 0
                if edges.get("l", False):
                    x2 = max(x2, self.MIN_W)
            if x2 > vw:
                x2 = vw
                if edges.get("r", False):
                    x1 = min(x1, vw - self.MIN_W)
            if y1 < 0:
                y1 = 0
                if edges.get("t", False):
                    y2 = max(y2, self.MIN_H)
            if y2 > vh:
                y2 = vh
                if edges.get("b", False):
                    y1 = min(y1, vh - self.MIN_H)

        return [int(x1), int(y1), int(x2), int(y2)]

    def inside_rect(self, x: int, y: int) -> bool:
        """座標がクロップ矩形の内側にあるかを判定する."""
        x1, y1, x2, y2 = self.crop_rect
        scaled_x1 = int(x1 * self.canvas_scale_x) + self.canvas_offset_x
        scaled_y1 = int(y1 * self.canvas_scale_y) + self.canvas_offset_y
        scaled_x2 = int(x2 * self.canvas_scale_x) + self.canvas_offset_x
        scaled_y2 = int(y2 * self.canvas_scale_y) + self.canvas_offset_y
        return scaled_x1 <= x <= scaled_x2 and scaled_y1 <= y <= scaled_y2

    def near_edge(self, x: int, y: int, m: int = 20) -> dict[str, bool]:
        """座標がクロップ矩形のエッジ近くにあるかを判定する."""
        x1, y1, x2, y2 = self.crop_rect
        scaled_x1 = int(x1 * self.canvas_scale_x) + self.canvas_offset_x
        scaled_y1 = int(y1 * self.canvas_scale_y) + self.canvas_offset_y
        scaled_x2 = int(x2 * self.canvas_scale_x) + self.canvas_offset_x
        scaled_y2 = int(y2 * self.canvas_scale_y) + self.canvas_offset_y
        
        left = abs(x - scaled_x1) < m and (scaled_y1 - m) <= y <= (scaled_y2 + m)
        right = abs(x - scaled_x2) < m and (scaled_y1 - m) <= y <= (scaled_y2 + m)
        top = abs(y - scaled_y1) < m and (scaled_x1 - m) <= x <= (scaled_x2 + m)
        bottom = abs(y - scaled_y2) < m and (scaled_x1 - m) <= x <= (scaled_x2 + m)
        
        return {"l": left, "r": right, "t": top, "b": bottom}

    def canvas_mouse_to_image_coords(
        self, canvas_x: float, canvas_y: float
    ) -> tuple[float, float]:
        """キャンバス上のマウス座標を元の画像座標系に変換する."""
        image_x = (canvas_x - self.canvas_offset_x) / \
            self.canvas_scale_x if self.canvas_scale_x > 0 else canvas_x
        image_y = (canvas_y - self.canvas_offset_y) / \
            self.canvas_scale_y if self.canvas_scale_y > 0 else canvas_y
        return image_x, image_y

    def move_crop_by(self, dx: int, dy: int) -> None:
        """クロップ矩形をdx,dyだけ移動する（ピクセル単位）."""
        x1, y1, x2, y2 = self.crop_rect
        new_x1 = x1 + dx
        new_y1 = y1 + dy
        new_x2 = x2 + dx
        new_y2 = y2 + dy
        
        # 範囲チェック
        vw = getattr(self, "vid_w", self.CANVAS_W)
        vh = getattr(self, "vid_h", self.CANVAS_H)
        allow_oversize = getattr(self, 'allow_oversize_var', None) and self.allow_oversize_var.get()
        
        if not allow_oversize:
            if new_x1 < 0:
                new_x1 = 0
                new_x2 = x2 - x1
            if new_y1 < 0:
                new_y1 = 0
                new_y2 = y2 - y1
            if new_x2 > vw:
                new_x2 = vw
                new_x1 = vw - (x2 - x1)
            if new_y2 > vh:
                new_y2 = vh
                new_y1 = vh - (y2 - y1)
        
        self.crop_rect = [int(new_x1), int(new_y1), int(new_x2), int(new_y2)]
        self._sync_crop_rect_ui()
