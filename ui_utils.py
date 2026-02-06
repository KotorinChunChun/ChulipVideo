"""UIユーティリティモジュール.

ツールチップ、ウィンドウジオメトリ管理など、UI関連の共通機能を提供。
"""
from __future__ import annotations

import tkinter as tk
from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    pass


def add_tooltip(widget: tk.Widget, text: str) -> None:
    """ウィジェットにツールチップを追加する.
    
    フォーカスやホバー時にツールチップを表示する。
    
    Args:
        widget: ツールチップを追加するウィジェット
        text: ツールチップに表示するテキスト
    """
    def show(e=None):
        try:
            if getattr(widget, '_tooltip_win', None):
                return
            # ウィジェットの直下に表示するように座標を計算（重なりを防止）
            x = widget.winfo_rootx() + 10
            y = widget.winfo_rooty() + widget.winfo_height() + 5
            tw = tk.Toplevel(widget)
            tw.wm_overrideredirect(True)
            tw.wm_geometry(f"+{x}+{y}")
            tw.attributes("-topmost", True)
            lbl = tk.Label(tw, text=text, background="#ffffe0", relief='solid', borderwidth=1, padx=3, pady=1)
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


def fix_button_active_colors(container: tk.Widget) -> None:
    """コンテナ内の全ボタンの activebackground を背景色に合わせる.
    
    tkのボタンはクリック時にactivebackgroundが使われるが、デフォルトだと
    背景色と異なる色になるため、この関数で統一する。
    
    Args:
        container: 対象のコンテナウィジェット
    """
    for child in container.winfo_children():
        if isinstance(child, tk.Button):
            bg = child.cget("bg")
            if bg and bg != "SystemButtonFace":
                child.config(activebackground=bg)
        elif isinstance(child, (tk.Frame, tk.LabelFrame)):
            fix_button_active_colors(child)


class WindowGeometryMixin:
    """ウィンドウジオメトリ管理のMixinクラス.
    
    ウィンドウの位置・サイズ・最大化状態の保存と復元を提供する。
    使用するクラスは以下の属性を持つ必要がある:
    - root: tk.Tk インスタンス
    - _geometry_prefix: str (設定キーのプレフィックス、例: "recorder_" or "")
    """
    
    # サブクラスで上書き可能
    _geometry_prefix: str = ""
    _default_geometry: str = "1000x700"
    _min_width: int = 800
    _min_height: int = 600
    
    def load_window_geometry(self) -> None:
        """設定ファイルからウィンドウの位置とサイズを復元する."""
        from config import load_global_config
        config = load_global_config()
        prefix = self._geometry_prefix
        
        # ジオメトリ文字列がある場合はそれを使用
        geom = config.get(f"{prefix}geometry")
        if geom:
            try:
                self.root.geometry(geom)
                self.root.update_idletasks()
            except Exception:
                self.root.geometry(self._default_geometry)
        else:
            # 個別のwidth/height/x/yがある場合
            width = config.get(f"{prefix}window_width", self._min_width)
            height = config.get(f"{prefix}window_height", self._min_height)
            x = config.get(f"{prefix}window_x")
            y = config.get(f"{prefix}window_y")
            
            if x is not None and y is not None:
                x_sign = '+' if x >= 0 else ''
                y_sign = '+' if y >= 0 else ''
                geometry = f"{width}x{height}{x_sign}{x}{y_sign}{y}"
            else:
                geometry = f"{width}x{height}"
            
            try:
                self.root.geometry(geometry)
                self.root.update_idletasks()
            except Exception as e:
                print(f"ジオメトリ設定エラー: {e}")
                self.root.geometry(self._default_geometry)
        
        # 最大化状態の復元
        if config.get(f"{prefix}maximized", False):
            try:
                self.root.state('zoomed')
            except Exception:
                pass
    
    def save_window_geometry(self) -> None:
        """ウィンドウの位置とサイズを設定ファイルに保存する."""
        from config import load_global_config, save_global_config
        config = load_global_config()
        prefix = self._geometry_prefix
        
        # 最大化状態
        is_zoomed = (self.root.state() == 'zoomed')
        config[f"{prefix}maximized"] = is_zoomed
        
        if not is_zoomed:
            # 最大化されていない場合のみ位置とサイズを保存
            config[f"{prefix}window_x"] = self.root.winfo_x()
            config[f"{prefix}window_y"] = self.root.winfo_y()
            config[f"{prefix}window_width"] = self.root.winfo_width()
            config[f"{prefix}window_height"] = self.root.winfo_height()
            # geometry文字列も保存（screen_recorder互換用）
            config[f"{prefix}geometry"] = self.root.geometry()
        
        save_global_config(config)
