"""設定管理モジュール.

アプリケーション設定と動画個別設定の読み書きを担当。
"""
from __future__ import annotations

import json
import os
from typing import TYPE_CHECKING, Any

from utils import get_base_dir, ratio_label_from_wh

# 定数
PROJECT_NAME = "ChulipVideo"

# 設定ファイル名を定数化
CONFIG_FILENAME = f"{PROJECT_NAME}_config.json"

# デフォルト解像度プリセット
DEFAULT_PRESETS = {
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


def get_config_path() -> str:
    """設定ファイルのパスを返す. 実行ファイルと同じディレクトリ。"""
    import sys
    # utils.get_base_dir() と同等のロジックで実行ファイルディレクトリを取得
    if getattr(sys, 'frozen', False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_dir, CONFIG_FILENAME)


def load_global_config() -> dict[str, Any]:
    """グローバル設定ファイル（アプリ共通設定）を読み込む."""
    config_path = get_config_path()
    theme_defaults = {
        "theme": {
            "main_color": "#ffcccc",
            "active_color": "#ff9999",
            "crop_width": 2,
            
            # 背景色
            "canvas_bg": "#f5f5f5",
            
            # クロップ枠線の設定 (命名規則: crop_[状態]_[linecolor/linestyle])
            "crop_default_linecolor": "#FF8C00", # 橙
            "crop_default_linestyle": (5, 5),
            "crop_focused_linecolor": "#FF0000",   # 赤
            "crop_focused_linestyle": "",          # 実線
            "crop_hover_linecolor": "#FF0000",     # 赤
            "crop_hover_linestyle": (5, 5),        # 破線
            
            # 四つ角のドット設定
            "handle_color": "#FFFFFF",
            "handle_size": 8,
            "edge_margin": 20,
            
            # ボタンの色 (パステル配色)
            "button_play_bg": "#B3E5FC",    # 薄い青 (button_undo_bgと共通)
            "button_stop_bg": "#EF9A9A",    # 薄い赤
            "button_export_bg": "#F48FB1",  # 薄いピンク
            "button_video_bg": "#81C784",   # パステル緑
            "button_gif_bg": "#FFCC80",     # 薄い橙
            "button_copy_bg": "#80DEEA",    # 薄いシアン
            "button_normal_bg": "#E0E0E0",  # 明るいグレー
            
            # 追加ボタンの色
            "button_help_bg": "#FFF59D",    # 薄い黄色
            "button_reload_bg": "#FFCC80",  # 薄いオレンジ
            "button_save_bg": "#90CAF9",    # 薄い青色
            "button_trim_start_bg": "#77BB77", # 開始色 (start_color_bgと共通)
            "button_trim_end_bg": "#FF7777",   # 終了色 (end_color_bgと共通)
            "button_undo_bg": "#B3E5FC",    # 薄い青
            "button_redo_bg": "#B3E5FC",    # 薄い青
            "button_locked_bg": "#FFAB91",   # 薄い赤（ロック中）
            "button_unlocked_bg": "#B9F6CA", # 薄い緑（解除中）
            "start_color_bg": "#77BB77",     # 開始カラー
            "end_color_bg": "#FF7777",       # 終了カラー

            # マウス軌跡オーバーレイ設定
            "mouse_overlay": {
                # ポインタ（通常時）の設定
                "pointer": {
                    "shape": "cursor",        # ポインタ位置を示す形状 (circle, square, cursor)
                    "color": "#000000",       # ポインタの罫線の色 (既定: 黒)
                    "width": 2,               # ポインタの罫線の太さ
                    "fill": "",               # ポインタの塗りつぶしの色（空文字なら透明）
                    "radius": 6               # ポインタの半径/サイズ
                },
                # 左クリック時の設定
                "click_left": {
                    "color": "#FF0000",       # 左クリック時の強調色 (既定: 赤)
                    "shape": "circle",        # 左クリック時の形状
                    "width": 3,               # 左クリック時の太さ
                    "ripple_duration": 0.5,   # 左クリックを離した時の波紋が消えるまでの時間（秒）
                    "ripple_range": 20        # 左クリックを離した時の波紋の広がる範囲（ピクセル）
                },
                # 右クリック時の設定
                "click_right": {
                    "color": "#0000FF",       # 右クリック時の強調色 (既定: 青)
                    "shape": "circle",        # 右クリック時の形状
                    "width": 3,               # 右クリック時の太さ
                    "ripple_duration": 0.5,   # 右クリックを離した時の波紋が消えるまでの時間（秒）
                    "ripple_range": 20        # 右クリックを離した時の波紋の広がる範囲（ピクセル）
                },
                # 中クリック時の設定
                "click_middle": {
                    "color": "#00FF00",       # 中クリック時の強調色 (既定: 緑)
                    "shape": "circle",        # 中クリック時の形状
                    "width": 3,               # 中クリック時の太さ
                    "ripple_duration": 0.5,   # 中クリックを離した時の波紋が消えるまでの時間（秒）
                    "ripple_range": 20        # 中クリックを離した時の波紋の広がる範囲（ピクセル）
                }
            },
            
            # 入力（キー・マウス操作）オーバーレイ設定
            "input_overlay": {
                "position": "center",         # 文字の位置（left, right, center）
                "v_position": "bottom",       # 上下位置（top, center, bottom）
                "font_family": "msgothic.ttc",# フォントの種類（msgothic.ttc, meiryo.ttc 等、OS標準フォント推奨）
                "font_size": 24,              # 文字の大きさ
                "font_weight": "bold",        # 文字の太さ
                "font_color": "#000000",      # 文字の前景色
                "outline_color": "",          # 文字のアウトライン色（空文字ならOFF）
                "bg_color": "#FFFFFF",        # 入力を描く四角形の背景塗りつぶし色
                "offset_x": 0,                # 表示領域の横オフセット
                "offset_y": 0,                # 表示領域の縦オフセット
                "text_offset_y": -2,          # 文字の描画位置の微調整（負の数で上に移動）
                "max_stack": 3,               # 最大表示数
                "fade_duration": 1.0          # 入力を離してから文字が消えるまでの時間（秒）
            }
        },
        "window_x": None,
        "window_y": None,
        "window_width": 1000,
        "window_height": 700,
        "window_maximized": False,
        "last_video_path": "",
        "resolution_presets": {},
        "selected_ratio": "未指定",
        "png_compression": 3,
        "check_prev_next": True,
        "check_duplicate": True,
        "play_speed": 1.0,
        "play_range": False,
        "play_loop": False,
        "play_pingpong": False,
        "show_trajectory": True
    }

    config_loaded = {}
    should_update_file = False

    if os.path.exists(config_path):
        try:
            if os.path.getsize(config_path) > 0:
                with open(config_path, "r", encoding="utf-8") as f:
                    config_loaded = json.load(f)
            else:
                # 0バイトファイルは破損とみなして上書きフラグを立てる
                should_update_file = True
        except Exception as e:
            print(f"設定ファイルの読み込みに失敗しました: {e}")
            # 破損している場合は初期化するために上書きフラグを立てる
            should_update_file = True

    # 古い設定名から新しい設定名への移行マッピング
    migration_map = {
        "crop_color_unfocused": "crop_default_linecolor",
        "crop_unfocused_linecolor": "crop_default_linecolor",
        "crop_color_focused": "crop_focused_linecolor",
        "crop_dash_unfocused": "crop_default_linestyle",
        "crop_unfocused_linestyle": "crop_default_linestyle",
        "crop_dash_focused": "crop_focused_linestyle",
    }

    # テーマ設定内の移行と補完
    loaded_theme = config_loaded.get("theme", {})
    default_theme = theme_defaults["theme"]

    for old_key, new_key in migration_map.items():
        if old_key in loaded_theme and new_key not in loaded_theme:
            loaded_theme[new_key] = loaded_theme.pop(old_key)
            should_update_file = True

    for k, v in default_theme.items():
        if k not in loaded_theme:
            loaded_theme[k] = v
            should_update_file = True
    
    config_loaded["theme"] = loaded_theme

    # グローバル項目の補完
    for k, v in theme_defaults.items():
        if k == "theme": continue
        if k not in config_loaded:
            config_loaded[k] = v
            should_update_file = True

    # 新規項目があった場合、または移行が行われた場合は保存する
    if should_update_file:
        save_global_config(config_loaded)

    return config_loaded


def save_global_config(config: dict[str, Any]) -> bool:
    """グローバル設定ファイルに保存する."""
    config_path = get_config_path()
    try:
        with open(config_path, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        print(f"設定ファイルの保存に失敗しました: {e}")
        return False


def load_video_settings(video_filepath: str) -> dict[str, Any] | None:
    """動画個別の設定ファイルを読み込む（動画ファイルと同階層）."""
    if not video_filepath:
        return None
    settings_path = os.path.splitext(video_filepath)[0] + '.settings.json'
    if os.path.exists(settings_path):
        try:
            with open(settings_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            pass
    return None


def save_video_settings(
    video_filepath: str,
    crop_rect: list[int],
    start_time: float,
    end_time: float,
    current_time: float = 0.0,
    allow_oversize: bool = False,
    additional_data: dict[str, Any] | None = None
) -> str | None:
    """動画個別の設定ファイルに保存する."""
    if not video_filepath:
        return None
    
    save_path = os.path.splitext(video_filepath)[0] + '.settings.json'
    data = {
        'video_file': video_filepath,
        'crop_rect': {
            'x1': int(crop_rect[0]),
            'y1': int(crop_rect[1]),
            'x2': int(crop_rect[2]),
            'y2': int(crop_rect[3])
        },
        'start_time': start_time,
        'end_time': end_time,
        'current_time': current_time,
        'allow_oversize': allow_oversize
    }
    if additional_data:
        data.update(additional_data)
        
    try:
        with open(save_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        return save_path
    except Exception:
        return None


def normalize_presets(presets: dict[str, list[int]]) -> dict[str, list[int]]:
    """プリセットのキーに比率ラベルがなければ付与する."""
    import re
    new_presets = {}
    for k, v in presets.items():
        if re.match(r'^\d+:\d+\s', str(k)):
            new_presets[k] = v
        else:
            try:
                w, h = int(v[0]), int(v[1])
                lbl = ratio_label_from_wh(w, h)
                new_key = f"{lbl} {k}"
                # avoid collision
                if new_key in new_presets:
                    idx = 1
                    while f"{new_key}#{idx}" in new_presets:
                        idx += 1
                    new_key = f"{new_key}#{idx}"
                new_presets[new_key] = v
            except Exception:
                new_presets[k] = v
    return new_presets


def get_default_presets_with_labels() -> dict[str, list[int]]:
    """比率ラベル付きのデフォルトプリセットを返す."""
    new_defaults = {}
    for k, v in DEFAULT_PRESETS.items():
        try:
            w, h = int(v[0]), int(v[1])
            lbl = ratio_label_from_wh(w, h)
            new_key = f"{lbl} {k}"
        except Exception:
            new_key = k
        new_defaults[new_key] = v
    return new_defaults
