"""設定管理モジュール.

アプリケーション設定と動画個別設定の読み書きを担当。
"""
from __future__ import annotations

import json
import os
from typing import TYPE_CHECKING, Any

from utils import get_base_dir, ratio_label_from_wh

# 設定ファイル名を定数化
CONFIG_FILENAME = "video_frame_cropper_config.json"

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
    """設定ファイルのパスを返す."""
    return os.path.join(get_base_dir(), CONFIG_FILENAME)


def load_global_config() -> dict[str, Any]:
    """グローバル設定ファイルを読み込む."""
    config_path = get_config_path()
    if os.path.exists(config_path):
        try:
            with open(config_path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            print(f"設定ファイルの読み込みに失敗しました: {e}")
    return {}


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
    """動画個別の設定ファイルを読み込む.
    
    Args:
        video_filepath: 動画ファイルのパス
        
    Returns:
        設定辞書、またはファイルが存在しない/読み込みエラーの場合None
    """
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
    start_time: int,
    end_time: int
) -> str | None:
    """動画個別の設定ファイルに保存する.
    
    Args:
        video_filepath: 動画ファイルのパス
        crop_rect: クロップ矩形 [x1, y1, x2, y2]
        start_time: 開始時間（秒）
        end_time: 終了時間（秒）
        
    Returns:
        保存先パス、失敗時はNone
    """
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
        'start_time': int(start_time),
        'end_time': int(end_time)
    }
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
