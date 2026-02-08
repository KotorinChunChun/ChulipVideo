"""オーバーレイ表示（マウス軌跡・キー入力）の共通ユーティリティ.
PILを使用して画像に直接描画することで透過表示に対応します。
"""
from __future__ import annotations

from typing import Dict, Any, List, Optional, Tuple
from PIL import Image, ImageDraw, ImageFont


def convert_tk_event_to_key_name(keysym: str) -> str:
    """Tkinterのkeysymを、録画データで使用する共通形式のキー名に変換する."""
    # recorder_core.py で定義されている WinAPI ベースのキー名に合わせる
    key_map = {
        "Return": "Enter",
        "BackSpace": "BS",
        "Escape": "Esc",
        "space": "Space",
        "Delete": "Del",
        "Prior": "PageUp",
        "Next": "PageDown",
        "Control_L": "Ctrl", "Control_R": "Ctrl",
        "Shift_L": "Shift", "Shift_R": "Shift",
        "Alt_L": "Alt", "Alt_R": "Alt",
        "Win_L": "Win", "Win_R": "Win",
        "Meta_L": "Win", "Meta_R": "Win", # Windowsキーの別名対応
        "Super_L": "Win", "Super_R": "Win"
    }
    
    # マッピングがあればそれを使う
    if keysym in key_map:
        return key_map[keysym]
    
    # 1文字の英字キーは大文字にする
    if len(keysym) == 1 and keysym.isalpha():
        return keysym.upper()
        
    return keysym


def get_input_display_text(click: str, keys: str) -> str:
    """マウスのクリックとキー入力文字列から、表示用のテキストを生成する。"""
    modifiers_found = []
    mouse_found = []
    other_keys_found = []
    modifier_order = ["Ctrl", "Shift", "Alt", "Win"]
    
    if keys and keys != "None":
        k_list = [k.strip() for k in keys.split(",") if k.strip()]
        for k in k_list:
            if k in modifier_order:
                modifiers_found.append(k)
            else:
                # 1文字の英字キーは大文字に固定
                if len(k) == 1 and k.isalpha():
                    k = k.upper()
                other_keys_found.append(k)
    # 修飾キーを規定の順序でソートし、重複を排除
    modifiers_found = sorted(list(set(modifiers_found)), key=lambda x: modifier_order.index(x) if x in modifier_order else 99)
    
    if click and click != "None":
        if "L" in click: mouse_found.append("左クリック")
        if "R" in click: mouse_found.append("右クリック")
        if "M" in click: mouse_found.append("中クリック")

    combined_parts = modifiers_found + mouse_found + other_keys_found
    
    # 修飾キーのみの場合は表示しない（何か具体的操作が含まれる場合のみ表示）
    has_non_modifier = (len(mouse_found) > 0) or (len(other_keys_found) > 0)
    if combined_parts and has_non_modifier:
        return " + ".join(combined_parts)
    return ""


def _hex_to_rgba(hex_color: str, alpha: int = 255) -> Tuple[int, int, int, int]:
    """16進数カラーコードをRGBAに変換する."""
    if not hex_color or hex_color == "":
        return (0, 0, 0, 0)
    hex_color = hex_color.lstrip('#')
    try:
        if len(hex_color) == 6:
            r, g, b = tuple(int(hex_color[i:i+2], 16) for i in range(0, 6, 2))
        elif len(hex_color) == 3:
            r, g, b = tuple(int(hex_color[i:i+1]*2, 16) for i in range(0, 3))
        else:
            return (0, 0, 0, alpha)
        return (r, g, b, alpha)
    except:
        return (0, 0, 0, alpha)


def _get_star_points(cx: float, cy: float, r: float, rotation_deg: float = 0) -> List[Tuple[float, float]]:
    """5外角の星形の頂点を計算する."""
    points = []
    import math
    for i in range(10):
        # 偶数インデックスが外側、奇数インデックスが内側
        angle = math.radians(90 - 36 * i - rotation_deg)
        dist = r if i % 2 == 0 else r * 0.4
        points.append((cx + dist * math.cos(angle), cy - dist * math.sin(angle)))
    return points


def draw_mouse_overlay(
    img: Image.Image,
    x: int, y: int,
    click_info: str,
    scale_x: float, scale_y: float,
    theme: Dict[str, Any],
    ripple_age: float = 0.0,
    ripple_type: str = ""
):
    """マウスのポインタとクリック箇所を画像に描画する."""
    # RGBAのオーバーレイレイヤーを作成してそこに描画
    overlay = Image.new("RGBA", img.size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay, "RGBA")
    m_cfg = theme.get("mouse_overlay", {})
    
    # 画像内の座標
    ix = x * scale_x
    iy = y * scale_y

    # 1. 進行中のクリック表現 (押しっぱなし)
    if click_info and click_info != "None":
        c_cfg = None
        if "L" in click_info: c_cfg = m_cfg.get("click_left")
        elif "R" in click_info: c_cfg = m_cfg.get("click_right")
        elif "M" in click_info: c_cfg = m_cfg.get("click_middle")

        if c_cfg:
            color_hex = c_cfg.get("color", "#FFFF00")
            color = _hex_to_rgba(color_hex, 180) # 若干透過
            width = c_cfg.get("width", 3)
            shape = c_cfg.get("shape", "circle")
            
            r = 12
            if shape == "circle":
                draw.ellipse([ix - r, iy - r, ix + r, iy + r], outline=color, width=width)
            elif shape == "square":
                draw.rectangle([ix - r, iy - r, ix + r, iy + r], outline=color, width=width)
            elif shape == "star":
                draw.polygon(_get_star_points(ix, iy, r), outline=color, width=width)
            elif shape == "rote_star":
                draw.polygon(_get_star_points(ix, iy, r, 36), outline=color, width=width)

    # 2. 波紋 (離した瞬間の演出、フェードアウト)
    if ripple_type and ripple_age > 0:
        c_cfg = m_cfg.get(f"click_{ripple_type}")
        if c_cfg:
            duration = c_cfg.get("ripple_duration", 0.5)
            max_range = c_cfg.get("ripple_range", 20)
            if ripple_age < duration:
                progress = ripple_age / duration
                alpha = int(255 * (1.0 - progress)) # 徐々に透明に
                color_hex = c_cfg.get("color", "#FFFF00")
                color = _hex_to_rgba(color_hex, alpha)
                width = c_cfg.get("width", 2)
                shape = c_cfg.get("shape", "circle")
                
                # 半径は 12 から 12 + max_range まで広がる
                r = 12 + progress * max_range
                if shape == "circle":
                    draw.ellipse([ix - r, iy - r, ix + r, iy + r], outline=color, width=width)
                elif shape == "square":
                    draw.rectangle([ix - r, iy - r, ix + r, iy + r], outline=color, width=width)
                elif shape == "star":
                    draw.polygon(_get_star_points(ix, iy, r), outline=color, width=width)
                elif shape == "rote_star":
                    # 時間経過(age)に応じて回転させる (1秒で180度)
                    spin_angle = 36 + ripple_age * 180
                    draw.polygon(_get_star_points(ix, iy, r, spin_angle), outline=color, width=width)

    # 3. ポインタ本体
    p_cfg = m_cfg.get("pointer", {})
    p_color_hex = p_cfg.get("color", "#FF0000")
    p_color = _hex_to_rgba(p_color_hex, 255) if p_color_hex else None
    p_fill_hex = p_cfg.get("fill", "")
    p_fill = _hex_to_rgba(p_fill_hex, 100) if p_fill_hex else None
    
    p_width = p_cfg.get("width", 2)
    p_shape = p_cfg.get("shape", "circle")
    p_r = p_cfg.get("radius", 6)

    if p_shape == "circle":
        draw.ellipse([ix - p_r, iy - p_r, ix + p_r, iy + p_r], outline=p_color, width=p_width, fill=p_fill)
        if p_color: draw.ellipse([ix - 2, iy - 2, ix + 2, iy + 2], fill=p_color)
    elif p_shape == "square":
        draw.rectangle([ix - p_r, iy - p_r, ix + p_r, iy + p_r], outline=p_color, width=p_width, fill=p_fill)
        if p_color: draw.rectangle([ix - 2, iy - 2, ix + 2, iy + 2], fill=p_color)
    elif p_shape == "star":
        draw.polygon(_get_star_points(ix, iy, p_r), outline=p_color, width=p_width, fill=p_fill)
    elif p_shape == "rote_star":
        draw.polygon(_get_star_points(ix, iy, p_r, 36), outline=p_color, width=p_width, fill=p_fill)
    elif p_shape == "cursor":
        # 基準サイズ 10 としたときのスケーリング
        s = p_r / 10.0
        arrow_points = [
            (ix, iy),                   # 先端
            (ix, iy + 17 * s),          # 左下
            (ix + 4 * s, iy + 13 * s),  # くぼみ（左）
            (ix + 9 * s, iy + 22 * s),  # 足の先端（右側）
            (ix + 12 * s, iy + 20 * s), # 足の先端（左側）
            (ix + 7 * s, iy + 11 * s),  # くぼみ（右）
            (ix + 12 * s, iy + 12 * s), # 右端
        ]
        # 外枠（既定: 黒）と中身（既定: 白）
        fill_color = p_fill if p_fill else (255, 255, 255, 255)
        outline_c = p_color if p_color else (0, 0, 0, 255)
        draw.polygon(arrow_points, fill=fill_color, outline=outline_c)

    # 描画済みレイヤーを元の画像に合成
    if img.mode == "RGBA":
        img.alpha_composite(overlay)
    else:
        img.paste(overlay, (0, 0), overlay)


def draw_input_overlay(
    img: Image.Image,
    inputs_with_age: List[Tuple[str, float]],
    scale_x: float,
    scale_y: float,
    theme: Dict[str, Any],
    target_rect: Optional[List[int]] = None
):
    """入力操作（キー・マウス）の履歴を画像上に描画する.
    
    inputs_with_age: (文字列, 経過時間sec) のリスト (新しいものが先頭)
    target_rect: [x1, y1, x2, y2] 描画対象の範囲 (Noneなら画像全体)
    """
    if not inputs_with_age:
        return

    # RGBAのオーバーレイレイヤーを作成してそこに描画
    overlay = Image.new("RGBA", img.size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay, "RGBA")
    k_cfg = theme.get("input_overlay", {})
    m_cfg = theme.get("mouse_overlay", {})
    
    # スケールを考慮 (縦横比維持のフォントサイズなら scale_y を基準にするなど)
    scale = (scale_x + scale_y) / 2.0 if (scale_x > 0 and scale_y > 0) else 1.0
    
    pos_type = k_cfg.get("position", "center")
    font_family = k_cfg.get("font_family", "msgothic.ttc")
    base_font_size = k_cfg.get("font_size", 24)
    font_size = max(8, int(base_font_size * scale))
    
    font_color_hex = k_cfg.get("font_color", "#000000")
    bg_color_hex = k_cfg.get("bg_color", "#FFFFFF")
    # 字幕全体のアウトライン設定
    base_outline_color_hex = k_cfg.get("outline_color", "")
    
    area_offset_x = int(k_cfg.get("offset_x", 0) * scale)
    area_offset_y = int(k_cfg.get("offset_y", 0) * scale)
    text_offset_y = int(k_cfg.get("text_offset_y", -2) * scale)
    max_stack = k_cfg.get("max_stack", 3)
    
    fade_duration = k_cfg.get("fade_duration", 1.0)
    # フェードアウトの開始を 0.5 秒前に変更して滑らかにする
    fade_out_range = 0.5
    fade_out_start = max(0.0, fade_duration - fade_out_range)

    # フォントの読み込み
    font = None
    for f_name in [font_family, "msgothic.ttc", "meiryo.ttc", "arialbd.ttf"]:
        try:
            font = ImageFont.truetype(f_name, font_size)
            break
        except:
            continue
    if not font:
        font = ImageFont.load_default()

    img_w, img_h = img.size
    
    # 描画範囲の決定
    if target_rect:
        tx1, ty1, tx2, ty2 = target_rect
        tw_area, th_area = tx2 - tx1, ty2 - ty1
    else:
        tx1, ty1 = 0, 0
        tw_area, th_area = img_w, img_h

    padding_x = 8
    padding_y = k_cfg.get("padding_y", 5)
    margin = 8
    
    max_stack = k_cfg.get("max_stack", 3)
    display_items = inputs_with_age[:max_stack]
    
    unit_h = (font_size + padding_y * 2 + margin)
    v_pos = k_cfg.get("v_position", "bottom")

    # 開始位置調整 (max_stack に応じて画面中央方向へ押し戻す)
    # 1件なら 20px (端), 5件なら 20 + 4*unit_h (中央寄り)
    # dynamic_margin は「最新の1件（最も内側）」の端からの距離
    dynamic_margin = 20 + (max_stack - 1) * unit_h
    
    if v_pos == "top":
        # 上部表示: 1件目を dynamic_margin (内側) に置く
        base_y = ty1 + dynamic_margin + area_offset_y
    elif v_pos == "center":
        # 全スタック分を考慮した中央
        total_h = unit_h * len(display_items)
        base_y = ty1 + (th_area - total_h) // 2 + area_offset_y
    else: # bottom
        # 下部表示: 1件目を dynamic_margin (内側) に置く
        base_y = ty1 + th_area - dynamic_margin - unit_h + area_offset_y
    
    current_y = base_y
    for i, (raw_text, age) in enumerate(display_items):
        # ... (中略: 色やテキストの準備)
        
        # --- (ここからループ内描画) ---
        # 修飾キーのみの場合は非表示にする
        keys = [k.strip() for k in raw_text.replace(',', ' ').split()]
        modifiers = {"Ctrl", "Alt", "Shift", "Win"}
        has_content = any("クリック" in k for k in keys) or any(k not in modifiers for k in keys)
        if not has_content:
            continue

        display_text = f" {raw_text.replace(',', ' + ')} "
        
        # 透明度の計算
        alpha_factor = 1.0
        if age > fade_out_start:
            alpha_factor = max(0.0, 1.0 - (age - fade_out_start) / fade_out_range)
        if alpha_factor <= 0: continue

        font_alpha = int(255 * alpha_factor)
        bg_alpha = int(200 * alpha_factor)
        c_font = _hex_to_rgba(font_color_hex, font_alpha)
        c_bg = _hex_to_rgba(bg_color_hex, bg_alpha)
        
        box_outline_color_hex = base_outline_color_hex
        if "左クリック" in raw_text: box_outline_color_hex = m_cfg.get("click_left", {}).get("color", "#FF0000")
        elif "右クリック" in raw_text: box_outline_color_hex = m_cfg.get("click_right", {}).get("color", "#0000FF")
        elif "中クリック" in raw_text: box_outline_color_hex = m_cfg.get("click_middle", {}).get("color", "#00FF00")
        
        c_box_outline = _hex_to_rgba(box_outline_color_hex, font_alpha) if box_outline_color_hex else None
        c_text_outline = _hex_to_rgba(base_outline_color_hex, font_alpha) if base_outline_color_hex else None

        try:
            left, top, right, bottom = draw.textbbox((0, 0), display_text, font=font)
            tw, th = right - left, bottom - top
        except:
            tw, th = font.getsize(display_text)

        if pos_type == "left": draw_x = tx1 + 50 + area_offset_x
        elif pos_type == "right": draw_x = tx1 + tw_area - 50 - tw + area_offset_x
        else: draw_x = tx1 + (tw_area - tw) // 2 + area_offset_x

        # 背景ボックス
        final_box_outline = c_box_outline if c_box_outline else c_font
        draw.rectangle([draw_x - padding_x, current_y - padding_y, draw_x + tw + padding_x, current_y + th + padding_y], 
                       fill=c_bg, outline=final_box_outline)

        if c_text_outline:
            for dx, dy in [(-1,0), (1,0), (0,-1), (0,1)]:
                draw.text((draw_x + dx, current_y + dy + text_offset_y), display_text, fill=c_text_outline, font=font)
        draw.text((draw_x, current_y + text_offset_y), display_text, fill=c_font, font=font)
        
        # --- 次の段へ ---
        if v_pos == "top":
            # 上部表示なら、古いものは上に逃がす
            current_y -= (th + padding_y * 2 + margin)
        else:
            # 中央や下部なら、古いものは下に逃がす
            current_y += (th + padding_y * 2 + margin)
        
        # 画面外に十分出たら終了
        if current_y < -unit_h or current_y > img_h + unit_h:
            break

    # 描画済みレイヤーを元の画像に合成
    if img.mode == "RGBA":
        img.alpha_composite(overlay)
    else:
        img.paste(overlay, (0, 0), overlay)


class InputHistoryManager:
    def __init__(self):
        self.history = []  # (text, timestamp)
        self.last_text = None
        self.last_time = -1.0
        # 連続入力とみなす時間は呼び出し元で調整してもよいが、ここでは固定
        self.stack_timeout = 1.0 

    def update(self, current_time, check_text):
        """現在の入力テキストに基づいて履歴を更新する"""
        # 入力がない（離された）場合
        if not check_text:
            # 入力が途切れたので直前のキー情報をリセット
            self.last_text = None
            return

        # 直前の入力と同じなら履歴に追加しない（押しっぱなし対策）
        # 以前は0.2秒などの猶予を持たせていたが、
        # 「離されるまでは追加表示しない」という要望のため、時間経過に関わらず
        # 同一テキストが継続している間は無視する。
        if check_text == self.last_text:
            # 表示期間（フェードアウト開始時間）の延長を行いたい場合は
            # ここで last_time を更新する手もあるが、今回は「追加表示の抑制」が主眼。
            self.last_time = current_time
            return

        # 新しい入力を追加
        self.history.append((check_text, current_time))
        self.last_text = check_text
        self.last_time = current_time
        
        # 古すぎる履歴は削除 (描画時にフィルタリングするのでここでは緩めに)
        # 10秒以上前のものは不要
        self.history = [h for h in self.history if current_time - h[1] < 10.0]

    def get_active_inputs(self, current_time, fade_duration):
        """現在表示すべき入力リスト (テキスト, 経過時間) を返す. 新しい順"""
        active = []
        # 新しい順に取得して、古いものは描画しない
        # historyは古い順に入っているため reversed する
        for text, timestamp in reversed(self.history):
            age = current_time - timestamp
            if 0 <= age <= fade_duration:
                active.append((text, age))
        return active
