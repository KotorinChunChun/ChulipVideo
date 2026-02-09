"""ショートカットキー管理モジュール.

記録対象とするショートカットキーのホワイトリスト管理、
TSVファイルへの読み書き、入力が許可されているかの判定を行う。
全ての物理可能なキー組み合わせを網羅する。
"""
from __future__ import annotations

import csv
import itertools
import os
from typing import Dict, List, Set, Tuple

def generate_all_shortcuts() -> List[Dict]:
    """全ての可能なショートカットキー組み合わせを生成する。"""
    
    # 1. 構成要素の定義
    
    # 修飾キー (順序: Ctrl, Shift, Alt, Win) - recorder_core.py と合わせる
    modifiers_base = ["Ctrl", "Shift", "Alt", "Win"]
    
    # 基本キー
    base_keys = []
    
    # A-Z
    base_keys.extend([chr(c) for c in range(ord('A'), ord('Z') + 1)])
    
    # 0-9
    base_keys.extend([chr(c) for c in range(ord('0'), ord('9') + 1)])
    
    # F1-F12
    base_keys.extend([f"F{i}" for i in range(1, 13)])
    
    # 特殊キー (recorder_core.py _get_input_state にあるもの)
    special_keys = [
        "Enter", "Space", "Esc", "BS", "Tab", "Del",
        "PageUp", "PageDown", "End", "Home", "Insert",
        "Left", "Up", "Right", "Down",
        "PrintScreen", "Pause", "CapsLock", "ScrollLock"
    ]
    base_keys.extend(special_keys)
    
    # マウスアクション (これらは単独 または 修飾キーとの組み合わせ)
    mouse_actions = ["L-Click", "R-Click", "M-Click"]
    
    # 2. 組み合わせ生成
    
    all_shortcuts = []
    
    # すべての修飾キーの部分集合を生成 (空集合含む)
    # 0個〜4個
    mod_combinations = []
    for r in range(len(modifiers_base) + 1):
        for combo in itertools.combinations(modifiers_base, r):
            # 組み合わせ内の順序も固定する (Ctrl, Shift, Alt, Win)
            # itertools.combinations は入力順序を保持するので modifiers_base の順序になる
            mod_combinations.append(list(combo))
            
    # キーボードキーとの組み合わせ
    for key in base_keys:
        for mods in mod_combinations:
            # 組み合わせ文字列生成
            # 例: Ctrl+Shift+A, A (mods=[]の場合)
            parts = mods + [key]
            combo_str = "+".join(parts)
            
            # 説明文生成 (簡易)
            # 既知のものは上書きしてもいいが、ここでは自動生成
            # desc = combo_str (Descriptionは動的に解決するので不要だが、デフォルト値として持っておく)
            default_desc = combo_str
            
            # デフォルト有効状態:
            # - アルファベット、数字のキー単体(修飾キーなし)は無効
            # - それ以外は有効
            default_enabled = True
            if not mods:
                # A-Z
                if len(key) == 1 and 'A' <= key <= 'Z':
                    default_enabled = False
                # 0-9
                elif len(key) == 1 and '0' <= key <= '9':
                    default_enabled = False
            
            all_shortcuts.append({
                'key': combo_str,
                'enabled': default_enabled,
                'default_desc': default_desc
            })
            
    # マウスとの組み合わせ
    for mouse in mouse_actions:
        for mods in mod_combinations:
            if not mods:
                # マウス単体 (修飾キーなし)
                parts = [mouse]
            else:
                parts = mods + [mouse]
            
            combo_str = "+".join(parts)
            
            all_shortcuts.append({
                'key': combo_str,
                'enabled': True, # マウス操作は基本記録
                # 'desc': combo_str # Descriptionは動的に結合するため、ここでは空にしておくか、初期値としてKeyを入れるか。
                # 基本的にアプリ定義がない場合はKeyを表示したい。
                # ここではdefault_descとして入れておく
                'default_desc': combo_str
            })

    # 修飾キー単体 (Ctrl, Shift, Alt, Win)
    # 修飾キーのみの入力も記録できるようにするが、既定では無効
    for m in modifiers_base:
        all_shortcuts.append({
            'key': m,
            'enabled': False,
            'default_desc': m
        })
            
    return all_shortcuts


class ShortcutManager:
    """ショートカットキーの有効/無効を管理するクラス."""

    def __init__(self, filepath: str):
        self.filepath = filepath
        # Key: key_combo, Value: {'record_enabled': bool, 'play_enabled': bool}
        # Descriptionはここには持たず、UI表示時に動的に解決するが、
        # 便宜上 get_all_shortcuts で結合して渡す。
        self.shortcuts: Dict[str, Dict] = {} 
        
        # 内部で全リストを持っておく
        self.all_possible_shortcuts = generate_all_shortcuts()
        
        # アプリ定義キャッシュ
        self.app_definitions_cache: Dict[str, Dict[str, str]] = {} # filename -> {key: desc}
        
        self.load()

    def get_shortcuts_dir(self) -> str:
        """shortcutsフォルダのパスを返す"""
        return os.path.join(os.path.dirname(self.filepath), "shortcuts")

    def get_app_definitions_list(self) -> List[str]:
        """利用可能なアプリ定義ファイル名のリストを返す"""
        d = self.get_shortcuts_dir()
        if not os.path.exists(d):
            return []
        files = [f for f in os.listdir(d) if f.endswith(".tsv")]
        return files

    def load_app_definitions(self, filenames: List[str]) -> Dict[str, str]:
        """指定された複数の定義ファイルから、キーごとの説明（結合・重複排除済み）を返す"""
        descriptions: Dict[str, List[str]] = {}
        
        base_dir = self.get_shortcuts_dir()
        
        for fname in filenames:
            path = os.path.join(base_dir, fname)
            if not os.path.exists(path):
                continue
                
            try:
                with open(path, 'r', encoding='utf-8', newline='') as f:
                    # ヘッダーなし、または Key\tDesc 形式を想定
                    # タブ区切り
                    reader = csv.reader(f, delimiter='\t')
                    for row in reader:
                        if len(row) >= 2:
                            key = row[0].strip()
                            desc = row[1].strip()
                            if key and desc:
                                if key not in descriptions:
                                    descriptions[key] = []
                                # 重複排除して追加
                                if desc not in descriptions[key]:
                                    descriptions[key].append(desc)
            except Exception as e:
                print(f"Error loading app definition {fname}: {e}")
        
        # リストを結合文字列に変換
        result = {}
        for key, desc_list in descriptions.items():
            result[key] = "、".join(desc_list)
            
        return result

    def load(self):
        """設定ファイル(TSV)から読み込む。ファイルがない場合はデフォルトを作成。"""
        self.shortcuts = {}
        
        # 全リストで初期化
        for item in self.all_possible_shortcuts:
            default_enabled = item['enabled']
            self.shortcuts[item['key']] = {
                'record_enabled': default_enabled,
                'play_enabled': default_enabled,
                'default_desc': item['default_desc']
            }

        if not os.path.exists(self.filepath):
            self.save() # デフォルトファイルを作成
            return

        try:
            with open(self.filepath, 'r', encoding='utf-8', newline='') as f:
                reader = csv.reader(f, delimiter='\t')
                header = next(reader, None) # ヘッダー読み飛ばし
                
                # ヘッダーチェックで旧形式かどうか判定
                # 新形式: Key, RecordEnabled, PlayEnabled
                # 旧形式: Key, Enabled, Description
                
                is_legacy = False
                if header:
                    if len(header) >= 3 and header[1] == "Enabled":
                        is_legacy = True
                
                if header:
                    for row in reader:
                        if not row: continue
                        key = row[0]
                        
                        if key in self.shortcuts:
                            if is_legacy:
                                # 旧形式: Enabledの値をRec/Play両方に適用
                                enabled = (row[1] == '1')
                                self.shortcuts[key]['record_enabled'] = enabled
                                self.shortcuts[key]['play_enabled'] = enabled
                            else:
                                # 新形式: Rec, Play
                                if len(row) >= 3:
                                    self.shortcuts[key]['record_enabled'] = (row[1] == '1')
                                    self.shortcuts[key]['play_enabled'] = (row[2] == '1')
        except Exception as e:
            print(f"Shortcut config load error: {e}")
            # エラー時はデフォルトのまま

    def save(self):
        """現在の設定をTSVファイルに保存する。"""
        try:
            with open(self.filepath, 'w', encoding='utf-8', newline='') as f:
                writer = csv.writer(f, delimiter='\t')
                writer.writerow(['Key', 'RecordEnabled', 'PlayEnabled'])
                
                save_order = [item['key'] for item in self.all_possible_shortcuts]
                
                for key in save_order:
                    if key in self.shortcuts:
                        info = self.shortcuts[key]
                        writer.writerow([
                            key,
                            '1' if info['record_enabled'] else '0',
                            '1' if info['play_enabled'] else '0'
                        ])
        except Exception as e:
            print(f"Shortcut config save error: {e}")

    def is_allowed_record(self, combo_string: str) -> bool:
        """録画時: 指定されたキー組み合わせが有効かどうかを判定する。"""
        if combo_string == "None" or not combo_string:
            return True 
        if combo_string in self.shortcuts:
            return self.shortcuts[combo_string]['record_enabled']
        return False

    def is_allowed_playback(self, combo_string: str) -> bool:
        """再生時: 指定されたキー組み合わせが有効かどうかを判定する。"""
        if combo_string == "None" or not combo_string:
            return True
        if combo_string in self.shortcuts:
            return self.shortcuts[combo_string]['play_enabled']
        return False
        
    # 旧メソッド互換用（削除してもいいが念のため）
    def is_allowed(self, combo_string: str) -> bool:
        return self.is_allowed_record(combo_string)

    def get_all_shortcuts(self, app_descriptions: Dict[str, str] = None) -> List[Dict]:
        """UI表示用に全てのショートカット情報をリストで返す。
        app_descriptions: アプリ定義から読み込んだ {Key: Desc} マップ。
                          指定された場合、default_descの代わりにこれを使う。
        """
        result = []
        order = [item['key'] for item in self.all_possible_shortcuts]
        for key in order:
            if key in self.shortcuts:
                info = self.shortcuts[key]
                
                desc = info['default_desc']
                # アプリ定義があれば上書き
                if app_descriptions and key in app_descriptions:
                    desc = app_descriptions[key]
                elif app_descriptions is not None:
                    # アプリ定義が指定されている（=フィルタ中）が、定義がない場合
                    # 「説明がないものは除外」フィルタのために、空文字にしておく手もあるが、
                    # UI側で default_desc を使うか判定できるように、ここではそのまま返すか、
                    # あるいはUI側ロジックに任せる。
                    # Task: "概要が空欄のものをリストから除外する"
                    # -> ここでは定義がなければ空文字を返すのが自然か？
                    # しかし通常のキー名（"Ctrl+C"など）は最低限表示したい場合もある。
                    # いや、ユーザー意図は「アプリの機能名」を知りたい。
                    # 定義にないキーは「ただのキー」なので、"概要"は空であるべき。
                    desc = ""

                result.append({
                    'key': key,
                    'record_enabled': info['record_enabled'],
                    'play_enabled': info['play_enabled'],
                    'desc': desc
                })
        return result

    def update_shortcut(self, key: str, record_enabled: bool, play_enabled: bool):
        """ショートカットの有効状態を更新する。"""
        if key in self.shortcuts:
            self.shortcuts[key]['record_enabled'] = record_enabled
            self.shortcuts[key]['play_enabled'] = play_enabled
            
    def set_all(self, record_enabled: Optional[bool] = None, play_enabled: Optional[bool] = None):
        """全てのショートカットを一括設定。Noneの場合は変更しない。"""
        for key in self.shortcuts:
            if record_enabled is not None:
                self.shortcuts[key]['record_enabled'] = record_enabled
            if play_enabled is not None:
                self.shortcuts[key]['play_enabled'] = play_enabled
