"""ショートカットキー設定ダイアログ."""
import tkinter as tk
from tkinter import ttk, messagebox
from typing import Optional, TYPE_CHECKING
import re

if TYPE_CHECKING:
    from shortcut_manager import ShortcutManager


class ShortcutSettingsDialog(tk.Toplevel):
    """ショートカットキーの有効/無効を設定するダイアログ."""

    def __init__(self, parent: tk.Widget, shortcut_manager: 'ShortcutManager'):
        super().__init__(parent)
        self.title("マウス・ショートカットキー操作の記録・再生の切り替え")
        self.geometry("900x600")
        
        # モーダル設定
        self.transient(parent)
        self.grab_set()
        
        self.shortcut_manager = shortcut_manager
        
        # アプリ定義のロード
        self.app_files = self.shortcut_manager.get_app_definitions_list()
        
        # 内部データ (初期表示用) - フィルタ適用時に再取得する
        self.current_app_descriptions = {}
        
        # UI構築
        self._build_ui()
        
        # 初期表示
        self._on_app_selection_change(None)
        
        # 中央配置
        self.update_idletasks()
        w = self.winfo_width()
        h = self.winfo_height()
        x = parent.winfo_rootx() + (parent.winfo_width() // 2) - (w // 2)
        y = parent.winfo_rooty() + (parent.winfo_height() // 2) - (h // 2)
        self.geometry(f"+{x}+{y}")
        
        # 呼び出し元をブロック
        self.wait_window(self)

    def _build_ui(self):
        # 左右ペイン
        paned = tk.PanedWindow(self, orient=tk.HORIZONTAL, sashwidth=4)
        paned.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # --- 左ペイン: アプリ一覧 ---
        left_frame = tk.Frame(paned)
        paned.add(left_frame, minsize=150, width=200)
        
        # グループ化: ショートカット定義アプリ
        lf_apps = tk.LabelFrame(left_frame, text="アプリのショートカット定義")
        lf_apps.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
        
        # mk: 見出しはLabelFrameになったのでLabel削除
        # tk.Label(left_frame, text="説明を表示するアプリ", font=("", 10, "bold")).pack(anchor="w", padx=2, pady=2)
        
        # アプリ選択操作ボタン
        app_btn_frame = tk.Frame(lf_apps)
        app_btn_frame.pack(fill=tk.X, padx=2, pady=2)
        tk.Button(app_btn_frame, text="全選択", command=self._select_all_apps, font=("", 8)).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=1)
        tk.Button(app_btn_frame, text="全解除", command=self._deselect_all_apps, font=("", 8)).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=1)

        # 複数選択可能なリストボックス
        self.app_listbox = tk.Listbox(lf_apps, selectmode=tk.EXTENDED, exportselection=False)
        self.app_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scroll_app = tk.Scrollbar(lf_apps, orient=tk.VERTICAL, command=self.app_listbox.yview)
        scroll_app.pack(side=tk.RIGHT, fill=tk.Y)
        self.app_listbox.config(yscrollcommand=scroll_app.set)
        
        # ファイル一覧投入 & 全選択
        for f in self.app_files:
            self.app_listbox.insert(tk.END, f)
        
        self.app_listbox.select_set(0, tk.END) # 初期状態で全選択
            
        self.app_listbox.bind("<<ListboxSelect>>", self._on_app_selection_change)
        
        # --- 右ペイン: ショートカット設定 ---
        right_frame = tk.Frame(paned)
        paned.add(right_frame, minsize=500)
        
        # カラム幅定義 (ピクセル)
        W_KEY = 200
        W_REC = 80
        W_PLAY = 80
        W_DESC = 300
        
        # = フィルタエリア =
        filter_frame = tk.LabelFrame(right_frame, text="フィルタ")
        filter_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # 1行目: テキスト検索
        f_row1 = tk.Frame(filter_frame)
        f_row1.pack(fill=tk.X, padx=5, pady=2)
        
        tk.Label(f_row1, text="検索:").pack(side=tk.LEFT)
        self.var_filter_text = tk.StringVar()
        self.var_filter_text.trace_add("write", lambda *args: self._refresh_list())
        tk.Entry(f_row1, textvariable=self.var_filter_text).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        # 2行目: 修飾キー
        f_row2 = tk.Frame(filter_frame)
        f_row2.pack(fill=tk.X, padx=5, pady=2)
        
        self.var_filter_ctrl = tk.BooleanVar(value=False)
        self.var_filter_shift = tk.BooleanVar(value=False)
        self.var_filter_alt = tk.BooleanVar(value=False)
        self.var_filter_win = tk.BooleanVar(value=False)
        self.var_filter_none = tk.BooleanVar(value=False)

        for text, var in [("Ctrl", self.var_filter_ctrl), ("Shift", self.var_filter_shift), 
                          ("Alt", self.var_filter_alt), ("Win", self.var_filter_win), ("なし", self.var_filter_none)]:
            tk.Checkbutton(f_row2, text=text, variable=var, command=self._refresh_list).pack(side=tk.LEFT)
            
        # 3行目: Rec/Play状態フィルタ (カラム位置に合わせる)
        f_row3 = tk.Frame(filter_frame)
        f_row3.pack(fill=tk.X, padx=5, pady=2)
        
        # Keyカラム相当 (説明ありのみチェックボックスをここに配置 - 左端)
        f_r3_key = tk.Frame(f_row3, width=W_KEY, height=25)
        f_r3_key.pack_propagate(False) # 固定幅
        f_r3_key.pack(side=tk.LEFT)
        
        self.var_filter_has_desc = tk.BooleanVar(value=False)
        # 文言変更: 説明ありのみ表示
        tk.Checkbutton(f_r3_key, text="説明がある操作のみ表示", variable=self.var_filter_has_desc, command=self._refresh_list).pack(side=tk.LEFT, anchor="w")
        
        # Recカラム相当
        f_r3_rec = tk.Frame(f_row3, width=W_REC, height=25)
        f_r3_rec.pack_propagate(False)
        f_r3_rec.pack(side=tk.LEFT)
        self.combo_filter_rec = ttk.Combobox(f_r3_rec, values=["全て", "有効", "無効"], state="readonly")
        self.combo_filter_rec.current(0)
        self.combo_filter_rec.pack(fill=tk.BOTH, expand=True, padx=5) # 左右に余白少々
        self.combo_filter_rec.bind("<<ComboboxSelected>>", lambda e: self._refresh_list())
        
        # Playカラム相当
        f_r3_play = tk.Frame(f_row3, width=W_PLAY, height=25)
        f_r3_play.pack_propagate(False)
        f_r3_play.pack(side=tk.LEFT)
        self.combo_filter_play = ttk.Combobox(f_r3_play, values=["全て", "有効", "無効"], state="readonly")
        self.combo_filter_play.current(0)
        self.combo_filter_play.pack(fill=tk.BOTH, expand=True, padx=5)
        self.combo_filter_play.bind("<<ComboboxSelected>>", lambda e: self._refresh_list())

        # = グループ化: マウス・キー操作一覧 =
        lf_list = tk.LabelFrame(right_frame, text="マウス・キー操作一覧（記録と再生のON/OFF）")
        lf_list.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # = リストエリア =
        list_frame = tk.Frame(lf_list)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=0) 
        
        cols = ("Key", "Rec", "Play", "Desc")
        self.tree = ttk.Treeview(list_frame, columns=cols, show="headings", selectmode="extended")
        
        self.tree.heading("Key", text="キー", command=lambda: self._sort_column("Key", False))
        self.tree.heading("Rec", text="記録", command=lambda: self._sort_column("Rec", False))
        self.tree.heading("Play", text="再生", command=lambda: self._sort_column("Play", False))
        self.tree.heading("Desc", text="説明", command=lambda: self._sort_column("Desc", False))
        
        self.tree.column("Key", width=W_KEY, anchor="w", stretch=False) 
        self.tree.column("Rec", width=W_REC, anchor="center", stretch=False)
        self.tree.column("Play", width=W_PLAY, anchor="center", stretch=False)
        self.tree.column("Desc", width=W_DESC, anchor="w", stretch=True) 
        
        scroll = tk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.tree.bind("<Double-1>", self._on_double_click)
        self.tree.bind("<space>", self._on_space_key)
        self.tree.bind("<Control-a>", self._select_all_shortcuts)
        self.tree.bind("<ButtonPress-1>", self._on_mouse_down)
        self.tree.bind("<B1-Motion>", self._on_mouse_drag)

        # = アクションボタンエリア (リストと同じグループ内) =
        btn_frame = tk.Frame(lf_list)
        btn_frame.pack(fill=tk.X, padx=5, pady=5)
        
        BTN_AREA_H = 70
        
        # Keyカラム相当 (スペーサー)
        f_btn_key = tk.Frame(btn_frame, width=W_KEY, height=BTN_AREA_H)
        f_btn_key.pack_propagate(False)
        f_btn_key.pack(side=tk.LEFT)
        
        # Recカラム相当
        f_btn_rec = tk.Frame(btn_frame, width=W_REC, height=BTN_AREA_H)
        f_btn_rec.pack_propagate(False) # 80px固定
        f_btn_rec.pack(side=tk.LEFT) 
        # widthがリストと揃わない問題を解消するため、fill=tk.X しつつ padxを入れる
        # 録画無効の色: #e0e0ff (rec_disabled)
        tk.Button(f_btn_rec, text="有効", command=lambda: self._set_selection("rec", True), bg="#ccffcc").pack(fill=tk.X, padx=5, pady=1)
        tk.Button(f_btn_rec, text="無効", command=lambda: self._set_selection("rec", False), bg="#e0e0ff").pack(fill=tk.X, padx=5, pady=1)
        
        # Playカラム相当
        f_btn_play = tk.Frame(btn_frame, width=W_PLAY, height=BTN_AREA_H)
        f_btn_play.pack_propagate(False)
        f_btn_play.pack(side=tk.LEFT)
        # 再生無効の色: #ffe0e0 (play_disabled)
        tk.Button(f_btn_play, text="有効", command=lambda: self._set_selection("play", True), bg="#ccffcc").pack(fill=tk.X, padx=5, pady=1)
        tk.Button(f_btn_play, text="無効", command=lambda: self._set_selection("play", False), bg="#ffe0e0").pack(fill=tk.X, padx=5, pady=1)
        
        # 右側: 全て有効/無効
        # 垂直位置を揃えるため anchor="n" を指定
        f_btn_all = tk.Frame(btn_frame)
        f_btn_all.pack(side=tk.RIGHT, padx=10, anchor="n")
        
        # 全無効の色: #f0f0f0 (disabled)
        tk.Button(f_btn_all, text="表示中のキーをすべて有効", command=lambda: self._set_all_visible(True), width=24, bg="#ccffcc").pack(fill=tk.X, pady=1)
        tk.Button(f_btn_all, text="表示中のキーをすべて無効", command=lambda: self._set_all_visible(False), width=24, bg="#cccccc").pack(fill=tk.X, pady=1)

        # = フッター =
        footer_frame = tk.Frame(self)
        footer_frame.pack(fill=tk.X, padx=10, pady=10)
        
        tk.Button(footer_frame, text="キャンセル", command=self.destroy, width=10).pack(side=tk.RIGHT, padx=5)
        tk.Button(footer_frame, text="保存", command=self._save, width=10, bg="#ccffcc").pack(side=tk.RIGHT, padx=5)

    def _select_all_apps(self):
        self.app_listbox.select_set(0, tk.END)
        self._on_app_selection_change(None)

    def _deselect_all_apps(self):
        self.app_listbox.selection_clear(0, tk.END)
        self._on_app_selection_change(None)

    def _on_app_selection_change(self, event):
        # 選択されたファイル名を取得
        indices = self.app_listbox.curselection()
        filenames = [self.app_listbox.get(i) for i in indices]
        
        if filenames:
            self.current_app_descriptions = self.shortcut_manager.load_app_definitions(filenames)
        else:
            self.current_app_descriptions = None
            
        self._refresh_list()

    def _refresh_list(self):
        # 現在の選択状態保存
        selected_keys = []
        for item_id in self.tree.selection():
            vals = self.tree.item(item_id, "values")
            if vals: selected_keys.append(vals[0]) # Key

        for item in self.tree.get_children():
            self.tree.delete(item)
            
        # フィルタ値取得
        f_text = self.var_filter_text.get().lower()
        f_has_desc = self.var_filter_has_desc.get()
        
        f_rec = self.combo_filter_rec.get()
        f_play = self.combo_filter_play.get()
        
        # Modifiers
        use_ctrl = self.var_filter_ctrl.get()
        use_shift = self.var_filter_shift.get()
        use_alt = self.var_filter_alt.get()
        use_win = self.var_filter_win.get()
        use_none = self.var_filter_none.get()
        is_mod_filter = use_ctrl or use_shift or use_alt or use_win or use_none
        
        # マネージャーから全データ取得 (現在のDescriptionマップを適用)
        all_data = self.shortcut_manager.get_all_shortcuts(self.current_app_descriptions)
        
        # 内部で保持している一時的な変更があれば反映する必要があるが、
        # 今回はManagerのデータを直接書き換えるのではなく、保存時にManagerに書き込む?
        # いや、UI操作で enabled を変えたら、それは Manager のメモリ上のデータを変えてしまって良い。
        # 保存ボタンでファイルに書き込むだけ。
        # キャンセル時はファイル書き込まず終了するが、Managerのメモリは汚れる。
        # -> Modalなので、アプリ再起動するか、再度ロードしない限りメモリは汚れたままになる可能性がある。
        # -> 本来は deepcopy して作業し、保存時に反映すべき。
        # -> しかし簡易実装として、キャンセル時は `self.shortcut_manager.load()` を呼んでリセットさせると安全。
        pass # 下記destroyで実装
        
        for item in all_data:
            key = item['key']
            rec = item['record_enabled']
            play = item['play_enabled']
            desc = item['desc']
            
            # --- フィルタリング ---
            
            # 1. 説明ありのみ
            if f_has_desc and not desc:
                continue
                
            # 2. テキスト検索
            if f_text:
                if f_text not in key.lower() and f_text not in desc.lower():
                    continue
                    
            # 3. 修飾キー
            if is_mod_filter:
                matched = False
                has_mod = any(m in key for m in ["Ctrl", "Shift", "Alt", "Win"])
                
                if use_ctrl and "Ctrl" in key: matched = True
                if use_shift and "Shift" in key: matched = True
                if use_alt and "Alt" in key: matched = True
                if use_win and "Win" in key: matched = True
                if use_none and not has_mod: matched = True
                
                if not matched:
                    continue
            
            # 4. Rec状態
            if f_rec == "有効" and not rec: continue
            if f_rec == "無効" and rec: continue
            
            # 5. Play状態
            if f_play == "有効" and not play: continue
            if f_play == "無効" and play: continue
            
            # 表示
            rec_sym = "✔" if rec else "✘"
            play_sym = "✔" if play else "✘"
            
            # タグ (Rec/Play両方有効なら enabled, どちらか無効なら... 色分けが難しいので単純化)
            # 全体が無効(RecもPlayも)ならグレーにするか？
            tag = "enabled"
            if not rec and not play:
                tag = "disabled"
            elif not rec:
                tag = "rec_disabled"
            elif not play:
                tag = "play_disabled"
                
            item_id = self.tree.insert("", "end", values=(key, rec_sym, play_sym, desc), tags=(tag,))
            
            if key in selected_keys:
                self.tree.selection_add(item_id)
                
        # タグ設定
        self.tree.tag_configure("enabled", foreground="black", background="white")
        self.tree.tag_configure("disabled", foreground="gray", background="#cccccc")
        self.tree.tag_configure("rec_disabled", foreground="blue", background="#e0e0ff") # Recだけ無効（Play有効）はほぼないはずだが
        self.tree.tag_configure("play_disabled", foreground="red", background="#ffe0e0") # Rec有効Play無効

    # --- Event Handlers ---

    def _select_all_shortcuts(self, event):
        # Treeviewの全項目を選択
        children = self.tree.get_children()
        self.tree.selection_set(children)
        return "break" # 標準イベントを止める

    def _on_mouse_down(self, event):
        # ドラッグ選択の開始
        # 標準のクリック動作を妨げないようにするが、ドラッグ開始位置を記録
        # Treeviewの標準バインディングと競合する可能性があるので注意
        self._drag_start_item = self.tree.identify_row(event.y)

    def _on_mouse_drag(self, event):
        # ドラッグ中に選択範囲を更新
        # 簡易実装: ドラッグした先のアイテムを選択に追加していく(paint select)
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_add(item) 

    def _on_space_key(self, event):
        # 選択中のアイテムに対し、共通のトグル操作を行う
        # ロジック: ダブルクリックの「その他カラム」と同じ（Rec基準で両方トグル）
        selected_items = self.tree.selection()
        if not selected_items: return
        
        # 最初の1つのRec状態を見て、反転させるか決める（統一操作のため）
        # あるいは個別に反転させるか？ -> エクセルなどは一括設定が多い。
        # ここでは「最初のアイテムのRec」の逆を、全選択アイテムに適用する。
        
        first_id = selected_items[0]
        vals = self.tree.item(first_id, "values")
        first_key = vals[0]
        is_rec_first = self.shortcut_manager.is_allowed_record(first_key)
        
        target_enabled = not is_rec_first
        
        for item_id in selected_items:
            vals = self.tree.item(item_id, "values")
            key = vals[0]
            # Rec/Play両方を target_enabled にする
            self.shortcut_manager.update_shortcut(key, target_enabled, target_enabled)
            
        self._refresh_list()

    def _on_double_click(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell": return
        
        col_id = self.tree.identify_column(event.x)
        # col_id is like "#1", "#2"...
        # Columns: #1=Key, #2=Rec, #3=Play, #4=Desc
        
        item_id = self.tree.identify_row(event.y)
        if not item_id: return
        
        vals = self.tree.item(item_id, "values")
        key = vals[0]
        
        # 現在の状態を取得
        is_rec = self.shortcut_manager.is_allowed_record(key)
        is_play = self.shortcut_manager.is_allowed_playback(key)
        
        if col_id == "#2": # Rec列
            self.shortcut_manager.update_shortcut(key, not is_rec, is_play)
        elif col_id == "#3": # Play列
            self.shortcut_manager.update_shortcut(key, is_rec, not is_play)
        else:
            # キーや説明をクリック -> Recを基準に両方をトグル
            # Recが有効なら両方無効に。Recが無効なら両方有効に。
            target = not is_rec
            self.shortcut_manager.update_shortcut(key, target, target)
            
        # リフレッシュ (selection維持)
        self._refresh_list()

    def _set_selection(self, target_type: str, enabled: bool):
        """target_type: 'rec' or 'play'"""
        selected_items = self.tree.selection()
        if not selected_items: return
        
        for item_id in selected_items:
            vals = self.tree.item(item_id, "values")
            key = vals[0]
            
            is_rec = self.shortcut_manager.is_allowed_record(key)
            is_play = self.shortcut_manager.is_allowed_playback(key)
            
            if target_type == "rec":
                self.shortcut_manager.update_shortcut(key, enabled, is_play)
            elif target_type == "play":
                self.shortcut_manager.update_shortcut(key, is_rec, enabled)
                
        self._refresh_list()

    def _set_all_visible(self, enabled: bool):
        """表示中の全項目のRec/Playをまとめて設定"""
        for item_id in self.tree.get_children():
            vals = self.tree.item(item_id, "values")
            key = vals[0]
            self.shortcut_manager.update_shortcut(key, enabled, enabled)
        self._refresh_list()

    def _save(self):
        self.shortcut_manager.save()
        self.destroy()

    def destroy(self):
        # キャンセル（または×ボタン）の場合、保存せずに閉じる
        # メモリ上の変更を破棄するためにリロードする
        # (Saveボタン経由以外でここに来た場合(=保存されていない)のみリロードしたいが、
        #  _save内で save() -> destroy() と呼ぶと同じこと。
        #  _save で flag を立てるか、リロードはこのクラスの外で行うか。
        #  一番簡単なのは、変更をメモリに即反映せず、保存時に反映することだが、
        #  リスト更新のためにメモリ反映が必要だった。
        #  よって、保存されなかった場合はリロードして元に戻すのが安全。)
        if self.shortcut_manager:
             # saveメソッドが呼ばれるとファイル更新される。
             # ここでは簡易判定が難しいので、常にリロードしてしまうと、
             # Save -> File Saved -> Reload (from File) -> OK.
             # Cancel -> No Save -> Reload (from File) -> Revert Change. -> OK.
             # なので常にリロードして整合性を取るのが良い。
             self.shortcut_manager.load()
             
        super().destroy()

    def _sort_column(self, col, reverse):
        l = [(self.tree.set(k, col), k) for k in self.tree.get_children('')]
        l.sort(reverse=reverse)
        for index, (val, k) in enumerate(l):
            self.tree.move(k, '', index)
        self.tree.heading(col, command=lambda: self._sort_column(col, not reverse))
