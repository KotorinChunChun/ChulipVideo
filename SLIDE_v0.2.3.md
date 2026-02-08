# ChulipVideo v0.2.3 Update Highlights

## 🛠️ 新機能の仕組み (The Mechanism)

### 1. 🖱️ 入力オーバーレイ (Input Overlay)
- **記録 (Recording)**: 
  - 録画中に `keyboard` / `mouse` イベントをフックし、タイムスタンプ付きで `TSVファイル` に保存。
  - 映像そのものには焼き付けないため、後から編集が可能！
- **再生・出力 (Playback/Export)**:
  - 映像フレームに合わせて TSV を読み込み、`PIL (Pillow)` で動的に描画。
  - **スタッキング処理**: 短時間の連続入力は自動的に積み上げて表示し、視認性を確保。

### 2. 🪟 ウィンドウキャプチャ (WGC Impl)
- **技術**: `Windows Graphics Capture` API を採用。
- **メリット**:
  - 対象ウィンドウの **GPUテクスチャ** を直接取得。
  - 手前に他のウィンドウが被っても、**対象ウィンドウだけ** を綺麗に録画可能。

### 3. 🎨 テーマ＆設定 (Global Config)
- **統一化**: `ChulipVideo_config.json` に設定を集約。
- **反映**: 詳細設定ダイアログでの変更は即座にメモリ上の `theme` 辞書に反映され、`Canvas` 再描画トリガーによってリアルタイムにプレビュー更新。

---
*Created by ChulipVideo Dev Team*
