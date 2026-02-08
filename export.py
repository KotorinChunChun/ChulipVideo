"""エクスポート機能モジュール.

PNG連番出力と動画出力のロジックを提供。
"""
from __future__ import annotations

import os
import subprocess
import sys
import time
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import filedialog, messagebox
from typing import TYPE_CHECKING
import csv

import cv2
import numpy as np
from PIL import Image

import overlay_utils
from utils import imwrite_jp, sec_to_hhmmss, open_folder_with_selection

if TYPE_CHECKING:
    pass


def open_folder(path: str) -> None:
    """プラットフォーム依存でフォルダを開く."""
    try:
        if os.name == 'nt':
            os.startfile(path)
        elif sys.platform == 'darwin':
            subprocess.Popen(['open', path])
        else:
            # Linux 等
            subprocess.Popen(['xdg-open', path])
    except Exception as e:
        messagebox.showerror("Error", f"フォルダを開けませんでした:\n{e}")


def open_file(path: str) -> None:
    """プラットフォーム依存でファイルを開く."""
    try:
        if os.name == 'nt':
            os.startfile(path)
        elif sys.platform == 'darwin':
            subprocess.Popen(['open', path])
        else:
            subprocess.Popen(['xdg-open', path])
    except Exception as e:
        messagebox.showerror('Error', f'ファイルを開けませんでした:\n{e}')


class ExportMixin:
    """エクスポート機能のメソッドを提供するMixinクラス.
    
    VideoCropperAppクラスにミックスインして使用する。
    """
    
    # 以下の属性は VideoCropperApp から継承される想定
    # cap: cv2.VideoCapture
    # video_filepath: str
    # video_filename: str
    # crop_rect: list[int]
    # start_time, end_time: float
    # fps: float
    # playing: bool
    # _play_after_id: str | None
    # root: tk.Tk
    # CANVAS_W, CANVAS_H: int
    # png_compression: int
    # check_prev_next: tk.BooleanVar
    # check_duplicate: tk.BooleanVar
    
    def open_folder(self, path: str) -> None:
        """プラットフォーム依存でフォルダを開く."""
        open_folder(path)

    def open_video_settings(self) -> None:
        """当該動画に紐づく設定ファイルを既定のプログラムで開く."""
        if not self.video_filepath:
            messagebox.showinfo("Info", "設定ファイルを開く対象の動画が選択されていません")
            return
        settings_path = os.path.splitext(self.video_filepath)[0] + '.settings.json'
        if not os.path.exists(settings_path):
            messagebox.showinfo("Info", f"設定ファイルが見つかりません:\n{settings_path}")
            return
        open_file(settings_path)

    def export_png(self) -> None:
        """クロップ範囲をPNG連番として出力する."""
        if not self.cap:
            messagebox.showerror("Error", "動画なし")
            return
        video_name = os.path.basename(self.video_filepath) if self.video_filepath else None
        video_dir = os.path.dirname(self.video_filepath) if self.video_filepath else None
        base_dir = filedialog.askdirectory(initialdir=video_dir)
        if not base_dir:
            return

        # 実行日時のフォルダを作成
        now = time.strftime("%Y%m%d_%H%M%S")
        save_dir = os.path.join(base_dir, f"{video_name}_crops_{now}")
        os.makedirs(save_dir, exist_ok=True)

        # PNG圧縮レベルを設定
        save_params = [int(cv2.IMWRITE_PNG_COMPRESSION), self.png_compression]

        # 座標変換
        x1, y1, x2, y2 = self.crop_rect
        vid_w = int(self.cap.get(cv2.CAP_PROP_FRAME_WIDTH))
        vid_h = int(self.cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
        vx1, vy1, vx2, vy2 = [int(v) for v in self.crop_rect]
        
        # 枠外選択が許可されていない場合は安全のためクランプ
        if not (getattr(self, 'allow_oversize_var', None) and self.allow_oversize_var.get()):
            vx1 = max(0, min(vid_w, vx1))
            vy1 = max(0, min(vid_h, vy1))
            vx2 = max(0, min(vid_w, vx2))
            vy2 = max(0, min(vid_h, vy2))

        was_playing = self.playing
        self.playing = False
        if self._play_after_id:
            self.root.after_cancel(self._play_after_id)

        # Create modal progress dialog
        try:
            progress_win = tk.Toplevel(self.root)
            progress_win.title("PNG 書き出し...")
            progress_win.transient(self.root)
            progress_win.attributes("-topmost", True)
            progress_win.grab_set()
            progress_win.resizable(False, False)
            
            # ウィンドウを一時的に隠してサイズ計算
            progress_win.withdraw()
            
            tk.Label(progress_win, text="PNG を出力中...").pack(padx=20, pady=(15, 5))
            pb = ttk.Progressbar(progress_win, orient=tk.HORIZONTAL, length=400, mode='determinate')
            pb.pack(padx=20, pady=(0, 10))
            prog_label = tk.Label(progress_win, text="0 / 0")
            prog_label.pack(padx=20, pady=(0, 15))
            
            # メインウィンドウの中央に配置
            progress_win.update_idletasks()
            win_w = progress_win.winfo_width()
            win_h = progress_win.winfo_height()
            root_x = self.root.winfo_x()
            root_y = self.root.winfo_y()
            root_w = self.root.winfo_width()
            root_h = self.root.winfo_height()
            pos_x = root_x + (root_w // 2) - (win_w // 2)
            pos_y = root_y + (root_h // 2) - (win_h // 2)
            progress_win.geometry(f"+{pos_x}+{pos_y}")
            progress_win.deiconify()
        except Exception:
            progress_win = None
            pb = None
            prog_label = None

        try:
            count = 0
            t = self.start_time
            limit = self.end_time
            frame_interval = 1.0 / self.fps
            
            # estimate total steps for progressbar
            try:
                total_steps = max(0, int((limit - t) / frame_interval) + 1)
            except Exception:
                total_steps = 0
            if pb is not None and total_steps > 0:
                pb['maximum'] = total_steps
                try:
                    prog_label.config(text=f"0 / {total_steps}")
                except Exception:
                    pass
            
            prev_crop = None
            next_crop = None
            last_saved_img = None

            # Overlay用
            history_manager = overlay_utils.InputHistoryManager()
            last_search_idx = 0
            
            from utils import get_safe_crop

            step_idx = 0
            while t <= limit:
                self.cap.set(cv2.CAP_PROP_POS_MSEC, t * 1000)
                ret, frm = self.cap.read()
                if ret and frm is not None:
                    # get_safe_crop を使用してクロップ (枠外は黒)
                    crop = get_safe_crop(frm, (vx1, vy1, vx2, vy2), (0, 0, 0))
                    
                    if crop.size > 0:
                        # 次のフレームを取得
                        next_t = t + frame_interval
                        if next_t <= limit:
                            self.cap.set(cv2.CAP_PROP_POS_MSEC, next_t * 1000)
                            ret_next, frm_next = self.cap.read()
                            if ret_next and frm_next is not None:
                                next_crop = get_safe_crop(frm_next, (vx1, vy1, vx2, vy2), (0, 0, 0))
                            else:
                                next_crop = None
                        else:
                            next_crop = None

                        # 前のフレーム、現在のフレーム、次のフレームが全て同じかチェック
                        is_matches_prev_next = False
                        if self.check_prev_next.get():
                            if prev_crop is not None and next_crop is not None:
                                diff1 = cv2.cvtColor(cv2.absdiff(prev_crop, crop), cv2.COLOR_BGR2GRAY)
                                diff2 = cv2.cvtColor(cv2.absdiff(crop, next_crop), cv2.COLOR_BGR2GRAY)
                                if cv2.countNonZero(diff1) == 0 and cv2.countNonZero(diff2) == 0:
                                    is_matches_prev_next = True
                            elif prev_crop is None and next_crop is not None:
                                diff2 = cv2.cvtColor(cv2.absdiff(crop, next_crop), cv2.COLOR_BGR2GRAY)
                                if cv2.countNonZero(diff2) == 0:
                                    is_matches_prev_next = True
                            elif prev_crop is not None and next_crop is None:
                                diff1 = cv2.cvtColor(cv2.absdiff(prev_crop, crop), cv2.COLOR_BGR2GRAY)
                                if cv2.countNonZero(diff1) == 0:
                                    is_matches_prev_next = True
                        else:
                            is_matches_prev_next = True

                        # オーバーレイ適用 (判定用に作成)
                        crop_overlay = crop.copy() # デフォルトはそのまま
                        if is_matches_prev_next:
                             rgb = cv2.cvtColor(crop, cv2.COLOR_BGR2RGB)
                             pil = Image.fromarray(rgb)
                             last_search_idx = self._draw_overlay_on_image(pil, t, history_manager, last_search_idx, vx1, vy1)
                             crop_overlay = cv2.cvtColor(np.array(pil), cv2.COLOR_RGB2BGR)

                        # 直前に出力したフレームとも比較 (比較はオーバーレイ無しの raw frame で行う)
                        is_same_as_last_saved = False
                        if self.check_duplicate.get() and last_saved_img is not None:
                            # 以前は last_saved_img (overlayあり) と crop_overlay を比較していたが、
                            # マウスが動くだけで不一致扱いになるのを防ぐため、生データ(crop)と比較する
                            # last_saved_img には生の crop を保存しておく必要がある
                            diff_last = cv2.cvtColor(cv2.absdiff(last_saved_img, crop), cv2.COLOR_BGR2GRAY)
                            if cv2.countNonZero(diff_last) == 0:
                                is_same_as_last_saved = True

                        # チェックボックスの設定に応じて出力判定
                        if is_matches_prev_next and not is_same_as_last_saved:
                            time_str = sec_to_hhmmss(t)
                            frame_in_sec = int((t - int(t)) * self.fps)
                            filepath = os.path.join(save_dir, f"{self.video_filename}_{time_str}_{frame_in_sec:03d}.png")
                            imwrite_jp(filepath, crop_overlay, params=save_params)
                            
                            # 重複判定用に生の crop を保存
                            last_saved_img = crop.copy()
                            count += 1

                        prev_crop = crop.copy()
                t += frame_interval
                
                # update progress
                step_idx += 1
                if pb is not None:
                    try:
                        pb['value'] = step_idx
                        prog_label.config(text=f"{step_idx} / {total_steps}")
                        progress_win.update()
                    except Exception:
                        pass

            # 設定を保存
            try:
                self.save_config()
            except Exception:
                pass

            # 完了ダイアログの前にウィンドウを消す
            try:
                if progress_win is not None:
                    progress_win.grab_release()
                    progress_win.destroy()
                    progress_win = None
            except Exception:
                pass

            # 完了ダイアログとフォルダを開くかの確認
            open_now = messagebox.askyesno("完了", f"{count} images saved.\nフォルダを開きますか？")
            if open_now:
                # フォルダを選択状態で開く
                open_folder_with_selection(save_dir)
        finally:
            # Destroy progress dialog
            try:
                if progress_win is not None:
                    progress_win.grab_release()
                    progress_win.destroy()
            except Exception:
                pass
            self.playing = was_playing
            if self.playing:
                self.play_step()

    def export_gif(self) -> None:
        """クロップ範囲をGIFとして出力する(ffmpeg使用)."""
        if not self.cap or not self.video_filepath:
            messagebox.showerror("Error", "動画なし")
            return

        # 保存先の決定
        now = time.strftime("%Y%m%d_%H%M%S")
        default_name = f"{self.video_filename}_{now}.gif"
        video_dir = os.path.dirname(self.video_filepath)
        save_path = filedialog.asksaveasfilename(
            defaultextension='.gif',
            initialfile=default_name,
            initialdir=video_dir,
            filetypes=[('GIF', '*.gif')],
            title='GIFの保存先を選択'
        )
        if not save_path:
            return

        # 座標変換
        x1, y1, x2, y2 = self.crop_rect
        vid_w = int(self.cap.get(cv2.CAP_PROP_FRAME_WIDTH))
        vid_h = int(self.cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
        vx1, vy1, vx2, vy2 = [int(v) for v in self.crop_rect]
        
        # 安全のためクランプ
        vx1 = max(0, min(vid_w, vx1))
        vy1 = max(0, min(vid_h, vy1))
        vx2 = max(0, min(vid_w, vx2))
        vy2 = max(0, min(vid_h, vy2))

        vw = vx2 - vx1
        vh = vy2 - vy1

        if vw <= 0 or vh <= 0:
            messagebox.showerror("Error", "有効な範囲を選択してください")
            return

        # プログレスバー表示用
        progress_win = tk.Toplevel(self.root)
        progress_win.title("GIF 書き出し中...")
        progress_win.transient(self.root)
        progress_win.grab_set()
        tk.Label(progress_win, text="ffmpeg で高品質 GIF を生成しています...").pack(padx=20, pady=20)
        pb = ttk.Progressbar(progress_win, mode='indeterminate', length=250)
        pb.pack(padx=20, pady=(0, 20))
        pb.start()

        def _run():
            try:
                # 高品質GIF生成のためのコマンド (パレット生成 -> パレット使用)
                # パレット生成パス
                palette_path = os.path.join(os.path.dirname(save_path), ".palette.png")

                # クロップ設定
                # 枠外選択がある場合、padフィルターを挟む必要がある
                allow_oversize = getattr(self, 'allow_oversize_var', None) and self.allow_oversize_var.get()
                
                # パディング計算
                pad_l = max(0, -vx1) if allow_oversize else 0
                pad_t = max(0, -vy1) if allow_oversize else 0
                pad_r = max(0, vx2 - vid_w) if allow_oversize else 0
                pad_b = max(0, vy2 - vid_h) if allow_oversize else 0
                
                vf_filters = []
                if pad_l > 0 or pad_t > 0 or pad_r > 0 or pad_b > 0:
                    # パディングが必要な場合
                    new_w = vid_w + pad_l + pad_r
                    new_h = vid_h + pad_t + pad_b
                    vf_filters.append(f"pad={new_w}:{new_h}:{pad_l}:{pad_t}:black")
                    
                    # パディング後の座標系でのクロップ位置
                    # 元の (vx1, vy1) は、パディング後は (vx1 + pad_l, vy1 + pad_t) になる
                    # vx1 が負の場合、vx1 + (-vx1) = 0 となり整合する
                    crop_x = vx1 + pad_l
                    crop_y = vy1 + pad_t
                else:
                    # パディング不要
                    crop_x = vx1
                    crop_y = vy1

                vf_filters.append(f"crop={vw}:{vh}:{crop_x}:{crop_y}")
                crop_filter_chain = ",".join(vf_filters)

                # 時間指定
                ss = self.start_time
                t_dur = self.end_time - self.start_time
                
                # 1. パレット作成
                cmd_palette = [
                    'ffmpeg', '-y', '-ss', str(ss), '-t', str(t_dur),
                    '-i', self.video_filepath,
                    '-vf', f"{crop_filter_chain},palettegen",
                    palette_path
                ]
                subprocess.run(cmd_palette, capture_output=True, check=True)

                # 2. GIF生成
                cmd_gif = [
                    'ffmpeg', '-y', '-ss', str(ss), '-t', str(t_dur),
                    '-i', self.video_filepath,
                    '-i', palette_path,
                    '-filter_complex', f"[0:v]{crop_filter_chain}[v];[v][1:v]paletteuse",
                    save_path
                ]
                subprocess.run(cmd_gif, capture_output=True, check=True)

                if os.path.exists(palette_path):
                    os.remove(palette_path)

                def _done():
                    progress_win.destroy()
                    if messagebox.askyesno("完了", "GIF保存が完了しました。フォルダを開きますか？"):
                        open_folder_with_selection(save_path)
                
                self.root.after(0, _done)

            except Exception as e:
                def _error():
                    progress_win.destroy()
                    messagebox.showerror("Error", f"GIFの生成に失敗しました:\n{e}")
                self.root.after(0, _error)

        import threading
        threading.Thread(target=_run, daemon=True).start()

    def export_video(self) -> None:
        """赤枠範囲をstart時間からend時間まで動画ファイルとして出力する."""
        if not self.cap:
            messagebox.showerror("Error", "動画なし")
            return

        # 保存先とファイル名をユーザーに選ばせる
        start_h = sec_to_hhmmss(int(self.start_time))
        start_frame = int((self.start_time - int(self.start_time)) * max(1, self.fps))
        end_h = sec_to_hhmmss(int(self.end_time))
        end_frame = int((self.end_time - int(self.end_time)) * max(1, self.fps))
        default_name = f"{self.video_filename}_trim_{start_h}_{start_frame:03d}_{end_h}_{end_frame:03d}.mp4"
        video_dir = os.path.dirname(self.video_filepath) if self.video_filepath else None
        save_path = filedialog.asksaveasfilename(
            defaultextension='.mp4',
            initialfile=default_name,
            initialdir=video_dir,
            filetypes=[('MP4', '*.mp4')],
            title='保存先とファイル名を選択'
        )
        if not save_path:
            return
        save_dir = os.path.dirname(save_path)

        # 座標変換
        x1, y1, x2, y2 = self.crop_rect
        vid_w = int(self.cap.get(cv2.CAP_PROP_FRAME_WIDTH))
        vid_h = int(self.cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
        vx1, vy1, vx2, vy2 = [int(v) for v in self.crop_rect]

        # 枠外選択が許可されていない場合は安全のためクランプ
        if not (getattr(self, 'allow_oversize_var', None) and self.allow_oversize_var.get()):
            vx1 = max(0, min(vid_w, vx1))
            vy1 = max(0, min(vid_h, vy1))
            vx2 = max(0, min(vid_w, vx2))
            vy2 = max(0, min(vid_h, vy2))

        # クロップ後のサイズ
        crop_w = vx2 - vx1
        crop_h = vy2 - vy1

        if crop_w <= 0 or crop_h <= 0:
            messagebox.showerror("Error", "クロップ範囲が無効です")
            return

        # 再生を一時停止
        was_playing = self.playing
        self.playing = False
        if self._play_after_id:
            self.root.after_cancel(self._play_after_id)

        # Create progress dialog
        progress_win = tk.Toplevel(self.root)
        progress_win.title("動画を書き出し中...")
        progress_win.transient(self.root)
        progress_win.grab_set()
        
        tk.Label(progress_win, text="動画を保存しています...").pack(padx=20, pady=(15, 5))
        pb = ttk.Progressbar(progress_win, orient=tk.HORIZONTAL, length=400, mode='determinate')
        pb.pack(padx=20, pady=(0, 10))
        prog_label = tk.Label(progress_win, text="0 / 0")
        prog_label.pack(padx=20, pady=(0, 15))

        def _run_export():
            try:
                # VideoWriter の設定
                fourcc = cv2.VideoWriter_fourcc(*'mp4v')
                out = cv2.VideoWriter(save_path, fourcc, self.fps, (crop_w, crop_h))

                if not out.isOpened():
                    self.root.after(0, lambda: messagebox.showerror("Error", "動画ファイルを作成できませんでした"))
                    self.root.after(0, progress_win.destroy)
                    return

                # start_time から end_time までのフレームを処理
                t = self.start_time
                limit = self.end_time
                frame_interval = 1.0 / self.fps
                total_frames = int((limit - t) / frame_interval) + 1
                
                self.root.after(0, lambda: pb.config(maximum=total_frames))
                
                frame_count = 0
                
                # Overlay用
                history_manager = overlay_utils.InputHistoryManager()
                last_search_idx = 0
                
                from utils import get_safe_crop

                while t <= limit:
                    self.cap.set(cv2.CAP_PROP_POS_MSEC, t * 1000)
                    ret, frm = self.cap.read()
                    if ret and frm is not None:
                        # get_safe_crop を使用してクロップ
                        crop = get_safe_crop(frm, (vx1, vy1, vx2, vy2), (0, 0, 0))
                        
                        if crop.size > 0:
                            # Overlay適用 (埋め込みチェックボックスがONの場合のみ)
                            if getattr(self, 'embed_overlay_var', None) and self.embed_overlay_var.get():
                                rgb = cv2.cvtColor(crop, cv2.COLOR_BGR2RGB)
                                pil = Image.fromarray(rgb)
                                last_search_idx = self._draw_overlay_on_image(pil, t, history_manager, last_search_idx, vx1, vy1)
                                crop = cv2.cvtColor(np.array(pil), cv2.COLOR_RGB2BGR)

                            # サイズが合わない場合はリサイズ (端数の関係でズレることがある)
                            if crop.shape[1] != crop_w or crop.shape[0] != crop_h:
                                crop = cv2.resize(crop, (crop_w, crop_h))
                            out.write(crop)
                            frame_count += 1
                    t += frame_interval
                    
                    # プログレス更新
                    curr_c = frame_count
                    self.root.after(0, lambda c=curr_c: pb.step(1) or prog_label.config(text=f"{c} / {total_frames}"))

                out.release()
                
                # 操作ログを抽出してTSV保存 (trajectory_data が存在する場合は常に実行)
                # 操作ログを抽出してTSV保存 (元ファイルベースの加工)
                # 元のTSVファイルパス
                original_tsv_path = os.path.splitext(self.video_filepath)[0] + ".tsv"
                
                # ベースファイル名 (拡張子なし)
                base_name = os.path.splitext(save_path)[0]
                tsv_save_path = f"{base_name}.tsv"

                if os.path.exists(original_tsv_path):
                    try:
                        with open(original_tsv_path, 'r', encoding='utf-8', newline='') as f_in, \
                             open(tsv_save_path, 'w', encoding='utf-8', newline='') as f_out:
                            
                            reader = csv.reader(f_in, delimiter='\t')
                            writer = csv.writer(f_out, delimiter='\t')
                            
                            header = next(reader, None)
                            if header:
                                writer.writerow(header)
                                
                                # 時間列のインデックスを探す
                                time_col_idx = -1
                                for i, col_name in enumerate(header):
                                    if col_name.lower() in ["timestamp", "time"]:
                                        time_col_idx = i
                                        break
                                
                                if time_col_idx >= 0:
                                    for row in reader:
                                        try:
                                            # 時間を取得
                                            t_val = float(row[time_col_idx])
                                            
                                            # フィルタリング
                                            if self.start_time <= t_val <= self.end_time:
                                                # 時間シフト
                                                new_t = t_val - self.start_time
                                                # 値を更新 (小数点3桁まで)
                                                row[time_col_idx] = f"{new_t:.3f}"
                                                writer.writerow(row)
                                        except ValueError:
                                            # 数値変換できない行はそのまま書き出すか、スキップするか...
                                            # ここでは安全のためそのまま書き出す（ヘッダー以外のメタ行の可能性）
                                            # ただしフィルタリングできないので、明らかにデータ行なら除外したいが
                                            # 形式不明なので書き出しておくのが無難か
                                            pass
                                else:
                                    # 時間列が見つからない場合は単純にコピーするしかない（トリミング不可）
                                    # あるいは警告を出すか。ここでは全コピー。
                                    for row in reader:
                                        writer.writerow(row)

                    except Exception as e:
                        print(f"TSV process failed: {e}")
                        # 失敗時はフォールバックへ

                # 元ファイルがない、または読み込みに失敗した場合のフォールバック
                # メモリ内の trajectory_data から生成
                elif hasattr(self, 'trajectory_data') and self.trajectory_data:
                    extracted_log = []
                    # (time, x, y, click, keys)
                    for row in self.trajectory_data:
                        r_t = row[0]
                        if self.start_time <= r_t <= self.end_time:
                            rel_t = r_t - self.start_time
                            extracted_log.append([rel_t] + list(row[1:]))
                    
                    if extracted_log:
                        try:
                            with open(tsv_save_path, 'w', newline='', encoding='utf-8') as f:
                                writer = csv.writer(f, delimiter='\t')
                                # recoder_core.py の形式に合わせる ("frame" はメモリデータにはないので省略するか、0で埋めるか)
                                # ここでは既存の video_frame_cropper 実装に合わせておくが、ヘッダー名を timestamp にする
                                writer.writerow(['timestamp', 'frame', 'x', 'y', 'click', 'keys'])
                                for r in extracted_log:
                                    # (rel_t, frame_idx, x, y, click, keys) 
                                    # trajectory_data は (time, frame_idx, x, y, click, keys) なので
                                    # r は [rel_t, frame_idx, x, y, click, keys] となっているはず
                                    writer.writerow(r)
                        except Exception as e:
                            print(f"Log export fallback failed: {e}")

                def _finish():
                    progress_win.destroy()
                    if messagebox.askyesno("完了", f"動画を保存しました。\n{frame_count} フレーム\nフォルダを開きますか？"):
                        open_folder_with_selection(save_path)
                
                self.root.after(0, _finish)
                
            except Exception as e:
                def _err():
                    progress_win.destroy()
                    messagebox.showerror("Error", f"動画保存中にエラーが発生しました:\n{e}")
                self.root.after(0, _err)
            finally:
                self.root.after(0, lambda: setattr(self, 'playing', was_playing))
                if was_playing:
                    self.root.after(0, self.play_step)

        import threading
        threading.Thread(target=_run_export, daemon=True).start()

    def _draw_overlay_on_image(self, img_pil, current_time, history_manager, last_idx, offset_x=0, offset_y=0):
        """画像にマウスと入力履歴のオーバーレイを描画する helper"""
        if not hasattr(self, 'trajectory_data') or not self.trajectory_data:
            return last_idx

        # 設定の取得
        show_mouse = self.show_trajectory_var.get() if hasattr(self, 'show_trajectory_var') else True
        show_sub = self.show_subtitle_var.get() if hasattr(self, 'show_subtitle_var') else True
        
        if not show_mouse and not show_sub:
            return last_idx

        # 1. 該当するマウスデータの検索 (last_idx から開始)
        mouse_data = None
        current_row_idx = -1
        
        # タイムスタンプが current_time に最も近いものを探す
        # ただし、過去のものはスキップして高速化
        start_search = max(0, last_idx)
        n = len(self.trajectory_data)
        
        best_diff = float('inf')
        best_idx = -1
        
        # 少し先まで探す (データの粒度にもよるが、1秒分くらい見れば十分)
        # ここでは簡易的に、current_time を超えて diff が増え始めたら終了
        for i in range(start_search, n):
            row = self.trajectory_data[i]
            t = row[0]
            diff = abs(t - current_time)
            
            if diff < best_diff:
                best_diff = diff
                best_idx = i
                mouse_data = row
            else:
                # 差が広がり始めたら終了 (ただし t < current_time の場合は近づいてる途中かもしれないので注意)
                if t > current_time:
                    break
        
        # マッチした行があればそこを新しい last_idx とする
        new_last_idx = best_idx if best_idx >= 0 else last_idx
        
        # 閾値判定 (フレームレート依存だが、例えば 1/fps 以内なら採用)
        valid_mouse = False
        if mouse_data and best_diff < (1.0 / self.fps if self.fps > 0 else 0.033):
            valid_mouse = True
            
        t_curr = current_time
        
        # --- マウス描画 ---
        if show_mouse and valid_mouse:
            t, x, y, click, keys = mouse_data
            # クロップ座標系に変換
            draw_x = x - offset_x
            draw_y = y - offset_y
            
            # 波紋用
            ripple_age = 0.0
            ripple_type = ""
            if best_idx > 0:
                 # 少し遡ってクリック開始を探す
                 lookback = 0.5
                 for k in range(best_idx, max(0, best_idx - 30), -1):
                     tr = self.trajectory_data
                     if t_curr - tr[k][0] > lookback: break
                     if k > 0:
                         c_c = tr[k][3]
                         c_p = tr[k-1][3]
                         for btn_char, btn_name in [("L", "left"), ("R", "right"), ("M", "middle")]:
                            if btn_char in c_p and btn_char not in c_c:
                                ripple_type = btn_name
                                ripple_age = t_curr - tr[k][0]
                                break
                         if ripple_type: break

            overlay_utils.draw_mouse_overlay(
                img_pil, draw_x, draw_y, click, 
                1.0, 1.0, # 画像解像度そのままなのでスケールは1.0
                self.theme,
                ripple_age=ripple_age,
                ripple_type=ripple_type
            )

        # --- 字幕データの履歴更新 (last_idx から best_idx まで全て走査して反映) ---
        if show_sub and best_idx >= 0:
            for k in range(start_search, best_idx + 1):
                row = self.trajectory_data[k]
                t_row, _, _, click_row, keys_row = row
                disp_text = overlay_utils.get_input_display_text(click_row, keys_row)
                history_manager.update(t_row, disp_text)

        # 字幕描画 (履歴があれば valid_mouse がなくても描画する)
        if show_sub:
            fade_duration = self.theme.get("input_overlay", {}).get("fade_duration", 1.0)
            active_inputs = history_manager.get_active_inputs(t_curr, fade_duration)
            if active_inputs:
                overlay_utils.draw_input_overlay(
                    img_pil, active_inputs,
                    1.0, 1.0, # スケール1.0
                    self.theme
                )

        return new_last_idx
