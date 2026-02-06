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

import cv2

from utils import imwrite_jp, sec_to_hhmmss

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
        
        # 安全のためクランプ
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
            last_saved_crop = None

            step_idx = 0
            while t <= limit:
                self.cap.set(cv2.CAP_PROP_POS_MSEC, t * 1000)
                ret, frm = self.cap.read()
                if ret and frm is not None:
                    crop = frm[vy1:vy2, vx1:vx2]
                    if crop.size > 0:
                        # 次のフレームを取得
                        next_t = t + frame_interval
                        if next_t <= limit:
                            self.cap.set(cv2.CAP_PROP_POS_MSEC, next_t * 1000)
                            ret_next, frm_next = self.cap.read()
                            if ret_next and frm_next is not None:
                                next_crop = frm_next[vy1:vy2, vx1:vx2]
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

                        # 直前に出力したフレームとも比較
                        is_same_as_last_saved = False
                        if self.check_duplicate.get() and last_saved_crop is not None:
                            diff_last = cv2.cvtColor(cv2.absdiff(last_saved_crop, crop), cv2.COLOR_BGR2GRAY)
                            if cv2.countNonZero(diff_last) == 0:
                                is_same_as_last_saved = True

                        # チェックボックスの設定に応じて出力判定
                        if is_matches_prev_next and not is_same_as_last_saved:
                            time_str = sec_to_hhmmss(t)
                            frame_in_sec = int((t - int(t)) * self.fps)
                            filepath = os.path.join(save_dir, f"{self.video_filename}_{time_str}_{frame_in_sec:03d}.png")
                            imwrite_jp(filepath, crop, params=save_params)
                            last_saved_crop = crop.copy()
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
                open_folder(save_dir)
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
                # 一旦ワークディレクトリにパレットを出力
                palette_path = os.path.join(os.path.dirname(save_path), ".palette.png")
                # クロップ設定
                crop_str = f"crop={vw}:{vh}:{vx1}:{vy1}"
                # 時間指定
                ss = self.start_time
                t_dur = self.end_time - self.start_time
                
                # 1. パレット作成
                cmd_palette = [
                    'ffmpeg', '-y', '-ss', str(ss), '-t', str(t_dur),
                    '-i', self.video_filepath,
                    '-vf', f"{crop_str},palettegen",
                    palette_path
                ]
                subprocess.run(cmd_palette, capture_output=True, check=True)

                # 2. GIF生成
                cmd_gif = [
                    'ffmpeg', '-y', '-ss', str(ss), '-t', str(t_dur),
                    '-i', self.video_filepath,
                    '-i', palette_path,
                    '-filter_complex', f"[0:v]{crop_str}[v];[v][1:v]paletteuse",
                    save_path
                ]
                subprocess.run(cmd_gif, capture_output=True, check=True)

                if os.path.exists(palette_path):
                    os.remove(palette_path)

                def _done():
                    progress_win.destroy()
                    if messagebox.askyesno("完了", "GIF保存が完了しました。フォルダを開きますか？"):
                        open_folder(os.path.dirname(save_path))
                
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

        # 安全のためクランプ
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
                while t <= limit:
                    self.cap.set(cv2.CAP_PROP_POS_MSEC, t * 1000)
                    ret, frm = self.cap.read()
                    if ret and frm is not None:
                        crop = frm[vy1:vy2, vx1:vx2]
                        if crop.size > 0:
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

                def _finish():
                    progress_win.destroy()
                    if messagebox.askyesno("完了", f"動画を保存しました。\n{frame_count} フレーム\nフォルダを開きますか？"):
                        open_folder(save_dir)
                
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
