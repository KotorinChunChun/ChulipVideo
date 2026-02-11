import tkinter as tk
import tkinter.ttk as ttk
import cv2
import math
from PIL import Image, ImageTk
import os

class VideoPlayerWidget(tk.Frame):
    def __init__(self, parent, path, fps_callback=None):
        super().__init__(parent, relief=tk.RAISED, borderwidth=1)
        self.path = path
        self.cap = None
        self.total_frames = 0
        self.fps = 30.0
        self.current_frame_idx = 0
        self.is_finished = False
        self.photo_image = None
        
        # Callbacks
        self.fps_callback = fps_callback # To notify manager of my FPS/Duration
        
        # UI Elements
        self.lbl_title = tk.Label(self, text=os.path.basename(path), anchor="w", bg="#ddd")
        self.lbl_title.pack(fill=tk.X)
        
        self.canvas = tk.Canvas(self, bg="black", highlightthickness=0)
        self.canvas.pack(fill=tk.BOTH, expand=True)
        self.canvas.bind("<Configure>", self._on_resize)
        
        self.ctrl_frame = tk.Frame(self)
        self.ctrl_frame.pack(fill=tk.X, padx=2, pady=2)
        
        self.btn_play = tk.Button(self.ctrl_frame, text="▶", command=self.toggle_play, width=3, font=("Arial", 8))
        self.btn_play.pack(side=tk.LEFT)
        
        self.var_seek = tk.DoubleVar()
        self.slider = ttk.Scale(self.ctrl_frame, from_=0, to=100, variable=self.var_seek, orient=tk.HORIZONTAL, command=self._on_seek)
        self.slider.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)
        
        self.lbl_time = tk.Label(self.ctrl_frame, text="00:00", font=("Arial", 8), width=5)
        self.lbl_time.pack(side=tk.LEFT)
        
        self._load_video()
        
        # Internal state
        self.is_playing = False
        self.is_manual_seeking = False

    def _load_video(self):
        self.cap = cv2.VideoCapture(self.path)
        if self.cap.isOpened():
            self.total_frames = int(self.cap.get(cv2.CAP_PROP_FRAME_COUNT))
            self.fps = self.cap.get(cv2.CAP_PROP_FPS)
            if self.fps <= 0: self.fps = 30.0
            
            self.slider.config(to=self.total_frames - 1)
            # Show first frame
            self.seek(0)
            
            if self.fps_callback:
                self.fps_callback(self.total_frames, self.fps)

    def _on_resize(self, event):
        # Re-draw current frame with new size
        if self.image_cache:
             self._display_image(self.image_cache)

    def _on_seek(self, value):
        # Called when USER moves the slider
        val = int(float(value))
        self.seek(val)
        
    def toggle_play(self):
        self.is_playing = not self.is_playing
        self.btn_play.config(text="||" if self.is_playing else "▶")

    def set_playing(self, playing):
        self.is_playing = playing
        self.btn_play.config(text="||" if self.is_playing else "▶")

    def seek(self, frame_idx):
        if not self.cap: return
        
        frame_idx = max(0, min(frame_idx, self.total_frames - 1))
        self.current_frame_idx = frame_idx
        self.cap.set(cv2.CAP_PROP_POS_FRAMES, frame_idx)
        ret, frame = self.cap.read()
        if ret:
            frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            self._display_image(Image.fromarray(frame))
            self.is_finished = False
        else:
            # End of video
            self.is_finished = True
            self._show_finished_overlay()
            
        self.var_seek.set(frame_idx)
        self._update_time_label(frame_idx)

    def step(self):
        """Advance one frame if playing"""
        if not self.is_playing or not self.cap: return
        
        if self.current_frame_idx < self.total_frames - 1:
            ret, frame = self.cap.read()
            if ret:
                self.current_frame_idx += 1
                frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                self._display_image(Image.fromarray(frame))
                self.var_seek.set(self.current_frame_idx)
                self._update_time_label(self.current_frame_idx)
            else:
                self.is_finished = True
                self._show_finished_overlay()
        else:
            self.is_finished = True
            self._show_finished_overlay()

    def image_cache(self):
        return self._last_pil_image

    def _display_image(self, pil_image):
        self._last_pil_image = pil_image
        w = self.canvas.winfo_width()
        h = self.canvas.winfo_height()
        if w < 10 or h < 10: return
        
        # Aspect fit
        iw, ih = pil_image.size
        scale = min(w/iw, h/ih)
        nw, nh = int(iw*scale), int(ih*scale)
        if nw > 0 and nh > 0:
            resized = pil_image.resize((nw, nh), Image.Resampling.LANCZOS)
            self.photo_image = ImageTk.PhotoImage(resized)
            self.canvas.delete("all")
            self.canvas.create_image(w//2, h//2, image=self.photo_image, anchor=tk.CENTER)

    def _show_finished_overlay(self):
        w = self.canvas.winfo_width()
        h = self.canvas.winfo_height()
        self.canvas.create_rectangle(0, 0, w, h, fill="#888888", stipple="gray50", outline="")
        self.canvas.create_text(w//2, h//2, text="Finished", fill="white", font=("Arial", 16, "bold"))

    def _update_time_label(self, frame):
        sec = int(frame / self.fps)
        m, s = divmod(sec, 60)
        self.lbl_time.config(text=f"{m:02d}:{s:02d}")
        
    def release(self):
        if self.cap:
            self.cap.release()
            self.cap = None

class MultiVideoManager:
    def __init__(self, container, paths):
        self.container = container
        self.paths = paths
        self.widgets = []
        self.max_duration_sec = 0.0
        self.base_fps = 30.0 # Assumed base FPS for global sync
        
        self._setup_ui()
        
    def _setup_ui(self):
        # 1. Clear container
        for w in self.container.winfo_children():
            w.destroy()
            
        # 2. Calculate grid size
        n = len(self.paths)
        if n <= 2: rows, cols = 1, 2
        elif n <= 4: rows, cols = 2, 2
        elif n <= 6: rows, cols = 2, 3
        elif n <= 9: rows, cols = 3, 3
        elif n <= 12: rows, cols = 3, 4
        else: rows, cols = 4, 4 # Max 16
        
        # 3. Create widgets
        for i, path in enumerate(self.paths):
            if i >= 16: break
            
            r = i // cols
            c = i % cols
            
            # Configure grid weight
            self.container.grid_rowconfigure(r, weight=1)
            self.container.grid_columnconfigure(c, weight=1)
            
            widget = VideoPlayerWidget(self.container, path, self._on_video_info)
            widget.grid(row=r, column=c, sticky="nsew", padx=2, pady=2)
            self.widgets.append(widget)
            
    def _on_video_info(self, total_frames, fps):
        duration = total_frames / fps if fps > 0 else 0
        if duration > self.max_duration_sec:
            self.max_duration_sec = duration
            # Only adopt the FPS of the longest video if not set, or just stick to standard?
            # For global slider, let's use the longest video's frame count as reference?
            # Or use seconds for global slider.
            # Let's say global slider is 0..100% or 0..MaxFrames?
            # User wants "Global slider managed based on the longest video time".
            pass

    def get_max_duration(self):
        return self.max_duration_sec

    def play_all(self):
        for w in self.widgets:
            w.set_playing(True)

    def stop_all(self):
        for w in self.widgets:
            w.set_playing(False)

    def seek_all_by_ratio(self, ratio):
        """ratio: 0.0 to 1.0"""
        for w in self.widgets:
            # Each video has different length. 
            # "Simultaneous playback" usually means they play at same Real Time.
            # So if global is at 10s, all videos should be at 10s.
            target_sec = self.max_duration_sec * ratio
            target_frame = int(target_sec * w.fps)
            w.seek(target_frame)

    def on_global_step(self):
        """Called by main loop"""
        # We need to decide who drives the loop.
        # Ideally, we call w.step() for each widget.
        any_playing = False
        for w in self.widgets:
            # We trust individual widgets to handle their frame stepping
            if not w.is_finished and w.is_playing:
                w.step()
                any_playing = True
        return any_playing
        
    def release(self):
        for w in self.widgets:
            w.release()
        self.widgets = []
