import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, Checkbutton
import subprocess
import os
import cv2
from PIL import Image, ImageTk
import threading
import shutil
import uuid
import time
import json
from datetime import datetime
from tkinterdnd2 import DND_FILES, TkinterDnD
import re
import numpy as np
from queue import Queue

class MosaicRemoverApp:
    def __init__(self, root):
        self.root = root
        self.root.title("動画モザイク除去 GUI (VR対応 20251002-6)")
        self.root.geometry("1000x1000")
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        self.script_dir = os.path.dirname(os.path.abspath(__file__))
        self.ps_script_path = os.path.join(self.script_dir, "LADA_LAUNCHER_FOR_GUI.ps1")
        self.output_dir = os.path.join(self.script_dir, "output")
        self.log_file = os.path.join(self.script_dir, "LOG_LADA_GUI.txt")
        
        self.root.drop_target_register(DND_FILES)
        self.root.dnd_bind('<<Drop:DND_Files>>', self.drop_file)
        
        self.bind_keys(self.root)
        self.root.bind('<Configure>', self.on_window_resize)
        self.after_id = None
        
        self.cap = None
        self.paused = True
        self.current_frame = 0
        self.video_fps = 30.0
        self.actual_fps = 30.0
        self.video_total_frames = 0
        self.process = None
        self.start_frame = 0
        self.end_frame = 0
        self.video_path = ""
        
        self.config_file = "config.ini"
        self.queue_file = "processing_queue.json"
        self.cli_options = {
            "model_choice": "1",
            "tvai_choice": "2",
            "quality": "15",
            "crf_value": "19"
        }
        self.processing_queue = self.load_queue()
        self.is_batch_processing = False
        self.is_running = False
        
        self.frame_queue = Queue(maxsize=3)
        self.frame_buffer_thread = None
        self.buffer_running = False
        self.cap_lock = threading.Lock()
        self.last_frame_time = time.time()
        
        if not os.path.exists(self.ps_script_path):
            messagebox.showerror("エラー", "PowerShellスクリプト 'LADA_LAUNCHER_FOR_GUI.ps1' が見つかりません。")
            self.ps_script_path = None
        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir)
        
        self.create_widgets()
        self.load_config()
        self.root.after(100, self.update_preview)
        
        self.fullscreen_window = None
        self.fullscreen_progress_canvas = None
        self.fullscreen_progress_bar = None
        self.fullscreen_start_marker = None
        self.fullscreen_end_marker = None
        self.fullscreen_progress_text = None

    def bind_keys(self, window):
        window.bind('<Right>', self.move_frame)
        window.bind('<Left>', self.move_frame)
        window.bind('<Shift-Right>', self.move_frame)
        window.bind('<Shift-Left>', self.move_frame)
        window.bind('<space>', self.toggle_play_pause)
        window.bind('<Up>', self.jump_to_start)
        window.bind('<Down>', self.jump_to_end)
        window.bind('<Control-Up>', self.set_start_point_by_key)
        window.bind('<Control-Down>', self.set_end_point_by_key)
        window.bind('<Control-e>', self.add_to_queue)
        window.bind('<Control-q>', lambda e: self.open_queue_window())
        window.bind('<Control-r>', lambda e: self.reset_points())
        window.bind('f', self.toggle_fullscreen)
        window.bind('j', self.move_one_frame_backward)
        window.bind('k', self.toggle_play_pause)
        window.bind('l', self.move_one_frame_forward)
        window.bind('h', self.move_one_second_backward)
        window.bind(';', self.move_one_second_forward)
        window.bind('<Home>', lambda e: self.set_frame_and_start(0))
        window.bind('<End>', lambda e: self.set_frame_and_end(self.video_total_frames))
        window.bind('s', self.jump_to_video_start)
        window.bind('e', self.jump_to_video_end)
        for i in range(1, 10):
            window.bind(str(i), lambda e, percentage=i*10: self.jump_to_percentage(percentage))

    def jump_to_video_start(self, event=None):
        if not self.cap or not self.cap.isOpened():
            return
        self.current_frame = 0
        self.clear_frame_queue()
        with self.cap_lock:
            try:
                self.cap.set(cv2.CAP_PROP_POS_FRAMES, 0)
                ret, frame = self.cap.read()
                if ret:
                    self.display_frame(frame)
                    if self.fullscreen_window:
                        self.display_frame_fullscreen(frame)
                self.on_progress_update()
                self.update_time_labels()
            except Exception as e:
                self.write_log(f"動画先頭ジャンプエラー: {e}")

    def jump_to_video_end(self, event=None):
        if not self.cap or not self.cap.isOpened():
            return
        self.current_frame = max(0, self.video_total_frames - 1)
        self.clear_frame_queue()
        with self.cap_lock:
            try:
                self.cap.set(cv2.CAP_PROP_POS_FRAMES, self.current_frame)
                ret, frame = self.cap.read()
                if ret:
                    self.display_frame(frame)
                    if self.fullscreen_window:
                        self.display_frame_fullscreen(frame)
                self.on_progress_update()
                self.update_time_labels()
            except Exception as e:
                self.write_log(f"動画末尾ジャンプエラー: {e}")

    def exit_fullscreen(self, event=None):
        if self.fullscreen_window:
            self.buffer_running = False
            try:
                self.fullscreen_window.destroy()
            except:
                pass
            self.fullscreen_window = None
            self.fullscreen_progress_canvas = None
            self.fullscreen_progress_bar = None
            self.fullscreen_start_marker = None
            self.fullscreen_end_marker = None
            self.fullscreen_progress_text = None
            self.clear_frame_queue()

    def set_frame_and_start(self, frame):
        if self.cap and self.cap.isOpened():
            self.current_frame = frame
            self.clear_frame_queue()
            with self.cap_lock:
                try:
                    self.cap.set(cv2.CAP_PROP_POS_FRAMES, frame)
                    ret, frame_data = self.cap.read()
                    if ret:
                        self.display_frame(frame_data)
                        if self.fullscreen_window:
                            self.display_frame_fullscreen(frame_data)
                    self.on_progress_update()
                    self.update_time_labels()
                    self.set_start_point_by_key()
                except Exception as e:
                    self.write_log(f"先頭フレーム設定エラー: {e}")

    def set_frame_and_end(self, frame):
        if self.cap and self.cap.isOpened():
            self.current_frame = frame
            self.clear_frame_queue()
            with self.cap_lock:
                try:
                    self.cap.set(cv2.CAP_PROP_POS_FRAMES, frame)
                    ret, frame_data = self.cap.read()
                    if ret:
                        self.display_frame(frame_data)
                        if self.fullscreen_window:
                            self.display_frame_fullscreen(frame_data)
                    self.on_progress_update()
                    self.update_time_labels()
                    self.set_end_point_by_key()
                except Exception as e:
                    self.write_log(f"末尾フレーム設定エラー: {e}")

    def load_queue(self):
        if os.path.exists(self.queue_file):
            try:
                with open(self.queue_file, 'r', encoding='utf-8') as f:
                    queue = json.load(f)
                    self.write_log(f"キューを読み込みました: {len(queue)} 項目")
                    return queue
            except Exception as e:
                self.write_log(f"キュー読み込みエラー: {e}")
                messagebox.showwarning("警告", f"キュー読み込みに失敗しました: {e}。空のキューで続行します。")
                return []
        self.write_log("キューが存在しません。新規作成します。")
        return []

    def save_queue(self):
        try:
            with open(self.queue_file, 'w', encoding='utf-8') as f:
                json.dump(self.processing_queue, f, ensure_ascii=False, indent=2)
        except Exception as e:
            self.write_log(f"キュー保存エラー: {e}")
            messagebox.showwarning("警告", f"キュー保存に失敗しました: {e}。手動で確認してください。")

    def generate_unique_filepath(self, base_path):
        if not os.path.exists(base_path):
            return base_path
        base, ext = os.path.splitext(base_path)
        counter = 1
        while True:
            new_path = f"{base}_{counter}{ext}"
            if not os.path.exists(new_path):
                return new_path
            counter += 1

    def create_widgets(self):
        main_frame = tk.Frame(self.root, padx=10, pady=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        main_frame.grid_rowconfigure(2, weight=4)
        main_frame.grid_rowconfigure(5, weight=1)
        main_frame.grid_columnconfigure(0, weight=1)

        file_frame = tk.LabelFrame(main_frame, text="1. 動画ファイル選択", padx=10, pady=10)
        file_frame.grid(row=0, column=0, sticky="ew", pady=5)
        
        self.file_path_entry = tk.Entry(file_frame)
        self.file_path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        self.browse_button = tk.Button(file_frame, text="参照...", command=self.browse_file)
        self.browse_button.pack(side=tk.LEFT, padx=5)

        options_frame = tk.LabelFrame(main_frame, text="2. CLIオプション設定", padx=10, pady=10)
        options_frame.grid(row=1, column=0, sticky="ew", pady=5)

        model_label = tk.Label(options_frame, text="検出モデル:")
        model_label.pack(side=tk.LEFT, padx=(0, 5))
        self.model_var = tk.StringVar(value=self.cli_options["model_choice"])
        self.model_var.trace_add("write", self.save_config_callback)
        self.model_menu = tk.OptionMenu(options_frame, self.model_var, "1", "2", "3")
        self.model_menu.pack(side=tk.LEFT, padx=5)
        
        tvai_label = tk.Label(options_frame, text="TVAIで画質向上:")
        tvai_label.pack(side=tk.LEFT, padx=(15, 5))
        self.tvai_var = tk.StringVar(value=self.cli_options["tvai_choice"])
        self.tvai_var.trace_add("write", self.save_config_callback)
        self.tvai_menu = tk.OptionMenu(options_frame, self.tvai_var, "1", "2")
        self.tvai_menu.pack(side=tk.LEFT, padx=5)
        
        quality_label = tk.Label(options_frame, text="映像品質(5-30):")
        quality_label.pack(side=tk.LEFT, padx=(15, 5))
        self.quality_var = tk.StringVar(value=self.cli_options["quality"])
        self.quality_var.trace_add("write", self.save_config_callback)
        quality_values = [str(i) for i in range(5, 31)]
        self.quality_menu = tk.OptionMenu(options_frame, self.quality_var, *quality_values)
        self.quality_menu.pack(side=tk.LEFT, padx=5)

        preview_frame = tk.LabelFrame(main_frame, text="3. 処理範囲の指定", padx=10, pady=10)
        preview_frame.grid(row=2, column=0, sticky="nsew", pady=5)
        preview_frame.grid_rowconfigure(0, weight=1)
        preview_frame.grid_columnconfigure(0, weight=1)

        self.video_label = tk.Label(preview_frame, bg="black")
        self.video_label.grid(row=0, column=0, sticky="nsew")
        self.video_label.bind("<Button-1>", self.toggle_play_pause)
        self.video_label.bind("<Double-Button-1>", self.toggle_fullscreen)
        self.video_label.bind("<MouseWheel>", self.on_mouse_wheel)

        self.progress_canvas = tk.Canvas(preview_frame, height=20, bg="grey")
        self.progress_canvas.grid(row=1, column=0, sticky="ew", pady=2)
        self.progress_canvas.bind("<Button-1>", self.on_progress_click)
        self.progress_canvas.bind("<MouseWheel>", self.on_mouse_wheel)
        self.progress_bar = self.progress_canvas.create_rectangle(0, 0, 0, 20, fill="green")
        self.start_marker = self.progress_canvas.create_line(0, 0, 0, 20, fill="red", width=2)
        self.end_marker = self.progress_canvas.create_line(0, 0, 0, 20, fill="blue", width=2)
        
        control_range_frame = tk.Frame(preview_frame)
        control_range_frame.grid(row=2, column=0, pady=2)

        self.play_pause_button = tk.Button(control_range_frame, text="▶ 再生", command=self.toggle_play_pause)
        self.play_pause_button.pack(side=tk.LEFT)

        self.current_time_label = tk.Label(control_range_frame, text="00:00:00 / 00:00:00")
        self.current_time_label.pack(side=tk.LEFT, padx=10)

        self.set_start_button = tk.Button(control_range_frame, text="開始点を指定", command=self.set_start_point)
        self.set_start_button.pack(side=tk.LEFT, padx=5)
        
        self.set_end_button = tk.Button(control_range_frame, text="終了点を指定", command=self.set_end_point)
        self.set_end_button.pack(side=tk.LEFT, padx=5)

        time_display_frame = tk.Frame(preview_frame)
        time_display_frame.grid(row=3, column=0, pady=2, padx=100)
        tk.Label(time_display_frame, text="開始時間:").pack(side=tk.LEFT)
        self.start_time_label = tk.Label(time_display_frame, text="00:00:00", width=12, relief="sunken")
        self.start_time_label.pack(side=tk.LEFT)
        tk.Label(time_display_frame, text="終了時間:").pack(side=tk.LEFT, padx=(10, 0))
        self.end_time_label = tk.Label(time_display_frame, text="00:00:00", width=12, relief="sunken")
        self.end_time_label.pack(side=tk.LEFT)
        
        self.reset_button = tk.Button(time_display_frame, text="範囲リセット", command=self.reset_points)
        self.reset_button.pack(side=tk.LEFT, padx=10)

        ffmpeg_frame = tk.LabelFrame(main_frame, text="4. 動画切り出し設定", padx=10, pady=10)
        ffmpeg_frame.grid(row=3, column=0, sticky="ew", pady=5)
        
        self.ffmpeg_option_var = tk.StringVar(value="re_encode")
        
        tk.Radiobutton(ffmpeg_frame, text="-c copy (高速)", variable=self.ffmpeg_option_var, value="copy").pack(side=tk.LEFT, padx=5)
        tk.Radiobutton(ffmpeg_frame, text="-c copy +genpts (タイムスタンプ修正)", variable=self.ffmpeg_option_var, value="copy_genpts").pack(side=tk.LEFT, padx=5)
        tk.Radiobutton(ffmpeg_frame, text="再エンコード (NVENC)", variable=self.ffmpeg_option_var, value="re_encode").pack(side=tk.LEFT, padx=5)
        
        crf_label = tk.Label(ffmpeg_frame, text="映像品質(5-30):")
        crf_label.pack(side=tk.LEFT, padx=(15, 5))
        self.crf_var = tk.StringVar(value=self.cli_options["crf_value"])
        self.crf_var.trace_add("write", self.save_config_callback)
        crf_values = [str(i) for i in range(5, 31)]
        self.crf_menu = tk.OptionMenu(ffmpeg_frame, self.crf_var, *crf_values)
        self.crf_menu.pack(side=tk.LEFT, padx=5)
        
        self.batch_count_label = tk.Label(ffmpeg_frame, text="", fg="blue")
        self.batch_count_label.pack(side=tk.RIGHT, padx=5)

        # VR処理チェックボックス追加
        vr_frame = tk.LabelFrame(main_frame, text="5. VR映像処理", padx=10, pady=10)
        vr_frame.grid(row=4, column=0, sticky="ew", pady=5)
        
        self.vr_processing_var = tk.BooleanVar(value=False)
        self.vr_processing_check = Checkbutton(vr_frame, text="VR処理(180度SBS形式)", variable=self.vr_processing_var, command=self.on_vr_mode_toggle)
        self.vr_processing_check.pack(side=tk.LEFT, padx=5)
        
        self.vr_simple_mode_var = tk.BooleanVar(value=True)  # デフォルトをTrueに変更
        self.vr_simple_mode_check = Checkbutton(vr_frame, text="簡易処理モード(中央70%のみ)", variable=self.vr_simple_mode_var, state=tk.DISABLED)  # 変更不可に設定
        self.vr_simple_mode_check.pack(side=tk.LEFT, padx=5)

        control_frame = tk.Frame(main_frame, pady=10)
        control_frame.grid(row=5, column=0, sticky="ew")
        
        self.save_trimmed_video_var = tk.BooleanVar(value=False)
        self.save_trimmed_video_check = Checkbutton(control_frame, text="切り出し動画を保存", variable=self.save_trimmed_video_var)
        self.save_trimmed_video_check.pack(side=tk.LEFT, padx=5)
        
        self.show_completion_dialog_var = tk.BooleanVar(value=True)
        self.show_completion_dialog_check = Checkbutton(control_frame, text="完了ダイアログを表示", variable=self.show_completion_dialog_var)
        self.show_completion_dialog_check.pack(side=tk.LEFT, padx=5)
        
        self.queue_view_button = tk.Button(control_frame, text="キュー確認", command=self.open_queue_window)
        self.queue_view_button.pack(side=tk.LEFT, padx=5)
        
        self.start_button = tk.Button(control_frame, text="処理開始 (単一)", command=self.start_processing)
        self.start_button.pack(side=tk.LEFT, padx=5)
        
        self.batch_button = tk.Button(control_frame, text="一括開始", command=lambda: self.start_batch_processing(control_frame))
        self.batch_button.pack(side=tk.LEFT, padx=5)
        
        self.status_label = tk.Label(control_frame, text="準備完了", fg="blue")
        self.status_label.pack(side=tk.LEFT, padx=5)
        
        self.abort_button = tk.Button(control_frame, text="中断", command=self.abort_processing, bg="orange", fg="white")
        self.abort_button.pack(side=tk.RIGHT, padx=5)
        
        self.log_button = tk.Button(control_frame, text="ログ", command=self.open_log_file)
        self.log_button.pack(side=tk.RIGHT, padx=5)

        time_display_frame = tk.Frame(preview_frame)
        time_display_frame.grid(row=5, column=0, pady=2, padx=100)
        self.queue_add_button = tk.Button(time_display_frame, text="キューに追加", command=self.add_to_queue, bg="#87CEEB")
        self.queue_add_button.pack(side=tk.LEFT, padx=5)

        self.suppress_queue_message_var = tk.BooleanVar(value=True)
        self.suppress_queue_message_check = Checkbutton(time_display_frame, text="キュー追加時メッセージなし", variable=self.suppress_queue_message_var)
        self.suppress_queue_message_check.pack(side=tk.RIGHT, padx=5)

        lada_info_frame = tk.LabelFrame(main_frame, text="LADA処理情報", padx=10, pady=10)
        lada_info_frame.grid(row=6, column=0, sticky="nsew", pady=5)

        self.console_text = scrolledtext.ScrolledText(lada_info_frame, height=5, state=tk.DISABLED)
        self.console_text.pack(fill=tk.BOTH, expand=True, pady=10)
        
    def on_vr_mode_toggle(self):
        """VRモード切り替え時の処理"""
        if self.vr_processing_var.get():
            # VRモードON: 簡易処理モードを強制ON
            self.vr_simple_mode_var.set(True)
            self.vr_simple_mode_check.config(state=tk.DISABLED)
            self.write_log("VRモード有効化: 簡易処理モード(中央70%)で動作")
        else:
            # VRモードOFF: 簡易処理モードのチェックは維持するが操作不可のまま
            pass
    
    def apply_vr_undistortion(self, input_file, output_file, unique_id):
        """180度SBS映像 - 中央領域を抽出（面積約70%）"""
        self.console_text.config(state=tk.NORMAL)
        self.console_text.insert(tk.END, f"VR映像の中央領域を抽出中（面積70%）...\n")
        self.console_text.config(state=tk.DISABLED)
        self.write_log("VR中央領域抽出開始（面積70%）")
        
        # 面積70%なら縦横83.7%（√0.7 ≈ 0.837）
        crop_center_cmd = [
            'ffmpeg', '-y', '-i', input_file,
            '-vf', 'crop=iw*0.837:ih*0.837:iw*0.0815:ih*0.0815',
            '-c:v', 'h264_nvenc', '-preset', 'p4', '-cq', '18',
            '-an',
            output_file
        ]
        subprocess.run(crop_center_cmd, check=True, creationflags=subprocess.CREATE_NO_WINDOW)
        
        self.write_log("VR中央領域抽出完了")

    def apply_vr_distortion(self, input_file, output_file, unique_id):
        """LADA処理済み中央領域を元動画の同じ位置に合成"""
        self.console_text.config(state=tk.NORMAL)
        self.console_text.insert(tk.END, f"処理済み領域を元動画に合成中...\n")
        self.console_text.config(state=tk.DISABLED)
        self.write_log("元動画への合成開始")
        
        # 元の切り出しファイルを探す
        trimmed_file = None
        for file in os.listdir(self.output_dir):
            if file.startswith(f'trimmed_{unique_id}') and file.endswith('.mp4'):
                trimmed_file = os.path.join(self.output_dir, file)
                break
        
        if not trimmed_file or not os.path.exists(trimmed_file):
            self.write_log("エラー: 元の切り出し動画が見つかりません")
            raise Exception("元の切り出し動画が見つかりません")
        
        # LADA処理済み中央領域を元動画の中央に重ねる
        overlay_cmd = [
            'ffmpeg', '-y',
            '-i', trimmed_file,  # 元動画（背景）
            '-i', input_file,    # LADA処理済み中央部
            '-filter_complex', '[0:v][1:v]overlay=(W-w)/2:(H-h)/2[v]',
            '-map', '[v]',
            '-c:v', 'h264_nvenc', '-preset', 'p4', '-cq', '18',
            output_file
        ]
        subprocess.run(overlay_cmd, check=True, creationflags=subprocess.CREATE_NO_WINDOW)
        
        self.write_log("元動画への合成完了")

    def split_vr_video(self, input_file, unique_id):
        """VR映像処理 - 中央領域のみ抽出（簡易モード専用）"""
        
        # 1. 音声を抽出
        audio_file = os.path.join(self.output_dir, f'{unique_id}_audio.aac')
        extract_audio_command = [
            'ffmpeg', '-y', '-i', input_file,
            '-vn', '-acodec', 'copy', audio_file
        ]
        
        self.console_text.config(state=tk.NORMAL)
        self.console_text.insert(tk.END, f"音声を抽出中...\n")
        self.console_text.config(state=tk.DISABLED)
        self.write_log("VR映像から音声抽出開始")
        
        try:
            subprocess.run(extract_audio_command, check=True, creationflags=subprocess.CREATE_NO_WINDOW)
            self.write_log("音声抽出完了")
        except subprocess.CalledProcessError:
            self.write_log("音声抽出失敗または音声トラックなし")
            audio_file = None
        
        # 2. 中央領域を抽出
        self.console_text.config(state=tk.NORMAL)
        self.console_text.insert(tk.END, f"VR簡易モード: 中央領域のみ抽出\n")
        self.console_text.config(state=tk.DISABLED)
        self.write_log("VR簡易モード: 中央領域抽出開始")
        
        center_file = os.path.join(self.output_dir, f'{unique_id}_center.mp4')
        self.apply_vr_undistortion(input_file, center_file, unique_id)
        
        parts = ['center']
        
        return parts, audio_file

    def merge_vr_video(self, unique_id, parts, output_file, audio_file):
        """VR処理済み中央領域を元動画に合成して音声を追加"""
        
        # LADA処理済みの中央領域ファイルを探す
        center_processed = None
        for file in os.listdir(self.output_dir):
            if file.startswith(f'{unique_id}_center') and 'lada' in file and file.endswith('.mp4'):
                center_processed = os.path.join(self.output_dir, file)
                break
        
        if not center_processed or not os.path.exists(center_processed):
            self.write_log("エラー: LADA処理済みファイルが見つかりません")
            raise Exception("LADA処理済みファイルが見つかりません")
        
        # 中央領域を元動画に合成
        temp_video = os.path.join(self.output_dir, f'{unique_id}_temp_composited.mp4')
        self.apply_vr_distortion(center_processed, temp_video, unique_id)
        
        # 音声を合成
        if audio_file and os.path.exists(audio_file):
            final_merge_cmd = [
                'ffmpeg', '-y', '-i', temp_video, '-i', audio_file,
                '-c:v', 'copy', '-c:a', 'aac', '-b:a', '192k',
                '-shortest', output_file
            ]
            subprocess.run(final_merge_cmd, check=True, creationflags=subprocess.CREATE_NO_WINDOW)
            self.write_log("音声合成完了")
            
            # 音声ファイル削除
            os.remove(audio_file)
            self.write_log(f"音声ファイル削除: {audio_file}")
        else:
            # 音声がない場合はそのまま移動
            os.rename(temp_video, output_file)
        
        # 一時ファイルを削除
        if os.path.exists(temp_video):
            try:
                os.remove(temp_video)
            except:
                pass
        
        # LADA処理済みファイルを削除
        if os.path.exists(center_processed):
            os.remove(center_processed)
            self.write_log(f"LADA処理済みファイル削除: {os.path.basename(center_processed)}")
        
        # 中央抽出ファイルを削除
        center_file = os.path.join(self.output_dir, f'{unique_id}_center.mp4')
        if os.path.exists(center_file):
            os.remove(center_file)
            self.write_log(f"中央抽出ファイル削除: center.mp4")
        
        self.write_log("VR合成処理完了")

    def abort_processing(self):
        """処理を中断し、LADAプロセスをKILLしてバッチループも中止する"""
        if not (hasattr(self, 'is_running') and self.is_running) and \
           not (hasattr(self, 'is_batch_processing') and self.is_batch_processing):
            messagebox.showinfo("情報", "現在、処理は実行されていません。")
            return
        
        if not messagebox.askyesno("確認", "現在実行中の処理を中断しますか?"):
            return
        
        # 1. バッチ処理とメイン実行フラグを即座にFalseに設定してループを停止
        self.is_batch_processing = False
        self.is_running = False
        self.buffer_running = False
        
        # 2. LADAプロセス(PowerShell)を強制終了
        if self.process and self.process.poll() is None:
            try:
                self.process.kill()
                self.process.wait(timeout=3)  # 最大3秒待機
                self.write_log("LADAプロセスを強制終了しました")
            except subprocess.TimeoutExpired:
                self.write_log("LADAプロセス終了タイムアウト")
            except Exception as e:
                self.write_log(f"LADAプロセス終了エラー: {e}")
            finally:
                self.process = None
        
        # 3. FFMPEGプロセスも確実に終了させる
        # output_dirから実行中の可能性のあるファイルに対応するFFMPEGプロセスを検索して終了
        try:
            import psutil
            for proc in psutil.process_iter(['name', 'cmdline']):
                try:
                    if proc.info['name'] and 'ffmpeg' in proc.info['name'].lower():
                        # コマンドラインにoutput_dirが含まれる場合のみ終了
                        if proc.info['cmdline'] and any(self.output_dir in arg for arg in proc.info['cmdline']):
                            proc.kill()
                            self.write_log(f"FFMPEGプロセスを強制終了しました: PID {proc.pid}")
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    pass
        except ImportError:
            # psutilがない場合は警告のみ
            self.write_log("警告: psutilが利用できないため、FFMPEGプロセスの完全な終了を保証できません")
        except Exception as e:
            self.write_log(f"FFMPEGプロセス検索エラー: {e}")
        
        # 4. UI要素を元に戻す
        self.start_button.config(state=tk.NORMAL, text="処理開始 (単一)")
        self.batch_button.config(state=tk.NORMAL)
        self.queue_add_button.config(state=tk.NORMAL)
        self.queue_view_button.config(state=tk.NORMAL)
        self.root.bind('<Control-e>', self.add_to_queue)
        
        # 5. ステータス更新
        self.status_label.config(text="処理を中断しました", fg="red")
        self.batch_count_label.config(text="")
        self.console_text.config(state=tk.NORMAL)
        self.console_text.insert(tk.END, "処理を中断しました。\n")
        self.console_text.config(state=tk.DISABLED)
        
        self.write_log("処理を中断しました")
        messagebox.showinfo("中断完了", "処理を中断しました。")
        
    def add_to_queue(self, event=None):
        if self.is_batch_processing:
            messagebox.showwarning("警告", "一括処理中はキューに追加できません。")
            return
        input_file = self.file_path_entry.get()
        if not input_file or not os.path.exists(input_file):
            messagebox.showerror("エラー", "有効な動画ファイルを選択してください。")
            return
        
        with self.cap_lock:
            cap_temp = cv2.VideoCapture(input_file)
            if not cap_temp.isOpened():
                messagebox.showerror("エラー", "動画ファイルを開けませんでした。")
                cap_temp.release()
                return
            total_frames = int(cap_temp.get(cv2.CAP_PROP_FRAME_COUNT))
            fps = cap_temp.get(cv2.CAP_PROP_FPS) or 30.0
            cap_temp.release()
        
        queue_entry = {
            'video_path': input_file,
            'model': self.model_var.get(),
            'tvai': self.tvai_var.get(),
            'quality': int(self.quality_var.get()),
            'start_frame': self.start_frame,
            'end_frame': min(self.end_frame, total_frames),
            'ffmpeg_option': self.ffmpeg_option_var.get(),
            'save_trimmed': self.save_trimmed_video_var.get(),
            'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            'fps': fps,
            'crf_value': int(self.crf_var.get()),
            'vr_processing': self.vr_processing_var.get(),
            'vr_simple_mode': self.vr_simple_mode_var.get()  # 追加
        }
        
        self.processing_queue.append(queue_entry)
        self.save_queue()
        self.write_log(f"キューに追加: {os.path.basename(input_file)}")
        if not self.suppress_queue_message_var.get():
            messagebox.showinfo("追加完了", f"キューに {os.path.basename(input_file)} を追加しました。\n総キュー数: {len(self.processing_queue)}")

    def open_queue_window(self):
        if not self.processing_queue:
            self.write_log("キュー確認: キューは空です")
            messagebox.showinfo("情報", "キューは空です。")
            return
        
        queue_window = tk.Toplevel(self.root)
        queue_window.title("処理キュー確認")
        queue_window.geometry("900x400")
        
        list_frame = tk.Frame(queue_window)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.queue_listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, font=("MS Gothic", 10), width=100)
        self.queue_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.queue_listbox.yview)
        
        self.update_queue_listbox()
        
        queue_status_label = tk.Label(queue_window, text=f"現在のキュー数: {len(self.processing_queue)}", fg="blue")
        queue_status_label.pack(pady=(5, 10))
        
        btn_frame = tk.Frame(queue_window)
        btn_frame.pack(fill=tk.X, padx=10, pady=5)
        
        up_btn = tk.Button(btn_frame, text="↑ 上へ", command=lambda: self.move_queue_item(-1))
        up_btn.pack(side=tk.LEFT, padx=5)
        
        down_btn = tk.Button(btn_frame, text="↓ 下へ", command=lambda: self.move_queue_item(1))
        down_btn.pack(side=tk.LEFT, padx=5)
        
        delete_btn = tk.Button(
            btn_frame, 
            text="削除", 
            command=lambda: self.delete_queue_item(
                queue_window=queue_window, 
                queue_status_label=queue_status_label
            )
        )
        delete_btn.pack(side=tk.LEFT, padx=5)
        
        clear_all_btn = tk.Button(btn_frame, text="すべて削除", command=lambda: self.clear_all_queue(queue_window, queue_status_label))
        clear_all_btn.pack(side=tk.LEFT, padx=5)
        
        close_btn = tk.Button(btn_frame, text="閉じる", command=queue_window.destroy)
        close_btn.pack(side=tk.RIGHT, padx=5)
        
        queue_window.lift()

    def clear_all_queue(self, queue_window, queue_status_label):
        if not self.processing_queue:
            queue_status_label.config(text="キューはすでに空です。", fg="blue")
            queue_window.lift()
            return
        
        self.processing_queue.clear()
        self.save_queue()
        self.update_queue_listbox()
        queue_status_label.config(text="キューをすべて削除しました。現在のキュー数: 0", fg="blue")
        queue_window.lift()
        self.write_log("キューをすべて削除しました。")

    def update_queue_listbox(self):
        self.queue_listbox.delete(0, tk.END)
        ffmpeg_display_map = {
            'copy': '高速',
            'copy_genpts': 'タイムスタンプ修正',
            're_encode': '再エンコード (NVENC)'
        }
        for i, entry in enumerate(self.processing_queue):
            try:
                filename = os.path.basename(entry['video_path'])
                model = entry['model']
                tvai = entry['tvai']
                quality = entry['quality']
                fps = entry.get('fps', 30.0)
                start_time = self.format_time(entry['start_frame'] / fps if fps > 0 else 0)
                end_time = self.format_time(entry['end_frame'] / fps if fps > 0 else 0)
                ffmpeg_option = ffmpeg_display_map.get(entry['ffmpeg_option'], entry['ffmpeg_option'])
                save_trimmed = '保存する' if entry['save_trimmed'] else '保存しない'
                crf_value = entry.get('crf_value', 19)
                vr_mode = 'VR' if entry.get('vr_processing', False) else '2D'
                simple_mode = '簡易' if entry.get('vr_simple_mode', False) else '通常'  # 追加
                display_text = (f"{i+1}. {filename}, Model:{model}, TVAI:{tvai}, Quality:{quality}, "
                               f"Range:{start_time}-{end_time}, FFmpeg:{ffmpeg_option}, CRF:{crf_value}, "
                               f"SaveTrim:{save_trimmed}, Mode:{vr_mode}, VRMode:{simple_mode}")  # 修正
                self.queue_listbox.insert(tk.END, display_text)
            except Exception as e:
                self.queue_listbox.insert(tk.END, f"{i+1}. 表示エラー: {e}")

    def move_queue_item(self, direction):
        sel = self.queue_listbox.curselection()
        if not sel:
            messagebox.showwarning("警告", "項目を選択してください。")
            return
        idx = sel[0]
        new_idx = idx + direction
        if 0 <= new_idx < len(self.processing_queue):
            self.processing_queue[idx], self.processing_queue[new_idx] = self.processing_queue[new_idx], self.processing_queue[idx]
            self.save_queue()
            self.update_queue_listbox()
            self.queue_listbox.selection_set(new_idx)

    def delete_queue_item(self, queue_window, queue_status_label):
        sel = self.queue_listbox.curselection()
        if not sel:
            queue_status_label.config(text="削除する項目を選択してください。", fg="red")
            queue_window.lift()
            return

        deleted_index = sel[0]
        items_deleted = 0

        if 0 <= deleted_index < len(self.processing_queue):
            del self.processing_queue[deleted_index]
            self.save_queue()
            self.update_queue_listbox()
            items_deleted = 1

        queue_status_label.config(
            text=f"{items_deleted}件の項目を削除しました。現在のキュー数: {len(self.processing_queue)}",
            fg="blue" if items_deleted > 0 else "red"
        )
        queue_window.lift()
        self.write_log(f"キューから {items_deleted}件の項目を削除しました。")

    def save_config_callback(self, *args):
        self.save_config()

    def load_config(self):
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r') as f:
                    lines = f.readlines()
                    for line in lines:
                        line = line.strip()
                        if line.startswith("model="):
                            model_value = line.split("=")[1]
                            if model_value in ["1", "2", "3"]:
                                self.cli_options["model_choice"] = model_value
                                self.model_var.set(model_value)
                        elif line.startswith("tvai="):
                            tvai_value = line.split("=")[1]
                            if tvai_value in ["1", "2"]:
                                self.cli_options["tvai_choice"] = tvai_value
                                self.tvai_var.set(tvai_value)
                        elif line.startswith("quality="):
                            quality = line.split("=")[1]
                            if quality.isdigit() and 5 <= int(quality) <= 30:
                                self.cli_options["quality"] = quality
                                self.quality_var.set(quality)
                            else:
                                self.write_log(f"無効な品質値: {quality}、デフォルト15を使用")
                                self.cli_options["quality"] = "15"
                                self.quality_var.set("15")
                        elif line.startswith("crf="):
                            crf = line.split("=")[1]
                            if crf.isdigit() and 5 <= int(crf) <= 30:
                                self.cli_options["crf_value"] = crf
                                self.crf_var.set(crf)
                            else:
                                self.write_log(f"無効なCRF値: {crf}、デフォルト19を使用")
                                self.cli_options["crf_value"] = "19"
                                self.crf_var.set("19")
            except Exception as e:
                self.write_log(f"設定ファイルの読み込みに失敗しました: {e}")
                messagebox.showwarning("警告", f"設定ファイルの読み込みに失敗しました: {e}。デフォルト値で続行します。")

    def save_config(self):
        try:
            with open(self.config_file, 'w') as f:
                f.write(f"model={self.model_var.get()}\n")
                f.write(f"tvai={self.tvai_var.get()}\n")
                f.write(f"quality={self.quality_var.get()}\n")
                f.write(f"crf={self.crf_var.get()}\n")
        except Exception as e:
            self.write_log(f"設定ファイルの保存に失敗しました: {e}")
            messagebox.showwarning("警告", f"設定ファイルの保存に失敗しました: {e}。手動で確認してください。")
            
    def write_log(self, message):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_entry = f"{timestamp} {message}\n"
        try:
            with open(self.log_file, 'a', encoding='utf-8') as f:
                f.write(log_entry)
        except Exception as e:
            print(f"ログの書き込みに失敗しました: {e}")
            
    def open_log_file(self):
            """ログファイルをメモ帳で開く"""
            if os.path.exists(self.log_file):
                try:
                    subprocess.Popen(['notepad.exe', self.log_file])
                    self.write_log("ログファイルを開きました")
                except Exception as e:
                    self.write_log(f"ログファイルを開けませんでした: {e}")
                    messagebox.showerror("エラー", f"ログファイルを開けませんでした: {e}")
            else:
                messagebox.showwarning("警告", "ログファイルが存在しません。")

    def browse_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("動画ファイル", "*.mp4 *.avi *.mkv *.mov")]
        )
        if file_path:
            self.file_path_entry.delete(0, tk.END)
            self.file_path_entry.insert(0, file_path)
            self.load_video(file_path)

    def drop_file(self, event):
        file_paths_str = event.data
        self.write_log(f"D&D raw data: {repr(file_paths_str)}")
        
        file_paths = []
        valid_extensions = ('.mp4', '.avi', '.mkv', '.mov', '.ts', '.wmv', '.flv')
        
        import re
        
        brace_pattern = r'\{([^}]+)\}'
        braced_paths = re.findall(brace_pattern, file_paths_str)
        
        temp_str = file_paths_str
        placeholder_map = {}
        for i, path in enumerate(braced_paths):
            placeholder = f"__PLACEHOLDER_{i}__"
            temp_str = temp_str.replace(f"{{{path}}}", placeholder)
            placeholder_map[placeholder] = path
        
        ideographic_space = '\u3000'
        parts = re.split(r'[\s\u3000]+', temp_str)
        
        potential_paths = []
        for part in parts:
            if part.strip():
                if part.startswith("__PLACEHOLDER_"):
                    original_path = placeholder_map[part]
                    potential_paths.append(original_path)
                else:
                    potential_paths.append(part)
        
        i = 0
        while i < len(potential_paths):
            current_path = potential_paths[i]
            
            if os.path.exists(current_path) and current_path.lower().endswith(valid_extensions):
                file_paths.append(current_path)
                self.write_log(f"Valid path found: {current_path}")
                i += 1
                continue
            
            found = False
            for j in range(i + 1, len(potential_paths) + 1):
                combined_path = ideographic_space.join(potential_paths[i:j])
                if os.path.exists(combined_path) and combined_path.lower().endswith(valid_extensions):
                    file_paths.append(combined_path)
                    self.write_log(f"Valid path found (reconstructed): {combined_path}")
                    i = j
                    found = True
                    break
            
            if not found:
                i += 1
        
        if not file_paths:
            self.write_log("D&Dエラー: 有効なファイルが見つかりません")
            messagebox.showerror("エラー", "有効な動画ファイルがドロップされませんでした。")
            return
        
        self.write_log(f"Final parsed file paths: {file_paths}")
        
        if len(file_paths) == 1:
            file_path = file_paths[0]
            self.write_log(f"単一ファイル処理: {file_path}")
            self.file_path_entry.delete(0, tk.END)
            self.file_path_entry.insert(0, file_path)
            self.load_video(file_path)
            self.current_frame = 0
            self.on_progress_update()
            self.reset_points()
            self.write_log(f"単一ファイルD&D: {os.path.basename(file_path)} をプレビューにロード")
            return
        
        if not messagebox.askyesno("確認", f"{len(file_paths)}個のファイルが指定されました。キューに登録しますか?"):
            self.write_log("D&Dキャンセル: ユーザーがキュー登録を拒否しました")
            return
        
        added_files = 0
        for file_path in file_paths:
            self.write_log(f"処理対象ファイル: {file_path}")
            
            with self.cap_lock:
                cap_temp = cv2.VideoCapture(file_path)
                if not cap_temp.isOpened():
                    self.write_log(f"D&Dエラー: 動画ファイルを開けませんでした: {file_path}")
                    cap_temp.release()
                    continue
                total_frames = int(cap_temp.get(cv2.CAP_PROP_FRAME_COUNT))
                fps = cap_temp.get(cv2.CAP_PROP_FPS) or 30.0
                cap_temp.release()
            
            queue_entry = {
                'video_path': file_path,
                'model': self.model_var.get(),
                'tvai': self.tvai_var.get(),
                'quality': int(self.quality_var.get()),
                'start_frame': 0,
                'end_frame': total_frames,
                'ffmpeg_option': self.ffmpeg_option_var.get(),
                'save_trimmed': self.save_trimmed_video_var.get(),
                'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                'fps': fps,
                'crf_value': int(self.crf_var.get()),
                'vr_processing': self.vr_processing_var.get(),
                'vr_simple_mode': self.vr_simple_mode_var.get()
            }
            
            self.processing_queue.append(queue_entry)
            added_files += 1
            self.write_log(f"キューに追加: {os.path.basename(file_path)}")
        
        if added_files > 0:
            self.save_queue()
            if not self.suppress_queue_message_var.get():
                messagebox.showinfo("追加完了", f"{added_files}件のファイルをキューに追加しました。\n総キュー数: {len(self.processing_queue)}")
            
            last_file_path = file_paths[-1]
            self.file_path_entry.delete(0, tk.END)
            self.file_path_entry.insert(0, last_file_path)
            self.load_video(last_file_path)
            self.current_frame = 0
            self.on_progress_update()
            self.reset_points()
        else:
            self.write_log("D&Dエラー: 有効な動画ファイルがありません")
            messagebox.showerror("エラー", "有効な動画ファイルがドロップされませんでした。")

    def start_batch_processing(self, control_frame):
        if not self.processing_queue:
            messagebox.showinfo("情報", "キューは空です。")
            return
        if hasattr(self, 'is_running') and self.is_running:
            messagebox.showwarning("警告", "処理中です。完了後に実行してください。")
            return
        
        self.queue_add_button.config(state=tk.DISABLED)
        self.queue_view_button.config(state=tk.DISABLED)
        self.start_button.config(state=tk.DISABLED)
        self.batch_button.config(state=tk.DISABLED)
        self.root.bind('<Control-e>', lambda e: None)

        self.is_batch_processing = True
        self.is_running = True
        self.batch_thread = threading.Thread(target=self.batch_process_main)
        self.batch_thread.daemon = True
        self.batch_thread.start()

    def batch_process_main(self):
        original_batch_count = len(self.processing_queue)
        processed_items = 0

        if original_batch_count == 0:
            self.root.after(0, lambda: self.status_label.config(text="キューは空です", fg="blue"))
            self.is_batch_processing = False
            self.is_running = False
            self.root.after(0, lambda: self.batch_count_label.config(text=""))
            self.root.after(0, lambda: self.queue_add_button.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.queue_view_button.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.start_button.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.batch_button.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.root.bind('<Control-e>', self.add_to_queue))
            return
        
        self.root.after(0, lambda: self.batch_count_label.config(
            text=f"バッチ処理中: 1/{original_batch_count}",
            fg="red"
        ))
        
        while self.processing_queue and self.is_batch_processing:
            entry = self.processing_queue[0]
            processed_items += 1
            current_count = processed_items
            
            self.root.after(0, lambda idx=current_count: self.batch_count_label.config(
                text=f"バッチ処理中: {idx}/{original_batch_count}",
                fg="red"
            ))
            
            self.root.after(0, lambda: self.status_label.config(text=f"処理中: {os.path.basename(entry['video_path'])}"))
            self.write_log(f"処理中: {os.path.basename(entry['video_path'])}")
            self.console_text.config(state=tk.NORMAL)
            self.console_text.insert(tk.END, f"処理中: {os.path.basename(entry['video_path'])}\n")
            self.console_text.config(state=tk.DISABLED)
            self.root.update()
            
            processing_success = False  # 処理成功フラグを追加
            
            try:
                self.processing_main(
                    entry['video_path'], 
                    entry['start_frame'] / entry['fps'], 
                    entry['end_frame'] / entry['fps'],
                    entry.get('vr_simple_mode', False)
                )
                
                processing_success = True  # 処理が正常完了
                
            except Exception as e:
                self.write_log(f"処理中にエラー発生: {os.path.basename(entry['video_path'])}, エラー: {e}")
                self.root.after(0, lambda: self.status_label.config(text=f"エラー中断: {os.path.basename(entry['video_path'])}", fg="red"))
                self.root.after(0, lambda: messagebox.showerror("処理エラー", f"{os.path.basename(entry['video_path'])} の処理中にエラーが発生し、バッチ処理を中断しました。\n未処理の項目はキューに残っています。"))
                break
            
            # 処理が正常完了した場合のみキューから削除
            if processing_success and self.is_batch_processing:
                del self.processing_queue[0]
                self.save_queue()

                self.root.after(0, lambda: self.status_label.config(text=f"完了: {os.path.basename(entry['video_path'])}"))
                self.write_log(f"完了: {os.path.basename(entry['video_path'])}")
                self.console_text.config(state=tk.NORMAL)
                self.console_text.insert(tk.END, f"完了: {os.path.basename(entry['video_path'])}\n")
                self.console_text.config(state=tk.DISABLED)
                self.root.update()
            elif not self.is_batch_processing:
                # 中断された場合はキューから削除せず、ループを抜ける
                self.write_log(f"中断により未完了: {os.path.basename(entry['video_path'])}")
                break

        self.root.after(0, lambda: self.status_label.config(text="バッチ処理完了"))
        if self.show_completion_dialog_var.get():
            self.root.after(0, lambda: messagebox.showinfo("バッチ完了", f"バッチ処理が完了しました!\n総処理数: {original_batch_count} (処理済み: {processed_items})"))
        
        self.is_batch_processing = False
        self.is_running = False
        self.root.after(0, lambda: self.batch_count_label.config(text=""))
        self.root.after(0, lambda: self.queue_add_button.config(state=tk.NORMAL))
        self.root.after(0, lambda: self.queue_view_button.config(state=tk.NORMAL))
        self.root.after(0, lambda: self.start_button.config(state=tk.NORMAL))
        self.root.after(0, lambda: self.batch_button.config(state=tk.NORMAL))
        self.root.after(0, lambda: self.root.bind('<Control-e>', self.add_to_queue))
        self.write_log("バッチ処理全体完了")

    def start_processing(self):
        if not self.validate_inputs():
            return
        
        self.is_batch_processing = False
        self.is_running = True
        
        self.batch_button.config(state=tk.DISABLED)
        self.start_button.config(state=tk.DISABLED, text="処理中...")
        
        self.status_label.config(text="動画を切り出し中...", fg="orange")
        self.console_text.config(state=tk.NORMAL)
        self.console_text.delete('1.0', tk.END)
        self.console_text.insert(tk.END, "処理を開始します...\n")
        self.console_text.config(state=tk.DISABLED)
        self.root.update()
        
        input_file = self.file_path_entry.get()
        start_time_sec = self.start_frame / self.video_fps
        end_time_sec = self.end_frame / self.video_fps
        input_filename = os.path.basename(input_file)
        self.write_log(f"LADA処理を開始しました {input_filename}")
        
        self.processing_thread = threading.Thread(target=self.processing_main, args=(input_file, start_time_sec, end_time_sec))
        self.processing_thread.daemon = True
        self.processing_thread.start()

    def validate_inputs(self):
        if not self.file_path_entry.get():
            messagebox.showerror("エラー", "動画ファイルを選択してください。")
            return False
        if not os.path.exists(self.file_path_entry.get()):
            messagebox.showerror("エラー", "指定された動画ファイルが存在しません。")
            return False
        if self.start_frame >= self.end_frame:
            messagebox.showerror("エラー", "開始フレームが終了フレーム以上です。")
            return False
        return True

    def processing_main(self, input_file, start_time_sec, end_time_sec, vr_simple_mode=None):
        # vr_simple_modeがNoneの場合は現在のGUI設定を使用(単一処理用)
        if vr_simple_mode is None:
            vr_simple_mode = self.vr_simple_mode_var.get()
        
        if self.is_batch_processing:
            self.write_log(f"処理前リソースチェック: {os.path.basename(input_file)}")
        
        time.sleep(1)

        unique_id = uuid.uuid4().hex
        input_ext = os.path.splitext(input_file)[1]
        trimmed_base_name = f"trimmed_{unique_id}"
        trimmed_file_ext = '.mp4' if self.ffmpeg_option_var.get() == "re_encode" else input_ext
        trimmed_file_path = os.path.join(self.output_dir, f"{trimmed_base_name}{trimmed_file_ext}")
        processed_file_path = None
        
        try:
            option = self.ffmpeg_option_var.get()
            start_time_str = self.format_time(start_time_sec)
            end_time_str = self.format_time(end_time_sec)
            
            if option == "re_encode":
                crf_value = self.crf_var.get()
                ffmpeg_command = [
                    "ffmpeg", "-y", "-ss", start_time_str, "-to", end_time_str, "-i", input_file,
                    "-c:v", "h264_nvenc", "-c:a", "aac", "-preset", "fast", "-rc", "vbr_hq", "-cq", crf_value,
                    trimmed_file_path
                ]
            elif option == "copy":
                ffmpeg_command = [
                    "ffmpeg", "-y", "-ss", start_time_str, "-to", end_time_str, "-i", input_file,
                    "-c", "copy", trimmed_file_path
                ]
            elif option == "copy_genpts":
                ffmpeg_command = [
                    "ffmpeg", "-y", "-ss", start_time_str, "-to", end_time_str, "-i", input_file,
                    "-c", "copy", "-fflags", "+genpts", trimmed_file_path
                ]
            else:
                raise ValueError("無効なFFmpegオプションです。")
            
            self.console_text.config(state=tk.NORMAL)
            self.console_text.insert(tk.END, f"動画を切り出し中...\n実行コマンド: {' '.join(ffmpeg_command)}\n")
            self.console_text.config(state=tk.DISABLED)
            self.write_log(f"動画を切り出し中...\n実行コマンド: {' '.join(ffmpeg_command)}")
            
            process = subprocess.run(ffmpeg_command, check=True, creationflags=subprocess.CREATE_NO_WINDOW)

            self.console_text.config(state=tk.NORMAL)
            self.console_text.insert(tk.END, "動画の切り出しが完了しました。\n")
            self.console_text.config(state=tk.DISABLED)
            self.write_log("動画の切り出しが完了しました。")

            self.root.update()
            
            self.status_label.config(text="切り出し完了。モザイク除去を開始します...", fg="green")
            self.save_config()

            # VR処理の判定
            is_vr_mode = self.vr_processing_var.get()
            
            if is_vr_mode:
                # VR処理モード（簡易専用）
                self.console_text.config(state=tk.NORMAL)
                self.console_text.insert(tk.END, "VR処理モードで実行します\n")
                self.console_text.config(state=tk.DISABLED)
                self.write_log("VR処理モード開始")
                
                # 1. VR映像を処理（中央領域抽出、音声抽出）
                parts, audio_file = self.split_vr_video(trimmed_file_path, unique_id)
                
                # 2. 中央領域をLADA処理（1回のみ）
                center_file = os.path.join(self.output_dir, f'{unique_id}_center.mp4')
                
                if not os.path.exists(center_file):
                    raise Exception("中央領域ファイルが見つかりません")
                
                self.console_text.config(state=tk.NORMAL)
                self.console_text.insert(tk.END, f"VR中央領域を処理中...\n")
                self.console_text.config(state=tk.DISABLED)
                self.status_label.config(text="VR中央領域を処理中")
                self.write_log("VR中央領域LADA処理開始")
                
                if self.ps_script_path:
                    ps_command = ["powershell.exe", "-ExecutionPolicy", "Bypass", "-File", self.ps_script_path]
                    input_data = f"{center_file}\n{self.model_var.get()}\n{self.tvai_var.get()}\n{self.quality_var.get()}\n"

                    self.process = subprocess.Popen(
                        ps_command,
                        stdin=subprocess.PIPE,
                        stdout=subprocess.PIPE,
                        stderr=subprocess.STDOUT,
                        text=True,
                        bufsize=1,
                        creationflags=subprocess.CREATE_NO_WINDOW
                    )
                    
                    self.process.stdin.write(input_data)
                    self.process.stdin.flush()
                    self.process.stdin.close()

                    for line in iter(self.process.stdout.readline, ''):
                        self.console_text.config(state=tk.NORMAL)
                        self.console_text.insert(tk.END, line)
                        self.console_text.see(tk.END)
                        self.console_text.config(state=tk.DISABLED)
                        if not line.strip().startswith("Processing frames:"):
                            self.write_log(line.strip())
                        self.root.update()

                    self.process.stdout.close()
                    self.process.wait()

                    if self.process.returncode != 0:
                        raise Exception("VR中央領域の処理に失敗しました")
                
                # 3. 処理済みファイルを結合して音声合成
                base_name = os.path.splitext(os.path.basename(input_file))[0]
                start_time_str_renamed = self.format_time(start_time_sec).replace(':', '')
                end_time_str_renamed = self.format_time(end_time_sec).replace(':', '')
                timestamp_tag = f"{start_time_str_renamed}-{end_time_str_renamed}"
                cli_options_tag = f"model{self.model_var.get()}_tvai{self.tvai_var.get()}_quality{self.quality_var.get()}"

                saved_processed_name = f"{base_name}_{timestamp_tag}_{cli_options_tag}_VR_unmosaiced.mp4"
                saved_processed_path = os.path.join(self.output_dir, saved_processed_name)
                saved_processed_path = self.generate_unique_filepath(saved_processed_path)

                self.merge_vr_video(unique_id, parts, saved_processed_path, audio_file)

                self.status_label.config(text=f"VR処理完了: {os.path.basename(saved_processed_path)}", fg="blue")
                self.write_log(f"VR処理完了: {os.path.basename(saved_processed_path)}")

                if not self.is_batch_processing and self.show_completion_dialog_var.get():
                    messagebox.showinfo("完了", f"動画のモザイク除去が完了しました! (モード: VR)\n\nファイル名: " + os.path.basename(saved_processed_path))
                self.write_log("動画のモザイク除去が完了しました!")
                
            else:
                # 通常の2D処理モード
                if self.ps_script_path:
                    ps_command = ["powershell.exe", "-ExecutionPolicy", "Bypass", "-File", self.ps_script_path]
                    input_data = f"{trimmed_file_path}\n{self.model_var.get()}\n{self.tvai_var.get()}\n{self.quality_var.get()}\n"

                    self.process = subprocess.Popen(
                        ps_command,
                        stdin=subprocess.PIPE,
                        stdout=subprocess.PIPE,
                        stderr=subprocess.STDOUT,
                        text=True,
                        bufsize=1,
                        creationflags=subprocess.CREATE_NO_WINDOW
                    )
                    
                    self.process.stdin.write(input_data)
                    self.process.stdin.flush()
                    self.process.stdin.close()

                    for line in iter(self.process.stdout.readline, ''):
                        self.console_text.config(state=tk.NORMAL)
                        self.console_text.insert(tk.END, line)
                        self.console_text.see(tk.END)
                        self.console_text.config(state=tk.DISABLED)
                        if not line.strip().startswith("Processing frames:"):
                            self.write_log(line.strip())
                        self.root.update()

                    self.process.stdout.close()
                    self.process.wait()

                    if self.process.returncode != 0:
                        self.status_label.config(text="PowerShellスクリプト実行失敗", fg="red")
                        self.console_text.config(state=tk.NORMAL)
                        self.console_text.insert(tk.END, "PowerShellスクリプトの実行に失敗しました。\n")
                        self.console_text.config(state=tk.DISABLED)
                        self.write_log("PowerShellスクリプトの実行に失敗しました。")
                        messagebox.showerror("実行エラー", "PowerShellスクリプトの実行に失敗しました。詳細はコンソールログをご確認ください。")
                    else:
                        for file_name in os.listdir(self.output_dir):
                            if trimmed_base_name in file_name and file_name != os.path.basename(trimmed_file_path):
                                processed_file_path = os.path.join(self.output_dir, file_name)
                                break
                        
                        if not processed_file_path or not os.path.exists(processed_file_path):
                            self.status_label.config(text="処理済み動画ファイルが見つかりません。", fg="red")
                            self.console_text.config(state=tk.NORMAL)
                            self.console_text.insert(tk.END, "エラー: LADAの出力ファイルが見つかりませんでした。\n")
                            self.console_text.config(state=tk.DISABLED)
                            self.write_log("エラー: LADAの出力ファイルが見つかりませんでした。")
                            return
                        
                        base_name = os.path.splitext(os.path.basename(input_file))[0]
                        start_time_str_renamed = self.format_time(start_time_sec).replace(':', '')
                        end_time_str_renamed = self.format_time(end_time_sec).replace(':', '')
                        timestamp_tag = f"{start_time_str_renamed}-{end_time_str_renamed}"
                        cli_options_tag = f"model{self.model_var.get()}_tvai{self.tvai_var.get()}_quality{self.quality_var.get()}"
                        
                        saved_processed_name = f"{base_name}_{timestamp_tag}_{cli_options_tag}_unmosaiced.mp4"
                        saved_processed_path = os.path.join(self.output_dir, saved_processed_name)
                        saved_processed_path = self.generate_unique_filepath(saved_processed_path)
                        saved_processed_name = os.path.basename(saved_processed_path)
                        try:
                            os.rename(processed_file_path, saved_processed_path)
                            self.status_label.config(text=f"処理済み動画を保存しました: {saved_processed_name}", fg="blue")
                            self.write_log(f"処理済み動画を保存しました: {saved_processed_name}")
                        except Exception as e:
                            self.status_label.config(text=f"ファイル名の変更に失敗しました: {e}", fg="red")
                            self.console_text.config(state=tk.NORMAL)
                            self.console_text.insert(tk.END, f"ファイル名の変更に失敗しました: {e}\n")
                            self.console_text.config(state=tk.DISABLED)
                            self.write_log(f"ファイル名の変更に失敗しました: {e}")
                
                # 切り出し動画の保存処理
                if self.save_trimmed_video_var.get():
                    base_name = os.path.splitext(os.path.basename(input_file))[0]
                    start_time_str_renamed = self.format_time(start_time_sec).replace(':', '')
                    end_time_str_renamed = self.format_time(end_time_sec).replace(':', '')
                    timestamp_tag = f"{start_time_str_renamed}-{end_time_str_renamed}"
                    saved_trimmed_name = f"{base_name}_{timestamp_tag}_trimmed{trimmed_file_ext}"
                    saved_trimmed_path = os.path.join(self.output_dir, saved_trimmed_name)
                    saved_trimmed_path = self.generate_unique_filepath(saved_trimmed_path)
                    saved_trimmed_name = os.path.basename(saved_trimmed_path)
                    try:
                        os.rename(trimmed_file_path, saved_trimmed_path)
                        self.status_label.config(text=f"切り出し動画を保存しました: {saved_trimmed_name}", fg="blue")
                        self.write_log(f"切り出し動画を保存しました: {saved_trimmed_name}")
                    except Exception as e:
                        self.status_label.config(text=f"切り出し動画が見つからず、保存できませんでした: {e}", fg="red")
                        self.console_text.config(state=tk.NORMAL)
                        self.console_text.insert(tk.END, f"切り出し動画が見つからず、保存できませんでした: {e}\n")
                        self.console_text.config(state=tk.DISABLED)
                        self.write_log(f"切り出し動画が見つからず、保存できませんでした: {e}")
                else:
                    if os.path.exists(trimmed_file_path):
                        os.remove(trimmed_file_path)
                        self.write_log(f"一時ファイル削除: {trimmed_file_path}")
                
                if not self.is_batch_processing and self.show_completion_dialog_var.get():
                    mode_text = "VR" if is_vr_mode else "2D"
                    messagebox.showinfo("完了", f"動画のモザイク除去が完了しました! (モード: {mode_text})\n\nファイル名: " + os.path.basename(saved_processed_path))
                self.write_log("動画のモザイク除去が完了しました!")
            
        except subprocess.CalledProcessError as e:
            self.status_label.config(text="エラーが発生しました", fg="red")
            self.console_text.config(state=tk.NORMAL)
            self.console_text.insert(tk.END, f"コマンド実行に失敗しました。\nエラーコード: {e.returncode}\n")
            self.console_text.config(state=tk.DISABLED)
            self.write_log(f"コマンド実行に失敗しました。エラーコード: {e.returncode}")
        except Exception as e:
            self.status_label.config(text="予期せぬエラーが発生しました", fg="red")
            self.console_text.config(state=tk.NORMAL)
            self.console_text.insert(tk.END, f"エラー: {e}\n")
            self.console_text.config(state=tk.DISABLED)
            self.write_log(f"エラー: {e}")
        finally:
            if not self.is_batch_processing:
                self.start_button.config(state=tk.NORMAL, text="処理開始 (単一)")
                self.batch_button.config(state=tk.NORMAL)
            self.is_running = False
            input_filename = os.path.basename(input_file)
            self.write_log(f"LADA処理を終了しました {input_filename}")
            if 'trimmed_file_path' in locals() and os.path.exists(trimmed_file_path):
                if not self.save_trimmed_video_var.get():
                    os.remove(trimmed_file_path)
                    self.write_log(f"一時ファイル削除: {trimmed_file_path}")
                    
    def merge_vr_video(self, unique_id, parts, output_file, audio_file):
        """VR処理済み中央領域を元動画に合成して音声を追加"""
        
        # LADA処理済みの中央領域ファイルを探す
        center_processed = None
        for file in os.listdir(self.output_dir):
            # より柔軟な検索パターン
            if (f'{unique_id}_center' in file and 
                'lada' in file.lower() and 
                file.endswith('.mp4')):
                center_processed = os.path.join(self.output_dir, file)
                self.write_log(f"LADA処理済みファイル検出: {file}")
                break
        
        if not center_processed or not os.path.exists(center_processed):
            # デバッグ: outputディレクトリ内のファイル一覧を表示
            self.write_log("outputディレクトリ内のファイル一覧:")
            for file in os.listdir(self.output_dir):
                if unique_id in file:
                    self.write_log(f"  - {file}")
            
            self.write_log("エラー: LADA処理済みファイルが見つかりません")
            raise Exception(f"LADA処理済みファイルが見つかりません (unique_id: {unique_id})")
        
        self.console_text.config(state=tk.NORMAL)
        self.console_text.insert(tk.END, f"LADA処理済みファイル: {os.path.basename(center_processed)}\n")
        self.console_text.config(state=tk.DISABLED)
        
        # 中央領域を元動画に合成
        temp_video = os.path.join(self.output_dir, f'{unique_id}_temp_composited.mp4')
        self.apply_vr_distortion(center_processed, temp_video, unique_id)
        
        # 音声を合成
        if audio_file and os.path.exists(audio_file):
            final_merge_cmd = [
                'ffmpeg', '-y', '-i', temp_video, '-i', audio_file,
                '-c:v', 'copy', '-c:a', 'aac', '-b:a', '192k',
                '-shortest', output_file
            ]
            subprocess.run(final_merge_cmd, check=True, creationflags=subprocess.CREATE_NO_WINDOW)
            self.write_log("音声合成完了")
            
            # 音声ファイル削除
            os.remove(audio_file)
            self.write_log(f"音声ファイル削除: {audio_file}")
        else:
            # 音声がない場合はそのまま移動
            if os.path.exists(temp_video):
                os.rename(temp_video, output_file)
        
        # 一時ファイルを削除
        if os.path.exists(temp_video) and temp_video != output_file:
            try:
                os.remove(temp_video)
                self.write_log(f"一時ファイル削除: temp_composited.mp4")
            except:
                pass
        
        # LADA処理済みファイルを削除
        if os.path.exists(center_processed):
            os.remove(center_processed)
            self.write_log(f"LADA処理済みファイル削除: {os.path.basename(center_processed)}")
        
        # 中央抽出ファイルを削除
        center_file = os.path.join(self.output_dir, f'{unique_id}_center.mp4')
        if os.path.exists(center_file):
            os.remove(center_file)
            self.write_log(f"中央抽出ファイル削除: center.mp4")
        
        self.write_log("VR合成処理完了")

    def load_video(self, file_path):
        with self.cap_lock:
            if self.cap:
                self.cap.release()
                self.cap = None
            
            try:
                self.cap = cv2.VideoCapture(file_path)
                if not self.cap.isOpened():
                    messagebox.showerror("エラー", "動画ファイルを開けませんでした。別のファイルを選択してください。")
                    self.cap = None
                    self.video_path = ""
                    return
                self.video_path = file_path
                
                self.video_total_frames = int(self.cap.get(cv2.CAP_PROP_FRAME_COUNT))
                raw_fps = self.cap.get(cv2.CAP_PROP_FPS)
                if raw_fps <= 0 or raw_fps > 120:
                    total_frames = self.video_total_frames
                    duration = self.cap.get(cv2.CAP_PROP_FRAME_COUNT) / self.cap.get(cv2.CAP_PROP_FPS) if self.cap.get(cv2.CAP_PROP_FPS) > 0 else None
                    if duration and duration > 0:
                        self.video_fps = total_frames / duration
                    else:
                        self.video_fps = 30.0
                else:
                    self.video_fps = raw_fps
                
                self.actual_fps = self.video_fps
                
                self.reset_points()
                
                self.cap.set(cv2.CAP_PROP_POS_FRAMES, 0)
                ret, frame = self.cap.read()
                if ret:
                    self.display_frame(frame)
                
                self.paused = True
                self.play_pause_button.config(text="▶ 再生")
                self.on_progress_update()
                self.write_log(f"動画読み込み成功: {file_path}, FPS: {self.video_fps}, 総フレーム: {self.video_total_frames}")
            except Exception as e:
                self.write_log(f"動画読み込みエラー: {e}")
                messagebox.showerror("エラー", f"動画読み込みに失敗しました: {e}")
                if self.cap:
                    self.cap.release()
                self.cap = None

    def toggle_play_pause(self, event=None):
        if not self.video_path or not self.cap or not self.cap.isOpened():
            return
        
        if self.paused:
            self.paused = False
            self.play_pause_button.config(text="|| 一時停止")
            self.buffer_running = True
            self.last_frame_time = time.time()
            self.start_frame_buffer()
            self.update_frame()
        else:
            self.paused = True
            self.buffer_running = False
            self.play_pause_button.config(text="▶ 再生")
            self.clear_frame_queue()
            if self.frame_buffer_thread:
                self.frame_buffer_thread = None

    def start_frame_buffer(self):
        if not self.buffer_running or not self.cap or not self.cap.isOpened():
            return
        if not self.frame_buffer_thread or not self.frame_buffer_thread.is_alive():
            self.buffer_running = True
            self.frame_buffer_thread = threading.Thread(target=self.buffer_frames)
            self.frame_buffer_thread.daemon = True
            self.frame_buffer_thread.start()

    def buffer_frames(self):
        while self.buffer_running and self.cap and self.cap.isOpened():
            with self.cap_lock:
                try:
                    ret, frame = self.cap.read()
                    if ret and not self.frame_queue.full():
                        self.frame_queue.put(frame, block=False)
                    elif not ret:
                        self.buffer_running = False
                        self.root.after(0, self.toggle_play_pause)
                        break
                except Exception as e:
                    self.buffer_running = False
                    self.root.after(0, self.toggle_play_pause)
                    self.write_log(f"フレームバッファエラー: {e}")
                    break
            target_interval = 1.0 / self.actual_fps if self.actual_fps > 0 else 1.0 / 30.0
            time.sleep(target_interval * 0.5)

    def update_frame(self):
        if not self.cap or not self.cap.isOpened() or self.paused or not self.root.winfo_exists():
            return
        
        current_time = time.time()
        target_interval = 1.0 / self.actual_fps if self.actual_fps > 0 else 1.0 / 30.0
        
        if current_time - self.last_frame_time < target_interval:
            remaining_time = target_interval - (current_time - self.last_frame_time)
            self.root.after(max(1, int(remaining_time * 1000)), self.update_frame)
            return
        
        try:
            if not self.frame_queue.empty():
                frame = self.frame_queue.get_nowait()
                with self.cap_lock:
                    self.current_frame = int(self.cap.get(cv2.CAP_PROP_POS_FRAMES))
                
                if frame is not None:
                    self.display_frame(frame)
                    if self.fullscreen_window:
                        self.display_frame_fullscreen(frame)
                    self.update_time_labels()
                    self.on_progress_update()
                    if self.fullscreen_window:
                        self.update_fullscreen_progress()
                    
                    self.last_frame_time = current_time
                
                if self.current_frame >= self.video_total_frames - 1:
                    self.toggle_play_pause()
                    return
                
                self.root.after(max(1, int(target_interval * 1000)), self.update_frame)
            else:
                self.root.after(10, self.update_frame)
        except Exception as e:
            self.write_log(f"フレーム更新エラー: {e}")
            self.root.after(max(1, int(target_interval * 1000)), self.update_frame)

    def clear_frame_queue(self):
        with self.cap_lock:
            while not self.frame_queue.empty():
                try:
                    self.frame_queue.get_nowait()
                except:
                    pass

    def on_progress_update(self):
        if self.video_total_frames > 0:
            width = self.progress_canvas.winfo_width()
            progress_width = (self.current_frame / self.video_total_frames) * width
            self.progress_canvas.coords(self.progress_bar, 0, 0, progress_width, 20)
            
            start_pos = (self.start_frame / self.video_total_frames) * width
            end_pos = (self.end_frame / self.video_total_frames) * width
            self.progress_canvas.coords(self.start_marker, start_pos, 0, start_pos, 20)
            self.progress_canvas.coords(self.end_marker, end_pos, 0, end_pos, 20)

    def on_progress_click(self, event):
        if self.video_total_frames > 0 and self.cap and self.cap.isOpened():
            width = self.progress_canvas.winfo_width()
            click_pos = event.x / width
            new_frame = int(click_pos * self.video_total_frames)
            self.current_frame = new_frame
            self.clear_frame_queue()
            with self.cap_lock:
                try:
                    self.cap.set(cv2.CAP_PROP_POS_FRAMES, new_frame)
                    ret, frame = self.cap.read()
                    if ret:
                        self.display_frame(frame)
                        if self.fullscreen_window:
                            self.display_frame_fullscreen(frame)
                    self.on_progress_update()
                    self.update_time_labels()
                except Exception as e:
                    self.write_log(f"進捗クリックエラー: {e}")

    def move_frame(self, event):
        if not self.cap or not self.cap.isOpened():
            return
        
        steps = 300
        if event.state & 0x0001:
            steps = 30
            
        current_pos = self.current_frame
        new_pos = current_pos
        
        if event.keysym == 'Right':
            new_pos = min(self.video_total_frames, current_pos + steps)
        elif event.keysym == 'Left':
            new_pos = max(0, current_pos - steps)
            
        self.current_frame = new_pos
        self.clear_frame_queue()
        with self.cap_lock:
            try:
                self.cap.set(cv2.CAP_PROP_POS_FRAMES, new_pos)
                ret, frame = self.cap.read()
                if ret:
                    self.display_frame(frame)
                    if self.fullscreen_window:
                        self.display_frame_fullscreen(frame)
                self.on_progress_update()
                self.update_time_labels()
            except Exception as e:
                self.write_log(f"フレーム移動エラー: {e}")

    def move_one_frame_backward(self, event=None):
        if not self.cap or not self.cap.isOpened():
            return
        new_frame = max(0, self.current_frame - 1)
        self.current_frame = new_frame
        self.clear_frame_queue()
        with self.cap_lock:
            try:
                self.cap.set(cv2.CAP_PROP_POS_FRAMES, new_frame)
                ret, frame = self.cap.read()
                if ret:
                    self.display_frame(frame)
                    if self.fullscreen_window:
                        self.display_frame_fullscreen(frame)
                self.on_progress_update()
                self.update_time_labels()
            except Exception as e:
                self.write_log(f"1フレーム戻るエラー: {e}")

    def move_one_frame_forward(self, event=None):
        if not self.cap or not self.cap.isOpened():
            return
        new_frame = min(self.video_total_frames, self.current_frame + 1)
        self.current_frame = new_frame
        self.clear_frame_queue()
        with self.cap_lock:
            try:
                self.cap.set(cv2.CAP_PROP_POS_FRAMES, new_frame)
                ret, frame = self.cap.read()
                if ret:
                    self.display_frame(frame)
                    if self.fullscreen_window:
                        self.display_frame_fullscreen(frame)
                self.on_progress_update()
                self.update_time_labels()
            except Exception as e:
                self.write_log(f"1フレーム進むエラー: {e}")

    def move_one_second_backward(self, event=None):
        if not self.cap or not self.cap.isOpened():
            return
        step_frames = int(self.video_fps)
        new_frame = max(0, self.current_frame - step_frames)
        self.current_frame = new_frame
        self.clear_frame_queue()
        with self.cap_lock:
            try:
                self.cap.set(cv2.CAP_PROP_POS_FRAMES, new_frame)
                ret, frame = self.cap.read()
                if ret:
                    self.display_frame(frame)
                    if self.fullscreen_window:
                        self.display_frame_fullscreen(frame)
                self.on_progress_update()
                self.update_time_labels()
            except Exception as e:
                self.write_log(f"1秒戻るエラー: {e}")

    def move_one_second_forward(self, event=None):
        if not self.cap or not self.cap.isOpened():
            return
        step_frames = int(self.video_fps)
        new_frame = min(self.video_total_frames, self.current_frame + step_frames)
        self.current_frame = new_frame
        self.clear_frame_queue()
        with self.cap_lock:
            try:
                self.cap.set(cv2.CAP_PROP_POS_FRAMES, new_frame)
                ret, frame = self.cap.read()
                if ret:
                    self.display_frame(frame)
                    if self.fullscreen_window:
                        self.display_frame_fullscreen(frame)
                self.on_progress_update()
                self.update_time_labels()
            except Exception as e:
                self.write_log(f"1秒進むエラー: {e}")

    def jump_to_start(self, event=None):
        if not self.cap or not self.cap.isOpened():
            return
        self.current_frame = self.start_frame
        self.clear_frame_queue()
        with self.cap_lock:
            try:
                self.cap.set(cv2.CAP_PROP_POS_FRAMES, self.start_frame)
                ret, frame = self.cap.read()
                if ret:
                    self.display_frame(frame)
                    if self.fullscreen_window:
                        self.display_frame_fullscreen(frame)
                self.on_progress_update()
                self.update_time_labels()
            except Exception as e:
                self.write_log(f"開始点ジャンプエラー: {e}")

    def jump_to_end(self, event=None):
        if not self.cap or not self.cap.isOpened():
            return
        self.current_frame = self.end_frame
        self.clear_frame_queue()
        with self.cap_lock:
            try:
                self.cap.set(cv2.CAP_PROP_POS_FRAMES, self.end_frame)
                ret, frame = self.cap.read()
                if ret:
                    self.display_frame(frame)
                    if self.fullscreen_window:
                        self.display_frame_fullscreen(frame)
                self.on_progress_update()
                self.update_time_labels()
            except Exception as e:
                self.write_log(f"終了点ジャンプエラー: {e}")

    def set_start_point_by_key(self, event=None):
        self.start_frame = self.current_frame
        if self.end_frame < self.start_frame:
            self.end_frame = self.start_frame
        self.update_time_labels()
        self.on_progress_update()

    def set_end_point_by_key(self, event=None):
        self.end_frame = self.current_frame
        if self.start_frame > self.end_frame:
            self.start_frame = self.end_frame
        self.update_time_labels()
        self.on_progress_update()

    def jump_to_percentage(self, percentage):
        if not self.cap or not self.cap.isOpened():
            return
        new_frame = int((percentage / 100) * self.video_total_frames)
        new_frame = min(max(0, new_frame), self.video_total_frames - 1)
        self.current_frame = new_frame
        self.clear_frame_queue()
        with self.cap_lock:
            try:
                self.cap.set(cv2.CAP_PROP_POS_FRAMES, new_frame)
                ret, frame = self.cap.read()
                if ret:
                    self.display_frame(frame)
                    if self.fullscreen_window:
                        self.display_frame_fullscreen(frame)
                self.on_progress_update()
                self.update_time_labels()
                if self.fullscreen_window:
                    self.update_fullscreen_progress()
            except Exception as e:
                self.write_log(f"パーセントジャンプエラー: {e}")

    def on_mouse_wheel(self, event):
        if not self.cap or not self.cap.isOpened():
            return
        step_frames = int(5 * self.video_fps)
        if event.delta > 0:
            new_frame = max(0, self.current_frame - step_frames)
        else:
            new_frame = min(self.video_total_frames, self.current_frame + step_frames)
        self.current_frame = new_frame
        self.clear_frame_queue()
        with self.cap_lock:
            try:
                self.cap.set(cv2.CAP_PROP_POS_FRAMES, new_frame)
                ret, frame = self.cap.read()
                if ret:
                    self.display_frame(frame)
                    if self.fullscreen_window:
                        self.display_frame_fullscreen(frame)
                self.on_progress_update()
                self.update_time_labels()
            except Exception as e:
                self.write_log(f"マウスホイールエラー: {e}")

    def toggle_fullscreen(self, event=None):
        if self.fullscreen_window:
            self.exit_fullscreen()
        else:
            self.fullscreen_window = tk.Toplevel(self.root)
            self.fullscreen_window.attributes('-fullscreen', True)
            self.bind_keys(self.fullscreen_window)
            
            self.fullscreen_label = tk.Label(self.fullscreen_window, bg="black")
            self.fullscreen_label.pack(fill=tk.BOTH, expand=True)
            self.fullscreen_label.bind("<Button-1>", self.toggle_play_pause)
            self.fullscreen_label.bind("<Double-Button-1>", self.toggle_fullscreen)
            self.fullscreen_label.bind("<MouseWheel>", self.on_mouse_wheel)
            
            self.fullscreen_progress_canvas = tk.Canvas(self.fullscreen_window, height=30, bg="grey", highlightthickness=0)
            self.fullscreen_progress_canvas.place(relx=0, rely=1, anchor="sw", relwidth=1)
            self.fullscreen_progress_bar = self.fullscreen_progress_canvas.create_rectangle(0, 0, 0, 30, fill="green")
            self.fullscreen_start_marker = self.fullscreen_progress_canvas.create_line(0, 0, 0, 30, fill="red", width=3)
            self.fullscreen_end_marker = self.fullscreen_progress_canvas.create_line(0, 0, 0, 30, fill="blue", width=3)
            self.fullscreen_progress_text = self.fullscreen_progress_canvas.create_text(10, 15, anchor="w", fill="white", text="00:00:00 / 00:00:00")
            
            self.fullscreen_progress_canvas.bind("<Button-1>", self.on_fullscreen_progress_click)
            self.fullscreen_progress_canvas.bind("<MouseWheel>", self.on_mouse_wheel)
            
            with self.cap_lock:
                try:
                    if self.cap and self.cap.isOpened():
                        self.cap.set(cv2.CAP_PROP_POS_FRAMES, self.current_frame)
                        ret, frame = self.cap.read()
                        if ret:
                            self.display_frame_fullscreen(frame)
                        self.update_fullscreen_progress()
                    self.start_frame_buffer()
                except Exception as e:
                    self.write_log(f"フルスクリーン初期化エラー: {e}")
                    messagebox.showerror("エラー", f"フルスクリーン初期化に失敗しました: {e}")

    def on_fullscreen_resize(self, event):
        if self.fullscreen_window:
            self.root.after(100, self.update_fullscreen_preview)

    def update_fullscreen_preview(self):
        if self.fullscreen_window and self.cap and self.cap.isOpened():
            with self.cap_lock:
                try:
                    self.cap.set(cv2.CAP_PROP_POS_FRAMES, self.current_frame)
                    ret, frame = self.cap.read()
                    if ret:
                        self.display_frame_fullscreen(frame)
                except Exception as e:
                    self.write_log(f"フルスクリーンプレビュー更新エラー: {e}")

    def display_frame_fullscreen(self, frame):
        if self.fullscreen_window and frame is not None:
            try:
                screen_width = self.fullscreen_window.winfo_screenwidth()
                screen_height = self.fullscreen_window.winfo_screenheight()
                
                frame_height, frame_width = frame.shape[:2]
                aspect_ratio = frame_width / frame_height
                
                if screen_width / screen_height > aspect_ratio:
                    new_height = screen_height
                    new_width = int(new_height * aspect_ratio)
                else:
                    new_width = screen_width
                    new_height = int(new_width / aspect_ratio)
                
                resized_frame = cv2.resize(frame, (new_width, new_height), interpolation=cv2.INTER_NEAREST)
                
                black_bg = np.zeros((screen_height, screen_width, 3), dtype=np.uint8)
                offset_x = (screen_width - new_width) // 2
                offset_y = (screen_height - new_height) // 2
                black_bg[offset_y:offset_y+new_height, offset_x:offset_x+new_width] = resized_frame
                
                rgb_bg = cv2.cvtColor(black_bg, cv2.COLOR_BGR2RGB)
                img = Image.fromarray(rgb_bg)
                imgtk = ImageTk.PhotoImage(image=img)
                self.fullscreen_label.configure(image=imgtk)
                self.fullscreen_label.image = imgtk
            except Exception as e:
                self.write_log(f"フルスクリーン表示エラー: {e}")

    def on_fullscreen_progress_click(self, event):
        if not self.video_total_frames > 0 or not self.cap or not self.cap.isOpened():
            self.write_log("進捗クリック無効: 動画がロードされていません")
            return
        
        try:
            width = self.fullscreen_progress_canvas.winfo_width()
            if width <= 0:
                self.write_log("進捗クリックエラー: キャンバス幅が無効")
                return
            click_pos = event.x / width
            new_frame = int(click_pos * self.video_total_frames)
            new_frame = max(0, min(new_frame, self.video_total_frames - 1))
            self.current_frame = new_frame
            self.clear_frame_queue()
            with self.cap_lock:
                self.cap.set(cv2.CAP_PROP_POS_FRAMES, new_frame)
                ret, frame = self.cap.read()
                if ret:
                    self.display_frame(frame)
                    self.display_frame_fullscreen(frame)
                else:
                    self.write_log("フルスクリーン進捗クリック: フレーム読み込み失敗")
            self.on_progress_update()
            self.update_time_labels()
            self.update_fullscreen_progress()
            self.root.update_idletasks()
        except Exception as e:
            self.write_log(f"フルスクリーン進捗クリックエラー: {e}")
            messagebox.showerror("エラー", f"進捗バー操作中にエラーが発生しました: {e}")

    def update_fullscreen_progress(self):
        if self.fullscreen_progress_canvas and self.video_total_frames > 0:
            try:
                width = self.fullscreen_progress_canvas.winfo_width()
                if width <= 0:
                    self.write_log("フルスクリーン進捗更新エラー: キャンバス幅が無効")
                    return
                progress_width = (self.current_frame / self.video_total_frames) * width
                self.fullscreen_progress_canvas.coords(self.fullscreen_progress_bar, 0, 0, progress_width, 30)
                
                start_pos = (self.start_frame / self.video_total_frames) * width
                end_pos = (self.end_frame / self.video_total_frames) * width
                self.fullscreen_progress_canvas.coords(self.fullscreen_start_marker, start_pos, 0, start_pos, 30)
                self.fullscreen_progress_canvas.coords(self.fullscreen_end_marker, end_pos, 0, end_pos, 30)
                
                total_time_sec = self.video_total_frames / self.video_fps if self.video_fps > 0 else 0
                current_time_str = self.format_time(self.current_frame / self.video_fps if self.video_fps > 0 else 0)
                total_time_str = self.format_time(total_time_sec)
                self.fullscreen_progress_canvas.itemconfig(self.fullscreen_progress_text, text=f"{current_time_str} / {total_time_str}")
            except Exception as e:
                self.write_log(f"フルスクリーン進捗更新エラー: {e}")

    def display_frame(self, frame):
        if frame is None or not self.video_label.winfo_exists():
            self.display_black_frame()
            return
            
        try:
            self.root.update_idletasks()
            label_width = self.video_label.winfo_width()
            label_height = self.video_label.winfo_height()

            if label_width <= 0 or label_height <= 0:
                self.display_black_frame()
                return
                
            frame_height, frame_width = frame.shape[:2]
            aspect_ratio = frame_width / frame_height
            label_aspect_ratio = label_width / label_height
            
            if aspect_ratio > label_aspect_ratio:
                new_width = label_width
                new_height = int(new_width / aspect_ratio)
            else:
                new_height = label_height
                new_width = int(new_height * aspect_ratio)

            resized_frame = cv2.resize(frame, (new_width, new_height), interpolation=cv2.INTER_NEAREST)
            
            black_bg = np.zeros((label_height, label_width, 3), dtype=np.uint8)
            offset_x = (label_width - new_width) // 2
            offset_y = (label_height - new_height) // 2
            black_bg[offset_y:offset_y+new_height, offset_x:offset_x+new_width] = resized_frame
            
            rgb_bg = cv2.cvtColor(black_bg, cv2.COLOR_BGR2RGB)
            img = Image.fromarray(rgb_bg)
            imgtk = ImageTk.PhotoImage(image=img)
            self.video_label.configure(image=imgtk)
            self.video_label.image = imgtk
        except Exception as e:
            self.write_log(f"フレーム表示エラー: {e}")
            self.display_black_frame()

    def on_window_resize(self, event):
        if self.after_id:
            self.root.after_cancel(self.after_id)
        self.after_id = self.root.after(100, self.update_preview)

    def update_preview(self):
        if self.cap and self.cap.isOpened():
            with self.cap_lock:
                try:
                    current_frame = int(self.cap.get(cv2.CAP_PROP_POS_FRAMES))
                    self.cap.set(cv2.CAP_PROP_POS_FRAMES, current_frame)
                    ret, frame = self.cap.read()
                    if ret:
                        self.display_frame(frame)
                    else:
                        self.display_black_frame()
                except Exception as e:
                    self.write_log(f"プレビュー更新エラー: {e}")
                    self.display_black_frame()
        else:
            self.display_black_frame()

    def display_black_frame(self):
        try:
            self.root.update_idletasks()
            label_width = self.video_label.winfo_width()
            label_height = self.video_label.winfo_height()
            
            if label_width > 0 and label_height > 0:
                black_bg = np.zeros((label_height, label_width, 3), dtype=np.uint8)
                rgb_bg = cv2.cvtColor(black_bg, cv2.COLOR_BGR2RGB)
                img = Image.fromarray(rgb_bg)
                imgtk = ImageTk.PhotoImage(image=img)
                self.video_label.configure(image=imgtk)
                self.video_label.image = imgtk
        except Exception as e:
            self.write_log(f"黒フレーム表示エラー: {e}")

    def set_start_point(self):
        self.start_frame = self.current_frame
        if self.end_frame < self.start_frame:
            self.end_frame = self.start_frame
        self.update_time_labels()
        self.on_progress_update()

    def set_end_point(self):
        self.end_frame = self.current_frame
        if self.start_frame > self.end_frame:
            self.start_frame = self.end_frame
        self.update_time_labels()
        self.on_progress_update()

    def reset_points(self):
        self.start_frame = 0
        self.end_frame = self.video_total_frames
        self.update_time_labels()
        self.on_progress_update()

    def format_time(self, seconds):
        minutes, seconds = divmod(int(seconds), 60)
        hours, minutes = divmod(minutes, 60)
        return f"{hours:02d}:{minutes:02d}:{seconds:02d}"

    def update_time_labels(self):
        try:
            current_time_sec = self.current_frame / self.video_fps if self.video_fps > 0 else 0
            total_time_sec = self.video_total_frames / self.video_fps if self.video_fps > 0 else 0
            current_time_str = self.format_time(current_time_sec)
            total_time_str = self.format_time(total_time_sec)
            self.current_time_label.config(text=f"{current_time_str} / {total_time_str}")
            
            start_time_sec = self.start_frame / self.video_fps if self.video_fps > 0 else 0
            end_time_sec = self.end_frame / self.video_fps if self.video_fps > 0 else 0
            self.start_time_label.config(text=self.format_time(start_time_sec))
            self.end_time_label.config(text=self.format_time(end_time_sec))
        except Exception as e:
            self.write_log(f"時間ラベル更新エラー: {e}")

    def on_closing(self):
        self.buffer_running = False
        with self.cap_lock:
            if self.cap:
                try:
                    self.cap.release()
                    cv2.destroyAllWindows()
                except:
                    pass
                self.cap = None
        if (hasattr(self, 'is_running') and self.is_running) or \
           (hasattr(self, 'is_batch_processing') and self.is_batch_processing):
            if messagebox.askyesno("確認", "現在、処理が実行中です。中断して終了しますか?"):
                if self.process and self.process.poll() is None:
                    self.process.kill()
                    self.write_log("サブプロセスを強制終了しました")
                self.root.destroy()
        else:
            self.root.destroy()

if __name__ == "__main__":
    root = TkinterDnD.Tk()
    try:
        app = MosaicRemoverApp(root)
        if app.root is not None:
            root.mainloop()
    except Exception as e:
        messagebox.showerror("起動エラー", f"プログラムの起動中にエラーが発生しました。\n{e}")
    finally:
        if 'root' in locals():
            try:
                root.destroy()
            except:
                pass