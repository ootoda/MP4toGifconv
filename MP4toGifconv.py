import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import subprocess
import os
import shutil
import threading
import queue
from PIL import Image, ImageTk
import tempfile
import math
import json
import re

class Mp4ToGifConverter(tk.Tk):
    """
    MP4 をアニメーション GIF に変換するGUIアプリケーション
    - 品質プリセット（高画質/標準/軽量/超軽量）
    - FPS維持 / 解像度維持を独立チェック可能
    - 解像度半分縮小機能追加   
    - ビジュアルトリミング機能（サムネイル+ドラッグバー）
    """
    def __init__(self):
        super().__init__()
        self.withdraw()

        try:
            self.ffmpeg_path = self.find_ffmpeg_path()
            if not self.ffmpeg_path:
                self.destroy()
                return

            self.title("MP4 to GIF Converter")
            self.geometry("1200x1000")
            self.resizable(False, False)

            self.progress_queue = queue.Queue()
            self.video_duration = 0
            self.thumbnails = []
            self.thumbnail_times = []
            self.trim_start_ratio = 0.0
            self.trim_end_ratio = 1.0
            self.original_fps = None  # 追加: 元のFPSを保存する属性
            
            self.setup_ui()
            self.after(100, self.process_queue)

            # 中央表示
            self.update_idletasks()
            x = (self.winfo_screenwidth() - self.winfo_width()) // 2
            y = (self.winfo_screenheight() - self.winfo_height()) // 2
            self.geometry(f"+{x}+{y}")
            self.deiconify()
        except Exception as e:
            messagebox.showerror("初期化エラー", f"アプリケーションの初期化中にエラー:\n{str(e)}")
            self.destroy()

    def find_ffmpeg_path(self):
        """FFmpegのパスを検索"""
        config_file = "ffmpeg_path.txt"
        if os.path.exists(config_file):
            with open(config_file, 'r', encoding='utf-8') as f:
                path = f.read().strip()
            if os.path.isfile(path) and "ffmpeg" in os.path.basename(path).lower():
                return path

        path_from_env = shutil.which("ffmpeg")
        if path_from_env:
            with open(config_file, 'w', encoding='utf-8') as f:
                f.write(path_from_env)
            return path_from_env

        self.deiconify()
        messagebox.showinfo("FFmpegが見つかりません", 
                            "ffmpeg.exe の場所を指定してください。")
        user_path = filedialog.askopenfilename(
            title="ffmpeg.exeを選択してください", 
            filetypes=[("FFmpeg Executable", "ffmpeg.exe"), ("All files", "*.*")]
        )
        self.withdraw()

        if user_path and os.path.isfile(user_path):
            with open(config_file, 'w', encoding='utf-8') as f:
                f.write(user_path)
            return user_path
        else:
            messagebox.showerror("エラー", "FFmpegが指定されなかったため終了します。")
            return None

    def setup_ui(self):
        main_frame = ttk.Frame(self, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- 入力設定 ---
        input_frame = ttk.LabelFrame(main_frame, text="入力設定", padding="10")
        input_frame.pack(fill=tk.X, pady=5)

        self.input_path = tk.StringVar()
        self.batch_mode = tk.BooleanVar(value=False)

        input_frame.columnconfigure(0, weight=1)
        ttk.Label(input_frame, text="ファイル/フォルダ:").grid(row=0, column=0, columnspan=2, sticky=tk.W, padx=5, pady=5)
        self.path_entry = ttk.Entry(input_frame, textvariable=self.input_path, width=80, state="readonly")
        self.path_entry.grid(row=1, column=0, padx=5, sticky="ew")
        ttk.Button(input_frame, text="選択...", command=self.select_input).grid(row=1, column=1, padx=5)
        ttk.Checkbutton(input_frame, text="フォルダ内のすべてのMP4を変換する", 
                        variable=self.batch_mode, command=self.toggle_input_mode).grid(row=2, column=0, columnspan=2, sticky=tk.W, padx=5, pady=5)

        # --- トリミング設定 ---
        trim_frame = ttk.LabelFrame(main_frame, text="トリミング設定", padding="10")
        trim_frame.pack(fill=tk.X, pady=5)
        
        self.enable_trim = tk.BooleanVar(value=False)
        ttk.Checkbutton(trim_frame, text="トリミングを有効にする", 
                        variable=self.enable_trim, command=self.toggle_trim_mode).grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        
        # サムネイル表示エリア
        self.thumbnail_frame = tk.Frame(trim_frame, bg="black", height=120)
        self.thumbnail_frame.grid(row=1, column=0, columnspan=2, sticky="ew", padx=5, pady=10)
        self.thumbnail_frame.grid_propagate(False)
        self.thumbnail_frame.columnconfigure(0, weight=1)
        
        # トリミングバー
        self.trim_bar_frame = tk.Frame(trim_frame, height=30)
        self.trim_bar_frame.grid(row=2, column=0, columnspan=2, sticky="ew", padx=5)
        self.trim_bar_frame.grid_propagate(False)
        
        # トリミングバーのキャンバス
        self.trim_canvas = tk.Canvas(self.trim_bar_frame, height=30, bg="#f0f0f0", highlightthickness=1, highlightbackground="gray")
        self.trim_canvas.pack(fill=tk.X, padx=2, pady=2)
        
        # 初期状態では非表示
        self.thumbnail_frame.grid_remove()
        self.trim_bar_frame.grid_remove()

        # --- GIF出力設定 ---
        settings_frame = ttk.LabelFrame(main_frame, text="GIF出力設定", padding="10")
        settings_frame.pack(fill=tk.X, pady=5)

        # プリセット
        self.preset = tk.StringVar(value="高画質")
        ttk.Label(settings_frame, text="品質プリセット:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        preset_combo = ttk.Combobox(settings_frame, textvariable=self.preset, state="readonly",
                                    values=["高画質", "標準", "軽量", "超軽量"])
        preset_combo.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        preset_combo.bind("<<ComboboxSelected>>", self.apply_preset)

        # FPS
        self.fps = tk.StringVar(value="30")
        ttk.Label(settings_frame, text="FPS (フレーム/秒):").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        
        fps_outer_frame = ttk.Frame(settings_frame)
        fps_outer_frame.grid(row=1, column=1, sticky=tk.W, padx=5)
        self.fps_entry = ttk.Entry(fps_outer_frame, textvariable=self.fps, width=5)
        self.fps_entry.pack(side=tk.LEFT)
        
        self.fps_radio_frame = ttk.Frame(fps_outer_frame)
        self.fps_radio_frame.pack(side=tk.LEFT, padx=(5, 0))
        for val in ["10", "15", "20", "24", "30"]:
            ttk.Radiobutton(self.fps_radio_frame, text=f"{val}FPS", variable=self.fps, value=val).pack(side=tk.LEFT)

        self.keep_fps = tk.BooleanVar(value=False)
        ttk.Checkbutton(settings_frame, text="MP4のFPSを維持する", 
                        variable=self.keep_fps, command=self.toggle_original_mode).grid(row=2, column=0, columnspan=2, sticky=tk.W, padx=5, pady=5)

        # 色数
        self.colors = tk.StringVar(value="256")
        ttk.Label(settings_frame, text="色数:").grid(row=3, column=0, sticky=tk.W, padx=5, pady=5)

        colors_outer_frame = ttk.Frame(settings_frame)
        colors_outer_frame.grid(row=3, column=1, sticky=tk.W, padx=5)
        self.colors_entry = ttk.Entry(colors_outer_frame, textvariable=self.colors, width=5)
        self.colors_entry.pack(side=tk.LEFT)
        
        self.colors_radio_frame = ttk.Frame(colors_outer_frame)
        self.colors_radio_frame.pack(side=tk.LEFT, padx=(5, 0))
        for val in ["8", "16", "32", "64", "128", "256"]:
            ttk.Radiobutton(self.colors_radio_frame, text=f"{val}色", variable=self.colors, value=val).pack(side=tk.LEFT)
            
        # 解像度
        self.width = tk.StringVar(value="640")
        self.height = tk.StringVar(value="360")
        self.keep_aspect = tk.BooleanVar(value=True)

        ttk.Label(settings_frame, text="解像度 (幅 x 高さ):").grid(row=4, column=0, sticky=tk.W, padx=5, pady=5)
        size_frame = ttk.Frame(settings_frame)
        size_frame.grid(row=4, column=1, sticky=tk.W, padx=5)
        self.width_entry = ttk.Entry(size_frame, textvariable=self.width, width=6)
        self.width_entry.pack(side=tk.LEFT)
        ttk.Label(size_frame, text=" x ").pack(side=tk.LEFT)
        self.height_entry = ttk.Entry(size_frame, textvariable=self.height, width=6)
        self.height_entry.pack(side=tk.LEFT)

        # 解像度制御チェックボックス
        self.keep_res = tk.BooleanVar(value=True)
        self.half_res = tk.BooleanVar(value=False)

        ttk.Checkbutton(settings_frame, text="MP4の解像度を維持する",
                        variable=self.keep_res, command=self.on_keep_res_changed).grid(row=6, column=0, columnspan=2, sticky=tk.W, padx=5, pady=5)
        ttk.Checkbutton(settings_frame, text="MP4の解像度を半分に縮小する",
                        variable=self.half_res, command=self.on_half_res_changed).grid(row=7, column=0, columnspan=2, sticky=tk.W, padx=5, pady=5)

        # ループ
        self.is_loop = tk.BooleanVar(value=True)
        ttk.Checkbutton(settings_frame, text="無限ループさせる", variable=self.is_loop).grid(row=8, column=0, columnspan=2, sticky=tk.W, padx=5, pady=5)

        # --- 実行と進捗 ---
        self.progress_label = ttk.Label(main_frame, text="待機中...")
        self.progress_label.pack(fill=tk.X, pady=(10, 0), padx=5)
        self.progress_bar = ttk.Progressbar(main_frame, orient="horizontal", mode="determinate")
        self.progress_bar.pack(fill=tk.X, pady=5, padx=5)
        self.convert_button = ttk.Button(main_frame, text="変換開始", command=self.start_conversion_thread)
        self.convert_button.pack(pady=10)

        # --- ログ表示 ---
        log_frame = ttk.LabelFrame(main_frame, text="実行ログ", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        self.log_text = tk.Text(log_frame, height=6, state="disabled", wrap=tk.WORD, bg="#f0f0f0", font=("Meiryo UI", 9))
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)

        # 初期状態
        self.toggle_original_mode()
        self.toggle_trim_mode()

    def toggle_trim_mode(self):
        """トリミングモードのUI切替"""
        if self.enable_trim.get():
            self.thumbnail_frame.grid()
            self.trim_bar_frame.grid()
            if self.input_path.get() and not self.batch_mode.get() and self.input_path.get().lower().endswith('.mp4'):
                self.load_video_thumbnails()
        else:
            self.thumbnail_frame.grid_remove()
            self.trim_bar_frame.grid_remove()

    def load_video_thumbnails(self):
        """動画のサムネイルを生成・表示"""
        if not self.input_path.get() or self.batch_mode.get():
            return
            
        video_file = self.input_path.get()
        if not os.path.exists(video_file) or not video_file.lower().endswith('.mp4'):
            return
            
        try:
            # 動画の長さを取得
            self.get_video_duration(video_file)
            
            if self.video_duration <= 0:
                messagebox.showerror("エラー", "動画の長さを取得できませんでした。")
                return
            
            # サムネイルを生成
            self.generate_thumbnails(video_file)
            
            # サムネイルを表示
            self.display_thumbnails()
            
            # トリミングバーを初期化
            self.init_trim_bar()
            
        except Exception as e:
            messagebox.showerror("エラー", f"サムネイル生成エラー: {e}")
            print(f"サムネイル生成エラー: {e}")

    def get_video_info(self, video_file):
        """動画の情報（長さとFPS）を取得 - 改良版"""
        try:
            # ffprobeがない場合はffmpegパスからffprobeのパスを推定
            ffprobe_path = self.ffmpeg_path.replace('ffmpeg.exe', 'ffprobe.exe').replace('ffmpeg', 'ffprobe')
            
            # ffprobeコマンドを実行
            command = [
                ffprobe_path, 
                '-v', 'quiet', 
                '-print_format', 'json', 
                '-show_streams',
                '-show_format', 
                video_file
            ]
            
            startupinfo = None
            creationflags = 0
            if os.name == 'nt':
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                creationflags = subprocess.CREATE_NO_WINDOW
                
            # エンコーディング問題を修正
            result = subprocess.run(
                command, 
                capture_output=True, 
                text=True,
                encoding='utf-8',
                errors='ignore',  # デコードエラーを無視
                startupinfo=startupinfo, 
                creationflags=creationflags
            )
            
            if result.returncode == 0:
                try:
                    info = json.loads(result.stdout)
                    
                    # 動画の長さを取得
                    self.video_duration = float(info['format']['duration'])
                    
                    # FPSを取得
                    for stream in info['streams']:
                        if stream['codec_type'] == 'video':
                            fps_str = stream.get('r_frame_rate', '30/1')
                            if '/' in fps_str:
                                num, den = fps_str.split('/')
                                self.original_fps = float(num) / float(den)
                            else:
                                self.original_fps = float(fps_str)
                            break
                    
                    print(f"動画の長さ: {self.video_duration}秒, FPS: {self.original_fps}")
                    return
                except (json.JSONDecodeError, KeyError, ValueError) as e:
                    print(f"ffprobe JSON解析エラー: {e}")
                    
        except FileNotFoundError:
            print("ffprobeが見つかりません")
        except Exception as e:
            print(f"ffprobe実行エラー: {e}")
            
        # ffprobeが失敗した場合はffmpegで長さを取得を試す
        self.get_duration_with_ffmpeg(video_file)

    def get_video_duration(self, video_file):
        """動画の長さを取得 - 改良版"""
        self.get_video_info(video_file)

    def get_duration_with_ffmpeg(self, video_file):
        """ffmpegを使って動画の長さを取得（代替手段） - 改良版"""
        try:
            command = [self.ffmpeg_path, '-i', video_file, '-f', 'null', '-']
            
            startupinfo = None
            creationflags = 0
            if os.name == 'nt':
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                creationflags = subprocess.CREATE_NO_WINDOW
                
            # エンコーディング問題を修正
            result = subprocess.run(
                command, 
                capture_output=True, 
                text=True,
                encoding='utf-8',
                errors='ignore',  # デコードエラーを無視
                startupinfo=startupinfo, 
                creationflags=creationflags
            )
            
            # ffmpegの出力からDurationを探す
            output_text = result.stderr if result.stderr else result.stdout
            duration_match = re.search(r'Duration: (\d{2}):(\d{2}):(\d{2}\.\d{2})', output_text)
            
            if duration_match:
                hours = int(duration_match.group(1))
                minutes = int(duration_match.group(2))  
                seconds = float(duration_match.group(3))
                self.video_duration = hours * 3600 + minutes * 60 + seconds
                print(f"ffmpegで検出した動画長さ: {self.video_duration}秒")
                return self.video_duration
            
            print("Durationが見つかりませんでした")
            self.video_duration = 60  # デフォルト値
            return 60
            
        except Exception as e:
            print(f"ffmpegでの動画長さ取得エラー: {e}")
            self.video_duration = 60
            return 60

    def generate_thumbnails(self, video_file, count=8):
        """サムネイルを生成 - 改良版"""
        self.thumbnails = []
        self.thumbnail_times = []
        
        print(f"サムネイル生成開始: {count}枚, 動画長さ: {self.video_duration}秒")
        
        if self.video_duration <= 0:
            print("動画の長さが無効のため再取得を試行")
            self.get_video_duration(video_file)
        
        try:
            with tempfile.TemporaryDirectory() as temp_dir:
                for i in range(count):
                    # より保守的な時間配置（動画の長さの10%から90%の範囲で配置）
                    if count == 1:
                        time_pos = self.video_duration * 0.5
                    else:
                        start_ratio = 0.1
                        end_ratio = 0.9
                        ratio_range = end_ratio - start_ratio
                        ratio = start_ratio + (ratio_range / (count - 1)) * i
                        time_pos = self.video_duration * ratio
                    
                    # 時間位置が動画の長さを超えないように制限
                    time_pos = min(time_pos, self.video_duration - 0.5)  # 0.5秒のマージン
                    time_pos = max(time_pos, 0.1)  # 最小0.1秒
                    
                    self.thumbnail_times.append(time_pos)
                    
                    output_path = os.path.join(temp_dir, f"thumb_{i:03d}.jpg")
                    
                    # より確実なFFmpegコマンド（改良版）
                    command = [
                        self.ffmpeg_path, 
                        '-i', video_file,           # 入力ファイル
                        '-ss', str(time_pos),       # シーク位置（入力後に指定）
                        '-vframes', '1',            # 1フレームのみ
                        '-vf', 'scale=90:50:force_original_aspect_ratio=decrease,pad=90:50:(ow-iw)/2:(oh-ih)/2',
                        '-q:v', '2',                # 高品質
                        '-f', 'image2',             # 出力フォーマットを明示
                        '-y', output_path           # 上書き
                    ]
                    
                    startupinfo = None
                    creationflags = 0
                    if os.name == 'nt':
                        startupinfo = subprocess.STARTUPINFO()
                        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                        creationflags = subprocess.CREATE_NO_WINDOW
                    
                    try:
                        print(f"サムネイル{i}生成コマンド実行: 時間位置 {time_pos:.2f}秒")
                        
                        # エンコーディング問題を修正したPopen実行
                        process = subprocess.Popen(
                            command,
                            stdout=subprocess.PIPE,
                            stderr=subprocess.PIPE,
                            text=True,
                            encoding='utf-8',
                            errors='ignore',  # デコードエラーを無視
                            startupinfo=startupinfo,
                            creationflags=creationflags
                        )
                        
                        stdout, stderr = process.communicate(timeout=30)  # 30秒タイムアウト
                        
                        if process.returncode == 0 and os.path.exists(output_path) and os.path.getsize(output_path) > 100:
                            img = Image.open(output_path)
                            photo = ImageTk.PhotoImage(img)
                            self.thumbnails.append(photo)
                            print(f"サムネイル{i}生成成功: {time_pos:.2f}秒")
                        else:
                            print(f"サムネイル{i}生成失敗: 戻り値={process.returncode}")
                            if stderr:
                                print(f"エラー詳細: {stderr[:500]}")  # エラー出力を500文字に制限
                            
                            # ダミー画像
                            dummy_img = Image.new('RGB', (90, 50), color=(128, 128, 128))
                            self.thumbnails.append(ImageTk.PhotoImage(dummy_img))
                    
                    except subprocess.TimeoutExpired:
                        print(f"サムネイル{i}生成タイムアウト: {time_pos:.2f}秒")
                        process.kill()
                        dummy_img = Image.new('RGB', (90, 50), color=(255, 0, 0))
                        self.thumbnails.append(ImageTk.PhotoImage(dummy_img))
                    except Exception as e:
                        print(f"サムネイル{i}生成例外: {time_pos:.2f}秒, エラー: {e}")
                        dummy_img = Image.new('RGB', (90, 50), color=(255, 165, 0))
                        self.thumbnails.append(ImageTk.PhotoImage(dummy_img))
                        
        except Exception as e:
            print(f"サムネイル生成エラー: {e}")
            # エラー時はダミー画像で埋める
            for i in range(count):
                dummy_img = Image.new('RGB', (90, 50), color='gray')
                self.thumbnails.append(ImageTk.PhotoImage(dummy_img))
                if len(self.thumbnail_times) <= i:
                    self.thumbnail_times.append((self.video_duration / count) * i)

    def display_thumbnails(self):
        """サムネイルを表示"""
        # 既存のウィジェットをクリア
        for widget in self.thumbnail_frame.winfo_children():
            widget.destroy()
        
        print(f"サムネイル表示: {len(self.thumbnails)}枚")
        
        # サムネイルを横に並べる
        container = tk.Frame(self.thumbnail_frame, bg="black")
        container.pack(fill=tk.BOTH, expand=True)
            
        for i, thumbnail in enumerate(self.thumbnails):
            label = tk.Label(container, image=thumbnail, bg="black", bd=1, relief="solid")
            label.pack(side=tk.LEFT, padx=2, pady=10)
            
            # 時間表示ラベルを追加
            time_str = f"{int(self.thumbnail_times[i]//60):02d}:{int(self.thumbnail_times[i]%60):02d}"
            time_label = tk.Label(container, text=time_str, bg="black", fg="white", font=("Arial", 8))
            time_label.pack(side=tk.LEFT, padx=(0, 5))

    def init_trim_bar(self):
        """トリミングバーを初期化"""
        self.trim_canvas.delete("all")
        
        # キャンバスの実際の幅を取得
        self.trim_canvas.update_idletasks()
        canvas_width = self.trim_canvas.winfo_width()
        if canvas_width <= 1:
            canvas_width = 750
        
        # バーの背景
        self.trim_canvas.create_rectangle(0, 10, canvas_width, 20, fill="lightgray", outline="gray")
        
        # 選択範囲（初期は全体）
        start_x = canvas_width * self.trim_start_ratio
        end_x = canvas_width * self.trim_end_ratio
        
        self.selection_rect = self.trim_canvas.create_rectangle(
            start_x, 8, end_x, 22, fill="red", stipple="gray50", outline="red", width=2
        )
        
        # ドラッグハンドル
        self.start_handle = self.trim_canvas.create_rectangle(
            start_x-5, 5, start_x+5, 25, fill="red", outline="darkred", width=2
        )
        
        self.end_handle = self.trim_canvas.create_rectangle(
            end_x-5, 5, end_x+5, 25, fill="red", outline="darkred", width=2
        )
        
        # イベントバインド
        self.trim_canvas.bind("<Button-1>", self.on_trim_click)
        self.trim_canvas.bind("<B1-Motion>", self.on_trim_drag)
        self.trim_canvas.bind("<ButtonRelease-1>", self.on_trim_release)
        
        self.drag_item = None

    def on_trim_click(self, event):
        """トリミングバークリック時"""
        x = event.x
        canvas_width = self.trim_canvas.winfo_width()
        
        # クリックされたアイテムを特定
        clicked_items = self.trim_canvas.find_overlapping(x-5, event.y-5, x+5, event.y+5)
        
        if self.start_handle in clicked_items:
            self.drag_item = "start"
        elif self.end_handle in clicked_items:
            self.drag_item = "end"
        elif self.selection_rect in clicked_items:
            self.drag_item = "selection"
            self.drag_offset = x - (canvas_width * self.trim_start_ratio)
        else:
            self.drag_item = None

    def on_trim_drag(self, event):
        """トリミングバードラッグ時"""
        if not self.drag_item:
            return
            
        canvas_width = self.trim_canvas.winfo_width()
        x = max(0, min(canvas_width, event.x))
        
        if self.drag_item == "start":
            new_ratio = x / canvas_width
            if new_ratio < self.trim_end_ratio:
                self.trim_start_ratio = new_ratio
                self.update_trim_display()
                
        elif self.drag_item == "end":
            new_ratio = x / canvas_width
            if new_ratio > self.trim_start_ratio:
                self.trim_end_ratio = new_ratio
                self.update_trim_display()
                
        elif self.drag_item == "selection":
            selection_width = self.trim_end_ratio - self.trim_start_ratio
            new_start = (x - self.drag_offset) / canvas_width
            new_end = new_start + selection_width
            
            if new_start >= 0 and new_end <= 1:
                self.trim_start_ratio = new_start
                self.trim_end_ratio = new_end
                self.update_trim_display()

    def on_trim_release(self, event):
        """トリミングバーリリース時"""
        self.drag_item = None

    def update_trim_display(self):
        """トリミング表示を更新"""
        canvas_width = self.trim_canvas.winfo_width()
        start_x = canvas_width * self.trim_start_ratio
        end_x = canvas_width * self.trim_end_ratio
        
        # 選択範囲を更新
        self.trim_canvas.coords(self.selection_rect, start_x, 8, end_x, 22)
        
        # ハンドルを更新
        self.trim_canvas.coords(self.start_handle, start_x-5, 5, start_x+5, 25)
        self.trim_canvas.coords(self.end_handle, end_x-5, 5, end_x+5, 25)

    def on_keep_res_changed(self):
        """「解像度を維持する」チェックボックスの変更時処理"""
        if self.keep_res.get():
            self.half_res.set(False)
        self.toggle_original_mode()

    def on_half_res_changed(self):
        """「解像度を半分に縮小する」チェックボックスの変更時処理"""
        if self.half_res.get():
            self.keep_res.set(False)
        self.toggle_original_mode()

    def apply_preset(self, event=None):
        """品質プリセットの適用"""
        if self.keep_fps.get() and (self.keep_res.get() or self.half_res.get()):
            return

        preset = self.preset.get()
        if not self.keep_fps.get():
            if preset == "高画質":
                self.fps.set("30")
            elif preset == "標準":
                self.fps.set("20")
            elif preset == "軽量":
                self.fps.set("15")
            elif preset == "超軽量":
                self.fps.set("10")

        if not self.keep_res.get() and not self.half_res.get():
            if preset == "高画質":
                self.width.set("1280"); self.height.set("720")
            elif preset == "標準":
                self.width.set("640"); self.height.set("360")
            elif preset == "軽量":
                self.width.set("480"); self.height.set("270")
            elif preset == "超軽量":
                self.width.set("320"); self.height.set("180")

        if preset == "高画質":
            self.colors.set("256")
        elif preset == "標準":
            self.colors.set("128")
        elif preset == "軽量":
            self.colors.set("64")
        elif preset == "超軽量":
            self.colors.set("8")

    def toggle_original_mode(self):
        """入力値維持モードのUI切替"""
        fps_state = tk.DISABLED if self.keep_fps.get() else tk.NORMAL
        res_state = tk.DISABLED if (self.keep_res.get() or self.half_res.get()) else tk.NORMAL

        self.fps_entry.config(state=fps_state)
        for widget in self.fps_radio_frame.winfo_children():
            widget.configure(state=fps_state)
            
        self.width_entry.config(state=res_state)
        self.height_entry.config(state=res_state)

    def toggle_input_mode(self):
        self.input_path.set("")
        # サムネイルをクリア
        for widget in self.thumbnail_frame.winfo_children():
            widget.destroy()
        if hasattr(self, 'trim_canvas'):
            self.trim_canvas.delete("all")

    def select_input(self):
        if self.batch_mode.get():
            path = filedialog.askdirectory(title="MP4ファイルが含まれるフォルダを選択")
        else:
            path = filedialog.askopenfilename(title="MP4ファイルを選択", filetypes=[("MP4 files", "*.mp4")])
        if path:
            self.input_path.set(path)
            # 単一ファイルでトリミングが有効な場合、サムネイルを読み込み
            if not self.batch_mode.get() and self.enable_trim.get() and path.lower().endswith('.mp4'):
                self.load_video_thumbnails()

    def start_conversion_thread(self):
        if not self.input_path.get():
            messagebox.showwarning("警告", "変換するファイルまたはフォルダを選択してください。")
            return
            
        # トリミング設定の検証
        if self.enable_trim.get():
            if self.batch_mode.get():
                messagebox.showwarning("警告", "バッチモードではトリミング機能は使用できません。")
                return
                
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete("1.0", tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.convert_button.config(state=tk.DISABLED)
        self.progress_bar["value"] = 0
        thread = threading.Thread(target=self.run_conversion, daemon=True)
        thread.start()

    def run_conversion(self):
        try:
            input_path = self.input_path.get()
            is_batch = self.batch_mode.get()
            if is_batch:
                files_to_convert = [os.path.join(input_path, f) for f in os.listdir(input_path) if f.lower().endswith(".mp4")]
                output_dir = os.path.join(input_path, "converted_gifs")
            else:
                files_to_convert = [input_path]
                output_dir = os.path.join(os.path.dirname(input_path), "converted_gifs")

            if not files_to_convert:
                self.progress_queue.put(("info", "対象フォルダにMP4ファイルが見つかりませんでした。"))
                self.progress_queue.put(("enable_button", None))
                return

            os.makedirs(output_dir, exist_ok=True)
            total_files = len(files_to_convert)

            for i, file_path in enumerate(files_to_convert):
                # 各ファイルの変換前に動画情報を取得
                if not is_batch or i == 0:  # バッチ処理では最初のファイルのみ
                    self.get_video_info(file_path)
                
                self.progress_queue.put(("label", f"処理中: {i+1}/{total_files} - {os.path.basename(file_path)}"))
                base_name = os.path.splitext(os.path.basename(file_path))[0]
                
                # トリミング情報をファイル名に追加
                if self.enable_trim.get() and not is_batch:
                    start_time = self.trim_start_ratio * self.video_duration
                    end_time = self.trim_end_ratio * self.video_duration
                    start_str = f"{int(start_time//60):02d}m{int(start_time%60):02d}s"
                    end_str = f"{int(end_time//60):02d}m{int(end_time%60):02d}s"
                    base_name += f"_trim_{start_str}_{end_str}"
                
                output_path = os.path.join(output_dir, f"{base_name}.gif")
                command = self.build_ffmpeg_command(file_path, output_path)
                if not command: continue

                self.progress_queue.put(("log", f"--- 「{os.path.basename(file_path)}」の変換を開始 ---"))
                self.progress_queue.put(("log", f"実行コマンド: {' '.join(command)}"))

                startupinfo = None
                creationflags = 0
                if os.name == 'nt':
                    startupinfo = subprocess.STARTUPINFO()
                    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                    creationflags = subprocess.CREATE_NO_WINDOW

                # エンコーディング問題を修正したプロセス実行
                process = subprocess.Popen(
                    command, 
                    stdout=subprocess.PIPE, 
                    stderr=subprocess.STDOUT,
                    text=True, 
                    encoding='utf-8', 
                    errors='ignore',  # デコードエラーを無視
                    startupinfo=startupinfo, 
                    creationflags=creationflags
                )

                for line in iter(process.stdout.readline, ''):
                    self.progress_queue.put(("log", line.strip()))
                process.wait()

                if process.returncode == 0:
                    self.progress_queue.put(("log", f"--- 変換成功 ---\n"))
                else:
                    self.progress_queue.put(("log", f"--- 変換失敗 (エラーコード: {process.returncode}) ---\n"))

                self.progress_queue.put(("progress", ((i + 1) / total_files) * 100))

            self.progress_queue.put(("label", f"完了: {total_files}個のファイルの処理が完了しました。"))
            self.progress_queue.put(("done", output_dir))

        except Exception as e:
            self.progress_queue.put(("error", f"予期せぬエラーが発生しました:\n{e}"))
        finally:
            self.progress_queue.put(("enable_button", None))

    def build_ffmpeg_command(self, input_file, output_file):
        """FFmpegコマンドを構築 - 改良版"""
        try:
            loop = "0" if self.is_loop.get() else "-1"
            command = [self.ffmpeg_path, "-i", input_file]
            
            # トリミング設定の追加
            if self.enable_trim.get():
                start_time = self.trim_start_ratio * self.video_duration
                end_time = self.trim_end_ratio * self.video_duration
                duration = end_time - start_time
                
                command.extend(["-ss", str(start_time)])
                command.extend(["-t", str(duration)])
            
            vf_filters = []

            # FPS設定の改良版
            if not self.keep_fps.get():
                try:
                    fps_val = int(self.fps.get())
                    if fps_val <= 0: 
                        raise ValueError("FPSは正の数である必要があります。")
                    vf_filters.append(f"fps={fps_val}")
                except ValueError as e:
                    self.progress_queue.put(("warning", f"無効なFPS値: {self.fps.get()}。デフォルトの30FPSを使用します。"))
                    vf_filters.append("fps=30")
            else:
                # 元のFPSを維持する場合の処理
                if self.original_fps is not None:
                    # 明示的に元のFPSを指定してフレームレートの変更を防ぐ
                    vf_filters.append(f"fps={self.original_fps:.6f}")
                    self.progress_queue.put(("log", f"元のFPSを維持: {self.original_fps:.3f}fps"))
                else:
                    # 元のFPSが取得できない場合は、フレームレートフィルタを適用しない
                    self.progress_queue.put(("log", "元のFPS情報が取得できないため、入力ファイルのフレームレートをそのまま使用します"))

            # 解像度設定
            if self.half_res.get():
                vf_filters.append("scale=iw/2:ih/2:flags=lanczos")
            elif not self.keep_res.get():
                try:
                    width = int(self.width.get())
                    height = int(self.height.get())
                    if width <= 0 or height <= 0:
                        raise ValueError("解像度は正の数である必要があります。")
                    
                    if self.keep_aspect.get():
                        vf_filters.append(f"scale={width}:{height}:flags=lanczos:force_original_aspect_ratio=decrease")
                        vf_filters.append(f"pad={width}:{height}:(ow-iw)/2:(oh-ih)/2")
                    else:
                        vf_filters.append(f"scale={width}:{height}:flags=lanczos")
                        
                except ValueError as e:
                    self.progress_queue.put(("warning", f"無効な解像度値: {self.width.get()}x{self.height.get()}。元の解像度を維持します。"))

            # 色数設定の検証
            try:
                colors_val = int(self.colors.get())
                if not 2 <= colors_val <= 256:
                    raise ValueError("色数は2から256の間である必要があります。")
            except ValueError:
                self.progress_queue.put(("warning", f"無効な色数値: {self.colors.get()}。デフォルトの256色を使用します。"))
                colors_val = 256

            # パレット生成 + 適用（改良版）
            if vf_filters:
                filter_chain = ",".join(vf_filters)
                full_filter = f"{filter_chain},split[s0][s1];[s0]palettegen=max_colors={colors_val}:stats_mode=diff[p];[s1][p]paletteuse=dither=bayer:bayer_scale=5:diff_mode=rectangle"
            else:
                full_filter = f"split[s0][s1];[s0]palettegen=max_colors={colors_val}:stats_mode=diff[p];[s1][p]paletteuse=dither=bayer:bayer_scale=5:diff_mode=rectangle"
            
            command.extend(["-vf", full_filter, "-loop", loop, "-y", output_file])
            return command

        except Exception as e:
            self.progress_queue.put(("warning", f"コマンド構築エラー: {e}"))
            return None

    def process_queue(self):
        try:
            while True:
                msg_type, data = self.progress_queue.get_nowait()
                if msg_type == "label":
                    self.progress_label.config(text=data)
                elif msg_type == "progress":
                    self.progress_bar["value"] = data
                elif msg_type == "done":
                    messagebox.showinfo("完了", f"変換が完了しました。\n出力先: {data}")
                elif msg_type == "info":
                    messagebox.showinfo("情報", data)
                elif msg_type == "warning":
                    messagebox.showwarning("警告", data)
                elif msg_type == "error":
                    messagebox.showerror("エラー", data)
                    self.progress_label.config(text="エラーが発生しました")
                elif msg_type == "log":
                    self.log_text.config(state=tk.NORMAL)
                    self.log_text.insert(tk.END, data + "\n")
                    self.log_text.see(tk.END)
                    self.log_text.config(state=tk.DISABLED)
                elif msg_type == "enable_button":
                    self.convert_button.config(state=tk.NORMAL)
                    if "エラー" not in self.progress_label.cget("text"):
                        self.progress_label.config(text="待機中...")
        except queue.Empty:
            pass
        finally:
            self.after(100, self.process_queue)

def main():
    app = Mp4ToGifConverter()
    if app.winfo_exists():
        app.mainloop()

if __name__ == "__main__":
    main()