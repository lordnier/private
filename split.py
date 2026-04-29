import subprocess
import os
import tkinter as tk
from tkinter import filedialog, simpledialog

def split_video_lossless(input_file: str, segment_duration: int = 20) -> None:
    if not os.path.exists(input_file):
        print(f"エラー: ファイルが見つかりません: {input_file}")
        return

    ffmpeg_path = '/opt/homebrew/bin/ffmpeg'
    # Intel Macの場合:
    # ffmpeg_path = '/usr/local/bin/ffmpeg'

    if not os.path.exists(ffmpeg_path):
        print(f"エラー: FFmpegが見つかりません: {ffmpeg_path}")
        print("以下のコマンドで確認してください: which ffmpeg")
        return

    file_name = os.path.splitext(os.path.basename(input_file))[0]
    input_dir = os.path.dirname(input_file)
    output_dir = os.path.join(input_dir, f"{file_name}_split")
    os.makedirs(output_dir, exist_ok=True)

    output_pattern = os.path.join(output_dir, f"{file_name}_%03d.mp4")

    cmd = [
        ffmpeg_path,
        '-i', input_file,
        '-c', 'copy',
        '-f', 'segment',
        '-segment_time', str(segment_duration),
        '-reset_timestamps', '1',
        output_pattern
    ]

    print(f"処理開始: {input_file}")
    result = subprocess.run(cmd, capture_output=True, text=True)

    if result.returncode == 0:
        print(f"✅ 分割完了: {output_dir}")
    else:
        print("❌ エラーが発生しました:")
        print(result.stderr)

root = tk.Tk()
root.withdraw()

segment_duration = simpledialog.askinteger(
    "分割間隔の指定",
    "何秒間隔で区切りますか？",
    minvalue=1,
    initialvalue=15
)

if segment_duration is None:
    print("秒数が入力されなかったため終了しました")
else:
    print("MP4ファイルを選択してください...")
    file_path = filedialog.askopenfilename(
        title="分割するMP4ファイルを選択",
        filetypes=[("MP4ファイル", "*.mp4"), ("すべてのファイル", "*.*")],
        initialdir=os.path.expanduser("~/Downloads")
    )

    if file_path:
        print(f"選択されたファイル: {file_path}")
        split_video_lossless(file_path, segment_duration=segment_duration)
    else:
        print("ファイルが選択されませんでした")
