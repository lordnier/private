import os
import subprocess
import random
from pathlib import Path

def get_audio_files_sorted_by_time(audio_dir):
    """input_audio の mp3 をダウンロード日時順（古い順）に並べる"""
    audio_path = Path(audio_dir)
    audio_files = [f for f in audio_path.iterdir() if f.suffix == '.mp3']
    # Macでは stat().st_mtime（更新日時）で「ダウンロードした順」に近い並びになることが多い。[web:13]
    audio_files.sort(key=lambda f: f.stat().st_mtime)
    return audio_files

def batch_rename_audio(audio_files):
    """複数行入力で一気に命名してリネーム"""
    print("=== MP3ファイル命名モード（まとめて貼り付け） ===")
    print("作成日時（ダウンロード日時）順で並んだファイル:")
    for i, f in enumerate(audio_files, start=1):
        print(f"{i}: {f.name}")
    print()
    print("上から順に対応する新しい名前を、1行ずつまとめて貼り付けてください。")
    print("行数がファイル数より少ない場合、その分だけリネームします。")
    print("何も入力せずに Enter すると終了します。")
    print()
    print("例:")
    print("  hello_01")
    print("  hello_02")
    print("  hello_03")
    print("  ...")
    print()
    print("貼り終わったら、最後に空行（何も打たずに Enter）を入力してください。")
    print("--------------------------------------------------")

    # 複数行入力を受け取る。[web:15][web:20]
    new_names = []
    while True:
        line = input()
        if not line.strip():
            break
        new_names.append(line.strip())

    if not new_names:
        print("新しい名前が入力されなかったので、リネームはスキップします。")
        return audio_files

    # ファイル数と入力数の短い方に合わせる
    n = min(len(audio_files), len(new_names))
    print(f"\n{n}個のファイルをリネームします。")

    renamed_files = []
    for old_file, base_name in zip(audio_files[:n], new_names[:n]):
        # .mp3 が付いてなければ付ける
        if not base_name.lower().endswith(".mp3"):
            base_name += ".mp3"
        new_path = old_file.parent / base_name

        if new_path.exists():
            print(f"警告: {new_path.name} は既に存在するためスキップします。")
            renamed_files.append(old_file)
            continue

        old_file.rename(new_path)
        print(f"リネーム: {old_file.name} -> {new_path.name}")
        renamed_files.append(new_path)

    # リネームされなかった残りも後で合成に使えるように戻す
    if len(audio_files) > n:
        renamed_files.extend(audio_files[n:])

    return renamed_files

def merge_audio_video_random(video_dir, audio_dir, output_dir):
    ffmpeg_path = '/opt/homebrew/bin/ffmpeg'

    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # 1. 音声ファイルをダウンロード日時順で取得
    audio_files = get_audio_files_sorted_by_time(audio_dir)
    if not audio_files:
        print("input_audio に mp3 ファイルがありません。")
        return

    # 2. まとめて命名（ユーザーが複数行貼り付け）
    audio_files = batch_rename_audio(audio_files)

    # 3. 動画ファイルを取得し、ランダムで1つ選択
    video_files = [f for f in os.listdir(video_dir) if f.endswith('.mp4')]
    if not video_files:
        print("動画が見つかりません。")
        return

    selected_video = random.choice(video_files)
    video_path = os.path.join(video_dir, selected_video)
    print(f"\n選ばれた動画: {selected_video}")
    
    # 4. 合成
    for audio_path in audio_files:
        audio_name = audio_path.name
        # ここを変更
        # output_name = f"merged_{os.path.splitext(audio_name)[0]}.mp4"
        output_name = f"{os.path.splitext(audio_name)[0]}.mp4"  # 'merged_' を削除[web:24][web:28]
        output_path = os.path.join(output_dir, output_name)

        print(f"合成中: {audio_name} -> {output_name}")

        cmd = [
            ffmpeg_path, '-y',
            '-i', video_path,
            '-i', str(audio_path),
            '-filter_complex', 'amix=inputs=2:duration=first',
            '-c:v', 'copy',
            '-c:a', 'aac',
            output_path
        ]

        try:
            subprocess.run(cmd, check=True)
        except subprocess.CalledProcessError as e:
            print(f"Error: {e}")


if __name__ == "__main__":
    merge_audio_video_random('input_video', 'input_audio', 'output')
