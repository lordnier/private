import os
import subprocess
import random

def merge_audio_video_random(video_dir, audio_dir, output_dir):
    ffmpeg_path = '/opt/homebrew/bin/ffmpeg'
    
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # 1. 動画ファイルを取得し、ランダムに1つ選ぶ
    video_files = [f for f in os.listdir(video_dir) if f.endswith('.mp4')]
    if not video_files:
        print("動画が見つかりません。")
        return
    
    selected_video = random.choice(video_files)
    video_path = os.path.join(video_dir, selected_video)
    print(f"選ばれた動画: {selected_video}")

    # 2. 全ての音声ファイルを取得し、それぞれ合成する
    audio_files = [f for f in os.listdir(audio_dir) if f.endswith('.mp3')]
    
    for audio_file in audio_files:
        audio_path = os.path.join(audio_dir, audio_file)
        # 命名: merged_音声ファイル名.mp4
        output_name = f"merged_{os.path.splitext(audio_file)[0]}.mp4"
        output_path = os.path.join(output_dir, output_name)
        
        print(f"合成中: {audio_file} -> {output_name}")
        
        cmd = [
            ffmpeg_path, '-y', '-i', video_path, '-i', audio_path,
            '-filter_complex', 'amix=inputs=2:duration=first',
            '-c:v', 'copy', '-c:a', 'aac', output_path
        ]
        
        try:
            subprocess.run(cmd, check=True)
        except subprocess.CalledProcessError as e:
            print(f"Error: {e}")

if __name__ == "__main__":
    # パスは適宜合わせてください
    merge_audio_video_random('input_video', 'input_audio', 'output')
