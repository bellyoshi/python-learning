import os
import subprocess

def convert_avi_to_mp4(root_dir):
    # ルートディレクトリのサブフォルダを再帰的に検索
    for subdir, dirs, files in os.walk(root_dir):
        # サブディレクトリの深さを確認（2階層まで）
        if subdir.count(os.sep) - root_dir.count(os.sep) < 2:
            for file in files:
                # 拡張子がAVIのファイルを対象
                if file.lower().endswith(".avi"):
                    avi_path = os.path.join(subdir, file)
                    # 出力ファイル名を作成（同じディレクトリに同じ名前で.mp4）
                    mp4_path = os.path.splitext(avi_path)[0] + ".mp4"
                    
                    # FFmpegコマンドを構築
                    command = [
                        "c:\bin\ffmpeg.exe", "-i", avi_path, "-vcodec", "libx264", "-acodec", "aac", mp4_path
                    ]
                    
                    # FFmpegコマンドを実行
                    subprocess.run(command, check=True)
                    print(f"Converted: {avi_path} -> {mp4_path}")

# 実行する際に変換したいディレクトリのパスを指定
root_directory = "D:\20240910防犯カメラバックアップ\"
convert_avi_to_mp4(root_directory)
