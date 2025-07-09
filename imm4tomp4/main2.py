import cv2

# 動画ファイルのパス
video_path = 'input.avi'

# ビデオキャプチャを開く
cap = cv2.VideoCapture(video_path)

# FPS（フレーム毎秒）を取得
fps = cap.get(cv2.CAP_PROP_FPS)

# 20秒ごとのフレーム番号を計算
frame_interval = int(fps * 20)

frame_count = 0
saved_image_count = 0

while cap.isOpened():
    ret, frame = cap.read()
    if not ret:
        break

    if frame_count % frame_interval == 0:
        # 画像を保存
        output_filename = f"output_{saved_image_count:03d}.jpg"
        cv2.imwrite(output_filename, frame)
        saved_image_count += 1

    frame_count += 1

cap.release()
