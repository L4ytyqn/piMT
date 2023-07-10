import cv2
import numpy as np
import tflite_runtime.interpreter as tflite
import time

# ゴミのカテゴリ
CLASSES = ['neutral', 'burn:A', 'pla:B', 'wood:D', 'recycl', 'dan:G', 'alm:E']

# モデルの読み込み
model_path = '/home/pi//bunbetu/model_unquant.tflite'
interpreter = tflite.Interpreter(model_path)
interpreter.allocate_tensors()

input_details = interpreter.get_input_details()
output_details = interpreter.get_output_details()

# タイトルテキスト
title_text = "Garbage Separation Program"

# タイトル画面の表示関数
def show_title_screen(text, duration):
    # ウィンドウのサイズと背景色を設定
    window_width = 800
    window_height = 600
    background_color = (255, 255, 255)  # 白色
    
    # ウィンドウを作成して背景色で塗りつぶす
    title_screen = np.zeros((window_height, window_width, 3), np.uint8)
    title_screen[:] = background_color
    
    # テキストのフォントとスタイルを指定
    font = cv2.FONT_HERSHEY_SIMPLEX
    font_scale = 1
    font_thickness = 2
    
    # テキストの描画位置を計算
    text_size, _ = cv2.getTextSize(text, font, font_scale, font_thickness)
    text_x = (window_width - text_size[0]) // 2
    text_y = (window_height + text_size[1]) // 2
    
    # テキストを描画
    cv2.putText(title_screen, text, (text_x, text_y), font, font_scale, (0, 250, 0), font_thickness, cv2.LINE_AA)
    
    # ウィンドウを表示して指定時間待機
    cv2.imshow('Title Screen', title_screen)
    cv2.waitKey(duration)
    
    # ウィンドウを閉じる
    cv2.destroyAllWindows()

# タイトル画面の表示と指定時間の待機
show_title_screen(title_text, 3000)  # 3秒間表示

# カメラの起動
cap = cv2.VideoCapture(0)

while True:
    ret, frame = cap.read()

    if not ret:
        break

    # 画像の前処理
    resized_frame = cv2.resize(frame, (224, 224))
    normalized_frame = resized_frame / 255.0
    input_data = np.expand_dims(normalized_frame, axis=0).astype(np.float32)

    # 入力データのセットアップ
    interpreter.set_tensor(input_details[0]['index'], input_data)

    # 推論の実行
    interpreter.invoke()

    # 出力データの取得
    output_data = interpreter.get_tensor(output_details[0]['index'])
    predictions = np.squeeze(output_data)

    class_index = np.argmax(predictions)
    confidence = predictions[class_index]

    # 確信度が70%以上の場合のみ分類結果を表示
    if confidence >= 0.7:
        predicted_class = CLASSES[class_index]
        cv2.putText(frame, predicted_class, (10, 30), cv2.FONT_HERSHEY_SIMPLEX, 1, (255, 0, 0), 2)
    else:
        cv2.putText(frame, 'please show the garbage', (10, 30), cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 0, 255), 2)

    # 分類結果の表示
    cv2.imshow('Gomi Classification', frame)

    # キー入力による終了
    if cv2.waitKey(1) & 0xFF == ord('q'):
        break

# 後片付け
cap.release()
cv2.destroyAllWindows()
