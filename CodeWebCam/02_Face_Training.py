import cv2
import numpy as np
from PIL import Image
import os
import json
from db import get_connection  # type: ignore

# Đường dẫn để lưu mô hình và nhãn
labels_file = 'trainer/labels.json'

# Kiểm tra module cv2.face
if not hasattr(cv2, 'face'):
    print("cv2.face module is not available. Cài đặt bằng lệnh: pip install opencv-contrib-python")
    exit()

# Tạo recognizer LBPH
try:
    recognizer = cv2.face.LBPHFaceRecognizer_create()
    print("LBPH Face Recognizer created successfully.")
except AttributeError:
    print("Không thể tạo LBPH Recognizer. Cần cài opencv-contrib-python.")
    exit()

# Load bộ cascade
detector = cv2.CascadeClassifier('haarcascade_frontalface_default.xml')

# Hàm lấy ảnh và nhãn từ database
def getImagesAndLabels():
    faceSamples = []
    ids = []
    label_ids = {}
    current_id = 0

    try:
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT MaSinhVien, DuLieuAnhKhuonMat FROM DuLieuKhuonMat")
        rows = cursor.fetchall()

        for row in rows:
            ma_sinh_vien, folder_path = row
            if not os.path.exists(folder_path):
                print(f"[WARNING] Folder does not exist: {folder_path}")
                continue

            # Gán ID số cho mã sinh viên
            if ma_sinh_vien not in label_ids:
                label_ids[ma_sinh_vien] = current_id
                current_id += 1

            id = label_ids[ma_sinh_vien]

            # Duyệt qua tất cả các ảnh trong thư mục
            for img_file in os.listdir(folder_path):
                img_path = os.path.join(folder_path, img_file)
                if os.path.isfile(img_path) and img_file.lower().endswith(('.jpg', '.jpeg', '.png')):
                    try:
                        PIL_img = Image.open(img_path).convert('L')  # grayscale
                        img_numpy = np.array(PIL_img, 'uint8')

                        # Phát hiện khuôn mặt
                        faces = detector.detectMultiScale(img_numpy)
                        for (x, y, w, h) in faces:
                            face_crop = img_numpy[y:y+h, x:x+w]
                            face_resized = cv2.resize(face_crop, (200, 200))  # Resize đồng bộ
                            faceSamples.append(face_resized)
                            ids.append(id)
                    except Exception as e:
                        print(f"[ERROR] Lỗi đọc ảnh {img_path}: {e}")

        return faceSamples, ids, label_ids
    finally:
        cursor.close()
        conn.close()

# Huấn luyện
print("\n[INFO] Training faces. Please wait ...")
faces, ids, label_ids = getImagesAndLabels()

# Nếu không có dữ liệu
if len(faces) == 0:
    print("Không tìm thấy ảnh khuôn mặt nào để train. Kiểm tra lại dữ liệu trong database.")
    exit()

recognizer.train(faces, np.array(ids))

# Tạo thư mục trainer nếu chưa có
if not os.path.exists('trainer'):
    os.makedirs('trainer')

# Lưu mô hình và ánh xạ
recognizer.write('trainer/trainer.yml')
with open(labels_file, 'w') as f:
    json.dump(label_ids, f)

print(f"\nTraining completed successfully. Trained on {len(label_ids)} users.")
print(f"Saved model to: trainer/trainer.yml")
print(f"Saved label mapping to: {labels_file}")