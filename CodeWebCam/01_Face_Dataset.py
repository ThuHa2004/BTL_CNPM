import cv2
import os
from db import get_connection # type: ignore

# Khởi động camera
cam = cv2.VideoCapture(0)
cam.set(3, 640)  # set video width
cam.set(4, 480)  # set video height

# Tải bộ phát hiện khuôn mặt
face_detector = cv2.CascadeClassifier('haarcascade_frontalface_default.xml')

# Nhập mã người dùng
face_id = input('\nNhập mã người dùng và nhấn <Enter>: ')

# Tạo thư mục riêng cho người dùng nếu chưa có
user_folder = os.path.join('dataset', face_id)
if not os.path.exists(user_folder):
    os.makedirs(user_folder)

print("\n[INFO] Đang khởi động chụp khuôn mặt. Nhìn vào camera và đợi...")

count = 0
while True:
    ret, img = cam.read()
    img = cv2.flip(img, 1)  # lật ảnh theo chiều ngang
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    faces = face_detector.detectMultiScale(gray, 1.3, 5)

    for (x, y, w, h) in faces:
        cv2.rectangle(img, (x, y), (x+w, y+h), (0, 255, 0), 2)
        count += 1

        # Lưu ảnh có màu vào thư mục riêng
        face_img = img[y:y+h, x:x+w]
        file_path = os.path.join(user_folder, f"{count}.jpg")
        cv2.imwrite(file_path, face_img)

        # Lưu đường dẫn vào database
        conn = None
        cursor = None
        try:
            conn = get_connection()
            cursor = conn.cursor()
            cursor.execute("EXEC sp_ThemDuLieuKhuonMat ?, ?", (face_id, user_folder))
            conn.commit()
            print(f"[INFO] Đã lưu ảnh {count} và đường dẫn vào database.")
        except Exception as e:
            print(f"[ERROR] Lỗi khi lưu vào database: {str(e)}")
        finally:
            if cursor:
                cursor.close()
            if conn:
                conn.close()

        cv2.imshow('image', img)

    k = cv2.waitKey(10) & 0xff
    if k == 27:  # Nhấn ESC để thoát
        break
    elif count >= 50:  # Chụp 50 ảnh thì dừng
        break

print(f"\n[INFO] Đã lưu {count} ảnh vào thư mục: {user_folder}")
cam.release()
cv2.destroyAllWindows()