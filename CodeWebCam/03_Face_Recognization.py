import cv2
import numpy as np
import json
import os
from datetime import datetime, timedelta
from db import get_connection

CONFIDENCE_THRESHOLD = 75  # hoặc 65 tùy bạn điều chỉnh thử nghiệm

# Đường dẫn model và label
MODEL_PATH = "trainer/trainer.yml"
LABELS_PATH = "trainer/labels.json"
CASCADE_PATH = "haarcascade_frontalface_default.xml"
# Thư mục lưu ảnh chụp khi điểm danh (có thể sửa đường dẫn)
CAPTURED_DIR = os.path.join(os.getcwd(), "Captured")

# Load mô hình đã huấn luyện
recognizer = cv2.face.LBPHFaceRecognizer_create()
recognizer.read(MODEL_PATH)

# Load bộ nhận diện khuôn mặt
face_cascade = cv2.CascadeClassifier(CASCADE_PATH)

# Load mapping ID -> Tên người dùng
if os.path.exists(LABELS_PATH):
    with open(LABELS_PATH, "r") as f:
        label_dict = json.load(f)
    # Đảo ngược lại: ID (int) -> Tên
    id_to_name = {v: k for k, v in label_dict.items()}
else:
    print("Không tìm thấy labels.json. Vui lòng huấn luyện trước.")
    exit()

# Đảm bảo thư mục Captured tồn tại
os.makedirs(CAPTURED_DIR, exist_ok=True)

print("\n[INFO] Chuẩn bị điểm danh.")

# Nhập mã lịch học
ma_lich_hoc = input("\nNhập mã lịch học (ví dụ: LH01): ")

# Kiểm tra và tự động thêm dữ liệu vào LichHoc_SinhVien
try:
    conn = get_connection()
    cursor = conn.cursor()

    # Kiểm tra xem MaLichHoc có tồn tại trong LichHoc không
    cursor.execute(
        "SELECT COUNT(*) FROM [dbo].[LichHoc] WHERE MaLichHoc = ?",
        (ma_lich_hoc,),
    )
    lich_hoc_count = cursor.fetchone()[0]
    if lich_hoc_count == 0:
        print(f"[ERROR] Mã lịch học {ma_lich_hoc} không tồn tại trong bảng LichHoc.")
        exit()

    # Kiểm tra xem LichHoc_SinhVien đã có dữ liệu cho MaLichHoc chưa
    cursor.execute(
        "SELECT COUNT(*) FROM [dbo].[LichHoc_SinhVien] WHERE MaLichHoc = ?",
        (ma_lich_hoc,),
    )
    lich_hoc_sinh_vien_count = cursor.fetchone()[0]

    if lich_hoc_sinh_vien_count == 0:
        # Lấy danh sách tất cả sinh viên từ bảng SinhVien
        cursor.execute("SELECT MaSinhVien FROM [dbo].[SinhVien]")
        sinh_vien_list = [row[0] for row in cursor.fetchall()]
        
        if not sinh_vien_list:
            print("[ERROR] Không có sinh viên nào trong bảng SinhVien.")
            exit()

        # Thêm sinh viên vào LichHoc_SinhVien
        for ma_sinh_vien in sinh_vien_list:
            cursor.execute(
                "INSERT INTO [dbo].[LichHoc_SinhVien] (MaLichHoc, MaSinhVien) VALUES (?, ?)",
                (ma_lich_hoc, ma_sinh_vien),
            )
        conn.commit()
        print(f"[INFO] Đã tự động thêm {len(sinh_vien_list)} sinh viên vào lịch học {ma_lich_hoc} trong bảng LichHoc_SinhVien.")
    else:
        print(f"[INFO] Lịch học {ma_lich_hoc} đã có dữ liệu sinh viên trong LichHoc_SinhVien.")

except Exception as e:
    print(f"[ERROR] Lỗi khi kiểm tra hoặc thêm dữ liệu vào LichHoc_SinhVien: {str(e)}")
    exit()
finally:
    cursor.close()
    conn.close()

# Khởi tạo hoặc lấy phiên điểm danh
try:
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(
        "SELECT ThoiGianBatDau, TrangThai FROM [dbo].[PhienDiemDanh] WHERE MaLichHoc = ?",
        (ma_lich_hoc,),
    )
    phien = cursor.fetchone()

    if phien is None:  # Nếu chưa có phiên điểm danh
        start_time = datetime.now()
        cursor.execute(
            "INSERT INTO [dbo].[PhienDiemDanh] (MaLichHoc, ThoiGianBatDau, TrangThai) "
            "VALUES (?, ?, ?)",
            (ma_lich_hoc, start_time, "Đang mở"),
        )
        conn.commit()
        print(
            f"[INFO] Đã tạo phiên điểm danh mới cho lịch học {ma_lich_hoc} "
            f"với thời gian bắt đầu {start_time}."
        )
    else:
        start_time = phien[0]
        trang_thai = phien[1]
        if trang_thai == "Hoàn tất":
            print(f"[WARNING] Phiên điểm danh cho lịch học {ma_lich_hoc} đã hoàn tất.")
            exit()

except Exception as e:
    print(f"[ERROR] Lỗi khi kiểm tra hoặc tạo phiên điểm danh: {str(e)}")
    exit()
finally:
    cursor.close()
    conn.close()

# Khởi tạo webcam sau khi nhập mã lịch học
cam = cv2.VideoCapture(0)
if not cam.isOpened():
    print("[ERROR] Không thể mở webcam.")
    exit()
cam.set(3, 640)  # Width
cam.set(4, 480)  # Height

# Kích thước tối thiểu để nhận là mặt người
minW = 0.1 * cam.get(3)
minH = 0.1 * cam.get(4)

font = cv2.FONT_HERSHEY_SIMPLEX

print("\n[INFO] Bắt đầu nhận diện. Nhấn ESC để thoát. Phiên sẽ tự động kết thúc sau 15 phút.")

# Tập hợp để theo dõi sinh viên đã được xử lý trong phiên này
processed_students = set()

# Thời gian hết hạn (15 phút từ khi bắt đầu)
end_time = start_time + timedelta(minutes=15)

while True:
    # Kiểm tra thời gian hết hạn
    current_time = datetime.now()
    if current_time > end_time:
        print(f"[INFO] Phiên điểm danh cho lịch học {ma_lich_hoc} đã hết thời gian (15 phút).")
        break

    ret, img = cam.read()
    if not ret:
        print("[ERROR] Không thể đọc từ webcam.")
        break

    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    faces = face_cascade.detectMultiScale(
        gray, scaleFactor=1.2, minNeighbors=5, minSize=(int(minW), int(minH))
    )

    for (x, y, w, h) in faces:
        roi_gray = gray[y : y + h, x : x + w]
        id_pred, confidence = recognizer.predict(roi_gray)

        # Lưu ảnh mới chụp
        file_name = (
            f"{id_to_name.get(id_pred, 'unknown')}_attendance_"
            f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.jpg"
        )
        file_path = os.path.join(CAPTURED_DIR, file_name)
        cv2.imwrite(file_path, img[y : y + h, x : x + w])

        # Confidence càng nhỏ thì càng chính xác
        if confidence < 75:
            ma_sinh_vien = id_to_name.get(id_pred, "Unknown")
            confidence_text = "  {0}%".format(round(60 - confidence))

            # Kiểm tra xem sinh viên đã được điểm danh trong phiên này chưa
            if ma_sinh_vien not in processed_students and ma_sinh_vien != "Unknown":
                try:
                    conn = get_connection()
                    cursor = conn.cursor()

                    # Kiểm tra xem sinh viên có trong LichHoc_SinhVien không
                    cursor.execute(
                        "SELECT COUNT(*) FROM [dbo].[LichHoc_SinhVien] WHERE MaLichHoc = ? AND MaSinhVien = ?",
                        (ma_lich_hoc, ma_sinh_vien),
                    )
                    exists = cursor.fetchone()[0]

                    if exists:
                        # Thêm bản ghi điểm danh vào bảng DiemDanh với đường dẫn ảnh
                        cursor.execute(
                            "INSERT INTO [dbo].[DiemDanh] (MaLichHoc, MaSinhVien, ThoiGianDiemDanh, TrangThai, DULieuAnhMoi, GhiChu) "
                            "VALUES (?, ?, ?, ?, ?, ?)",
                            (ma_lich_hoc, ma_sinh_vien, current_time, "Có mặt", file_path, None),
                        )
                        conn.commit()
                        print(f"[INFO] Đã điểm danh sinh viên {ma_sinh_vien} cho lịch học {ma_lich_hoc}.")
                        processed_students.add(ma_sinh_vien)
                    else:
                        print(f"[WARNING] Sinh viên {ma_sinh_vien} không thuộc lịch học {ma_lich_hoc}.")

                except Exception as e:
                    print(f"[ERROR] Lỗi khi điểm danh sinh viên {ma_sinh_vien}: {str(e)}")
                finally:
                    cursor.close()
                    conn.close()

            # Hiển thị thông tin trên frame
            cv2.rectangle(img, (x, y), (x + w, y + h), (0, 255, 0), 2)
            cv2.putText(
                img,
                f"{ma_sinh_vien} {confidence_text}",
                (x, y - 10),
                font,
                0.5,
                (0, 255, 0),
                1,
            )
        else:
            # Nếu không nhận diện được
            cv2.rectangle(img, (x, y), (x + w, y + h), (0, 0, 255), 2)
            cv2.putText(
                img,
                "Unknown",
                (x, y - 10),
                font,
                0.5,
                (0, 0, 255),
                1,
            )

    # Hiển thị frame
    cv2.imshow("Face Recognition", img)

    # Nhấn ESC để thoát
    key = cv2.waitKey(30) & 0xFF
    if key == 27:  # ESC
        print("[INFO] Thoát thủ công bởi người dùng.")
        break

# Thêm dữ liệu cho sinh viên vắng mặt sau khi kết thúc phiên
try:
    conn = get_connection()
    cursor = conn.cursor()

    # Lấy danh sách tất cả sinh viên của lịch học
    cursor.execute(
        "SELECT MaSinhVien FROM [dbo].[LichHoc_SinhVien] WHERE MaLichHoc = ?",
        (ma_lich_hoc,),
    )
    all_students = [row[0] for row in cursor.fetchall()]

    # Tìm sinh viên vắng mặt (không có trong processed_students)
    absent_students = [ma_sinh_vien for ma_sinh_vien in all_students if ma_sinh_vien not in processed_students]

    for ma_sinh_vien in absent_students:
        cursor.execute(
            "INSERT INTO [dbo].[DiemDanh] (MaLichHoc, MaSinhVien, ThoiGianDiemDanh, TrangThai, GhiChu) "
            "VALUES (?, ?, ?, ?, ?)",
            (ma_lich_hoc, ma_sinh_vien, datetime.now(), "Vắng mặt", "Nghỉ không phép"),
        )
        print(f"[INFO] Đã thêm sinh viên {ma_sinh_vien} vào danh sách vắng mặt cho lịch học {ma_lich_hoc}.")

    conn.commit()

except Exception as e:
    print(f"[ERROR] Lỗi khi thêm dữ liệu sinh viên vắng mặt: {str(e)}")
finally:
    cursor.close()
    conn.close()

# Cập nhật trạng thái phiên điểm danh thành "Hoàn tất"
try:
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(
        "UPDATE [dbo].[PhienDiemDanh] SET TrangThai = ? WHERE MaLichHoc = ?",
        ("Hoàn tất", ma_lich_hoc),
    )
    conn.commit()
    print(f"[INFO] Đã cập nhật trạng thái phiên điểm danh cho lịch học {ma_lich_hoc} thành 'Hoàn tất'.")
except Exception as e:
    print(f"[ERROR] Lỗi khi cập nhật trạng thái phiên điểm danh: {str(e)}")
finally:
    cursor.close()
    conn.close()

# Giải phóng tài nguyên
cam.release()
cv2.destroyAllWindows()
print("[INFO] Đã kết thúc phiên điểm danh.")