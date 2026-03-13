from ultralytics import YOLO
from sort import Sort
from skimage import io
import numpy as np
import cv2
import os

def main():
    path_to_model = os.path.join(os.getcwd(), 'Potato_model_best_weights.pt')
    path_to_video = os.path.join(os.getcwd(), 'Inference_data', 'potato.mp4')


    source = cv2.VideoCapture(path_to_video)
    cv2.namedWindow("Potato Detector", cv2.WINDOW_NORMAL)   

    model = YOLO(path_to_model)
    # Vùng đa giác bao trọn băng chuyền TRÁI (khi video được thu nhỏ về 1280x720)
    # Các điểm [x, y] tạo thành một hình chữ nhật nằm xiên theo băng chuyền
    polygon_zone = np.array([[3, 183], [306, 45], [472, 209], [113, 373]], np.int32)
    tracker = Sort(max_age=10)
    CountedIDs = set()
    
    # 7 màu sắc ngẫu nhiên chuẩn (B, G, R) trong OpenCV
    colors = [
        (255, 0, 0),   # Xanh dương
        (0, 255, 0),   # Xanh lá
        (0, 0, 255),   # Đỏ
        (255, 255, 0), # Cyan (lục lam)
        (255, 0, 255), # Tím
        (0, 255, 255), # Vàng
        (255, 128, 0)  # Cam nhạt
    ]


    while source.isOpened():
        detections = np.empty((0, 5))
        success, frame = source.read()

        if not success:
            print("Video frame is empty or processing is complete.")
            break
            
        # 1. TỐI ƯU TỐC ĐỘ: Thu nhỏ kích thước khung hình video lại thành HD (1280x720)
        # Video càng to thì máy xử lý càng chậm, việc thu nhỏ sẽ giúp FPS cải thiện đáng kể
        frame = cv2.resize(frame, (1280, 720))
        
        # 2. TỐI ƯU TỐC ĐỘ: Thêm tham số imgsz=480
        # Mặc định YOLO xài ảnh độ phân giải 640 để quét, giảm xuống 480 sẽ giúp iGPU/CPU chạy mượt hơn rất nhiều!
        result = model(frame, stream=True, imgsz=480)
        result = next(result)

        
        for box in result.boxes:
            x1, y1, x2, y2 = box.xyxy[0]
            x1, y1, x2, y2 = int(x1), int(y1), int(x2), int(y2)
            conf = box.conf[0]
            detections = np.vstack((detections, np.array([x1, y1, x2, y2, conf])))
            

        output = tracker.update(detections)
        for detection in output:
            x1, y1, x2, y2, track_id = detection
            x1, y1, x2, y2, track_id = int(x1), int(y1), int(x2), int(y2), int(track_id)
            cx, cy = int((x1 + x2) / 2), int((y1 + y2) / 2)
            
            # Chọn màu cho tracking box dựa trên track_id mod 7
            color = colors[track_id % 7]
            
            # Bây giờ mới vẽ hình: Ghi mỗi con số ID thay vì chữ "ID:" và dùng màu theo ID
            cv2.rectangle(frame, (x1, y1), (x2, y2), color, 2)
            cv2.putText(frame, f"{track_id}", (x1, y1-5), cv2.FONT_HERSHEY_SIMPLEX, 0.5, color, 2)
            
            # Kiểm tra xem tâm củ khoai (cx, cy) có rơi vào trong vùng đa giác không
            # Nếu >= 0 nghĩa là điểm đó nằm ở CẠNH hoặc TRONG đa giác
            if cv2.pointPolygonTest(polygon_zone, (cx, cy), False) >= 0:
                CountedIDs.add(track_id)
                

        cv2.putText(frame, f"Counted IDs: {len(CountedIDs)}", (10, 40), cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 255, 0), 2)
        # Vẽ đa giác lên màn hình thay vì đường thẳng
        cv2.polylines(frame, [polygon_zone], isClosed=True, color=(0, 255, 0), thickness=2)
        print(f"Counted IDs: {CountedIDs}")
        cv2.imshow("Potato Detector", frame)

        if cv2.waitKey(1) & 0xFF == ord('q'):
            break

    source.release()
    cv2.destroyAllWindows()


if __name__ == "__main__":
    main()



