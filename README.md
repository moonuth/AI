# 🥔 Hệ Thống Đếm Khoai Tây Tự Động — Potato Counter

Hệ thống đếm khoai tây theo thời gian thực trên băng chuyền, sử dụng AI phát hiện vật thể kết hợp thuật toán theo dõi xác suất.

<img width="1356" height="699" alt="image" src="https://github.com/user-attachments/assets/a46b7752-591f-46ee-ab1f-e82210ea9ca3" />

---

## 🧰 Công nghệ sử dụng

| Thành phần | Công nghệ | Vai trò |
|---|---|---|
| Phát hiện vật thể | YOLOv11m (custom trained) | Tìm bounding box khoai tây trong mỗi frame |
| Theo dõi đối tượng | SORT + Kalman Filter | Gán ID nhất quán cho mỗi củ qua nhiều frame |
| Tối ưu ghép cặp | Hungarian Algorithm | Ghép detection mới ↔ track đang tồn tại |
| Xử lý video | OpenCV | Đọc video, vẽ kết quả, hiển thị |
| Vùng đếm | Polygon Zone | Đếm khi tâm bounding box đi vào vùng định sẵn |
| Backend Deep Learning | PyTorch + Ultralytics | Chạy mạng neural YOLO |

---

## ⚙️ Kiến trúc & Cách hoạt động

```
Video (potato.mp4)
       │
       ▼
 [Resize → 1280×720]      ← tối ưu tốc độ xử lý
       │
       ▼
 [YOLOv11 (imgsz=480)]    ← phát hiện khoai → [x1,y1,x2,y2,conf] × N củ
       │
       ▼
 [SORT Tracker]
    ├── Kalman Predict     ← dự đoán vị trí M track hiện có
    ├── IOU Matrix         ← tính độ trùng lặp N×M
    ├── Hungarian Algo     ← ghép tối ưu detection ↔ track
    ├── Kalman Update      ← cập nhật track đã ghép
    └── Quản lý track      ← tạo mới / xóa track quá cũ
       │
       ▼
 [Polygon Zone Test]       ← cx,cy có trong vùng đa giác không?
       │
       ▼
 [CountedIDs (Set)]        ← thêm track_id → tự loại trùng
```

### 1. YOLOv11 — Phát hiện vật thể
Model YOLOv11m được huấn luyện riêng trên ảnh khoai tây. Mỗi frame, YOLO xử lý toàn bộ ảnh trong **một lần forward pass** duy nhất, trả về danh sách bounding box `[x1, y1, x2, y2, confidence]` cho mỗi củ khoai phát hiện được.

### 2. Kalman Filter — Dự đoán vị trí
Mỗi củ khoai đang theo dõi có **vector trạng thái 7 chiều**: `[cx, cy, s, r, vx, vy, vs]` (tọa độ tâm, diện tích, tỉ lệ, và các vận tốc tương ứng). Kalman Filter thực hiện 2 bước mỗi frame:
- **Predict**: Dự đoán vị trí tiếp theo dựa trên vận tốc hiện tại
- **Update**: Kết hợp dự đoán với quan sát thực từ YOLO

Nhờ đó, khi khoai bị che khuất vài frame, hệ thống vẫn duy trì track_id — không bị đếm trùng.

### 3. Hungarian Algorithm — Ghép tối ưu
Mỗi frame cần ghép N detection mới với M track cũ. Thuật toán tính ma trận IOU giữa tất cả cặp, sau đó dùng **Hungarian Algorithm** (qua thư viện `lap`) để tìm cách ghép 1-1 tối ưu nhất, minimize tổng chi phí (1 - IOU).

### 4. Polygon Zone + Set — Đếm không trùng
Kiểm tra tâm `(cx, cy)` của mỗi tracked object có nằm trong vùng đa giác đếm không bằng **Ray Casting**. Dùng Python `set()` để lưu track_id — tự động loại bỏ trùng lặp, mỗi củ khoai chỉ được đếm đúng **1 lần**.

---

## 🚀 Cài đặt & Chạy

### Yêu cầu
- Python **3.9 – 3.12** (không dùng 3.13)
- RAM tối thiểu 4 GB

### Bước 1 — Clone dự án
```bash
git clone https://github.com/moonuth/AI
cd Potato-Counter
```

### Bước 2 — Tạo môi trường ảo (khuyến nghị)
```bash
python -m venv .venv

# Windows
.venv\Scripts\activate

# Mac / Linux
source .venv/bin/activate
```

### Bước 3 — Cài thư viện
```bash
pip install -r requirements.txt
```

### Bước 4 — Đặt file model
Tải `Potato_model_best_weights.pt` từ Google Drive nhóm và đặt vào thư mục gốc dự án.

### Bước 5 — Chạy
```bash
python manual_counter.py
```

> Nhấn **`Q`** để thoát. Cửa sổ hiển thị video với bounding box, ID, vùng đếm và tổng số khoai.

---

## ☁️ Chạy trên Google Colab (GPU — 30+ FPS)

1. Vào [colab.research.google.com](https://colab.research.google.com) → Notebook mới
2. `Runtime` → `Change runtime type` → **T4 GPU** → Save
3. Upload lên Colab: `Potato_model_best_weights.pt`, `sort.py`, `ooo.mp4`
4. Chạy cell:

```python
!pip install ultralytics scikit-image filterpy lap

from ultralytics import YOLO
from sort import Sort
import cv2, numpy as np

model = YOLO('Potato_model_best_weights.pt')
tracker = Sort(max_age=10)
CountedIDs = set()
polygon_zone = np.array([[3,183],[306,45],[472,209],[113,373]], np.int32)
colors = [(255,0,0),(0,255,0),(0,0,255),(255,255,0),(255,0,255),(0,255,255),(255,128,0)]

cap = cv2.VideoCapture('ooo.mp4')
fps = int(cap.get(cv2.CAP_PROP_FPS))
w   = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH))
h   = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
out = cv2.VideoWriter('output_khoai_tay.mp4', cv2.VideoWriter_fourcc(*'mp4v'), fps, (w,h))

while cap.isOpened():
    ret, frame = cap.read()
    if not ret: break
    result = next(model(frame, stream=True, device=0))
    dets = np.empty((0,5))
    for box in result.boxes:
        x1,y1,x2,y2 = map(int, box.xyxy[0])
        dets = np.vstack((dets, [x1,y1,x2,y2,box.conf[0].item()]))
    for det in tracker.update(dets):
        x1,y1,x2,y2,tid = map(int, det)
        color = colors[tid % 7]
        cv2.rectangle(frame,(x1,y1),(x2,y2),color,2)
        cv2.putText(frame,str(tid),(x1,y1-5),cv2.FONT_HERSHEY_SIMPLEX,0.5,color,2)
        cx,cy = (x1+x2)//2,(y1+y2)//2
        if cv2.pointPolygonTest(polygon_zone,(cx,cy),False) >= 0:
            CountedIDs.add(tid)
    cv2.putText(frame,f'Counted: {len(CountedIDs)}',(10,40),cv2.FONT_HERSHEY_SIMPLEX,1,(0,255,0),2)
    cv2.polylines(frame,[polygon_zone],True,(0,255,0),2)
    out.write(frame)

cap.release(); out.release()
print(f'✅ Xong! Tổng: {len(CountedIDs)} củ khoai')
```

5. Sau khi xong → refresh Files panel → download `output_khoai_tay.mp4`

---

## 🔧 Thông số quan trọng

| Thông số | Vị trí | Mặc định | Khi nào đổi |
|---|---|---|---|
| `polygon_zone` | `manual_counter.py` dòng 19 | 4 điểm tại 1280×720 | Đổi góc camera / băng chuyền khác |
| `max_age` | `Sort(max_age=10)` | 10 frames | Tăng nếu khoai hay mất track |
| `imgsz` | `model(..., imgsz=480)` | 480 px | Giảm 320 nếu chậm; tăng 640 nếu cần chính xác hơn |
| `iou_threshold` | `Sort(iou_threshold=0.3)` | 0.3 | Tăng nếu đếm nhầm; giảm nếu mất track |

### Cách đổi vùng đếm (polygon)
1. Dùng ảnh `frame_1280x720.jpg` có sẵn trong thư mục
2. Upload lên [polygonzone.roboflow.com](https://polygonzone.roboflow.com)
3. Nhấn `P` → vẽ vùng → copy tọa độ numpy
4. Dán vào dòng 19 của `manual_counter.py`

---

## 📁 Cấu trúc dự án

```
Potato-Counter/
├── manual_counter.py          # File chính — chạy hệ thống
├── sort.py                    # Thuật toán SORT tracking (Kalman + Hungarian)
├── Potato_model_best_weights.pt  # Trọng số YOLOv11 (tải riêng, không commit Git)
├── requirements.txt           # Thư viện cần cài
├── frame_1280x720.jpg         # Ảnh mẫu để vẽ polygon
├── gen_doc.py                 # Script tạo tài liệu kỹ thuật Word
├── Inference_data/
│   ├── potato.mp4             # Video test chính
│   └── ooo.mp4                # Video dùng cho Colab
└── POTATO_COUNTER_TAI_LIEU_KY_THUAT.docx  # Tài liệu kỹ thuật đầy đủ
```

---

## 👥 Phân công nhóm 5 người

| Thành viên | Vai trò | Nhiệm vụ |
|---|---|---|
| **Lê Hoàng Phúc** | AI Trainer | Huấn luyện / fine-tune model YOLO, export `.pt` mới |
| **Nguyễn Tiến Đức** | Dataset Engineer | Tìm kiếm và kết hợp nhiều dataset khoai tây từ Roboflow Universe, lọc class potato, merge thành 1 dataset thống nhất để model nhận diện chính xác hơn |
| **Bùi Văn Ý** | Local Developer | Chạy & tối ưu code local, điều chỉnh polygon, fix bugs |
| **Thùy Trinh** | QA Tester | Test nhiều video, đo độ chính xác, viết báo cáo lỗi |
| **Phan Tấn Thuận** | Cloud Deployment | Deploy Colab GPU, xuất video demo, chuẩn bị thuyết trình |

---

## 📚 Tham chiếu lý thuyết

- **Russell & Norvig** — *Artificial Intelligence: A Modern Approach* (2021, Pearson)
  - Ch.17: Kalman Filter / Probabilistic Reasoning over Time
  - Ch.21: Deep Learning / CNN / Object Detection
- **Pourret, Naïm, Marcot** — *Bayesian Networks: A Practical Guide* (2008, Wiley)
  - Dynamic Bayes Networks — nền tảng lý thuyết của SORT tracker
- **Alex Bewley** — [SORT: Simple Online and Realtime Tracking](https://github.com/abewley/sort) (2016)
- **Ultralytics** — [YOLOv11 Documentation](https://docs.ultralytics.com)

