# TTCT Video Pipeline — Kịch bản `.docx` + Slide `.pptx` → Video bài giảng tiếng Việt

Pipeline tự động biến **kịch bản Word** và **bài trình chiếu PowerPoint** thành **video bài giảng hoàn
chỉnh** có giọng đọc tiếng Việt, dùng mô hình TTS mã nguồn mở
[**VieNeu-TTS**](https://github.com/pnnbao97/VieNeu-TTS) (Apache-2.0) — chạy hoàn toàn trên máy, không
cần API key, không tốn phí.

📖 **Trang hướng dẫn:** https://trungvumanh.github.io/ttct-video-pipeline/

```
kịch bản .docx ─┐
                ├─► [1] tách lời bình theo từng Slide, gom theo từng VIDEO
                │        │
slide .pptx ────┤        ├─► [2] VieNeu-TTS  ──► slide_NN.wav   (1 audio / slide)
                │        │
                └────────┼─► [3] xuất ảnh PNG 1920×1080 ──► slide_NN.png
                         │
                         ├─► [4] ffmpeg: ảnh + audio ──► slide_NN.mp4  (clip từng slide)
                         │
                         └─► [5] nối clip trong cùng VIDEO ──► "Video 1 - Tiêu đề.mp4"
```

**Điểm mạnh:** toàn bộ pipeline **resumable** — mọi bước đều bỏ qua file đã tồn tại, nên chạy lại
sau khi bị ngắt là an toàn và không mất công sinh lại từ đầu.

---

## Mục lục

- [1. Yêu cầu hệ thống](#1-yêu-cầu-hệ-thống)
- [2. Cài đặt](#2-cài-đặt)
- [3. Sắp xếp dữ liệu đầu vào](#3-sắp-xếp-dữ-liệu-đầu-vào)
- [4. Định dạng kịch bản `.docx`](#4-định-dạng-kịch-bản-docx)
- [5. Khai báo bài giảng trong `LESSONS`](#5-khai-báo-bài-giảng-trong-lessons)
- [6. Chọn giọng đọc](#6-chọn-giọng-đọc)
- [7. Chạy pipeline](#7-chạy-pipeline)
- [8. Kết quả đầu ra](#8-kết-quả-đầu-ra)
- [9. Chạy lại / sinh lại một phần](#9-chạy-lại--sinh-lại-một-phần)
- [10. Vá lỗi kịch bản bằng `MANUAL_PATCHES`](#10-vá-lỗi-kịch-bản-bằng-manual_patches)
- [11. Notebook thăm dò](#11-notebook-thăm-dò)
- [12. Xử lý sự cố](#12-xử-lý-sự-cố)
- [13. Cấu trúc mã nguồn](#13-cấu-trúc-mã-nguồn)
- [14. Giấy phép & ghi chú](#14-giấy-phép--ghi-chú)

---

## 1. Yêu cầu hệ thống

| Thành phần | Yêu cầu | Ghi chú |
|---|---|---|
| **Python** | 3.10 – 3.11 | Khuyến nghị 3.11 |
| **Hệ điều hành** | Windows 10/11 | Cần cho nhánh xuất ảnh bằng PowerPoint COM |
| **Microsoft PowerPoint** | Bản desktop bất kỳ | Cho ảnh slide đẹp nhất, đúng như PowerPoint hiển thị |
| **LibreOffice** | *(tuỳ chọn)* | Phương án dự phòng khi không có PowerPoint — chạy được trên Linux/macOS |
| **ffmpeg** | Không cần cài | Gói `imageio-ffmpeg` tự tải sẵn binary |
| **GPU NVIDIA** | Không bắt buộc | Mặc định chạy CPU qua ONNX Runtime |
| **Dung lượng đĩa** | ~1 GB / bài giảng | Audio WAV 48 kHz + ảnh PNG 1920×1080 + clip mp4 |

> **Về tốc độ:** trên CPU, TTS chiếm gần như toàn bộ thời gian chạy — mỗi slide mất khoảng vài chục
> giây tuỳ độ dài lời bình và cấu hình máy, nên một bài vài chục slide thường tính bằng chục phút.
> Có GPU CUDA thì nhanh hơn đáng kể. Vì pipeline resumable nên cứ để chạy nền, ngắt lúc nào cũng được.

---

## 2. Cài đặt

```bash
git clone https://github.com/TrungVuManh/ttct-video-pipeline.git
cd ttct-video-pipeline

# Tạo môi trường ảo (khuyến nghị)
python -m venv .venv
.venv\Scripts\activate          # Windows
# source .venv/bin/activate     # Linux / macOS

pip install -r requirements.txt
```

Lần chạy đầu tiên, `vieneu` sẽ **tự tải trọng số mô hình** về cache của Hugging Face — cần mạng, và
mất một lúc. Các lần sau nạp từ cache nên nhanh hơn nhiều.

**Tăng tốc bằng GPU (tuỳ chọn):** nếu máy có GPU NVIDIA với CUDA ≥ 12.8, cài `torch` **trước** khi
cài `vieneu`:

```bash
pip install torch==2.8.0 torchaudio==2.8.0 --index-url https://download.pytorch.org/whl/cu128
pip install -r requirements.txt
```

Kiểm tra cài đặt xong:

```bash
python -c "from vieneu import Vieneu; t=Vieneu(); print('OK, sample_rate =', t.sample_rate)"
```

---

## 3. Sắp xếp dữ liệu đầu vào

Repo **không chứa** file `.docx` / `.pptx` (tài liệu bài giảng, dung lượng lớn) — bạn tự đặt chúng vào
máy rồi trỏ đường dẫn trong `LESSONS`. Cấu trúc gợi ý:

```
ttct-video-pipeline/           ← thư mục repo, cũng là BASE mặc định
├── make_videos.py
├── export_all_slides.py
├── Làm Video CSCTCH/
│   └── Làm Video CSCTCH/                       ← thư mục "root" của nhóm bài
│       ├── CSCTCH_Kịch bản Video Bài giảng_4.1.docx
│       ├── CSCTCH chuong4_ 4.1.pptx
│       └── VIDEO_OUTPUT/                       ← script tự tạo
└── Làm Video NLTKCH/
    └── Làm Video NLTKCH/
        ├── 4.1 NLTKCH kịch bản cho 3 video.docx
        └── 4.1 NLTKCH_ Tiêu chuẩn thiết kế hầm Antigravity.pptx
```

`BASE` mặc định là **thư mục chứa `make_videos.py`**. Nếu dữ liệu nằm chỗ khác, đặt biến môi trường
`TTCT_BASE` thay vì sửa code:

```powershell
$env:TTCT_BASE = "D:\TTCT 2026\Audio"     # PowerShell
```
```bash
export TTCT_BASE="/mnt/d/TTCT 2026/Audio" # bash
```

---

## 4. Định dạng kịch bản `.docx`

`parse_script()` quét **từng đoạn văn (paragraph)** trong file Word và nhận diện 2 loại dòng:

**Dòng mở đầu một video** — mỗi video sẽ thành 1 file `.mp4` riêng:

```
VIDEO 1: Tổng quan và định nghĩa về không gian ngầm
KỊCH BẢN VIDEO 2: Phân loại công trình ngầm (SLIDE 11-20)
```

- Chữ `KỊCH BẢN` ở đầu là **tuỳ chọn**.
- Phần `(SLIDE 11-20)` ở cuối tiêu đề **tự động bị cắt bỏ** — nó chỉ là chú thích, không đưa vào
  tên file.

**Dòng lời bình của một slide** — mỗi dòng thành 1 file `.wav`:

```
Slide 5: "Xin chào các bạn, hôm nay chúng ta sẽ tìm hiểu về..."
Slide 6: Kết cấu vỏ hầm gồm ba lớp chính, lần lượt là...
```

- Nội dung sau dấu `:` chính là **lời bình được đọc thành tiếng**. Dấu ngoặc kép bao ngoài sẽ được
  gỡ bỏ.
- Số slide phải **khớp với số thứ tự slide trong file `.pptx`** — đây là cầu nối duy nhất giữa kịch
  bản và bài trình chiếu.

**Những dòng bị bỏ qua (không đọc thành tiếng):**

| Dòng | Vì sao bị bỏ |
|---|---|
| `Slide 3: (Slide thông tin môn học, chuyển qua slide sau)` | Toàn bộ nội dung nằm trong ngoặc đơn → ghi chú dẫn dắt, không phải lời bình |
| `Hiển thị: sơ đồ mặt cắt ngang hầm` | Không khớp mẫu `Slide n:` hay `VIDEO n:` |
| Dòng trống, tiêu đề chương, ghi chú khác | Không khớp mẫu |

> ⚠️ **Hệ quả quan trọng:** slide bị bỏ qua sẽ **không có audio, và không xuất hiện trong video**.
> Nếu bạn muốn một slide xuất hiện, nó bắt buộc phải có dòng `Slide n:` với lời bình thực sự.

**Kiểm tra kịch bản trước khi chạy cả pipeline** — in ra danh sách slide đã nhận diện được:

```python
# kiem_tra.py
from make_videos import LESSONS, load_lesson_script

KEY = "CSCTCH_4.1"
for r in load_lesson_script(KEY, LESSONS[KEY]["docx"]):
    print(f"V{r['video']} S{r['slide']:02d} | {len(r['text']):4d} ký tự | {r['text'][:60]}")
```

Đối chiếu số slide in ra với số slide trong PowerPoint; lệch nhau nghĩa là kịch bản thiếu dòng
`Slide n:` nào đó (xem [mục 10](#10-vá-lỗi-kịch-bản-bằng-manual_patches)).

---

## 5. Khai báo bài giảng trong `LESSONS`

Mỗi bài giảng là một mục trong `dict` **`LESSONS`** ở đầu [`make_videos.py`](make_videos.py). Thêm bài
mới = thêm một dòng:

```python
LESSONS = {
    "CSCTCH_4.1": dict(
        docx  = CSCTCH_DIR / "CSCTCH_Kịch bản Video Bài giảng_4.1.docx",
        pptx  = CSCTCH_DIR / "CSCTCH chuong4_ 4.1.pptx",
        root  = CSCTCH_DIR,          # thư mục chứa VIDEO_OUTPUT của bài
        voice = "Minh Đức",          # tên giọng, xem bảng ở mục 6
        style = "tu_nhien",          # phong cách đọc
    ),
    # ... thêm bài khác ở đây
}
```

| Khoá | Ý nghĩa |
|---|---|
| `docx` | Đường dẫn kịch bản Word |
| `pptx` | Đường dẫn bài trình chiếu PowerPoint |
| `root` | Thư mục gốc của bài — kết quả ghi vào `root/VIDEO_OUTPUT/<key>/` |
| `voice` | Tên giọng đọc VieNeu-TTS |
| `style` | Phong cách đọc: `tu_nhien`, `tin_tuc`, hoặc `doc_truyen` |

Các hằng số khác cũng nằm ở đầu file: `FPS = 24`, `IMG_W, IMG_H = 1920, 1080`.

---

## 6. Chọn giọng đọc

VieNeu-TTS v3 Turbo (48 kHz) có **14 giọng dựng sẵn**. `voice` và `style` là **hai tham số độc lập** —
cột "Phong cách gốc" dưới đây chỉ là phong cách mà giọng đó được thu, bạn vẫn có thể ghép giọng
`Minh Đức` với `style="tu_nhien"` (chính là cấu hình đang dùng cho phần lớn bài giảng trong repo này).

| Giọng | Giới tính | Vùng miền | Phong cách gốc |
|---|---|---|---|
| `Phạm Tuyên` *(mặc định của model)* | Nam | Bắc | `tu_nhien` |
| `Minh Đức` | Nam | Bắc | `tin_tuc` |
| `Thanh Bình` | Nam | Bắc | `doc_truyen` |
| `Trúc Ly` | Nữ | Bắc | `tu_nhien` |
| `Đoan Trang` | Nữ | Bắc | `tu_nhien` |
| `Ngọc Linh` | Nữ | Bắc | `doc_truyen` |
| `Mai Anh` | Nữ | Bắc | `tin_tuc` |
| `Xuân Vĩnh` | Nam | Nam | `tu_nhien` |
| `Thái Sơn` | Nam | Nam | `doc_truyen` |
| `Minh Triết` | Nam | Nam | `tin_tuc` |
| `Thục Đoan` | Nữ | Nam | `doc_truyen` |
| `Thùy Dung` | Nữ | Nam | `tin_tuc` |
| `Quang Sơn` | Nam | Trung | `tu_nhien` |
| `Ngọc Trân` | Nữ | Trung | `tu_nhien` |

**Nghe thử trước khi quyết định** — sinh một câu mẫu bằng vài giọng rồi so sánh:

```python
from vieneu import Vieneu

tts = Vieneu()
cau_mau = "Xin chào các bạn, hôm nay chúng ta sẽ tìm hiểu về kết cấu vỏ hầm."

for giong in ["Minh Đức", "Phạm Tuyên", "Trúc Ly", "Mai Anh"]:
    tts.save(tts.infer(cau_mau, voice=giong, style="tu_nhien"), f"thu_{giong}.wav")
```

Liệt kê toàn bộ giọng của model đang cài:

```bash
python -c "from vieneu import Vieneu; [print(v) for v in Vieneu().list_preset_voices()]"
```

---

## 7. Chạy pipeline

### Cách 1 — Chạy thẳng (đơn giản nhất)

```bash
python make_videos.py                       # làm TẤT CẢ các bài trong LESSONS
python make_videos.py CSCTCH_4.1            # chỉ làm 1 bài
python make_videos.py CSCTCH_4.1 CH2_damthep NLTKCH_4.2.1   # làm nhiều bài
```

Script chạy trọn 5 bước cho từng bài và in tiến độ theo thời gian thực:

```
==============================================================================
BAI: CSCTCH_4.1  | giong=Minh Đức | style=tu_nhien
==============================================================================
  32 slide, 3 video: V1=10, V2=10, V3=12
  [tts] 0 audio da co, can sinh 32 audio...
    [tts] 5/32 (slide 05) | 78s
    ...
  [slides] can xuat 32 anh (con lai da co 0)...
  [slides] xuat 32 anh bang PowerPoint COM.
  [clips] 32 clip moi (tong 32).
  [video] TAO: Video 1 - Tổng quan và định nghĩa về không gian ngầm.mp4 (10 slide)
  XONG BAI CSCTCH_4.1: 3 video -> ...\VIDEO_OUTPUT\CSCTCH_4.1
```

Bài nào lỗi thì script **vẫn chạy tiếp các bài còn lại**, rồi tổng kết ở cuối:

```
##############################################################################
TONG KET
##############################################################################
  CSCTCH_4.1       : ('OK', 3)
  CSCTCH_4.2       : ('LOI', 'Thieu pptx: ...')
```

### Cách 2 — Xuất ảnh slide trước, rồi mới chạy (khuyến nghị khi làm nhiều bài)

Mở/đóng PowerPoint COM lặp đi lặp lại dễ gây lỗi `Server execution failed`.
[`export_all_slides.py`](export_all_slides.py) mở PowerPoint **đúng một lần**, xuất ảnh cho **tất cả**
các bài, rồi thoát:

```bash
python export_all_slides.py     # bước 1: lấy toàn bộ ảnh slide
python make_videos.py           # bước 2: TTS + ghép video (ảnh đã có sẵn, được bỏ qua)
```

> ⚠️ `export_all_slides.py` gọi `taskkill /F /IM POWERPNT.EXE` để dọn tiến trình PowerPoint cũ trước
> khi mở phiên mới. **Hãy lưu và đóng mọi file PowerPoint đang mở tay trước khi chạy**, kẻo mất
> nội dung chưa lưu.

> 💡 Chạy với **sandbox tắt / quyền đầy đủ** — PowerPoint COM cần khởi động được tiến trình
> `POWERPNT.EXE` thật.

---

## 8. Kết quả đầu ra

Mỗi bài ghi vào `<root>/VIDEO_OUTPUT/<key>/`:

```
VIDEO_OUTPUT/CSCTCH_4.1/
├── _audio/
│   ├── slide_01.wav ... slide_32.wav      # giọng đọc từng slide (48 kHz)
├── _slide_images/
│   ├── slide_01.png ... slide_32.png      # ảnh slide 1920×1080
├── _clips/
│   ├── slide_01.mp4 ... slide_32.mp4      # clip từng slide (ảnh tĩnh + audio)
├── Video 1 - Tổng quan và định nghĩa về không gian ngầm.mp4    ← 🎬 SẢN PHẨM
├── Video 2 - Phân loại CTN và định nghĩa hầm giao thông.mp4    ← 🎬 SẢN PHẨM
├── Video 3 - Đặc thù vận hành và hệ thống an toàn sinh tồn.mp4 ← 🎬 SẢN PHẨM
└── manifest.json                          # nội dung lời bình + đường dẫn từng slide
```

**Thông số video:** H.264 (`libx264`, `-tune stillimage`) · 1920×1080 · 24 fps · audio AAC 192 kbps ·
`yuv420p` — phát được trên mọi trình duyệt, YouTube, LMS.

Tên file video lấy từ tiêu đề `VIDEO n:` trong kịch bản, sau khi loại bỏ các ký tự Windows cấm
(`\ / : * ? " < > |`).

`manifest.json` là danh sách bản ghi `{video, video_title, slide, text, audio_path}` — dùng để đối
chiếu, làm phụ đề, hoặc kiểm tra xem slide nào đã được đọc.

---

## 9. Chạy lại / sinh lại một phần

Mọi bước đều kiểm tra **file đã tồn tại chưa** trước khi làm. Muốn sinh lại thứ gì, **xoá file đó đi
rồi chạy lại** — script sẽ tự dựng lại đúng phần bị thiếu:

| Muốn sinh lại | Xoá | Ghi chú |
|---|---|---|
| Lời đọc của slide 07 | `_audio/slide_07.wav` **và** `_clips/slide_07.mp4` **và** file `Video n - ....mp4` chứa nó | Phải xoá cả clip, nếu không clip cũ vẫn được dùng |
| Ảnh slide 12 | `_slide_images/slide_12.png` **và** `_clips/slide_12.mp4` **và** video chứa nó | |
| Đổi giọng cả bài | Toàn bộ `_audio/`, `_clips/` và các file `Video *.mp4` | Giữ lại `_slide_images/` cho nhanh |
| Chỉ nối lại video | Các file `Video *.mp4` | Giữ `_clips/` — bước nối chỉ mất vài giây |

Ví dụ sinh lại slide 07 của bài `CSCTCH_4.1` (PowerShell):

```powershell
$d = "D:\TTCT 2026\Audio\Làm Video CSCTCH\Làm Video CSCTCH\VIDEO_OUTPUT\CSCTCH_4.1"
Remove-Item "$d\_audio\slide_07.wav", "$d\_clips\slide_07.mp4"
Remove-Item "$d\Video 1 - *.mp4"
python make_videos.py CSCTCH_4.1
```

---

## 10. Vá lỗi kịch bản bằng `MANUAL_PATCHES`

Khi file `.docx` **thiếu hẳn** một dòng `Slide n:` (lỗi soạn thảo), slide đó sẽ biến mất khỏi video.
Thay vì sửa file Word gốc, khai báo bản vá trong `MANUAL_PATCHES` ở [`make_videos.py`](make_videos.py):

```python
MANUAL_PATCHES = {
    "CH3_tru": [
        {"video": 3, "slide": 33,
         "text": "Cảm ơn các bạn đã theo dõi... Hẹn gặp lại ở video tiếp theo."},
    ],
}
```

Bản vá được `load_lesson_script()` chèn vào đúng vị trí theo số slide. Cơ chế này **cố ý làm thủ
công** — script không tự đoán nội dung còn thiếu, để tránh bịa lời bình không có trong kịch bản.

Ví dụ có sẵn trong repo: bài `CH3_tru` nhảy từ `Slide 32` sang `KỊCH BẢN VIDEO 4`, bỏ sót `Slide 33`
— là slide "Cảm ơn..." kết thúc Video 3. Bản vá dùng **đúng nguyên văn chữ trên slide đó**.

---

## 11. Notebook thăm dò

[`TTS_VieNeu_KichBan_4.1.ipynb`](TTS_VieNeu_KichBan_4.1.ipynb) là bản khám phá từng bước của cùng
pipeline, cho **một bài duy nhất**, có nghe thử audio ngay trong Jupyter.

```bash
jupyter lab TTS_VieNeu_KichBan_4.1.ipynb
```

| Dùng notebook khi | Dùng `make_videos.py` khi |
|---|---|
| Thử giọng, nghe so sánh trực tiếp | Sản xuất hàng loạt |
| Kịch bản mới, chưa chắc parse đúng | Kịch bản đã kiểm chứng |
| Cần xem kết quả từng bước | Cần chạy nền, resumable |

Notebook cũng xử lý được tình huống **số slide trong kịch bản lệch với số slide trong PPTX** qua biến
`MATCHED_SLIDES` (ghi đè tay danh sách slide khớp), và đặt hậu tố `_PARTIAL` vào tên video chưa đủ
slide.

Ngoài ra notebook hỗ trợ thêm **định dạng kịch bản có dòng `Lời bình (Voice-over):` riêng** (như bài
`4.1`), bên cạnh định dạng lời bình viết gộp ngay trên dòng `Slide n:` mà `make_videos.py` dùng.

> Notebook dùng `SILENCE_BETWEEN_SLIDES = 0.8` giây khi ghép audio cả video. `make_videos.py` không
> chèn khoảng lặng — nó nối ở mức clip video.

---

## 12. Xử lý sự cố

### `com_error: Server execution failed` khi xuất ảnh slide

PowerPoint COM không khởi động được. Theo thứ tự:

1. Đóng hết cửa sổ PowerPoint, rồi `taskkill /F /IM POWERPNT.EXE`.
2. Chạy [`export_all_slides.py`](export_all_slides.py) thay vì để `make_videos.py` tự xuất — nó mở
   PowerPoint 1 lần cho tất cả các bài và thử lại tối đa 5 lần, mỗi lần cách nhau 3 giây.
3. Chạy trong phiên desktop có tương tác (không phải SSH/service/sandbox).
4. Mở thử file `.pptx` bằng tay một lần — nếu PowerPoint hỏi kích hoạt bản quyền hay "Protected
   View", COM sẽ treo cho tới khi bạn bấm qua.

> Code **cố ý dùng `DispatchEx` chứ không phải `Dispatch`**: `DispatchEx` buộc tạo tiến trình
> out-of-process mới (`CLSCTX_LOCAL_SERVER`), còn `Dispatch` hay báo `Server execution failed` trong
> phiên không tương tác. Đừng đổi lại.

### `com_error: (-2147221021, 'Operation unavailable', ...)` trong Jupyter

Kernel Jupyter đã khởi tạo COM ở chế độ **MTA** (do event loop asyncio/ZMQ), trong khi PowerPoint COM
cần **STA**. Vì vậy mọi lời gọi COM trong repo này đều chạy trong **một `threading.Thread` riêng** —
thread mới chưa từng khởi tạo COM nên có "apartment" sạch. Nếu bạn viết thêm code COM, giữ nguyên
khuôn mẫu `worker()` + `pythoncom.CoInitialize()` đó.

### Không có PowerPoint / chạy trên Linux

Cài LibreOffice và bảo đảm `soffice` nằm trong `PATH`. Pipeline tự chuyển sang
`export_slides_libreoffice()`: chuyển `.pptx` → PDF rồi render bằng PyMuPDF ở `zoom=2.0`.
Chất lượng render có thể lệch chút so với PowerPoint (font, hiệu ứng).

### `ffmpeg loi:` khi nối clip

Bước nối dùng `-c copy` (không mã hoá lại) nên **mọi clip phải cùng thông số**. Lỗi thường gặp khi
trong `_clips/` còn sót clip cũ tạo bằng cấu hình khác (ví dụ trước khi thêm `-vf scale=1920:1080`).
Cách xử lý: **xoá cả thư mục `_clips/` của bài đó** rồi chạy lại.

### Kết quả ra ít slide hơn mong đợi

Kịch bản có dòng `Slide n:` bị bỏ qua. Chạy đoạn kiểm tra ở [mục 4](#4-định-dạng-kịch-bản-docx) và
đối chiếu với PPTX. Nguyên nhân hay gặp: nội dung nằm trọn trong ngoặc đơn, hoặc thiếu hẳn dòng
`Slide n:` (→ dùng [`MANUAL_PATCHES`](#10-vá-lỗi-kịch-bản-bằng-manual_patches)).

### `AssertionError: Thieu docx / Thieu pptx`

Đường dẫn trong `LESSONS` không đúng. Kiểm tra nhanh:

```python
from make_videos import LESSONS

for k, c in LESSONS.items():
    print("OK  " if c["docx"].exists() and c["pptx"].exists() else "THIẾU", k)
```

Lưu ý tên thư mục trong bộ dữ liệu gốc có **lặp hai lần** (`Làm Video CSCTCH/Làm Video CSCTCH/`) do
giải nén từ `.rar`, và một số tên chứa dấu tiếng Việt lẫn dấu `)` thừa — hãy copy y nguyên.

### Chữ tiếng Việt hiển thị lỗi trên console

`make_videos.py` đã tự bọc `sys.stdout`/`sys.stderr` bằng UTF-8. Nếu vẫn lỗi, đặt
`$env:PYTHONIOENCODING = "utf-8"` trước khi chạy.

### TTS quá chậm

Kiểm tra có đang chạy CPU hay không. Cài `torch` bản CUDA (xem [mục 2](#2-cài-đặt)) để dùng GPU.
Hoặc cứ để chạy nền — pipeline resumable, ngắt giữa chừng không mất dữ liệu đã sinh.

---

## 13. Cấu trúc mã nguồn

| File | Vai trò |
|---|---|
| [`make_videos.py`](make_videos.py) | Toàn bộ pipeline 5 bước + cấu hình `LESSONS`, `MANUAL_PATCHES` |
| [`export_all_slides.py`](export_all_slides.py) | Xuất ảnh slide hàng loạt trong 1 phiên PowerPoint |
| [`TTS_VieNeu_KichBan_4.1.ipynb`](TTS_VieNeu_KichBan_4.1.ipynb) | Notebook thăm dò từng bước |
| [`requirements.txt`](requirements.txt) | Danh sách thư viện |
| [`docs/`](docs/) | Trang hướng dẫn GitHub Pages |

Các hàm chính trong `make_videos.py`:

| Hàm | Nhiệm vụ |
|---|---|
| `parse_script(docx)` | Tách `.docx` → danh sách bản ghi `{video, slide, text}` |
| `load_lesson_script(key, docx)` | `parse_script()` + áp `MANUAL_PATCHES` |
| `export_slides(pptx, nums, dir)` | Xuất ảnh slide còn thiếu (COM → fallback LibreOffice) |
| `make_slide_clip(img, wav, out)` | ffmpeg: ảnh tĩnh + audio → mp4 |
| `concat_clips(clips, out)` | ffmpeg concat demuxer, `-c copy` |
| `process_lesson(key, tts, sr)` | Chạy trọn 5 bước cho 1 bài |

---

## 14. Giấy phép & ghi chú

- **Mô hình TTS:** [VieNeu-TTS](https://github.com/pnnbao97/VieNeu-TTS) — Apache-2.0.
- **Watermark:** VieNeu-TTS mặc định chèn **watermark âm thanh không nghe được**
  (`apply_watermark=True`) vào mọi audio sinh ra, để đánh dấu nguồn gốc AI-generated. Không ảnh hưởng
  chất lượng nghe.
- **Nội dung bài giảng** (`.docx`, `.pptx`, video xuất ra) thuộc về tác giả bài giảng và **không**
  được đưa lên repo này — `.gitignore` chặn sẵn các đuôi file đó.
