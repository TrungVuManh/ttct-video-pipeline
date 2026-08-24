# -*- coding: utf-8 -*-
"""
webapp/app.py — May chu FastAPI cho giao dien web local cua TTCT Video Pipeline.

Bao boc lai make_videos.process_lesson() bang HTTP API + stream log truc tiep
(Server-Sent Events) de trinh duyet co the: chon bai co san hoac upload
kich ban/slide moi, chon giong doc, bam chay, xem tien do theo thoi gian
thuc, roi tai/xem video ket qua.

Chi cho phep CHAY 1 TAC VU TAI 1 THOI DIEM (PowerPoint COM va model TTS
dung chung khong an toan khi chay song song).

Khoi dong: python run_webapp.py (xem file do o thu muc goc repo).
"""
import json
import queue
import re
import sys
import threading
import time
import traceback
import uuid
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional
from urllib.parse import urlencode

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import make_videos as mv  # noqa: E402  (can chen sys.path truoc)

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, StreamingResponse
from pydantic import BaseModel

app = FastAPI(title="TTCT Video Pipeline")

STATIC_DIR = Path(__file__).resolve().parent / "static"

# ----------------------------------------------------------------------------
# Danh sach giong doc (khop voi 14 giong dung san cua VieNeu-TTS v3 Turbo).
# Giu tinh o day (khong doc tu file noi bo cua goi vieneu) de khong phu thuoc
# vao cau truc file noi bo co the doi giua cac phien ban.
# ----------------------------------------------------------------------------
VOICES = [
    {"name": "Minh Đức",   "gender": "male",   "region": "Bắc",   "style": "tin_tuc",    "description": "Nam · Bắc · Phong cách tin tức"},
    {"name": "Phạm Tuyên", "gender": "male",   "region": "Bắc",   "style": "tu_nhien",   "description": "Nam · Bắc · Phong cách tự nhiên (mặc định)"},
    {"name": "Thanh Bình", "gender": "male",   "region": "Bắc",   "style": "doc_truyen", "description": "Nam · Bắc · Phong cách kể chuyện"},
    {"name": "Trúc Ly",    "gender": "female", "region": "Bắc",   "style": "tu_nhien",   "description": "Nữ · Bắc · Phong cách tự nhiên"},
    {"name": "Đoan Trang", "gender": "female", "region": "Bắc",   "style": "tu_nhien",   "description": "Nữ · Bắc · Phong cách tự nhiên"},
    {"name": "Ngọc Linh",  "gender": "female", "region": "Bắc",   "style": "doc_truyen", "description": "Nữ · Bắc · Phong cách kể chuyện"},
    {"name": "Mai Anh",    "gender": "female", "region": "Bắc",   "style": "tin_tuc",    "description": "Nữ · Bắc · Phong cách tin tức"},
    {"name": "Xuân Vĩnh",  "gender": "male",   "region": "Nam",   "style": "tu_nhien",   "description": "Nam · Nam · Phong cách tự nhiên"},
    {"name": "Thái Sơn",   "gender": "male",   "region": "Nam",   "style": "doc_truyen", "description": "Nam · Nam · Phong cách kể chuyện"},
    {"name": "Minh Triết", "gender": "male",   "region": "Nam",   "style": "tin_tuc",    "description": "Nam · Nam · Phong cách tin tức"},
    {"name": "Thục Đoan",  "gender": "female", "region": "Nam",   "style": "doc_truyen", "description": "Nữ · Nam · Phong cách kể chuyện"},
    {"name": "Thùy Dung",  "gender": "female", "region": "Nam",   "style": "tin_tuc",    "description": "Nữ · Nam · Phong cách tin tức"},
    {"name": "Quang Sơn",  "gender": "male",   "region": "Trung", "style": "tu_nhien",   "description": "Nam · Trung · Phong cách tự nhiên"},
    {"name": "Ngọc Trân",  "gender": "female", "region": "Trung", "style": "tu_nhien",   "description": "Nữ · Trung · Phong cách tự nhiên"},
]
VOICE_NAMES = {v["name"] for v in VOICES}
STYLES = {"tu_nhien", "tin_tuc", "doc_truyen"}

# ----------------------------------------------------------------------------
# Broadcast log theo dong, cho nhieu client SSE cung xem 1 tac vu.
# ----------------------------------------------------------------------------
class LineBroadcaster:
    def __init__(self, mirror):
        self._buf = ""
        self._mirror = mirror
        self.lines: list[str] = []
        self._subscribers: list[queue.Queue] = []
        self._lock = threading.Lock()

    def write(self, s):
        if not s:
            return 0
        try:
            self._mirror.write(s)
        except Exception:
            pass
        self._buf += s
        while "\n" in self._buf:
            line, self._buf = self._buf.split("\n", 1)
            self._push(line)
        return len(s)

    def flush(self):
        try:
            self._mirror.flush()
        except Exception:
            pass

    def _push(self, line):
        with self._lock:
            self.lines.append(line)
            for q in self._subscribers:
                q.put(line)

    def subscribe(self):
        q = queue.Queue()
        with self._lock:
            for line in self.lines:
                q.put(line)
            self._subscribers.append(q)
        return q

    def unsubscribe(self, q):
        with self._lock:
            if q in self._subscribers:
                self._subscribers.remove(q)

    def close(self):
        if self._buf:
            self._push(self._buf)
            self._buf = ""
        with self._lock:
            for q in self._subscribers:
                q.put(None)


@dataclass
class Job:
    id: str
    key: str
    voice: str
    style: str
    state: str = "running"  # running | done | error
    error: str = ""
    videos: list = field(default_factory=list)
    log: LineBroadcaster = None
    started_at: float = field(default_factory=time.time)
    finished_at: Optional[float] = None


JOBS: dict[str, Job] = {}
JOBS_LOCK = threading.Lock()
CURRENT_JOB_ID: Optional[str] = None

TTS_LOCK = threading.Lock()
TTS_MODEL = None
TTS_SR = None


def get_tts():
    global TTS_MODEL, TTS_SR
    with TTS_LOCK:
        if TTS_MODEL is None:
            print("Dang tai model VieNeu-TTS lan dau (co the mat vai chuc giay)...", flush=True)
            from vieneu import Vieneu
            TTS_MODEL = Vieneu()
            TTS_SR = TTS_MODEL.sample_rate
            print(f"Da tai xong model. sample_rate={TTS_SR}", flush=True)
    return TTS_MODEL, TTS_SR


def _slugify(s: str) -> str:
    s = re.sub(r"[^\w\s-]", "", s, flags=re.UNICODE).strip().lower()
    s = re.sub(r"[\s-]+", "_", s)
    return s[:40] or "bai_moi"


def _run_job(job: Job, cfg: dict):
    global CURRENT_JOB_ID
    broadcaster = job.log
    old_stdout, old_stderr = sys.stdout, sys.stderr
    sys.stdout = broadcaster
    sys.stderr = broadcaster
    try:
        tts, sr = get_tts()
        paths = mv.process_lesson(job.key, cfg, tts, sr)
        job.videos = [Path(p) for p in paths]
        job.state = "done"
    except Exception as e:
        job.error = f"{type(e).__name__}: {e}"
        traceback.print_exc()
        job.state = "error"
    finally:
        job.finished_at = time.time()
        sys.stdout, sys.stderr = old_stdout, old_stderr
        broadcaster.close()
        with JOBS_LOCK:
            CURRENT_JOB_ID = None


def start_job(key: str, cfg: dict) -> Job:
    global CURRENT_JOB_ID
    with JOBS_LOCK:
        if CURRENT_JOB_ID is not None:
            raise HTTPException(409, "Đang có 1 tác vụ chạy, đợi xong rồi thử lại.")
        job_id = uuid.uuid4().hex[:12]
        job = Job(id=job_id, key=key, voice=cfg["voice"], style=cfg["style"])
        job.log = LineBroadcaster(sys.__stdout__)
        JOBS[job_id] = job
        CURRENT_JOB_ID = job_id
    t = threading.Thread(target=_run_job, args=(job, cfg), daemon=True)
    t.start()
    return job


# ----------------------------------------------------------------------------
# Endpoints
# ----------------------------------------------------------------------------
def _video_info(p: Path) -> dict:
    st = p.stat()
    url = "/media?" + urlencode({"path": str(p.resolve())})
    return {"name": p.name, "url": url, "size_mb": round(st.st_size / 1_000_000, 1), "mtime": st.st_mtime}


@app.get("/", response_class=HTMLResponse)
def index():
    return (STATIC_DIR / "index.html").read_text(encoding="utf-8")


@app.get("/api/voices")
def api_voices():
    return VOICES


@app.get("/api/lessons")
def api_lessons():
    out = []
    for key, cfg in mv.LESSONS.items():
        out_root = cfg["root"] / "VIDEO_OUTPUT" / key
        videos = sorted(out_root.glob("Video *.mp4")) if out_root.exists() else []
        out.append({
            "key": key,
            "voice": cfg["voice"],
            "style": cfg["style"],
            "docx_exists": cfg["docx"].exists(),
            "pptx_exists": cfg["pptx"].exists(),
            "videos": [_video_info(p) for p in videos],
        })
    return out


@app.get("/api/current")
def api_current():
    with JOBS_LOCK:
        return {"job_id": CURRENT_JOB_ID}


class LessonJobRequest(BaseModel):
    voice: Optional[str] = None
    style: Optional[str] = None


@app.post("/api/jobs/lesson/{key}")
def start_lesson_job(key: str, req: LessonJobRequest = LessonJobRequest()):
    if key not in mv.LESSONS:
        raise HTTPException(404, f"Không có bài '{key}'")
    cfg = dict(mv.LESSONS[key])
    if req.voice:
        if req.voice not in VOICE_NAMES:
            raise HTTPException(400, f"Giọng không hợp lệ: {req.voice}")
        cfg["voice"] = req.voice
    if req.style:
        if req.style not in STYLES:
            raise HTTPException(400, f"Phong cách không hợp lệ: {req.style}")
        cfg["style"] = req.style
    if not cfg["docx"].exists() or not cfg["pptx"].exists():
        raise HTTPException(400, "Thiếu file docx/pptx của bài này trên máy.")
    job = start_job(key, cfg)
    return {"job_id": job.id}


@app.post("/api/jobs/upload")
async def start_upload_job(
    title: str = Form(...),
    voice: str = Form(...),
    style: str = Form(...),
    docx: UploadFile = File(...),
    pptx: UploadFile = File(...),
):
    if not (docx.filename or "").lower().endswith(".docx"):
        raise HTTPException(400, "File kịch bản phải là .docx")
    if not (pptx.filename or "").lower().endswith(".pptx"):
        raise HTTPException(400, "File slide phải là .pptx")
    if voice not in VOICE_NAMES:
        raise HTTPException(400, f"Giọng không hợp lệ: {voice}")
    if style not in STYLES:
        raise HTTPException(400, f"Phong cách không hợp lệ: {style}")

    job_id = uuid.uuid4().hex[:8]
    key = f"upload_{_slugify(title)}_{job_id}"
    upload_root = mv.BASE / "web_uploads" / key
    upload_root.mkdir(parents=True, exist_ok=True)
    docx_path = upload_root / "kich_ban.docx"
    pptx_path = upload_root / "slide.pptx"
    docx_path.write_bytes(await docx.read())
    pptx_path.write_bytes(await pptx.read())

    cfg = {"docx": docx_path, "pptx": pptx_path, "root": upload_root, "voice": voice, "style": style}
    job = start_job(key, cfg)
    return {"job_id": job.id}


@app.get("/api/jobs/{job_id}")
def job_status(job_id: str):
    job = JOBS.get(job_id)
    if not job:
        raise HTTPException(404, "Không tìm thấy tác vụ")
    return {
        "id": job.id,
        "key": job.key,
        "state": job.state,
        "error": job.error,
        "voice": job.voice,
        "style": job.style,
        "videos": [_video_info(p) for p in job.videos],
        "started_at": job.started_at,
        "finished_at": job.finished_at,
    }


@app.get("/api/jobs/{job_id}/stream")
def job_stream(job_id: str):
    job = JOBS.get(job_id)
    if not job:
        raise HTTPException(404, "Không tìm thấy tác vụ")

    def gen():
        q = job.log.subscribe()
        try:
            while True:
                line = q.get()
                if line is None:
                    payload = json.dumps({"state": job.state, "error": job.error}, ensure_ascii=False)
                    yield f"event: done\ndata: {payload}\n\n"
                    break
                yield f"data: {json.dumps(line, ensure_ascii=False)}\n\n"
        finally:
            job.log.unsubscribe(q)

    return StreamingResponse(
        gen(), media_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


@app.get("/media")
def media(path: str):
    p = Path(path).resolve()
    try:
        p.relative_to(mv.BASE.resolve())
    except ValueError:
        raise HTTPException(403, "Ngoài phạm vi cho phép")
    if not p.exists() or not p.is_file():
        raise HTTPException(404, "Không tìm thấy file")
    media_type = {
        ".mp4": "video/mp4", ".wav": "audio/wav", ".png": "image/png",
    }.get(p.suffix.lower(), "application/octet-stream")
    return FileResponse(str(p), media_type=media_type, filename=p.name, content_disposition_type="inline")
