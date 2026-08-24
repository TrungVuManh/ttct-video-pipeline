# -*- coding: utf-8 -*-
"""
make_videos.py — Chuyen kich ban (.docx) + PowerPoint (.pptx) thanh video bai giang.

Pipeline cho MOI bai:
  1. Doc kich ban .docx, tach loi binh theo tung Slide, gom theo tung VIDEO.
  2. Sinh giong doc tieng Viet bang VieNeu-TTS (giong/style rieng theo tung bai, xem LESSONS)
     -> 1 wav / slide.
  3. Xuat anh tung slide tu .pptx (PowerPoint COM 1920x1080, fallback LibreOffice+PyMuPDF).
  4. Ghep anh slide + audio -> clip mp4 tung slide.
  5. Noi cac clip trong cung 1 VIDEO -> file mp4 hoan chinh.

Resumable: bo qua audio/anh/clip/video da ton tai. Chay lai an toan.

Dung:
  python make_videos.py                # lam tat ca cac bai
  python make_videos.py CSCTCH_4.1     # chi lam 1 (hoac nhieu) bai theo key
"""
import sys, io, os, re, json, time, argparse, subprocess, threading
from pathlib import Path
from collections import defaultdict, Counter

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", line_buffering=True)
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", line_buffering=True)

# ----------------------------------------------------------------------------
# Cau hinh
# ----------------------------------------------------------------------------
# Thu muc goc chua cac thu muc bai giang. Doi bang bien moi truong TTCT_BASE
# neu ban dat du lieu o noi khac (mac dinh: thu muc chua chinh file nay).
BASE = Path(os.environ.get("TTCT_BASE") or Path(__file__).resolve().parent)
CSCTCH_DIR = BASE / "Làm Video CSCTCH" / "Làm Video CSCTCH"
NLTKCH_DIR = BASE / "Làm Video NLTKCH" / "Làm Video NLTKCH"
CH2_DIR = BASE / "Lam video_NLTKCH_chuong2" / "Lam video_NLTKCH_chuong2"
CH3_DIR = BASE / "Làm video_NLTKCH_chuong3 GIỌNG NỮ)" / "Làm video_NLTKCH_chuong3 GIỌNG NỮ)"

FPS = 24
IMG_W, IMG_H = 1920, 1080

# key -> dict(docx, pptx, root, voice, style)
LESSONS = {
    "CSCTCH_4.1":  dict(docx=CSCTCH_DIR / "CSCTCH_Kịch bản Video Bài giảng_4.1.docx",
                        pptx=CSCTCH_DIR / "CSCTCH chuong4_ 4.1.pptx", root=CSCTCH_DIR,
                        voice="Minh Đức", style="tu_nhien"),
    "CSCTCH_4.2":  dict(docx=CSCTCH_DIR / "CSCTCH_Kịch bản Video Bài giảng_4.2.docx",
                        pptx=CSCTCH_DIR / "CSCTCH chuong4_ 4.2.pptx", root=CSCTCH_DIR,
                        voice="Minh Đức", style="tu_nhien"),
    "CSCTCH_4.3":  dict(docx=CSCTCH_DIR / "CSCTCH_Kịch bản Video Bài giảng_4.3.docx",
                        pptx=CSCTCH_DIR / "CSCTCH chuong4_ 4.3.pptx", root=CSCTCH_DIR,
                        voice="Minh Đức", style="tu_nhien"),
    "NLTKCH_4.1":  dict(docx=NLTKCH_DIR / "4.1 NLTKCH kịch bản cho 3 video.docx",
                        pptx=NLTKCH_DIR / "4.1 NLTKCH_ Tiêu chuẩn thiết kế hầm Antigravity.pptx", root=NLTKCH_DIR,
                        voice="Minh Đức", style="tu_nhien"),
    "NLTKCH_4.2.1": dict(docx=NLTKCH_DIR / "4.2.1 NLTKCH_ kịch bản cho 3 video.docx",
                         pptx=NLTKCH_DIR / "4.2.1 NLTHCH_ Cấu tạo vỏ hầm Antigravity.pptx", root=NLTKCH_DIR,
                         voice="Minh Đức", style="tu_nhien"),
    "NLTKCH_4.2.2": dict(docx=NLTKCH_DIR / "4.2.2 NLTKCH_ kịch bản cho 3 video.docx",
                         pptx=NLTKCH_DIR / "4.2.2 NLTHCH_ Tính toán vỏ hầm Antigravity.pptx", root=NLTKCH_DIR,
                         voice="Minh Đức", style="tu_nhien"),
    # Chuong 2: giong nam Minh Duc, phong cach tu nhien
    "CH2_damthep": dict(docx=CH2_DIR / "NLTHCH_Kịch bản  _chương2_damthep.docx",
                        pptx=CH2_DIR / "NLTKCH_chuong2_damthep.pptx", root=CH2_DIR,
                        voice="Minh Đức", style="tu_nhien"),
    "CH2_btctdul": dict(docx=CH2_DIR / "NLTHCH_Kịch bản _chuong2_btctdul.docx",
                        pptx=CH2_DIR / "NLTKCH_chuong2_btctdul.pptx", root=CH2_DIR,
                        voice="Minh Đức", style="tu_nhien"),
    # Chuong 3: giong nu Mai Anh, phong cach tin tuc (theo yeu cau nguoi dung)
    "CH3_mo":      dict(docx=CH3_DIR / "NLTHCH_Kịch bản chuong3_mo.docx",
                        pptx=CH3_DIR / "NLTKCH_chuong3_mo.pptx", root=CH3_DIR,
                        voice="Mai Anh", style="tin_tuc"),
    "CH3_tru":     dict(docx=CH3_DIR / "NLTHCH_Kịch bản chuong3_tru.docx",
                        pptx=CH3_DIR / "NLTKCH_chuong3_tru.pptx", root=CH3_DIR,
                        voice="Mai Anh", style="tin_tuc"),
}

# Va cham thu cong cho cac loi soan thao kich ban da phat hien (khong tu dong doan):
# CH3_tru thieu han "Slide 33:" trong file .docx (nhay tu Slide 32 sang KICH BAN VIDEO 4 / Slide 34),
# trong khi PPTX slide 33 la slide "Cam on... hen gap lai o video tiep theo" ket thuc Video 3.
# Dung dung nguyen van tren slide do lam loi binh (khong bia noi dung moi).
MANUAL_PATCHES = {
    "CH3_tru": [
        {"video": 3, "slide": 33,
         "text": ("Cảm ơn các bạn đã theo dõi và rất mong các bạn sẽ hào hứng với những nội dung này. "
                  "Hẹn gặp lại các bạn ở video tiếp theo với nội dung của phần Công nghệ thi công trụ.")},
    ],
}

# ----------------------------------------------------------------------------
# 1. Parse kich ban
# ----------------------------------------------------------------------------
VIDEO_RE = re.compile(r"^\s*(?:KỊCH\s*BẢN\s*)?VIDEO\s+(\d+)\s*:\s*(.+)$", re.IGNORECASE)
SLIDE_RE = re.compile(r"^\s*Slide\s+(\d+)\s*:\s*(.+)$", re.IGNORECASE)
SLIDE_RANGE_RE = re.compile(r"\s*\(\s*SLIDE.*?\)\s*$", re.IGNORECASE)
# Ghi chu dan dat kieu "(Slide thong tin mon hoc, chuyen qua slide sau)" - khong phai loi binh,
# khong duoc doc thanh tieng.
NOTE_ONLY_RE = re.compile(r"^\(.*\)$")


def parse_script(docx_path: Path):
    from docx import Document
    doc = Document(str(docx_path))
    records = []
    cur_video, cur_video_title = None, None
    for p in doc.paragraphs:
        text = p.text.strip()
        if not text:
            continue
        m = VIDEO_RE.match(text)
        if m:
            cur_video = int(m.group(1))
            cur_video_title = SLIDE_RANGE_RE.sub("", m.group(2).strip()).strip()
            continue
        m = SLIDE_RE.match(text)
        if m:
            narration = m.group(2).strip().strip('"').strip()
            if NOTE_ONLY_RE.match(narration):
                continue  # ghi chu dan dat, khong phai loi binh - bo qua slide nay
            records.append({
                "video": cur_video,
                "video_title": cur_video_title,
                "slide": int(m.group(1)),
                "text": narration,
            })
    return records


def load_lesson_script(key: str, docx_path: Path):
    """parse_script() + ap dung MANUAL_PATCHES (bu cac loi soan thao da phat hien thu cong)."""
    records = parse_script(docx_path)
    patches = MANUAL_PATCHES.get(key, [])
    if patches:
        titles = {}
        for r in records:
            titles.setdefault(r["video"], r["video_title"])
        for p in patches:
            records.append({
                "video": p["video"],
                "video_title": p.get("video_title") or titles.get(p["video"], f"Video {p['video']}"),
                "slide": p["slide"],
                "text": p["text"],
            })
        records.sort(key=lambda r: r["slide"])
    return records


def sanitize(name: str) -> str:
    name = re.sub(r'[\\/:*?"<>|]', "", name)
    return re.sub(r"\s+", " ", name).strip()

# ----------------------------------------------------------------------------
# 3. Xuat anh slide
# ----------------------------------------------------------------------------
def export_slides_powerpoint(pptx_path, slide_numbers, out_dir, width=IMG_W, height=IMG_H):
    result, box = {}, {}

    def worker():
        import win32com.client, pythoncom
        pythoncom.CoInitialize()
        try:
            # DispatchEx buoc tao instance out-of-process moi (CLSCTX_LOCAL_SERVER);
            # Dispatch thuong bao 'Server execution failed' trong phien khong tuong tac.
            app = win32com.client.DispatchEx("PowerPoint.Application")
            pres = app.Presentations.Open(str(Path(pptx_path).resolve()), WithWindow=False)
            try:
                for n in slide_numbers:
                    out_path = Path(out_dir) / f"slide_{n:02d}.png"
                    pres.Slides(n).Export(str(out_path), "PNG", width, height)
                    result[n] = out_path
            finally:
                pres.Close(); app.Quit()
        except Exception as e:
            box["exc"] = e
        finally:
            pythoncom.CoUninitialize()

    t = threading.Thread(target=worker); t.start(); t.join()
    if "exc" in box:
        raise box["exc"]
    return result


def export_slides_libreoffice(pptx_path, slide_numbers, out_dir, zoom=2.0):
    import shutil, fitz
    soffice = shutil.which("soffice") or shutil.which("soffice.exe")
    if not soffice:
        raise RuntimeError("Khong tim thay LibreOffice (soffice).")
    subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir",
                    str(out_dir), str(pptx_path)], check=True)
    pdf_path = Path(out_dir) / (Path(pptx_path).stem + ".pdf")
    doc = fitz.open(str(pdf_path))
    mat = fitz.Matrix(zoom, zoom)
    paths = {}
    for n in slide_numbers:
        pix = doc[n - 1].get_pixmap(matrix=mat)
        out_path = Path(out_dir) / f"slide_{n:02d}.png"
        pix.save(str(out_path)); paths[n] = out_path
    doc.close()
    return paths


def export_slides(pptx_path, slide_numbers, out_dir):
    """Xuat cac slide con THIEU anh. Tra ve dict {n: path} cho toan bo slide_numbers."""
    out_dir = Path(out_dir); out_dir.mkdir(parents=True, exist_ok=True)
    existing = {n: out_dir / f"slide_{n:02d}.png" for n in slide_numbers
                if (out_dir / f"slide_{n:02d}.png").exists()}
    todo = [n for n in slide_numbers if n not in existing]
    if not todo:
        print(f"    [slides] da co du {len(slide_numbers)} anh, bo qua.")
        return {n: out_dir / f"slide_{n:02d}.png" for n in slide_numbers}
    print(f"    [slides] can xuat {len(todo)} anh (con lai da co {len(existing)})...")
    try:
        got = export_slides_powerpoint(pptx_path, todo, out_dir)
        print(f"    [slides] xuat {len(got)} anh bang PowerPoint COM.")
    except Exception as e1:
        print(f"    [slides] PowerPoint COM loi: {e1} -> thu LibreOffice...")
        got = export_slides_libreoffice(pptx_path, todo, out_dir)
        print(f"    [slides] xuat {len(got)} anh bang LibreOffice.")
    return {n: out_dir / f"slide_{n:02d}.png" for n in slide_numbers}

# ----------------------------------------------------------------------------
# 4-5. Ghep video bang ffmpeg
# ----------------------------------------------------------------------------
import imageio_ffmpeg
FFMPEG = imageio_ffmpeg.get_ffmpeg_exe()


def run_ffmpeg(args):
    cmd = [FFMPEG, "-y", "-loglevel", "error"] + [str(a) for a in args]
    r = subprocess.run(cmd, capture_output=True, text=True)
    if r.returncode != 0:
        raise RuntimeError("ffmpeg loi:\n" + r.stderr)


def make_slide_clip(image_path, audio_path, out_path, fps=FPS):
    run_ffmpeg([
        "-loop", "1", "-i", image_path, "-i", audio_path,
        "-c:v", "libx264", "-tune", "stillimage",
        "-c:a", "aac", "-b:a", "192k",
        "-pix_fmt", "yuv420p", "-vf", f"scale={IMG_W}:{IMG_H}",
        "-shortest", "-r", fps, out_path,
    ])


def concat_clips(clip_paths, out_path):
    list_file = Path(out_path).with_suffix(".txt")
    with open(list_file, "w", encoding="utf-8") as f:
        for p in clip_paths:
            f.write(f"file '{Path(p).resolve().as_posix()}'\n")
    run_ffmpeg(["-f", "concat", "-safe", "0", "-i", list_file, "-c", "copy", out_path])
    list_file.unlink()

# ----------------------------------------------------------------------------
# Xu ly 1 bai
# ----------------------------------------------------------------------------
def process_lesson(key, cfg, tts, sr):
    docx, pptx, root = cfg["docx"], cfg["pptx"], cfg["root"]
    voice, style = cfg["voice"], cfg["style"]
    print("\n" + "=" * 78)
    print(f"BAI: {key}  | giong={voice} | style={style}")
    print("=" * 78)
    assert docx.exists(), f"Thieu docx: {docx}"
    assert pptx.exists(), f"Thieu pptx: {pptx}"

    script = load_lesson_script(key, docx)
    videos = sorted(set(r["video"] for r in script))
    by_video = Counter(r["video"] for r in script)
    titles = {}
    for r in script:
        titles.setdefault(r["video"], r["video_title"])
    print(f"  {len(script)} slide, {len(videos)} video: " +
          ", ".join(f"V{v}={by_video[v]}" for v in videos))

    out_root = root / "VIDEO_OUTPUT" / key
    audio_dir = out_root / "_audio"
    img_dir = out_root / "_slide_images"
    clip_dir = out_root / "_clips"
    for d in (audio_dir, img_dir, clip_dir):
        d.mkdir(parents=True, exist_ok=True)

    # --- Buoc 2: TTS tung slide (resumable) ---
    t0 = time.time()
    need = [r for r in script if not (audio_dir / f"slide_{r['slide']:02d}.wav").exists()]
    print(f"  [tts] {len(script)-len(need)} audio da co, can sinh {len(need)} audio...")
    for i, r in enumerate(need, 1):
        out_path = audio_dir / f"slide_{r['slide']:02d}.wav"
        audio = tts.infer(r["text"], voice=voice, style=style)
        tts.save(audio, out_path)
        if i % 5 == 0 or i == len(need):
            el = time.time() - t0
            print(f"    [tts] {i}/{len(need)} (slide {r['slide']:02d}) | {el:.0f}s")
    for r in script:
        p = audio_dir / f"slide_{r['slide']:02d}.wav"
        r["audio_path"] = str(p)

    # --- Buoc 3: xuat anh slide ---
    slide_nums = sorted(r["slide"] for r in script)
    slide_images = export_slides(pptx, slide_nums, img_dir)

    # --- Buoc 4: clip tung slide ---
    audio_by_slide = {r["slide"]: r["audio_path"] for r in script}
    clip_by_slide = {}
    made = 0
    for n in slide_nums:
        clip = clip_dir / f"slide_{n:02d}.mp4"
        if not clip.exists():
            make_slide_clip(str(slide_images[n]), audio_by_slide[n], str(clip))
            made += 1
        clip_by_slide[n] = clip
    print(f"  [clips] {made} clip moi (tong {len(slide_nums)}).")

    # --- Buoc 5: noi clip theo tung video ---
    slides_by_video = defaultdict(list)
    for r in script:
        slides_by_video[r["video"]].append(r["slide"])
    final_paths = []
    for v in sorted(slides_by_video):
        nums = sorted(slides_by_video[v])
        title = sanitize(titles[v]) or f"Video {v}"
        out_mp4 = out_root / f"Video {v} - {title}.mp4"
        if out_mp4.exists():
            print(f"  [video] da co: {out_mp4.name}")
        else:
            concat_clips([clip_by_slide[n] for n in nums], str(out_mp4))
            print(f"  [video] TAO: {out_mp4.name} ({len(nums)} slide)")
        final_paths.append(out_mp4)

    # manifest
    with open(out_root / "manifest.json", "w", encoding="utf-8") as f:
        json.dump(script, f, ensure_ascii=False, indent=2)

    print(f"  XONG BAI {key}: {len(final_paths)} video -> {out_root}")
    return final_paths


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("lessons", nargs="*", help="key bai can lam (mac dinh: tat ca)")
    args = ap.parse_args()
    keys = args.lessons or list(LESSONS.keys())
    bad = [k for k in keys if k not in LESSONS]
    if bad:
        print("Key khong hop le:", bad, "| hop le:", list(LESSONS.keys())); sys.exit(1)

    print("Nap model VieNeu-TTS...", flush=True)
    t0 = time.time()
    from vieneu import Vieneu
    tts = Vieneu()
    sr = tts.sample_rate
    print(f"Nap xong sau {time.time()-t0:.0f}s | sample_rate={sr}")

    summary = {}
    for k in keys:
        try:
            paths = process_lesson(k, LESSONS[k], tts, sr)
            summary[k] = ("OK", len(paths))
        except Exception as e:
            import traceback; traceback.print_exc()
            summary[k] = ("LOI", str(e))

    print("\n" + "#" * 78 + "\nTONG KET\n" + "#" * 78)
    for k in keys:
        print(f"  {k:16} : {summary.get(k)}")


if __name__ == "__main__":
    main()
