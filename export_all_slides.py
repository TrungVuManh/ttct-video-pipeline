# -*- coding: utf-8 -*-
"""
export_all_slides.py — Xuat anh PNG (1920x1080) cho TAT CA cac bai trong 1 phien
PowerPoint COM duy nhat. Mo app 1 lan, lap qua tung pptx, dong tung presentation,
Quit o cuoi. Tranh loi 'Server execution failed' khi tao COM server lap lai.

Resumable: chi xuat slide con THIEU anh. Chay: python export_all_slides.py
(Bat buoc chay voi sandbox TAT de PowerPoint COM khoi dong duoc.)
"""
import sys, time, subprocess, threading
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent))
from make_videos import LESSONS, IMG_W, IMG_H  # module nay tu cau hinh sys.stdout utf-8
from pptx import Presentation


def kill_powerpoint():
    subprocess.run(["taskkill", "/F", "/IM", "POWERPNT.EXE"],
                   capture_output=True, text=True)


def worker(jobs, box):
    """jobs: list (key, pptx_path, out_dir, [slide_numbers_can_xuat])."""
    import win32com.client, pythoncom
    pythoncom.CoInitialize()
    app = None
    try:
        # thu tao app vai lan (server co the can vai giay de san sang)
        last = None
        for attempt in range(5):
            try:
                app = win32com.client.DispatchEx("PowerPoint.Application")
                break
            except Exception as e:
                last = e
                time.sleep(3)
        if app is None:
            raise last
        for key, pptx, out_dir, nums in jobs:
            print(f"  [{key}] mo {Path(pptx).name} ...", flush=True)
            pres = app.Presentations.Open(str(Path(pptx).resolve()), WithWindow=False)
            try:
                for n in nums:
                    out_path = Path(out_dir) / f"slide_{n:02d}.png"
                    pres.Slides(n).Export(str(out_path), "PNG", IMG_W, IMG_H)
                print(f"  [{key}] da xuat {len(nums)} anh -> {out_dir}", flush=True)
            finally:
                pres.Close()
        box["ok"] = True
    except Exception as e:
        box["exc"] = e
    finally:
        try:
            if app is not None:
                app.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()


def main():
    # gom cac job con thieu anh
    jobs = []
    for key, cfg in LESSONS.items():
        pptx, root = cfg["pptx"], cfg["root"]
        out_dir = root / "VIDEO_OUTPUT" / key / "_slide_images"
        out_dir.mkdir(parents=True, exist_ok=True)
        n_slides = len(Presentation(str(pptx)).slides)
        missing = [n for n in range(1, n_slides + 1)
                   if not (out_dir / f"slide_{n:02d}.png").exists()]
        if missing:
            jobs.append((key, pptx, out_dir, missing))
            print(f"{key}: thieu {len(missing)}/{n_slides} anh -> se xuat")
        else:
            print(f"{key}: du {n_slides} anh, bo qua")

    if not jobs:
        print("Tat ca da co du anh slide. Khong can lam gi.")
        return

    print(f"\nKill PowerPoint cu (neu co) va mo phien moi cho {len(jobs)} bai...")
    kill_powerpoint()
    time.sleep(2)

    box = {}
    t = threading.Thread(target=worker, args=(jobs, box))
    t.start(); t.join()
    kill_powerpoint()

    if "exc" in box:
        import traceback
        traceback.print_exception(type(box["exc"]), box["exc"], box["exc"].__traceback__)
        print("\nLOI khi xuat slide.")
        sys.exit(1)
    print("\nHOAN TAT xuat toan bo anh slide.")


if __name__ == "__main__":
    main()
