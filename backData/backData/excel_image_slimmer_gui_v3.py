import argparse
import io
import os
import shutil
import sys
import tempfile
import zipfile
import subprocess
import time
from pathlib import Path

# GUI
try:
    import tkinter as tk
    from tkinter import filedialog, messagebox
except Exception:
    tk = None
    filedialog = None
    messagebox = None

try:
    from PIL import Image, ImageOps
except Exception as e:
    print("[ERROR] Pillow is not installed. Install with: pip install pillow", file=sys.stderr)
    sys.exit(1)

SUPPORTED_IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff"}

def human_size(num_bytes: int) -> str:
    for unit in ["B", "KB", "MB", "GB"]:
        if num_bytes < 1024.0:
            return f"{num_bytes:.1f}{unit}"
        num_bytes /= 1024.0
    return f"{num_bytes:.1f}TB"

def log_write(log_path: Path, text: str):
    try:
        with log_path.open("a", encoding="utf-8") as f:
            f.write(text.rstrip() + "\n")
    except Exception:
        pass

def downscale_image(im, max_long_edge: int):
    w, h = im.size
    long_edge = max(w, h)
    if long_edge <= max_long_edge:
        return im
    scale = max_long_edge / float(long_edge)
    new_w = max(1, int(w * scale))
    new_h = max(1, int(h * scale))
    return im.resize((new_w, new_h), Image.LANCZOS)

def optimize_png(im, has_alpha: bool):
    out = io.BytesIO()
    save_params = dict(optimize=True, compress_level=9)
    try:
        if not has_alpha:
            im_q = im.convert("RGB").quantize(colors=256, method=Image.FASTOCTREE, kmeans=0)
            im_q.save(out, format="PNG", **save_params)
        else:
            im_rgba = im.convert("RGBA")
            im_rgba.save(out, format="PNG", **save_params)
    except Exception:
        im.save(out, format="PNG", **save_params)
    return out.getvalue()

def optimize_jpeg(im, jpeg_quality: int, progressive: bool):
    out = io.BytesIO()
    im_rgb = im.convert("RGB")
    im_rgb.save(out, format="JPEG", quality=jpeg_quality, optimize=True, progressive=progressive)
    return out.getvalue()

def process_media_file(path: Path, max_long_edge: int, jpeg_quality: int, progressive_jpeg: bool, log_path: Path) -> int:
    ext = path.suffix.lower()
    if ext not in SUPPORTED_IMAGE_EXTS:
        return 0
    try:
        with Image.open(path) as im:
            try:
                im = ImageOps.exif_transpose(im)
            except Exception:
                pass
            has_alpha = (im.mode in ("RGBA", "LA")) or (("transparency" in im.info) if hasattr(im, "info") else False)
            im2 = downscale_image(im, max_long_edge)

            original_bytes = path.read_bytes()
            if ext in (".jpg", ".jpeg"):
                new_bytes = optimize_jpeg(im2, jpeg_quality=jpeg_quality, progressive=progressive_jpeg)
            elif ext == ".png":
                new_bytes = optimize_png(im2, has_alpha)
            elif ext in (".bmp", ".tif", ".tiff"):
                out = io.BytesIO()
                try:
                    if ext in (".tif", ".tiff"):
                        im2.save(out, format="TIFF", compression="tiff_lzw")
                    else:
                        im2.convert("RGB").save(out, format="BMP")
                    new_bytes = out.getvalue()
                except Exception:
                    new_bytes = original_bytes
            else:
                new_bytes = original_bytes

            if len(new_bytes) < len(original_bytes):
                path.write_bytes(new_bytes)
                saved = len(original_bytes) - len(new_bytes)
                log_write(log_path, f"[OK] {path.name}: {human_size(len(original_bytes))} -> {human_size(len(new_bytes))} (saved {human_size(saved)})")
                return saved
            else:
                log_write(log_path, f"[SKIP] {path.name}: no smaller encoding found")
                return 0
    except Exception as e:
        log_write(log_path, f"[WARN] {path.name}: {e}")
        return 0

def slim_xlsx(input_path: Path, output_path: Path, max_long_edge: int, jpeg_quality: int, progressive_jpeg: bool, log_path: Path, ui=None) -> tuple[int, int, int]:
    tmpdir = Path(tempfile.mkdtemp(prefix="xlsx_slim_"))
    total_saved = 0
    image_count = 0
    try:
        with zipfile.ZipFile(input_path, 'r') as zf:
            zf.extractall(tmpdir)
        media_dir = tmpdir / "xl" / "media"

        if media_dir.exists():
            files = [p for p in media_dir.iterdir() if p.is_file()]
            image_count = len(files)
            for i, p in enumerate(files, 1):
                if ui:
                    ui.update_status(f"Processing images... {i}/{image_count}")
                total_saved += process_media_file(p, max_long_edge, jpeg_quality, progressive_jpeg, log_path)
        else:
            log_write(log_path, "[INFO] No xl/media directory found.")

        if ui:
            ui.update_status("Repacking workbook...")
        with zipfile.ZipFile(output_path, 'w', compression=zipfile.ZIP_DEFLATED) as zf_out:
            for folder, _, files in os.walk(tmpdir):
                for file in files:
                    full_path = Path(folder) / file
                    rel_path = full_path.relative_to(tmpdir)
                    zf_out.write(full_path, arcname=str(rel_path).replace(os.sep, "/"))

        return input_path.stat().st_size, output_path.stat().st_size, image_count
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)

class ProgressUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Excel Image Slimmer")
        self.root.geometry("420x140")
        self.root.resizable(False, False)

        self.label = tk.Label(self.root, text="Ready", anchor="w", justify="left")
        self.label.pack(fill="x", padx=14, pady=(16, 6))

        self.progress = tk.Label(self.root, text="...", anchor="w", justify="left")
        self.progress.pack(fill="x", padx=14, pady=(0, 6))

        self.note = tk.Label(self.root, text="창을 닫지 마세요. 완료 후 자동으로 안내됩니다.", fg="gray")
        self.note.pack(padx=14, pady=(0, 6))

        self.root.update()

    def update_status(self, text: str):
        self.progress.config(text=text)
        self.root.update_idletasks()
        # Small tick to keep UI responsive
        self.root.after(10)
        self.root.update()

    def close(self):
        try:
            self.root.destroy()
        except Exception:
            pass

def open_in_explorer_select(path: Path):
    try:
        subprocess.run(["explorer", "/select,", str(path)], check=False)
    except Exception:
        try:
            os.startfile(path.parent)  # type: ignore[attr-defined]
        except Exception:
            pass

def run_gui_flow(default_max_edge=1400, default_jpeg_quality=80, progressive=True):
    if tk is None or filedialog is None or messagebox is None:
        print("[ERROR] GUI components unavailable.", file=sys.stderr)
        sys.exit(2)

    root = tk.Tk()
    root.withdraw()

    filetypes = [("Excel Workbook", "*.xlsx"), ("Excel Macro-Enabled Workbook", "*.xlsm")]
    in_file = filedialog.askopenfilename(title="용량 줄일 엑셀 파일 선택", filetypes=filetypes)
    if not in_file:
        messagebox.showinfo("취소됨", "파일 선택이 취소되었습니다.")
        sys.exit(0)

    in_path = Path(in_file)
    if in_path.suffix.lower() not in (".xlsx", ".xlsm"):
        messagebox.showerror("에러", "지원하지 않는 형식입니다. .xlsx 또는 .xlsm 을 선택하세요.")
        sys.exit(2)

    log_path = in_path.with_suffix("")  # remove .xlsx
    log_path = log_path.with_name(log_path.name + "_slim_runtime.log")
    try:
        log_path.write_text("[START] Excel Image Slimmer runtime log\n", encoding="utf-8")
    except Exception:
        pass

    # Output name (avoid overwrite)
    base_out = in_path.with_stem(in_path.stem + "_slim")
    out_path = base_out
    idx = 1
    while out_path.exists():
        out_path = in_path.with_stem(in_path.stem + f"_slim({idx})")
        idx += 1

    ui = ProgressUI()
    ui.update_status("Preparing...")
    try:
        before, after, count = slim_xlsx(in_path, out_path, default_max_edge, default_jpeg_quality, progressive, log_path, ui=ui)
        saved = before - after
        pct = (saved / before * 100) if before > 0 else 0.0
        ui.close()
        messagebox.showinfo(
            "완료",
            f"이미지 개수: {count}\n\n"
            f"입력: {in_path.name}\n"
            f"출력: {out_path.name}\n\n"
            f"Before: {human_size(before)}\n"
            f"After : {human_size(after)}\n"
            f"Saved : {human_size(saved)} ({pct:.1f}%)\n\n"
            f"로그: {log_path}"
        )
        open_in_explorer_select(out_path)
    except PermissionError:
        ui.close()
        messagebox.showerror("권한 오류", "파일에 접근할 수 없습니다.\n엑셀에서 열려 있지 않은지 확인하고 다시 시도하세요.")
        sys.exit(3)
    except Exception as e:
        ui.close()
        # Ensure error is logged
        log_write(log_path, f"[FATAL] {e}")
        messagebox.showerror("에러", f"처리 중 오류가 발생했습니다.\n\n{e}\n\n자세한 내용은 로그를 확인하세요:\n{log_path}")
        open_in_explorer_select(log_path)
        sys.exit(3)

def main():
    parser = argparse.ArgumentParser(description="GUI v3 with live progress & runtime logging")
    parser.add_argument("input", nargs="?", help="(Optional) input .xlsx/.xlsm; if omitted, GUI picker is used.")
    parser.add_argument("--max-edge", type=int, default=1400)
    parser.add_argument("--jpeg-quality", type=int, default=80)
    parser.add_argument("--no-progressive", action="store_true")
    args = parser.parse_args()

    progressive = not args.no_progressive

    if not args.input:
        run_gui_flow(default_max_edge=args.max_edge, default_jpeg_quality=args.jpeg_quality, progressive=progressive)
        return

    # CLI path (kept for completeness)
    in_path = Path(args.input)
    if not in_path.exists():
        print(f"[ERROR] Input not found: {in_path}", file=sys.stderr)
        sys.exit(2)
    out_path = in_path.with_stem(in_path.stem + "_slim")
    before, after, count = slim_xlsx(in_path, out_path, args.max_edge, args.jpeg_quality, progressive, in_path.with_suffix(".log"))
    print(f"Done. Images: {count}, Before: {before}, After: {after}")

if __name__ == "__main__":
    main()