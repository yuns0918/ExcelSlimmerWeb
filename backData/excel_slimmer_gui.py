#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel Slimmer GUI — 정밀 모드 + 리사이즈 적용 + 완료 후 UI 리셋
- 공격 모드에서 실제로 이미지 "리사이즈 + 변환"을 수행하며,
  PNG→JPG 변환 시 파일명을 변경하고 모든 참조(.rels/VML/[Content_Types])를 동기화합니다.
- 원본은 절대 덮어쓰지 않고 *_slimmed.xlsx/.xlsm 로 저장합니다.
- GUI: 한국어, 전체/개별 진행률, 최종 요약, 완료 후 진행률/현재 파일 표시 리셋(로그 유지).
"""
import sys
import threading
import shutil
import tempfile
import zipfile
from pathlib import Path
import traceback

try:
    from PIL import Image, ImageOps
    PIL_OK = True
except Exception:
    PIL_OK = False

try:
    from lxml import etree
    LXML_OK = True
except Exception:
    LXML_OK = False

try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox, scrolledtext
except Exception:
    tk = None

# ---------- 설정 ----------
JPEG_QUALITY_SAFE = 85
JPEG_QUALITY_AGGRESSIVE = 70
PNG_OPTIMIZE = True
RECOMPRESS_ZIP_LEVEL = 9
MAX_IMAGE_DIM_AGGRESSIVE = (1600, 1600)  # 공격 모드 리사이즈 기준
# --------------------------

def ui_log(widget, msg):
    if widget is None:
        return
    widget.configure(state='normal')
    widget.insert('end', msg + "\n")
    widget.see('end')
    widget.configure(state='disabled')

class Progress:
    def __init__(self, bar, label):
        self.bar = bar
        self.label = label
        self.total = 100
        self.current = 0
        self._lock = threading.Lock()
        self.prefix = ""

    def reset(self, total_steps: int, label_text: str = None, prefix: str = ""):
        with self._lock:
            self.total = max(1, total_steps)
            self.current = 0
            self.prefix = prefix
        if label_text is not None and self.label is not None:
            try:
                self.label.after(0, lambda: self.label.configure(text=label_text))
            except Exception:
                pass
        self._apply()

    def add(self, steps: int = 1):
        with self._lock:
            self.current += steps
            if self.current > self.total:
                self.current = self.total
        self._apply()

    def finish(self):
        with self._lock:
            self.current = self.total
        self._apply()

    def _apply(self):
        if self.bar is not None:
            try:
                self.bar.after(0, lambda: (self.bar.configure(maximum=self.total),
                                           self.bar.configure(value=self.current)))
            except Exception:
                pass
        if self.label is not None:
            try:
                percent = int(self.current * 100 / self.total)
                text = f"{self.prefix} {percent}%" if self.prefix else f"{percent}%"
                self.label.after(0, lambda: self.label.configure(text=text))
            except Exception:
                pass

def reset_ui_widgets(widgets):
    """완료 후: 로그는 유지하고 진행률/현재 파일만 초기화"""
    try:
        overall_bar = widgets.get('overall_bar')
        overall_label = widgets.get('overall_label')
        file_bar = widgets.get('file_bar')
        file_label = widgets.get('file_label')
        run_btn = widgets.get('run_btn')

        if overall_bar is not None:
            overall_bar.after(0, lambda: (overall_bar.configure(maximum=100), overall_bar.configure(value=0)))
        if overall_label is not None:
            overall_label.after(0, lambda: overall_label.configure(text="0%"))
        if file_bar is not None:
            file_bar.after(0, lambda: (file_bar.configure(maximum=100), file_bar.configure(value=0)))
        if file_label is not None:
            file_label.after(0, lambda: file_label.configure(text="파일 진행률 — 0%"))
        if run_btn is not None:
            run_btn.after(0, lambda: run_btn.configure(state='normal'))
    except Exception:
        pass

def make_backup(src: Path, do_backup: bool = True, logger=None):
    if not do_backup:
        if logger: logger("백업 생성을 건너뜁니다 (--no-backup).")
        return
    if src.suffix.lower() not in [".xlsx", ".xlsm"]:
        raise ValueError("지원 확장자는 .xlsx / .xlsm 입니다.")
    backup = src.with_suffix(src.suffix + ".backup")
    shutil.copy2(src, backup)
    if logger: logger(f"백업 생성: {backup.name}")

def unzip_to_temp(src: Path, tempdir: Path) -> Path:
    work = tempdir / "unpacked"
    work.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(src, "r") as zf:
        zf.extractall(work)
    return work

def _replace_if_smaller(orig: Path, temp: Path):
    try:
        if not temp.exists():
            return False
        if temp.stat().st_size < orig.stat().st_size:
            orig.unlink(missing_ok=True)
            temp.rename(orig)
            return True
        else:
            temp.unlink(missing_ok=True)
            return False
    except Exception:
        temp.unlink(missing_ok=True)
        return False

# ---- 정밀 변환 + 참조 동기화 도우미 ----
def convert_png_to_jpg_with_rename_and_resize(p: Path, quality: int, max_dim: tuple[int, int]) -> str | None:
    """
    p: 원본 .png 경로
    - 투명도 없는 PNG만 JPEG으로 변환
    - 변환 전 EXIF 보정 + 리사이즈(thumbnail) 적용
    - 변환되면 파일명을 .jpg로 변경하여 저장, 더 작아질 때만 교체
    return: 새 파일명('image1.jpg') 또는 None
    """
    try:
        with Image.open(p) as im:
            has_alpha = im.mode in ("RGBA", "LA") or ('transparency' in im.info)
            if has_alpha:
                return None
            im = ImageOps.exif_transpose(im)
            im.thumbnail(max_dim, Image.LANCZOS)
            rgb = im.convert("RGB")
            new_name = p.stem + ".jpg"
            tmp_jpeg = p.with_name(new_name + ".tmp")
            rgb.save(tmp_jpeg, format="JPEG", quality=quality, optimize=True, progressive=True)
            if tmp_jpeg.stat().st_size < p.stat().st_size:
                p.unlink(missing_ok=True)
                final = p.with_name(new_name)
                if final.exists():
                    final.unlink(missing_ok=True)
                tmp_jpeg.rename(final)
                return new_name
            else:
                tmp_jpeg.unlink(missing_ok=True)
                return None
    except Exception:
        return None

def update_rels_targets_for_media(unpacked_dir: Path, rename_map: dict[str, str]) -> int:
    base = unpacked_dir / "xl"
    changed = 0
    for rels in base.rglob("_rels/*.rels"):
        try:
            parser = etree.XMLParser(remove_blank_text=True)
            tree = etree.parse(str(rels), parser)
            root = tree.getroot()
            dirty = False
            for rel in root.findall(".//{*}Relationship"):
                tgt = rel.get("Target") or ""
                for old_name, new_name in rename_map.items():
                    if "/media/" + old_name in tgt or tgt.endswith("media/" + old_name):
                        rel.set("Target", tgt.replace(old_name, new_name))
                        dirty = True
            if dirty:
                tree.write(str(rels), encoding="utf-8", xml_declaration=True, pretty_print=True)
                changed += 1
        except Exception:
            pass
    return changed

def update_vml_imagedata_sources(unpacked_dir: Path, rename_map: dict[str, str]) -> int:
    drawings = (unpacked_dir / "xl" / "drawings")
    if not drawings.exists():
        return 0
    changed = 0
    for vml in drawings.glob("vmlDrawing*.vml"):
        try:
            s = vml.read_text(encoding="utf-8", errors="ignore")
            s_new = s
            for old_name, new_name in rename_map.items():
                s_new = s_new.replace(f"/xl/media/{old_name}", f"/xl/media/{new_name}")
            if s_new != s:
                vml.write_text(s_new, encoding="utf-8")
                changed += 1
        except Exception:
            pass
    return changed

def update_content_types_for_renamed(unpacked_dir: Path, rename_map: dict[str, str]) -> int:
    ct_path = unpacked_dir / "[Content_Types].xml"
    if not ct_path.exists():
        return 0
    try:
        parser = etree.XMLParser(remove_blank_text=True)
        tree = etree.parse(str(ct_path), parser)
        root = tree.getroot()
        dirty = False
        for ov in root.findall(".//{*}Override"):
            part = ov.get("PartName") or ""
            for old_name, new_name in rename_map.items():
                if part.endswith("/xl/media/" + old_name):
                    ov.set("PartName", part.replace(old_name, new_name))
                    dirty = True
        if dirty:
            tree.write(str(ct_path), encoding="utf-8", xml_declaration=True, pretty_print=True)
            return 1
    except Exception:
        return 0
    return 0

# ---- 최적화 루틴 ----
def recompress_images_with_sync(unpacked_dir: Path, aggressive: bool, logger=None):
    """
    aggressive=True:
      - JPEG: EXIF 보정 + (필요시) RGB 변환 + 리사이즈 + 재압축
      - PNG(투명 無): EXIF 보정 + 리사이즈 + JPG로 변환(파일명 변경)
      - 참조(.rels/VML/[Content_Types])를 일괄 동기화
    aggressive=False:
      - JPEG: 품질 재압축(작아질 때만 교체)
      - PNG: 무손실 optimize
    return: (changed_count, rename_map)
    """
    if not PIL_OK:
        if logger: logger("Pillow가 없어 이미지 최적화를 건너뜁니다. (pip install pillow)")
        return 0, {}

    media_dir = unpacked_dir / "xl" / "media"
    if not media_dir.exists():
        return 0, {}

    changed = 0
    rename_map: dict[str, str] = {}

    for p in media_dir.iterdir():
        if not p.is_file():
            continue
        ext = p.suffix.lower()
        try:
            if ext in [".jpg", ".jpeg"]:
                with Image.open(p) as im:
                    if aggressive:
                        im = ImageOps.exif_transpose(im)
                        # 리사이즈
                        im.thumbnail(MAX_IMAGE_DIM_AGGRESSIVE, Image.LANCZOS)
                        if im.mode in ("RGBA", "P"):
                            im = im.convert("RGB")
                        tmp = p.with_suffix(p.suffix + ".tmp")
                        im.save(tmp, format="JPEG", quality=JPEG_QUALITY_AGGRESSIVE, optimize=True, progressive=True)
                        if _replace_if_smaller(p, tmp):
                            changed += 1
                    else:
                        tmp = p.with_suffix(p.suffix + ".tmp")
                        im.save(tmp, format="JPEG", quality=JPEG_QUALITY_SAFE, optimize=True, progressive=True)
                        if _replace_if_smaller(p, tmp):
                            changed += 1
            elif ext == ".png":
                if aggressive:
                    new_name = convert_png_to_jpg_with_rename_and_resize(p, quality=JPEG_QUALITY_AGGRESSIVE, max_dim=MAX_IMAGE_DIM_AGGRESSIVE)
                    if new_name:
                        rename_map[p.name] = new_name
                        changed += 1
                else:
                    with Image.open(p) as im:
                        tmp = p.with_suffix(p.suffix + ".tmp")
                        im.save(tmp, format="PNG", optimize=PNG_OPTIMIZE)
                        if _replace_if_smaller(p, tmp):
                            changed += 1
            else:
                continue
        except Exception as e:
            if logger: logger(f"이미지 처리 건너뜀: {p.name} ({e})")

    if rename_map:
        c1 = update_rels_targets_for_media(unpacked_dir, rename_map)
        c2 = update_vml_imagedata_sources(unpacked_dir, rename_map)
        c3 = update_content_types_for_renamed(unpacked_dir, rename_map)
        if logger:
            logger(f"[정밀 동기화] .rels: {c1}개, VML: {c2}개, Content_Types: {c3}개 갱신")

    if changed and logger:
        logger(f"이미지 최적화 완료: {changed}개 (리사이즈/변환/재압축 포함)")
    return changed, rename_map

# ---- 기타 최적화 ----
def disable_pivot_save_data(unpacked_dir: Path, logger=None) -> int:
    if not LXML_OK:
        if logger: logger("lxml이 없어 피벗 캐시 최적화를 건너뜁니다. (pip install lxml)")
        return 0
    piv_dir = unpacked_dir / "xl" / "pivotCache"
    if not piv_dir.exists():
        return 0
    changed = 0
    for p in piv_dir.glob("pivotCacheDefinition*.xml"):
        try:
            parser = etree.XMLParser(remove_comments=False, remove_blank_text=False)
            tree = etree.parse(str(p), parser)
            root = tree.getroot()
            if root.get("saveData") != "0":
                root.set("saveData", "0")
                tree.write(str(p), encoding="utf-8", standalone=True)
                changed += 1
        except Exception as e:
            if logger: logger(f"피벗 캐시 처리 건너뜀: {p.name} ({e})")
    if changed and logger: logger(f"피벗 캐시 saveData=0 적용: {changed}개")
    return changed

def clean_broken_defined_names(unpacked_dir: Path, logger=None) -> int:
    if not LXML_OK:
        if logger: logger("lxml이 없어 이름 정의 정리를 건너뜁니다. (pip install lxml)")
        return 0
    wb_xml = unpacked_dir / "xl" / "workbook.xml"
    if not wb_xml.exists():
        return 0
    try:
        parser = etree.XMLParser(remove_comments=False, remove_blank_text=False)
        tree = etree.parse(str(wb_xml), parser)
        ns = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
        defined_names = tree.xpath("//x:definedNames", namespaces=ns)
        if not defined_names:
            return 0
        dn = defined_names[0]
        removed = 0
        for node in list(dn):
            val = (node.text or "")
            if "#REF!" in val:
                dn.remove(node)
                removed += 1
        if removed:
            tree.write(str(wb_xml), encoding="utf-8", standalone=True)
            if logger: logger(f"#REF! 이름정의 제거: {removed}개")
        return removed
    except Exception as e:
        if logger: logger(f"이름정의 정리 실패(건너뜀): {e}")
        return 0

def remove_thumbnail(unpacked_dir: Path, logger=None) -> bool:
    thumb = unpacked_dir / "docProps" / "thumbnail.jpeg"
    if thumb.exists():
        thumb.unlink(missing_ok=True)
        if logger: logger("문서 썸네일 제거: docProps/thumbnail.jpeg")
        return True
    return False

def remove_customxml(unpacked_dir: Path, logger=None) -> int:
    custom = unpacked_dir / "xl" / "customXml"
    if not custom.exists():
        return 0
    total = sum(f.stat().st_size for f in custom.rglob("*") if f.is_file())
    try:
        shutil.rmtree(custom)
        if logger: logger(f"숨은 XML 데이터(customXml) 제거: {(total/1024/1024):.2f} MB 절감 예상")
        return 1
    except Exception as e:
        if logger: logger(f"customXml 제거 실패: {e}")
        return 0

def remove_docProps_core(unpacked_dir: Path, logger=None) -> bool:
    props_dir = unpacked_dir / "docProps"
    if not props_dir.exists():
        return False
    removed_any = False
    for name in ("custom.xml",):
        p = props_dir / name
        if p.exists():
            p.unlink(missing_ok=True)
            if logger: logger(f"문서 속성 파일 제거: docProps/{name}")
            removed_any = True
    return removed_any

def rezip_max_compress(unpacked_dir: Path, out_path: Path):
    with zipfile.ZipFile(out_path, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=RECOMPRESS_ZIP_LEVEL) as zf:
        for path in sorted(unpacked_dir.rglob("*")):
            if path.is_file():
                arcname = path.relative_to(unpacked_dir).as_posix()
                zf.write(path, arcname)

def get_new_output_path(src_path: Path) -> Path:
    stem = src_path.stem
    suffix = src_path.suffix  # .xlsx/.xlsm
    candidate = src_path.with_name(f"{stem}_slimmed{suffix}")
    i = 1
    while candidate.exists():
        candidate = src_path.with_name(f"{stem}_slimmed({i}){suffix}")
        i += 1
    return candidate

def process_file(src_path: Path, aggressive: bool, no_backup: bool, force_customxml_remove: bool, logger, overall_prog: Progress, file_prog: Progress, summary_dict):
    fname = src_path.name
    logger(f"처리 시작: {fname} (공격 모드={aggressive})")

    # per-file steps
    steps = 10 + (1 if aggressive else 0) + (1 if (aggressive and force_customxml_remove) else 0) + 1
    file_prog.reset(steps, label_text=f"{fname} — 0%", prefix=fname + " —")

    if not src_path.exists():
        logger("파일이 존재하지 않습니다.")
        overall_prog.add(steps); file_prog.finish()
        return

    if src_path.suffix.lower() not in [".xlsx", ".xlsm"]:
        logger("지원 확장자는 .xlsx / .xlsm 입니다. 건너뜁니다.")
        overall_prog.add(steps); file_prog.finish()
        return

    old_size = src_path.stat().st_size

    try:
        # 1) backup
        try:
            make_backup(src_path, do_backup=not no_backup, logger=logger)
        finally:
            overall_prog.add(1); file_prog.add(1)

        with tempfile.TemporaryDirectory() as td:
            tempdir = Path(td)

            # 2) unzip
            unpacked = unzip_to_temp(src_path, tempdir); overall_prog.add(1); file_prog.add(1)

            # 3) images (safe/aggressive) + 정밀 동기화
            if aggressive:
                changed, rename_map = recompress_images_with_sync(unpacked, aggressive=True, logger=logger)
            else:
                changed, rename_map = recompress_images_with_sync(unpacked, aggressive=False, logger=logger)
            overall_prog.add(1); file_prog.add(1)

            # 4) pivot
            disable_pivot_save_data(unpacked, logger=logger); overall_prog.add(1); file_prog.add(1)

            # 5) names
            clean_broken_defined_names(unpacked, logger=logger); overall_prog.add(1); file_prog.add(1)

            # 6) thumbnail
            remove_thumbnail(unpacked, logger=logger); overall_prog.add(1); file_prog.add(1)

            # 7) docProps
            remove_docProps_core(unpacked, logger=logger); overall_prog.add(1); file_prog.add(1)

            # 8) (공격 모드에서 추가 작업은 3번에서 처리)
            if aggressive:
                overall_prog.add(1); file_prog.add(1)

            # 9) customXml (옵션)
            if aggressive and force_customxml_remove:
                remove_customxml(unpacked, logger=logger); overall_prog.add(1); file_prog.add(1)

            # 10) rezip
            out_tmp = tempdir / ("slimmed" + src_path.suffix)
            rezip_max_compress(unpacked, out_tmp); overall_prog.add(1); file_prog.add(1)

            # 11) save as new *_slimmed file (항상 새 파일 생성)
            try:
                new_size = out_tmp.stat().st_size
                out_path = get_new_output_path(src_path)
                shutil.copy2(out_tmp, out_path)
                saved_mb = max(0.0, (old_size - new_size) / (1024*1024))
                pct = max(0.0, (1 - new_size/old_size) * 100) if old_size > 0 else 0.0
                if new_size < old_size:
                    logger(f"완료: {fname} → {out_path.name}  ⟶  {saved_mb:.2f} MB 절감 ({pct:.1f}%)")
                else:
                    logger(f"완료: {fname} → {out_path.name}  ⟶  절감 없음 (원본 대비 변동 없음 또는 증가)")
                summary_dict['files'].append((fname, out_path.name, old_size, new_size, saved_mb, pct))
                summary_dict['saved_bytes'] += max(0, (old_size - new_size))
                summary_dict['original_bytes'] += old_size
            finally:
                overall_prog.add(1); file_prog.add(1)

    except Exception:
        logger("오류 발생:\n" + traceback.format_exc())
        overall_prog.add(max(0, steps - (overall_prog.current % (steps+1))))
    finally:
        file_prog.finish()  # 100%

def run_processing(files, aggressive, no_backup, force_customxml, widgets):
    log_box = widgets['log']
    run_button = widgets['run_btn']
    overall_bar = widgets['overall_bar']
    overall_label = widgets['overall_label']
    file_bar = widgets['file_bar']
    file_label = widgets['file_label']

    overall = Progress(overall_bar, overall_label)
    perfile = Progress(file_bar, file_label)

    total_steps = len(files) * (10 + (1 if aggressive else 0) + (1 if (aggressive and force_customxml) else 0) + 1)
    overall.reset(total_steps, label_text="0%")

    summary = {'files': [], 'saved_bytes': 0, 'original_bytes': 0}

    try:
        for f in files:
            process_file(Path(f), aggressive, no_backup, force_customxml,
                         logger=lambda m: ui_log(log_box, m),
                         overall_prog=overall, file_prog=perfile, summary_dict=summary)
    finally:
        overall.finish()
        if run_button:
            try:
                run_button.configure(state='normal')
            except Exception:
                pass

        total_saved_mb = summary['saved_bytes'] / (1024*1024) if summary['saved_bytes'] else 0.0
        avg_pct = 0.0
        if summary['files']:
            avg_pct = sum(pct for _, _, _, _, _, pct in summary['files']) / len(summary['files'])

        ui_log(log_box, "----------------------------------------------------")
        ui_log(log_box, f"총 절감: {total_saved_mb:.2f} MB")
        ui_log(log_box, f"평균 절감율: {avg_pct:.1f}%")
        ui_log(log_box, "파일별 결과:")
        for fname, outname, old_b, new_b, saved_mb, pct in summary['files']:
            old_mb = old_b/(1024*1024)
            new_mb = new_b/(1024*1024)
            ui_log(log_box, f" - {fname} → {outname}: {old_mb:.2f} MB → {new_mb:.2f} MB  (절감 {saved_mb:.2f} MB, {pct:.1f}%)")

        try:
            messagebox.showinfo("완료", f"총 절감: {total_saved_mb:.2f} MB\n평균 절감율: {avg_pct:.1f}%")
        except Exception:
            pass

        # 완료 후 UI 리셋 (로그 유지)
        try:
            reset_ui_widgets(widgets)
        except Exception:
            pass

def choose_files_and_run(root, widgets):
    files = filedialog.askopenfilenames(title="엑셀 파일 선택 (.xlsx/.xlsm)", filetypes=[("Excel files", "*.xlsx *.xlsm")])
    if not files:
        return
    aggressive = bool(root.aggressive_var.get())
    no_backup = bool(root.nobackup_var.get())
    force_custom = bool(root.force_custom_var.get())
    widgets['run_btn'].configure(state='disabled')
    threading.Thread(target=run_processing,
                     args=(files, aggressive, no_backup, force_custom, widgets),
                     daemon=True).start()

def build_gui_and_run(initial_files=None):
    if tk is None:
        return

    root = tk.Tk()
    root.title("Excel Slimmer — 용량 최적화 도구 (정밀 모드)")
    root.geometry("900x640")

    frm = ttk.Frame(root, padding=10)
    frm.pack(fill='both', expand=True)

    opts = ttk.Frame(frm)
    opts.pack(fill='x', pady=(0,6))

    root.aggressive_var = tk.IntVar(value=0)
    root.force_custom_var = tk.IntVar(value=0)
    root.nobackup_var = tk.IntVar(value=0)

    ttk.Checkbutton(opts, text="공격 모드 (이미지 리사이즈 + 변환)", variable=root.aggressive_var).pack(side='left', padx=6)
    ttk.Checkbutton(opts, text="숨은 XML 데이터 삭제 (주의)", variable=root.force_custom_var).pack(side='left', padx=6)
    ttk.Checkbutton(opts, text="백업 안 만들기 (.backup) (비추천)", variable=root.nobackup_var).pack(side='left', padx=6)

    # 전체 진행률
    overall_frame = ttk.Frame(frm)
    overall_frame.pack(fill='x', pady=(2,4))
    overall_label = ttk.Label(overall_frame, text="0%")
    overall_label.pack(side='right')
    overall_bar = ttk.Progressbar(overall_frame, mode='determinate')
    overall_bar.pack(side='left', fill='x', expand=True, padx=(0,8))

    # 파일별 진행률
    file_frame = ttk.Frame(frm)
    file_frame.pack(fill='x', pady=(0,6))
    file_label = ttk.Label(file_frame, text="파일 진행률 — 0%")
    file_label.pack(side='right')
    file_bar = ttk.Progressbar(file_frame, mode='determinate')
    file_bar.pack(side='left', fill='x', expand=True, padx=(0,8))

    run_btn = ttk.Button(frm, text="파일 선택 후 실행 (정밀 모드)", command=lambda: choose_files_and_run(root, {
        'log': log_box, 'run_btn': run_btn,
        'overall_bar': overall_bar, 'overall_label': overall_label,
        'file_bar': file_bar, 'file_label': file_label
    }))
    run_btn.pack(fill='x', pady=(0,6))

    log_box = scrolledtext.ScrolledText(frm, state='disabled', height=20)
    log_box.pack(fill='both', expand=True)

    # 드래그앤드롭으로 전달된 파일이 있으면 자동 실행
    if initial_files:
        run_btn.configure(state='disabled')
        threading.Thread(target=run_processing, args=(initial_files, False, False, False, {
            'log': log_box, 'run_btn': run_btn,
            'overall_bar': overall_bar, 'overall_label': overall_label,
            'file_bar': file_bar, 'file_label': file_label
        }), daemon=True).start()

    root.mainloop()

def main():
    initial_files = [a for a in sys.argv[1:] if not a.startswith('-')]
    build_gui_and_run(initial_files if initial_files else None)

if __name__ == "__main__":
    main()
