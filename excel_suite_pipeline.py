import sys
import threading
import traceback
from pathlib import Path

try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox, scrolledtext
except Exception:  # tkinter may be unavailable in some environments (e.g. headless servers)
    tk = None
    ttk = None
    filedialog = None
    messagebox = None
    scrolledtext = None


def _ensure_module_paths() -> None:
    base = Path(__file__).resolve().parent
    # 다양한 배치에 대응하기 위해, 현재 파일 기준으로 위로 몇 단계 올라가며
    # ExcelCleaner / ExcelImageOptimization / ExcelByteReduce 폴더를 찾는다.
    search_roots: list[Path] = []
    for parent in [base, *base.parents[:3]]:  # base, parent, grand-parent 정도까지
        search_roots.append(parent)

    for root in search_roots:
        for name in ("ExcelCleaner", "ExcelImageOptimization", "ExcelByteReduce"):
            p = root / name
            if p.is_dir():
                sp = str(p)
                if sp not in sys.path:
                    sys.path.insert(0, sp)

    # 기존 Tk/GUI 기반 모듈들이 들어 있는 backData 폴더도 경로에 추가한다.
    backdata = base / "backData"
    if backdata.is_dir():
        sp = str(backdata)
        if sp not in sys.path:
            sys.path.insert(0, sp)


_ensure_module_paths()
try:
    from gui_clean_defined_names_desktop_date import process_file_gui
except ModuleNotFoundError:  # ExcelCleaner 모듈이 없는 환경(예: 웹 서버)에서도 import 가능하게
    process_file_gui = None

try:
    from excel_image_slimmer_gui_v3 import (
        slim_xlsx,
        human_size,
        open_in_explorer_select,
    )
except ModuleNotFoundError:
    slim_xlsx = None

    def human_size(num: int) -> str:
        for unit in ("B", "KB", "MB", "GB", "TB"):
            if num < 1024.0:
                return f"{num:.1f}{unit}"
            num /= 1024.0
        return f"{num:.1f}PB"

    def open_in_explorer_select(path) -> None:
        return

try:
    from excel_slimmer_precision_plus import process_file as precision_process, Progress
except ModuleNotFoundError:
    precision_process = None
    Progress = None
from settings import get_settings, save_settings


def run_image_slim(input_path: Path, max_edge: int, jpeg_quality: int, progressive: bool):
    if slim_xlsx is None:
        raise RuntimeError(
            "이미지 최적화 모듈이 이 환경에 설치되어 있지 않아 '이미지 최적화' 단계를 실행할 수 없습니다."
        )

    base_out = input_path.with_stem(input_path.stem + "_slim")
    out_path = base_out
    idx = 1
    while out_path.exists():
        out_path = input_path.with_stem(input_path.stem + f"_slim({idx})")
        idx += 1
    log_path = input_path.with_name(input_path.stem + "_image_slim.log")
    before, after, count = slim_xlsx(
        input_path,
        out_path,
        max_edge,
        jpeg_quality,
        progressive,
        log_path,
        ui=None,
    )
    return out_path, before, after, count, log_path


def run_precision_step(
    input_path: Path,
    aggressive: bool,
    no_backup: bool,
    do_xml_cleanup: bool,
    force_custom: bool,
    logger,
):
    if precision_process is None or Progress is None:
        raise RuntimeError(
            "Precision Plus 모듈이 이 환경에 설치되어 있지 않아 '정밀 슬리머' 단계를 실행할 수 없습니다."
        )

    overall = Progress(None, None)
    file_prog = Progress(None, None)
    summary = {"files": [], "saved_bytes": 0, "original_bytes": 0}
    precision_process(
        input_path,
        aggressive,
        no_backup,
        do_xml_cleanup,
        force_custom,
        logger,
        overall,
        file_prog,
        summary,
    )
    if summary["files"]:
        _, outname, old_b, new_b, saved_mb, pct = summary["files"][-1]
        out_path = input_path.with_name(outname)
        return out_path, saved_mb, pct, old_b, new_b
    size = input_path.stat().st_size
    return input_path, 0.0, 0.0, size, size


def run_pipeline_core(
    start_path: Path,
    use_clean: bool,
    use_image: bool,
    use_precision: bool,
    aggressive: bool,
    do_xml_cleanup: bool,
    force_custom: bool,
    log,
    set_status,
    show_error,
    on_finished,
) -> None:
    """UI-agnostic pipeline core shared by different front-ends.

    All UI interactions (로그 출력, 상태 표시, 메시지박스, 탐색기 열기 등)는
    콜백으로 주입받고 여기서는 순수하게 파이프라인 로직만 처리한다.
    """

    settings = get_settings()

    def log_info(message: str) -> None:
        """항상 출력하는 로그 (에러/요약 정보)."""

        log(message)

    def log_detail(message: str) -> None:
        """로그 모드가 verbose 일 때만 출력하는 상세 로그."""

        if settings.log_mode == "verbose":
            log(message)

    current = start_path
    intermediate_files = []
    log_files = []
    backup_files = []

    steps = []
    if use_clean:
        steps.append("clean")
    if use_image:
        steps.append("image")
    if use_precision:
        steps.append("precision")

    total = len(steps)
    log_info(f"[INFO] 파이프라인 시작: {start_path.name}, 단계 {total}개")

    for index, step in enumerate(steps, start=1):
        base = (index - 1) * 100.0 / total if total else 0.0
        next_p = index * 100.0 / total if total else 100.0
        try:
            if step == "clean":
                if process_file_gui is None:
                    raise RuntimeError("ExcelCleaner 모듈이 이 환경에 설치되어 있지 않아 '이름 정의 정리' 단계를 실행할 수 없습니다.")
                set_status("이름 정의 정리 중...", base)
                log_info(f"[{index}/{total}] 이름 정의 정리: {current.name}")
                (
                    backup_path,
                    cleaned_path,
                    stats,
                    ts_dir,
                    top_dir,
                ) = process_file_gui(str(current))
                current = Path(cleaned_path)
                if step != steps[-1]:
                    intermediate_files.append(current)
                try:
                    backup_files.append(Path(backup_path))
                except TypeError:
                    # 예상치 못한 타입인 경우에는 조용히 무시
                    pass
                log_detail(f" - 백업: {backup_path}")
                log_detail(f" - 정리본: {cleaned_path}")
                log_detail(
                    " - 통계: total="
                    + str(stats["total"])
                    + ", kept="
                    + str(stats["kept"])
                    + ", removed="
                    + str(stats["removed"])
                )
            elif step == "image":
                set_status("이미지 최적화 중...", base)
                log_info(f"[{index}/{total}] 이미지 최적화: {current.name}")
                # 설정에서 이미지 리사이즈/품질 값을 가져온다 (슬라이더와 연동).
                max_edge = max(200, min(settings.image_max_edge, 10000))
                jpeg_quality = max(10, min(settings.image_quality, 100))
                (
                    out_path,
                    before,
                    after,
                    count,
                    log_path,
                ) = run_image_slim(
                    current,
                    max_edge=max_edge,
                    jpeg_quality=jpeg_quality,
                    progressive=True,
                )
                current = out_path
                if step != steps[-1]:
                    intermediate_files.append(current)
                saved = before - after
                pct = (saved / before * 100.0) if before > 0 else 0.0
                log_detail(f" - 이미지 개수: {count}")
                log_detail(
                    " - Before: "
                    + human_size(before)
                    + ", After: "
                    + human_size(after)
                    + ", Saved: "
                    + human_size(saved)
                    + f" ({pct:.1f}%)"
                )
                log_detail(f" - 로그: {log_path}")
                log_files.append(log_path)
            elif step == "precision":
                set_status("정밀 슬리머 실행 중...", base)
                log_info(f"[{index}/{total}] 정밀 슬리머: {current.name}")
                has_clean_step = "clean" in steps
                no_backup = has_clean_step

                def logger(msg: str) -> None:
                    if settings.log_mode == "verbose":
                        log("[Precision] " + msg)

                (
                    out_path,
                    saved_mb,
                    pct,
                    old_b,
                    new_b,
                ) = run_precision_step(
                    current,
                    aggressive,
                    no_backup,
                    do_xml_cleanup,
                    force_custom,
                    logger,
                )
                current = out_path
                log_detail(f" - 결과: {current.name}")
                log_detail(
                    " - Before: "
                    + human_size(old_b)
                    + ", After: "
                    + human_size(new_b)
                    + f", Saved: {saved_mb:.2f} MB ({pct:.1f}%)"
                )

            set_status("진행 중...", next_p)
        except Exception as e:  # noqa: BLE001
            log_info(f"[ERROR] {step} 단계에서 오류: {e}")
            set_status("오류 발생", None)

            # 오류 시 로그 폴더 자동 열기 옵션 처리
            if log_files and settings.open_log_on_error:
                try:
                    log_file = log_files[-1]
                    settings.last_run_log_file = str(log_file)
                    save_settings(settings)
                    try:
                        open_in_explorer_select(log_file)
                    except Exception:
                        pass
                except Exception as inner:  # noqa: BLE001
                    log_info(f"[WARN] 오류 로그 처리 중 실패: {inner}")

            show_error(
                "오류",
                f"{step} 단계에서 오류가 발생했습니다.\n\n{e}",
            )
            return

    # 최종 파일 이름 정리: 어떤 조합이든 최종본은 원본 이름 + '_complete' 로 통일
    # 예: 원본.xlsx -> 원본_complete.xlsx
    try:
        orig_stem = start_path.stem
        parent = current.parent
        suffix = current.suffix
        desired = parent / f"{orig_stem}_complete{suffix}"

        if desired != current:
            candidate = desired
            idx = 1
            # 동일 이름이 이미 있으면 (1), (2) 를 붙여서 충돌 회피
            while candidate.exists():
                candidate = parent / f"{orig_stem}_complete({idx}){suffix}"
                idx += 1

            old = current
            old.rename(candidate)
            log(f"[INFO] 최종 파일 이름 변경: {old.name} -> {candidate.name}")
            current = candidate
    except Exception as e:  # noqa: BLE001
        log_info(f"[WARN] 최종 파일 이름 변경 실패: {e}")

    # 사용자 지정 출력 폴더가 설정된 경우, 최종 결과를 해당 폴더로 이동
    try:
        if settings.output_dir:
            target_dir = Path(settings.output_dir)
            target_dir.mkdir(parents=True, exist_ok=True)
            if target_dir.resolve() != current.parent.resolve():
                candidate = target_dir / current.name
                idx = 1
                while candidate.exists():
                    candidate = target_dir / f"{current.stem}({idx}){current.suffix}"
                    idx += 1

                old = current
                old.rename(candidate)
                log_info(f"[INFO] 최종 파일 이동: {old} -> {candidate}")
                current = candidate
    except Exception as e:  # noqa: BLE001
        log_info(f"[WARN] 사용자 지정 출력 폴더로 이동 실패: {e}")

    # 사용자가 백업 유지 옵션을 끈 경우, Clean 단계에서 생성된 백업 파일을 정리
    if not settings.keep_backup:
        for b in backup_files:
            try:
                if b.exists() and b.resolve() != current.resolve():
                    b.unlink()
                    log_detail(f"[INFO] 백업 파일 삭제: {b}")
            except Exception as e:  # noqa: BLE001
                log_info(f"[WARN] 백업 파일 삭제 실패: {b} ({e})")

    # 모든 단계가 성공적으로 끝난 경우에만 중간 산출물 및 로그 정리
    for tmp in intermediate_files:
        try:
            if tmp.exists() and tmp != current:
                tmp.unlink()
                log_detail(f"[INFO] 중간 결과 삭제: {tmp}")
        except Exception as e:  # noqa: BLE001
            log_info(f"[WARN] 중간 결과 삭제 실패: {tmp} ({e})")

    for log_path in log_files:
        try:
            if log_path.exists():
                log_path.unlink()
                log_detail(f"[INFO] 로그 파일 삭제: {log_path}")
        except Exception as e:  # noqa: BLE001
            log_info(f"[WARN] 로그 파일 삭제 실패: {log_path} ({e})")

    set_status("모든 작업 완료", 100.0)
    log_info(f"[INFO] 파이프라인 완료. 최종 파일: {current}")
    on_finished(current)


class ExcelSuiteApp:
    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title("ExcelSlimmer")
        self.root.geometry("1120x720")
        self.root.minsize(960, 640)

        self.file_var = tk.StringVar()
        self.clean_var = tk.IntVar(value=1)
        self.image_var = tk.IntVar(value=1)
        self.precision_var = tk.IntVar(value=0)

        self.prec_aggressive_var = tk.IntVar(value=0)
        self.prec_xmlcleanup_var = tk.IntVar(value=0)
        self.prec_force_custom_var = tk.IntVar(value=0)

        self.status_var = tk.StringVar(value="준비됨")
        self.progress_var = tk.DoubleVar(value=0.0)

        self._build_ui()

    def _build_ui(self) -> None:
        style = ttk.Style()
        try:
            style.theme_use("vista")
        except Exception:
            pass

        base_bg = "#ffffff"
        card_bg = "#ffffff"
        self.root.configure(bg=base_bg)

        # 기본 폰트를 맑은 고딕으로 설정 (사무용에 적합하면서도 너무 딱딱하지 않게)
        self.root.option_add("*Font", ("맑은 고딕", 10))

        style.configure("App.TFrame", background=base_bg)
        style.configure(
            "Card.TFrame",
            background=card_bg,
            borderwidth=1,
            relief="solid",
        )
        style.configure(
            "Card.TLabelframe",
            background=card_bg,
            borderwidth=1,
            relief="solid",
        )
        style.configure("Card.TLabelframe.Label", background=card_bg, font=("맑은 고딕", 10, "bold"))
        style.configure("TLabel", background=card_bg)
        style.configure("TCheckbutton", background=card_bg)
        style.configure("TNotebook", background=base_bg, borderwidth=0)
        style.configure("TNotebook.Tab", background=base_bg)
        style.configure("TButton", font=("맑은 고딕", 10), background=card_bg)
        style.map("TButton", background=[("active", "#f0f0f0")])
        style.configure("Header.TLabel", font=("맑은 고딕", 18, "bold"), background=base_bg)
        style.configure("SubHeader.TLabel", foreground="#666666", background=base_bg)
        style.configure("Section.TLabel", font=("맑은 고딕", 10, "bold"), background=base_bg)

        outer = ttk.Frame(self.root, style="App.TFrame", padding=(18, 14, 18, 18))
        outer.pack(fill="both", expand=True)

        header_frame = ttk.Frame(outer, style="App.TFrame")
        header_frame.pack(fill="x", pady=(0, 12))

        title_label = ttk.Label(header_frame, text="ExcelSlimmer", style="Header.TLabel")
        title_label.pack(side="left", anchor="w")

        notebook = ttk.Notebook(outer)
        notebook.pack(fill="both", expand=True)

        pipeline_page = ttk.Frame(notebook, style="App.TFrame")
        settings_page = ttk.Frame(notebook, style="App.TFrame")
        notebook.add(pipeline_page, text="슬리머 실행")
        notebook.add(settings_page, text="환경 설정")

        ttk.Label(
            settings_page,
            text="추후 업데이트 예정입니다.",
            style="SubHeader.TLabel",
        ).pack(pady=20, padx=20, anchor="w")

        pipeline_page.columnconfigure(0, weight=0, minsize=380)
        pipeline_page.columnconfigure(1, weight=1)
        pipeline_page.rowconfigure(0, weight=1)

        left_col = ttk.Frame(pipeline_page, style="App.TFrame")
        left_col.grid(row=0, column=0, sticky="nsew", padx=(8, 12))
        right_col = ttk.Frame(pipeline_page, style="App.TFrame")
        right_col.grid(row=0, column=1, sticky="nsew", padx=(0, 8))

        ttk.Label(left_col, text="대상 파일", style="Section.TLabel").pack(anchor="w", pady=(0, 4))
        file_card = ttk.Frame(
            left_col,
            style="Card.TFrame",
            padding=(12, 10, 12, 12),
        )
        file_card.pack(fill="x", pady=(0, 10))
        ttk.Label(file_card, text="파일 경로:").pack(anchor="w")
        entry = ttk.Entry(file_card, textvariable=self.file_var)
        entry.pack(fill="x", expand=True, pady=(4, 6))
        ttk.Button(file_card, text="찾기...", command=self._select_file).pack(anchor="e")

        ttk.Label(left_col, text="실행할 기능", style="Section.TLabel").pack(anchor="w", pady=(0, 4))
        pipeline_card = ttk.Frame(
            left_col,
            style="Card.TFrame",
            padding=(12, 8, 12, 10),
        )
        pipeline_card.pack(fill="x", pady=(0, 10))

        ttk.Checkbutton(
            pipeline_card,
            text="이름 정의 정리 (definedNames 클린)",
            variable=self.clean_var,
        ).pack(anchor="w", pady=(2, 2))
        ttk.Checkbutton(
            pipeline_card,
            text="이미지 최적화 (이미지 리사이즈/압축)",
            variable=self.image_var,
        ).pack(anchor="w", pady=(2, 2))
        self.precision_check = ttk.Checkbutton(
            pipeline_card,
            text="정밀 슬리머 (Precision Plus)",
            variable=self.precision_var,
            command=self._on_precision_toggle,
        )
        self.precision_check.pack(anchor="w", pady=(2, 0))
        self.precision_warning = ttk.Label(
            pipeline_card,
            text="주의: 정밀 슬리머 사용 시 엑셀에서 복구 여부를 물어볼 수 있습니다.",
            foreground="#aa0000",
            background=card_bg,
        )
        self.precision_warning.pack(anchor="w", padx=18, pady=(0, 4))

        ttk.Label(left_col, text="정밀 슬리머 옵션", style="Section.TLabel").pack(anchor="w", pady=(0, 4))
        precision_card = ttk.Frame(
            left_col,
            style="Card.TFrame",
            padding=(12, 8, 12, 10),
        )
        precision_card.pack(fill="x", pady=(0, 10))

        self.prec_xmlcleanup_cb = ttk.Checkbutton(
            precision_card,
            text="XML 정리 (calcChain, printerSettings 등)",
            variable=self.prec_xmlcleanup_var,
        )
        self.prec_xmlcleanup_cb.pack(anchor="w", pady=(2, 2))
        self.prec_force_custom_cb = ttk.Checkbutton(
            precision_card,
            text="숨은 XML 데이터 삭제 (customXml, 주의)",
            variable=self.prec_force_custom_var,
        )
        self.prec_force_custom_cb.pack(anchor="w", pady=(2, 2))
        self.prec_aggressive_cb = ttk.Checkbutton(
            precision_card,
            text="공격 모드 (이미지 리사이즈 + PNG→JPG)",
            variable=self.prec_aggressive_var,
        )
        self.prec_aggressive_cb.pack(anchor="w", pady=(2, 2))
        self.prec_force_custom_hint = ttk.Label(
            precision_card,
            text="주의: 숨은 XML 데이터 삭제는 일반적인 경우 사용하지 마세요.",
            foreground="#aa0000",
            background=card_bg,
        )
        self.prec_force_custom_hint.pack(anchor="w", padx=18, pady=(0, 4))

        run_card = ttk.Frame(left_col, style="Card.TFrame", padding=(12, 10, 12, 12))
        run_card.pack(fill="x")

        self.run_button = ttk.Button(
            run_card,
            text="선택한 기능 실행",
            command=self._on_run_clicked,
        )
        self.run_button.pack(anchor="w")

        # 상태/진행률 영역은 카드 안쪽이므로 별도 테두리 없이 App.TFrame 사용
        status_row = ttk.Frame(run_card, style="App.TFrame")
        status_row.pack(fill="x", pady=(8, 0))
        status_label = ttk.Label(status_row, textvariable=self.status_var)
        # 오른쪽이 살짝 잘려 보이지 않도록 약간의 내부 여백을 준다
        status_label.pack(side="right", padx=(4, 0))
        self.progress = ttk.Progressbar(
            status_row,
            maximum=100.0,
            variable=self.progress_var,
        )
        self.progress.pack(side="left", fill="x", expand=True, padx=(0, 8))

        ttk.Label(right_col, text="로그", style="Section.TLabel").pack(anchor="w", pady=(0, 4))
        log_card = ttk.Frame(
            right_col,
            style="Card.TFrame",
            padding=(12, 8, 12, 12),
        )
        log_card.pack(fill="both", expand=True)

        self.log_box = scrolledtext.ScrolledText(
            log_card,
            height=10,
            state="disabled",
            bg=card_bg,
            relief="flat",
            borderwidth=0,
            highlightthickness=0,
        )
        self.log_box.pack(fill="both", expand=True, padx=(0, 4), pady=(0, 4))

        # 스크롤바 트랙/테두리를 최소화해서 회색 줄처럼 보이는 영역을 줄인다.
        try:
            self.log_box.vbar.config(
                borderwidth=0,
                highlightthickness=0,
                bg=card_bg,
                troughcolor=card_bg,
                relief="flat",
            )
        except Exception:
            pass

        self._update_precision_options_state()

    def _on_precision_toggle(self) -> None:
        self._update_precision_options_state()

    def _update_precision_options_state(self) -> None:
        enabled = bool(self.precision_var.get())
        state = "normal" if enabled else "disabled"
        for cb in (
            self.prec_aggressive_cb,
            self.prec_xmlcleanup_cb,
            self.prec_force_custom_cb,
        ):
            cb.configure(state=state)

    def _select_file(self) -> None:
        filetypes = [("Excel 파일", "*.xlsx;*.xlsm"), ("모든 파일", "*.*")]
        path = filedialog.askopenfilename(
            title="대상 Excel 파일 선택",
            filetypes=filetypes,
        )
        if path:
            self.file_var.set(path)

    def _append_log(self, text: str) -> None:
        self.log_box.configure(state="normal")
        self.log_box.insert("end", text + "\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def log(self, text: str) -> None:
        self.root.after(0, lambda: self._append_log(text))

    def set_status(self, text: str, progress: float = None) -> None:
        def _update() -> None:
            self.status_var.set(text)
            if progress is not None:
                self.progress_var.set(progress)

        self.root.after(0, _update)

    def show_info(self, title: str, text: str) -> None:
        self.root.after(0, lambda: messagebox.showinfo(title, text))

    def show_error(self, title: str, text: str) -> None:
        self.root.after(0, lambda: messagebox.showerror(title, text))

    def _on_run_clicked(self) -> None:
        path_str = self.file_var.get().strip()
        if not path_str:
            messagebox.showwarning("안내", "대상 파일을 먼저 선택하세요.")
            return
        path = Path(path_str)
        if not path.exists():
            messagebox.showerror("오류", f"파일을 찾을 수 없습니다:\n{path}")
            return
        if path.suffix.lower() not in (".xlsx", ".xlsm"):
            messagebox.showerror("오류", "지원 형식은 .xlsx / .xlsm 입니다.")
            return
        if not (
            self.clean_var.get()
            or self.image_var.get()
            or self.precision_var.get()
        ):
            messagebox.showinfo("안내", "실행할 기능을 하나 이상 선택하세요.")
            return

        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")
        self.progress_var.set(0.0)
        self.status_var.set("작업 시작...")
        self.run_button.configure(state="disabled")

        t = threading.Thread(
            target=self._run_pipeline_worker,
            args=(path,),
            daemon=True,
        )
        t.start()

    def _run_pipeline_worker(self, start_path: Path) -> None:
        try:
            self._run_pipeline(start_path)
        except Exception as e:
            self.log(f"[ERROR] 예기치 못한 오류: {e}")
            traceback.print_exc()
            self.set_status("오류 발생", None)
            self.show_error("오류", f"예기치 못한 오류가 발생했습니다.\n\n{e}")
        finally:
            self.root.after(0, lambda: self.run_button.configure(state="normal"))

    def _reset_ui_after_finish(self) -> None:
        """파이프라인 완료 후 기본 상태로 되돌립니다 (로그는 유지)."""
        self.file_var.set("")
        self.clean_var.set(1)
        self.image_var.set(1)
        self.precision_var.set(0)
        self.prec_aggressive_var.set(0)
        self.prec_xmlcleanup_var.set(0)
        self.prec_force_custom_var.set(0)
        self._update_precision_options_state()
        self.progress_var.set(0.0)
        self.status_var.set("준비됨")

    def _run_pipeline(self, start_path: Path) -> None:
        def _on_finished(final_path: Path) -> None:
            def _after_msg() -> None:
                try:
                    open_in_explorer_select(final_path)
                except Exception:
                    pass
                # 로그는 유지하고 나머지 UI 상태만 초기화
                self._reset_ui_after_finish()

            self.root.after(
                0,
                lambda: (
                    messagebox.showinfo(
                        "완료",
                        f"모든 작업이 완료되었습니다.\n\n최종 결과 파일:\n{final_path}",
                    ),
                    _after_msg(),
                ),
            )

        run_pipeline_core(
            start_path=start_path,
            use_clean=bool(self.clean_var.get()),
            use_image=bool(self.image_var.get()),
            use_precision=bool(self.precision_var.get()),
            aggressive=bool(self.prec_aggressive_var.get()),
            do_xml_cleanup=bool(self.prec_xmlcleanup_var.get()),
            force_custom=bool(self.prec_force_custom_var.get()),
            log=self.log,
            set_status=self.set_status,
            show_error=self.show_error,
            on_finished=_on_finished,
        )

    def run(self) -> None:
        self.root.mainloop()


def main() -> None:
    app = ExcelSuiteApp()
    app.run()


if __name__ == "__main__":
    main()
