from __future__ import annotations

import shutil
import tempfile
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from starlette.middleware.cors import CORSMiddleware

from excel_suite_pipeline import run_pipeline_core
from settings import get_settings, save_settings

app = FastAPI(title="ExcelSlimmer Web")

# CORS 설정 (내부 사용이지만, 추후 확장을 고려해 허용 도메인을 조정할 수 있음)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.get("/", response_class=HTMLResponse)
async def index() -> HTMLResponse:
    """간단한 단일 페이지 업로드 UI를 제공한다."""

    html = Path(__file__).with_name("index.html").read_text(encoding="utf-8")
    return HTMLResponse(content=html)


@app.post("/api/slim")
async def slim_excel(
    file: UploadFile = File(...),
    use_clean: bool = Form(True),
    use_image: bool = Form(True),
    use_precision: bool = Form(False),
    aggressive: bool = Form(False),
    do_xml_cleanup: bool = Form(False),
    force_custom: bool = Form(False),
) -> FileResponse:
    """업로드된 Excel 파일을 슬림 처리 후 결과 파일을 반환한다."""

    if not file.filename:
        raise HTTPException(status_code=400, detail="파일 이름이 비어 있습니다.")

    suffix = Path(file.filename).suffix.lower()
    if suffix not in {".xlsx", ".xlsm"}:
        raise HTTPException(status_code=400, detail=".xlsx 또는 .xlsm 파일만 지원합니다.")

    # 임시 디렉토리 안에서 모든 작업을 수행하고, 요청 종료 후 자동 정리
    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir_path = Path(tmpdir)
        in_path = tmpdir_path / file.filename

        # 업로드 파일 저장
        with in_path.open("wb") as f:
            shutil.copyfileobj(file.file, f)

        logs: list[str] = []
        last_status: str = "준비됨"
        last_progress: float = 0.0
        result_path: Path | None = None
        error_message: str | None = None

        def log_cb(msg: str) -> None:
            logs.append(msg)

        def set_status_cb(text: str, progress: float | None) -> None:
            nonlocal last_status, last_progress
            last_status = text
            if progress is not None:
                last_progress = progress

        def show_error_cb(title: str, text: str) -> None:
            nonlocal error_message
            error_message = text

        def on_finished_cb(path: Path) -> None:
            nonlocal result_path
            result_path = path

        try:
            run_pipeline_core(
                start_path=in_path,
                use_clean=use_clean,
                use_image=use_image,
                use_precision=use_precision,
                aggressive=aggressive,
                do_xml_cleanup=do_xml_cleanup,
                force_custom=force_custom,
                log=log_cb,
                set_status=set_status_cb,
                show_error=show_error_cb,
                on_finished=on_finished_cb,
            )
        except Exception as exc:  # noqa: BLE001
            raise HTTPException(status_code=500, detail=f"서버 오류: {exc}") from exc

        if error_message is not None:
            # 파이프라인 내부에서 치명적 오류가 발생한 경우
            raise HTTPException(status_code=500, detail=error_message)

        if result_path is None or not result_path.exists():
            raise HTTPException(status_code=500, detail="결과 파일을 생성하지 못했습니다.")

        # 결과 파일을 다운로드로 반환
        return FileResponse(
            path=result_path,
            filename=result_path.name,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


@app.get("/api/health")
async def health() -> JSONResponse:
    return JSONResponse({"status": "ok"})
