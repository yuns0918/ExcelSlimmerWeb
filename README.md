# ExcelSlimmer Web

ExcelSlimmer Web은 ExcelSlimmer 데스크톱(EXE) 버전과 동일한 슬리밍 파이프라인을 **웹(업로드/다운로드)** 형태로 제공하는 프로젝트입니다.

- 백엔드: **FastAPI** (`web_app/main.py`)
- 프론트엔드: 단일 HTML/JS 페이지 (`web_app/index.html`)
- 코어 로직: 데스크톱과 공용으로 사용하는 `excel_suite_pipeline.py` + `backData/` 모듈

이 리포지토리는 **웹 서버용 코드만**을 포함하며, 데스크톱(EXE) 버전은 별도 리포지토리인
`ExcelSlimmerEXE`에서 관리합니다.

- 데스크톱/EXE: <https://github.com/yuns0918/ExcelSlimmerEXE>
- 웹 버전(이 리포): <https://github.com/yuns0918/ExcelSlimmerWeb>

---

## 폴더 구조

```text
ExcelSlimmerWeb/
  excel_suite_pipeline.py   # 공용 파이프라인 로직 (데스크톱/웹 공용)
  settings.py               # 설정 로드/저장 로직 (AppSettings)
  backData/                 # 기존 Tk 기반 도구 모듈
  web_app/                  # FastAPI 앱 + 웹 UI + 상세 README
    main.py
    index.html
    README.md
```

자세한 실행/배포 방법은 `web_app/README.md`를 참고하세요.

```text
web_app/README.md
  - 로컬 개발 서버 실행 방법 (uvicorn)
  - 의존성(pip install ...) 안내
  - API 엔드포인트(/, /api/slim, /api/health)
  - 배포 시 고려사항
```
