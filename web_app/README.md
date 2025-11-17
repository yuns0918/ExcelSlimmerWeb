# ExcelSlimmer Web

ExcelSlimmer Web은 ExcelSlimmer 데스크톱(EXE) 버전과 동일한 슬리밍 파이프라인을 **웹에서 업로드/다운로드 방식으로 사용할 수 있게 하는** 웹 버전입니다.

- 백엔드: **FastAPI** (`web_app/main.py`)
- 프론트엔드: 단일 HTML/JS 페이지 (`web_app/index.html`)
- 코어 로직: 데스크톱과 공용으로 사용하는 `excel_suite_pipeline.py` + `backData/` 모듈

사용자는 브라우저에서 엑셀 파일을 업로드하고, 사용할 기능을 선택한 뒤 서버에서 처리된 결과 파일을 다시 다운로드할 수 있습니다.

---

## 1. 주요 기능

웹 버전에서도 데스크톱 버전과 동일하게 다음 세 가지 기능을 선택적으로 사용할 수 있습니다.

1. **이름 정의 정리 (ExcelCleaner)**  
   사용되지 않는 `definedNames` 항목을 정리합니다.
2. **이미지 최적화 (Excel Image Slimmer)**  
   워크시트에 포함된 이미지의 해상도/품질을 낮춰 용량을 줄입니다.
3. **정밀 슬리머 (Precision Plus)**  
   이미지 재압축, XML 정리, 숨은 XML 데이터 삭제 등 고급 용량 최적화를 수행합니다.

FastAPI 핸들러는 내부적으로 `excel_suite_pipeline.run_pipeline_core()`를 호출하여, 데스크톱과 동일한 파이프라인 로직을 재사용합니다.

---

## 2. 폴더/파일 구조 (권장)

웹 전용 리포지토리 **ExcelSlimmerWeb**에서는 다음과 같은 구성을 권장합니다.

```text
ExcelSlimmerWeb/
  excel_suite_pipeline.py      # 공용 파이프라인 로직 (데스크톱/웹 공용)
  settings.py                  # 설정 로드/저장 로직 (AppSettings)
  backData/                    # 기존 Tk 기반 모듈 (이미지 슬리머, 정밀 슬리머, 이름 정리)
    excel_image_slimmer_gui_v3.py
    excel_slimmer_precision_plus.py
    gui_clean_defined_names_desktop_date.py
    ... (기타 필요한 파일)
  web_app/
    main.py                    # FastAPI 앱 엔트리포인트
    index.html                 # 단일 페이지 프론트엔드
    README.md                  # 이 문서
```

> **중요:** `web_app/main.py`는 `excel_suite_pipeline`과 `settings`를 `import` 하고, 
> `excel_suite_pipeline.py`는 `backData/` 안의 모듈들을 사용합니다.  
> 따라서 **웹 리포를 구성할 때 위 네 가지( `excel_suite_pipeline.py`, `settings.py`, `backData/`, `web_app/` )가 한 리포지토리 안에 함께 있어야** 합니다.

---

## 3. 실행 전 준비 (의존성 설치)

### 3.1 Python 버전

- Python **3.11 이상**을 권장합니다.

### 3.2 가상환경 생성 (선택)

```bash
python -m venv .venv
# Windows PowerShell 예시
.\.venv\Scripts\Activate.ps1
# 또는 CMD
.\.venv\Scripts\activate.bat
```

### 3.3 필수 패키지 설치

`FastAPI` + 파일 업로드 + 이미지/XML 처리에 필요한 최소 패키지는 다음과 같습니다.

```bash
pip install fastapi "uvicorn[standard]" python-multipart pillow lxml
```

- `fastapi`: 웹 프레임워크
- `uvicorn[standard]`: ASGI 서버 (로컬 실행 및 일부 배포 환경에서 사용)
- `python-multipart`: `UploadFile` / 폼 기반 파일 업로드 처리에 필요
- `pillow`: 이미지 리사이즈/압축
- `lxml`: XML 파싱 및 수정 (정밀 슬리머/이름 정리에서 사용)

프로젝트에 따라 추가로 사용하는 패키지가 있다면 `pip install ...` 로 함께 설치해 주세요.

---

## 4. 로컬 개발/테스트 실행

### 4.1 FastAPI 서버 실행

프로젝트 루트(예: `ExcelSlimmerWeb/`)에서 다음 명령으로 개발 서버를 실행합니다.

```bash
# 가상환경 활성화 후
uvicorn web_app.main:app --reload
```

- 기본 접속 주소: <http://127.0.0.1:8000/>
- `--reload` 옵션은 코드 변경 시 자동으로 서버를 재시작해 주므로 개발 시에 편리합니다.

### 4.2 브라우저에서 사용하기

1. 브라우저에서 <http://127.0.0.1:8000/> 접속
2. 화면 상단의 **"대상 파일"** 영역에서 엑셀 파일을 드래그 앤 드롭하거나, **"파일 선택"** 버튼으로 `.xlsx` 또는 `.xlsm` 파일 선택
3. **실행할 기능**에서 원하는 옵션 체크
   - 이름 정의 정리 (definedNames 클린)
   - 이미지 최적화 (이미지 리사이즈/압축)
   - 정밀 슬리머 (Precision Plus)
4. (선택) 정밀 슬리머 옵션 설정
   - XML 정리 (calcChain, printerSettings 등)
   - 숨은 XML 데이터 삭제 (customXml, 주의)
   - 이미지 포맷 변경 (PNG→JPG) + 참조 동기화 (고급)
5. **"슬리머 실행"** 버튼 클릭
6. 처리 완료 후 브라우저에서 결과 엑셀 파일이 자동으로 다운로드됩니다.

오른쪽 **로그 영역**에는 서버 측에서 생성한 로그 메시지가 순차적으로 표시됩니다.

---

## 5. API 엔드포인트

웹 프론트엔드는 내부적으로 다음 FastAPI 엔드포인트를 사용합니다.

### 5.1 `GET /`

- HTML 응답 (`web_app/index.html` 내용)
- 브라우저에서 접속하는 기본 페이지입니다.

### 5.2 `POST /api/slim`

업로드된 엑셀 파일을 슬림 처리한 뒤, 결과 파일을 직접 응답으로 반환합니다.

- 요청 형식: `multipart/form-data`
- 필드
  - `file`: 업로드할 엑셀 파일 (`UploadFile`)
  - `use_clean`: `true`/`false` (기본값: `true`)
  - `use_image`: `true`/`false` (기본값: `true`)
  - `use_precision`: `true`/`false` (기본값: `false`)
  - `aggressive`: `true`/`false`
  - `do_xml_cleanup`: `true`/`false`
  - `force_custom`: `true`/`false`
- 응답
  - 성공 시: 슬림 처리된 엑셀 파일 (`FileResponse`)
  - 실패 시: `HTTP 4xx/5xx` + JSON 바디(`{"detail": "..."}`)

웹 UI는 `fetch("/api/slim", { method: "POST", body: formData })`로 이 엔드포인트를 호출합니다.

### 5.3 `GET /api/health`

- 단순 헬스체크용 엔드포인트
- 응답 예: `{ "status": "ok" }`

배포 환경(예: 로드 밸런서/헬스 체크)에서 이 엔드포인트를 사용해 서버 상태를 확인할 수 있습니다.

---

## 6. 배포 시 참고 사항

### 6.1 일반적인 ASGI 배포 예시

예를 들어, 컨테이너/서버 환경에서 다음과 같이 실행할 수 있습니다.

```bash
uvicorn web_app.main:app --host 0.0.0.0 --port 8000
```

- 리버스 프록시(Nginx 등) 뒤에서 구동하거나, 
  PaaS 서비스(Railway, Render 등)의 **스타트 커맨드**에 위 명령을 넣어 사용할 수 있습니다.

### 6.2 파일 시스템 및 임시 디렉토리

- `web_app/main.py`는 업로드된 파일을 **임시 디렉토리**에 저장한 뒤, 
  `excel_suite_pipeline.run_pipeline_core()`를 호출합니다.
- 요청 처리 완료 후에는 `TemporaryDirectory()` 컨텍스트가 자동으로 정리되므로,
  긴 시간 동안 파일이 쌓이지 않습니다.

### 6.3 GUI 의존성

- 코어 모듈 중 일부는 원래 Tk 기반 GUI 코드를 포함하고 있지만,
  웹 서버 환경에서는 **GUI를 사용하지 않도록** `excel_suite_pipeline.py`에 방어 로직이 들어 있습니다.
- 따라서 일반적인 서버/컨테이너 환경(화면/디스플레이 없음)에서도 동작할 수 있도록 설계되어 있습니다.

---

## 7. 주의 사항 및 권장 사용 방법

- 지원 형식: 주로 `.xlsx`, `.xlsm`  
  (암호화된 파일이나 공유 편집 중인 파일은 먼저 해제/닫은 뒤 사용 권장)
- 정밀 슬리머 옵션은 엑셀 내부 구조를 적극적으로 정리합니다.
  - 일부 파일에서는 결과 파일을 열 때 엑셀이 **"복구" 안내 팝업**을 띄울 수 있습니다.
  - 실제 데이터 손상이 아닌 경우도 많지만, 중요한 문서는 반드시 결과를 확인해 주세요.
- 대용량 파일(수십 MB 이상)을 여러 명이 동시에 업로드하는 사용 패턴에서는
  서버 리소스(CPU/메모리/디스크 IO)를 고려해 인스턴스 스펙과 동시 처리 수를 조정해야 합니다.

---

## 8. GitHub 리포지토리

- 데스크톱/EXE 버전: <https://github.com/yuns0918/ExcelSlimmerEXE>
- 웹 버전(이 리포지토리): <https://github.com/yuns0918/ExcelSlimmerWeb>

웹 버전 리포에서 변경 사항을 반영할 때는

```bash
git add .
git commit -m "Update web app"
git push origin main
```

과 같이 일반적인 Git 워크플로를 사용하면 됩니다.
