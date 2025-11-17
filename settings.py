from __future__ import annotations

import json
import os
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Literal


APP_NAME = "ExcelSlimmer"


def _get_settings_path() -> Path:
    """Return the path to the JSON settings file.

    우선순위:
    1) Windows 환경에서는 %APPDATA%\ExcelSlimmer\settings.json
    2) 그 외에는 사용자 홈 아래 .ExcelSlimmer/settings.json
    """

    appdata = os.getenv("APPDATA")
    if appdata:
        root = Path(appdata) / APP_NAME
    else:
        root = Path.home() / f".{APP_NAME}"

    root.mkdir(parents=True, exist_ok=True)
    return root / "settings.json"


SETTINGS_FILE = _get_settings_path()


@dataclass
class AppSettings:
    """애플리케이션 전역 설정.

    앞으로 환경 설정 탭에서 제어할 옵션들을 여기에 추가한다.
    """

    # 결과 폴더에 backup 파일("*_backup")을 남길지 여부
    keep_backup: bool = False

    # 최종 결과 엑셀 파일을 옮겨 둘 사용자 지정 폴더 (빈 문자열이면 기본 위치 사용)
    output_dir: str = ""

    # 이미지 관련 기본값 (추후 슬라이더와 연동 예정)
    image_max_edge: int = 1400
    image_quality: int = 80

    # 로그/테마 관련 기본값 (추후 확장 예정)
    log_mode: Literal["minimal", "verbose"] = "verbose"
    open_log_on_error: bool = False
    theme: Literal["light", "dark"] = "light"
    last_run_log_file: str = ""


_settings_cache: AppSettings | None = None


def load_settings() -> AppSettings:
    """설정 파일을 로드하거나, 없으면 기본값을 반환한다."""

    default = AppSettings()
    if not SETTINGS_FILE.exists():
        return default

    try:
        data = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
        if not isinstance(data, dict):  # type: ignore[unreachable]
            return default
        base = asdict(default)
        base.update({k: v for k, v in data.items() if k in base})
        return AppSettings(**base)
    except Exception:
        # 파일이 깨졌거나 포맷이 변경된 경우에는 기본값으로 복구한다.
        return default


def save_settings(settings: AppSettings) -> None:
    """현재 설정을 JSON 파일로 저장한다."""

    SETTINGS_FILE.write_text(
        json.dumps(asdict(settings), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def get_settings() -> AppSettings:
    """프로세스 전체에서 공유되는 설정 객체를 반환한다."""

    global _settings_cache
    if _settings_cache is None:
        _settings_cache = load_settings()
    return _settings_cache
