#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel definedNames 정리 스크립트 (안전판)
- workbook.xml을 통째로 재직렬화하지 않고, <definedNames> 내부만 '외과수술'로 수정
- Print_Area / Print_Titles만 유지
- 바탕화면에 "Excel이름관리자정리완료" 최상위 폴더 고정 (없으면 생성, 있으면 재사용)
- 실행할 때마다 "YYYY-MM-DD-HH-MM-SS" 하위 폴더 생성 → 그 안에 "백업", "정리본" 분리 저장
- 완료 시 해당 날짜 폴더 자동 열기
- 네이티브 Win32 대화상자 사용(파일 선택/알림)
"""

import os, re, shutil, zipfile, sys, gc
from datetime import datetime
import ctypes
from ctypes import wintypes

KEEP_NAMES = {"_xlnm.Print_Area", "_xlnm.Print_Titles", "Print_Area", "Print_Titles"}
TOP_DIR_NAME = "ExcelSlimmed"
BACKUP_DIR = "백업"
RESULT_DIR = "정리본"

# --------- Windows native dialogs ---------
OFN_FILEMUSTEXIST = 0x00001000
OFN_PATHMUSTEXIST = 0x00000800

class OPENFILENAMEW(ctypes.Structure):
    _fields_ = [
        ("lStructSize", wintypes.DWORD),
        ("hwndOwner", wintypes.HWND),
        ("hInstance", wintypes.HINSTANCE),
        ("lpstrFilter", wintypes.LPWSTR),
        ("lpstrCustomFilter", wintypes.LPWSTR),
        ("nMaxCustFilter", wintypes.DWORD),
        ("nFilterIndex", wintypes.DWORD),
        ("lpstrFile", wintypes.LPWSTR),
        ("nMaxFile", wintypes.DWORD),
        ("lpstrFileTitle", wintypes.LPWSTR),
        ("nMaxFileTitle", wintypes.DWORD),
        ("lpstrInitialDir", wintypes.LPWSTR),
        ("lpstrTitle", wintypes.LPWSTR),
        ("Flags", wintypes.DWORD),
        ("nFileOffset", wintypes.WORD),
        ("nFileExtension", wintypes.WORD),
        ("lpstrDefExt", wintypes.LPWSTR),
        ("lCustData", wintypes.LPARAM),
        ("lpfnHook", wintypes.LPVOID),
        ("lpTemplateName", wintypes.LPWSTR),
        ("pvReserved", wintypes.LPVOID),
        ("dwReserved", wintypes.DWORD),
        ("FlagsEx", wintypes.DWORD),
    ]

def msg_box(text, title="알림", style=0x40):  # MB_ICONINFORMATION
    ctypes.windll.user32.MessageBoxW(0, str(text), str(title), style)

def open_file_dialog(title="정리할 Excel 파일(.xlsx) 선택"):
    # 준비: 1024문자 버퍼 (긴 경로 대비)
    buf = ctypes.create_unicode_buffer(1024)
    buf[0] = '\0'   # 버퍼 초기화 (이전 값 방지)

    ofn = OPENFILENAMEW()
    ofn.lStructSize = ctypes.sizeof(OPENFILENAMEW)
    ofn.hwndOwner = None

    # 필터는 \0으로 구분, 마지막은 \0\0 로 종결
    ofn.lpstrFilter = "Excel Workbook (*.xlsx)\0*.xlsx\0All Files (*.*)\0*.*\0\0"
    ofn.nFilterIndex = 1

    # 핵심: 버퍼를 LPWSTR(c_wchar_p) 포인터로 캐스팅
    ofn.lpstrFile = ctypes.cast(buf, ctypes.c_wchar_p)

    # 버퍼 길이는 문자 수 기준
    ofn.nMaxFile = ctypes.sizeof(buf) // ctypes.sizeof(ctypes.c_wchar)

    ofn.lpstrFileTitle = None
    ofn.nMaxFileTitle = 0
    ofn.lpstrInitialDir = None
    ofn.lpstrTitle = title
    ofn.Flags = OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST
    ofn.lpstrDefExt = "xlsx"   # 기본 확장자 힌트 추가 ✅

    # 대화상자 호출
    if ctypes.windll.comdlg32.GetOpenFileNameW(ctypes.byref(ofn)):
        return buf.value
    return None

# --------- Core helpers ---------
def get_desktop_path():
    try:
        CSIDL_DESKTOPDIRECTORY = 0x10
        SHGFP_TYPE_CURRENT = 0
        buf = ctypes.create_unicode_buffer(wintypes.MAX_PATH)
        ctypes.windll.shell32.SHGetFolderPathW(None, CSIDL_DESKTOPDIRECTORY, None, SHGFP_TYPE_CURRENT, buf)
        if buf.value and os.path.isdir(buf.value):
            return buf.value
    except Exception:
        pass
    return os.path.join(os.path.expanduser("~"), "Desktop")

def read_workbook_xml_from_zip(xlsx_path):
    with zipfile.ZipFile(xlsx_path, "r") as zf:
        for c in ("xl/workbook.xml", "xl/workBook.xml"):
            try:
                data = zf.read(c)
                return data, c
            except KeyError:
                continue
    raise FileNotFoundError("xl/workbook.xml not found in the .xlsx")

def surgical_filter_defined_names_text(xml_bytes: bytes):
    """
    workbook.xml의 원본 바이트는 최대한 보존하고,
    <definedNames> ... </definedNames> 내부의 <definedName>만 선별 유지.
    유지: Print_Area / Print_Titles (이름은 KEEP_NAMES에 정의)
    남길 게 없으면 <definedNames> 블록 자체 제거.
    """
    text = xml_bytes.decode("utf-8", errors="strict")

    m = re.search(r'(<definedNames\b[^>]*>)(.*?)(</definedNames>)', text, flags=re.S | re.I)
    if not m:
        return xml_bytes, {"total": 0, "kept": 0, "removed": 0}

    head, body, tail = m.group(1), m.group(2), m.group(3)

    total = len(re.findall(r'<definedName\b', body, flags=re.I))
    kept_chunks, kept = [], 0

    for dn in re.finditer(r'(<definedName\b[^>]*>.*?</definedName>)', body, flags=re.S | re.I):
        chunk = dn.group(1)
        nm = re.search(r'\bname\s*=\s*"([^"]*?)"', chunk, flags=re.I)
        if nm and nm.group(1) in KEEP_NAMES:
            kept_chunks.append(chunk)
            kept += 1

    removed = max(total - kept, 0)

    if kept_chunks:
        new_body = "".join(kept_chunks)
        new_block = f"{head}{new_body}{tail}"
        new_text = text[:m.start()] + new_block + text[m.end():]
    else:
        new_text = text[:m.start()] + text[m.end():]

    return new_text.encode("utf-8"), {"total": total, "kept": kept, "removed": removed}

def rewrite_xlsx_with_new_workbook_xml(src_path, dst_path, new_xml_bytes, workbook_xml_path):
    """원본 xlsx의 모든 항목을 복사하되, workbook.xml만 새 바이트로 교체."""
    with zipfile.ZipFile(src_path, "r") as zin, zipfile.ZipFile(dst_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == workbook_xml_path:
                zout.writestr(item, new_xml_bytes)
            else:
                zout.writestr(item, data)

def process_file_gui(xlsx_path):
    if not os.path.isfile(xlsx_path):
        raise FileNotFoundError(f"파일을 찾을 수 없습니다: {xlsx_path}")
    if not xlsx_path.lower().endswith(".xlsx"):
        raise ValueError("지원되는 형식은 .xlsx 입니다.")

    xml_bytes, workbook_xml_path = read_workbook_xml_from_zip(xlsx_path)
    new_xml, stats = surgical_filter_defined_names_text(xml_bytes)

    desktop = get_desktop_path()
    top_dir = os.path.join(desktop, TOP_DIR_NAME)
    os.makedirs(top_dir, exist_ok=True)  # 재사용

    ts = datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
    ts_dir = os.path.join(top_dir, ts)
    os.makedirs(ts_dir, exist_ok=True)

    stem, ext = os.path.splitext(os.path.basename(xlsx_path))
    if not ext:
        ext = ".xlsx"
    backup_path = os.path.join(ts_dir, f"{stem}_backup{ext}")
    cleaned_path = os.path.join(ts_dir, f"{stem}_clean{ext}")

    shutil.copy2(xlsx_path, backup_path)

    tmp_out = cleaned_path + ".tmp"
    rewrite_xlsx_with_new_workbook_xml(xlsx_path, tmp_out, new_xml, workbook_xml_path)
    if os.path.exists(cleaned_path):
        os.remove(cleaned_path)
    os.replace(tmp_out, cleaned_path)

    return backup_path, cleaned_path, stats, ts_dir, top_dir

def main():
    exit_code = 0
    try:
        file_path = open_file_dialog()
        if not file_path:
            return 0  # 취소

        try:
            backup_path, cleaned_path, stats, ts_dir, top_dir = process_file_gui(file_path)
            msg = (
                "정리가 완료되었습니다.\n\n"
                f"[최상위 폴더]\n{top_dir}\n"
                f"[오늘 폴더]\n{ts_dir}\n\n"
                f"- 백업: {backup_path}\n"
                f"- 정리본: {cleaned_path}\n\n"
                f"통계: total={stats['total']}, kept={stats['kept']}, removed={stats['removed']}\n\n"
                "확인을 누르면 저장된 'YYYY-MM-DD-HH-MM-SS' 폴더가 열립니다."
            )
            msg_box(msg, "정리 완료", 0x40)
            os.startfile(ts_dir)
            exit_code = 0
        except Exception as e:
            msg_box(f"오류가 발생했습니다:\n\n{e}", "오류", 0x10)
            exit_code = 2
        return exit_code
    finally:
        gc.collect()

if __name__ == "__main__":
    sys.exit(main())