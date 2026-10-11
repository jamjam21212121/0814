# -*- coding: utf-8 -*-
"""
설정 파일 — 경로, 시드, 표본 수 같은 값을 여기서 한 번에 관리합니다.
코드를 몰라도 이 파일의 값만 바꾸면 됩니다.
"""
from pathlib import Path

# 코드 버전 — 실행 화면 맨 위에 표시됩니다. 최신 코드를 받았는지 확인할 때 보세요.
TOOL_VERSION = "2026-10-15"

# ── 폴더 ────────────────────────────────────────────────
# 모든 입력(데이터)과 결과물(엑셀, 캡처, 백업, 보고서)은 이 폴더 안에서만 읽고 씁니다.
# tool 폴더를 어디에 두고 실행하든 이 경로가 기준입니다.
PROJECT_DIR = Path(r"C:\Users\kay85\OneDrive\바탕 화면\연구멘토링")
TOOL_DIR = Path(__file__).resolve().parent   # 코드가 있는 곳 (권장: PROJECT_DIR\tool)

DATA_DIR = PROJECT_DIR / "data"

# ── DICOM 영상 폴더 (train_images) ─────────────────────────
# 영상을 USB 등 다른 곳에 두었다면, 아래 순서로 찾습니다.
#   ① IMAGES_DIR_OVERRIDE 에 적은 경로 (예: Path(r"E:\RSNA\train_images"))
#   ② 노트북의 연구멘토링\data\train_images
#   ③ D: ~ Z: 드라이브에서 train_images 폴더 (드라이브 바로 아래 또는 2단계 아래 폴더까지)
# USB 드라이브 문자(E:, F: ...)가 바뀌어도 ③에서 자동으로 찾으므로 보통은 비워 두면 됩니다.
IMAGES_DIR_OVERRIDE = None
IMAGES_DIR_NAME = "train_images"


def _has_studies(folder: Path) -> bool:
    """study 폴더(숫자 이름)가 하나라도 있으면 영상 폴더로 인정."""
    try:
        return folder.is_dir() and any(p.is_dir() and p.name.isdigit() for p in folder.iterdir())
    except OSError:
        return False


def find_images_dir(override=None, local=None, roots=None, name=IMAGES_DIR_NAME):
    """영상 폴더 찾기. 못 찾으면 local 경로를 그대로 돌려줌(없다는 안내가 뜨도록)."""
    import string
    import sys
    local = local or DATA_DIR / name
    if override and _has_studies(Path(override)):
        return Path(override)
    if _has_studies(local):
        return local
    if roots is None:
        roots = [Path(f"{d}:\\") for d in string.ascii_uppercase[3:]] if sys.platform == "win32" else []
    for root in roots:
        try:
            if not root.is_dir():
                continue
            candidates = [root / name]
            for sub in sorted(root.iterdir()):
                if sub.is_dir() and not sub.name.startswith(("$", ".")) and sub.name != "System Volume Information":
                    candidates.append(sub / name)
                    try:
                        candidates += [s / name for s in sorted(sub.iterdir()) if s.is_dir()]
                    except OSError:
                        pass
        except OSError:
            continue
        for c in candidates:
            if _has_studies(c):
                return c
    return Path(override) if override else local


IMAGES_DIR = find_images_dir(IMAGES_DIR_OVERRIDE)


def _find_csv(name: str) -> Path:
    """data\\ 에 없으면 data 의 바로 아래 하위 폴더(예: data\\새 폴더\\)에서 찾습니다."""
    direct = DATA_DIR / name
    if direct.exists() or not DATA_DIR.is_dir():
        return direct
    for sub in sorted(DATA_DIR.iterdir()):
        if sub.is_dir() and sub.name != "train_images" and (sub / name).exists():
            return sub / name
    return direct


SERIES_CSV = _find_csv("train_series_descriptions.csv")
COORDS_CSV = _find_csv("train_label_coordinates.csv")

LABELING_DIR = PROJECT_DIR / "labeling"
EXCEL_PATH = LABELING_DIR / "LumbarDISC_Schmorl_labeling_v0.1.xlsx"
CAPTURE_DIR = LABELING_DIR / "captures"
BACKUP_DIR = LABELING_DIR / "backup"   # 엑셀을 처음 고치기 전 자동 백업 위치

# ── ① 표본 추출 ──────────────────────────────────────────
RANDOM_SEED = 2024     # 시드 고정: 같은 데이터면 몇 번 돌려도 같은 100명이 뽑힙니다.
SAMPLE_SIZE = 100

# series_description 값 (train_series_descriptions.csv 에 적힌 그대로)
DESC_SAG_T2 = "Sagittal T2/STIR"
DESC_SAG_T1 = "Sagittal T1"

# ── 엑셀 sample_list 의 status 값 ─────────────────────────
# README 시트에 다른 표기가 정해져 있으면 여기만 바꾸면 됩니다.
STATUS_TODO = "미완료"
STATUS_DONE = "완료"
STATUS_UNCERTAIN = "애매포함"

# ── ④ 판독 입력 ───────────────────────────────────────────
# N(정상)을 누르면 자동으로 다음 환자로 넘어갈지 (False 로 바꾸면 그대로 머묾)
AUTO_NEXT_AFTER_NORMAL = True

# ── ⑤ 이어하기 ────────────────────────────────────────────
# 마지막으로 보던 환자·슬라이스를 기억하는 파일 (지우면 첫 '미완료' 환자부터 시작)
STATE_PATH = LABELING_DIR / "labeler_state.json"

# ── ⑥ MicroDicom 자동 열기 ───────────────────────────────
MICRODICOM_PATH = Path(r"C:\Program Files\MicroDicom\mDicom.exe")
# MicroDicom 에 넘기는 인자. {folder} 자리에 현재 환자의 Sagittal T2 폴더가 들어갑니다.
# (공식 설명서: mDicom.exe -fd <파일 또는 폴더>. 버전에 따라 안 되면 ["{folder}"] 로 바꿔 보세요)
MICRODICOM_ARGS = ["-fd", "{folder}"]
AUTO_OPEN_MICRODICOM = True        # 환자를 바꾸면 자동으로 열기 (프로그램 안 체크박스로도 끄고 켤 수 있음)
MICRODICOM_DELAY_MS = 800          # 환자에 이 시간(ms) 이상 머물러야 열기 (빠르게 넘길 때 창이 쌓이지 않게)

# ── 화면 표시 ─────────────────────────────────────────────
# 엑셀 status 값은 그대로 두고, 프로그램 화면에만 다른 이름으로 보여 줄 때
STATUS_DISPLAY = {"애매포함": "확인필요"}
LEVEL_ZOOM = 3.0                   # Shift+1~5 (레벨로 바로 확대) 때 확대 배율
LEVEL_ZOOM_ANTERIOR_MM = 12.0      # 확대 중심을 초록점(척추관)에서 앞쪽(척추체 쪽)으로 옮기는 거리(mm)
