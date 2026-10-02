# -*- coding: utf-8 -*-
"""
설정 파일 — 경로, 시드, 표본 수 같은 값을 여기서 한 번에 관리합니다.
코드를 몰라도 이 파일의 값만 바꾸면 됩니다.
"""
from pathlib import Path

# ── 폴더 ────────────────────────────────────────────────
# 모든 입력(데이터)과 결과물(엑셀, 캡처, 백업, 보고서)은 이 폴더 안에서만 읽고 씁니다.
# tool 폴더를 어디에 두고 실행하든 이 경로가 기준입니다.
PROJECT_DIR = Path(r"C:\Users\kay85\OneDrive\바탕 화면\연구멘토링")
TOOL_DIR = Path(__file__).resolve().parent   # 코드가 있는 곳 (권장: PROJECT_DIR\tool)

DATA_DIR = PROJECT_DIR / "data"
IMAGES_DIR = DATA_DIR / "train_images"


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
