# -*- coding: utf-8 -*-
"""
설정 파일 — 경로, 시드, 표본 수 같은 값을 여기서 한 번에 관리합니다.
코드를 몰라도 이 파일의 값만 바꾸면 됩니다.
"""
from pathlib import Path

# ── 폴더 ────────────────────────────────────────────────
# 이 파일은 <프로젝트 폴더>\tool\config.py 에 있다고 가정합니다.
# 그래서 tool 폴더의 한 단계 위를 프로젝트 폴더로 씁니다.
# (다른 곳에 두었다면 아래 줄을 Path(r"C:\Users\...\연구멘토링") 처럼 바꾸세요.)
PROJECT_DIR = Path(__file__).resolve().parent.parent

DATA_DIR = PROJECT_DIR / "data"
IMAGES_DIR = DATA_DIR / "train_images"
SERIES_CSV = DATA_DIR / "train_series_descriptions.csv"
COORDS_CSV = DATA_DIR / "train_label_coordinates.csv"

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
