# -*- coding: utf-8 -*-
"""
RSNA 2024 Lumbar Spine 데이터(CSV, DICOM 폴더)를 읽는 공통 함수 모음.
study_id / series_id 는 모두 문자열로 다룹니다 (엑셀·폴더명과 비교하기 쉽도록).
"""
from __future__ import annotations

import csv
from pathlib import Path

import config


def _id(value) -> str:
    """'4003253', '4003253.0', 4003253 → '4003253' 로 통일."""
    s = str(value).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s


def id_sort_key(value):
    s = _id(value)
    return (0, int(s)) if s.isdigit() else (1, s)


def read_series_descriptions(path: Path = config.SERIES_CSV) -> dict[str, list[tuple[str, str]]]:
    """{study_id: [(series_id, series_description), ...]}"""
    result: dict[str, list[tuple[str, str]]] = {}
    with open(path, newline="", encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            study = _id(row["study_id"])
            result.setdefault(study, []).append(
                (_id(row["series_id"]), row["series_description"].strip())
            )
    return result


def series_dir(study_id, series_id) -> Path:
    return config.IMAGES_DIR / _id(study_id) / _id(series_id)


def list_dicom_files(folder: Path) -> list[Path]:
    """폴더 안의 .dcm 파일을 instance_number(파일명 숫자) 순서로 정렬해서 반환."""
    if not folder.is_dir():
        return []
    files = [p for p in folder.iterdir() if p.suffix.lower() == ".dcm"]
    return sorted(files, key=lambda p: id_sort_key(p.stem))


def downloaded_study_ids() -> list[str]:
    """train_images 아래에 폴더가 있는 study_id 목록."""
    if not config.IMAGES_DIR.is_dir():
        return []
    ids = [p.name for p in config.IMAGES_DIR.iterdir() if p.is_dir()]
    return sorted(ids, key=id_sort_key)


def studies_with_coordinates(path: Path = config.COORDS_CSV) -> set[tuple[str, str]]:
    """좌표 라벨이 있는 (study_id, series_id) 쌍."""
    pairs = set()
    with open(path, newline="", encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            pairs.add((_id(row["study_id"]), _id(row["series_id"])))
    return pairs


def pick_series(study_id: str, description: str,
                series_map: dict[str, list[tuple[str, str]]],
                labeled_pairs: set[tuple[str, str]] | None = None) -> tuple[str, bool]:
    """
    study 안에서 description 이 맞는 series 하나를 고릅니다.
    같은 종류가 여러 개면 ① 다운로드된 것 ② 좌표 라벨이 있는 것 ③ series_id 작은 것 순으로 우선.
    반환: (series_id 또는 "", 다운로드 여부)
    """
    candidates = [sid for sid, desc in series_map.get(study_id, []) if desc == description]
    if not candidates:
        return "", False

    def rank(sid):
        downloaded = bool(list_dicom_files(series_dir(study_id, sid)))
        labeled = labeled_pairs is not None and (study_id, sid) in labeled_pairs
        return (not downloaded, not labeled, id_sort_key(sid))

    best = min(candidates, key=rank)
    return best, bool(list_dicom_files(series_dir(study_id, best)))
