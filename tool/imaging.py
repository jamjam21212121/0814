# -*- coding: utf-8 -*-
"""
DICOM 시리즈 읽기와 밝기/대비(윈도우) 계산.

- 메모리 절약: 현재 환자의 시리즈만 SeriesData 하나로 들고 있고,
  슬라이스는 처음 볼 때 파일에서 읽어 그 시리즈 안에서만 보관합니다.
  환자를 바꾸면 SeriesData 를 새로 만들어 이전 것은 메모리에서 사라집니다.
- 좌표계: train_label_coordinates.csv 와 같은 '원본 픽셀' 기준.
  x = 열(가로, 왼쪽→오른쪽), y = 행(세로, 위→아래), 픽셀 (0,0)의 중심이 (0,0).
"""
from __future__ import annotations

from pathlib import Path

import numpy as np
import pydicom
from PIL import Image

import dataset


class SeriesData:
    def __init__(self, folder: Path):
        self.folder = folder
        self.files = dataset.list_dicom_files(folder)
        # 파일 이름의 숫자 = instance_number (RSNA 데이터 규칙)
        self.instances = [int(f.stem) if f.stem.isdigit() else i + 1
                          for i, f in enumerate(self.files)]
        self._cache: dict[int, np.ndarray] = {}

    def __len__(self) -> int:
        return len(self.files)

    def has_gaps(self) -> bool:
        """instance 번호 중간이 비어 있으면 True (다운로드 중단 의심)."""
        if not self.instances:
            return False
        return max(self.instances) - min(self.instances) + 1 != len(self.instances)

    def index_of_instance(self, instance_number: int) -> int | None:
        try:
            return self.instances.index(int(instance_number))
        except (ValueError, TypeError):
            return None

    def get(self, index: int) -> np.ndarray:
        """index 번째 슬라이스의 픽셀 배열(float32, 2차원)."""
        if index not in self._cache:
            self._cache[index] = read_pixels(self.files[index])
        return self._cache[index]


def read_pixels(path: Path) -> np.ndarray:
    ds = pydicom.dcmread(path)
    arr = ds.pixel_array.astype(np.float32)
    if arr.ndim == 3:                       # 컬러/다중 프레임이면 첫 장만 흑백으로
        arr = arr.mean(axis=-1) if arr.shape[-1] in (3, 4) else arr[0]
    slope = float(getattr(ds, "RescaleSlope", 1) or 1)
    intercept = float(getattr(ds, "RescaleIntercept", 0) or 0)
    arr = arr * slope + intercept
    if str(getattr(ds, "PhotometricInterpretation", "")).upper() == "MONOCHROME1":
        arr = arr.max() - arr
    return arr


def auto_window(arr: np.ndarray) -> tuple[float, float]:
    """기본 밝기/대비: 픽셀값 하위 1% ~ 상위 0.5% 를 화면 밝기 범위로. 반환 (level, width)"""
    lo, hi = np.percentile(arr, [1, 99.5])
    if hi <= lo:
        lo, hi = float(arr.min()), float(arr.max()) + 1
    return (lo + hi) / 2, hi - lo


def adjusted_window(level0: float, width0: float,
                    brightness: float, contrast: float) -> tuple[float, float]:
    """
    밝기·대비 슬라이더(-100~+100, 0=기본값)를 실제 윈도우로 바꿉니다.
    밝기 +  → 화면이 밝아짐 (level 을 내림)
    대비 +  → 대비가 강해짐 (width 를 좁힘, +100 이면 1/4)
    """
    level = level0 - brightness / 100 * width0 / 2
    width = width0 * 2 ** (-contrast / 50)
    return level, max(width, 1e-3)


def to_display(arr: np.ndarray, level: float, width: float) -> Image.Image:
    """픽셀 배열 → 화면용 8비트 흑백 이미지."""
    lo = level - width / 2
    img = np.clip((arr - lo) / width * 255.0, 0, 255).astype(np.uint8)
    return Image.fromarray(img)   # uint8 2차원 → 흑백(L)
