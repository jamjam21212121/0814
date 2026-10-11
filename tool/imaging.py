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
from PIL import Image, ImageDraw, ImageFont

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


def _font(size: int):
    for name in ("malgun.ttf", "arial.ttf", "DejaVuSans.ttf"):
        try:
            return ImageFont.truetype(name, size)
        except OSError:
            continue
    try:
        return ImageFont.load_default(size=size)
    except TypeError:
        return ImageFont.load_default()


def make_capture(arr: np.ndarray, level: float, width: float, lesion: tuple[float, float],
                 title: str, levels: list[dict], instance: int, min_size: int = 640) -> Image.Image:
    """
    애매 병변 캡처: 현재 밝기/대비의 슬라이스 + 클릭 위치(주황 원) + 레벨 표시 + 설명 글자.
    작은 영상은 min_size 픽셀 이상으로 키워서 저장합니다.
    """
    img = to_display(arr, level, width).convert("RGB")
    h, w = arr.shape
    s = max(1.0, min_size / max(h, w))
    if s > 1:
        img = img.resize((round(w * s), round(h * s)), Image.Resampling.BILINEAR)
    draw = ImageDraw.Draw(img)
    to_img = lambda x, y: ((x + 0.5) * s - 0.5, (y + 0.5) * s - 0.5)
    small = _font(max(12, img.width // 45))

    # 레벨 위치: 점은 원래 자리, 글자는 오른쪽으로 떨어진 곳에 모아 선으로 연결 (같은 슬라이스면 진하게)
    pts = [(*to_img(p["x"], p["y"]), p) for p in levels if 0 <= p["x"] < w and 0 <= p["y"] < h]
    if pts:
        label_x = max(x for x, _, _ in pts) + max(40, img.width * 0.15)
        widest = max(draw.textlength(p["level"], font=small) for _, _, p in pts)
        label_x = min(label_x, img.width - widest - 6)
        for cx, cy, p in pts:
            color = (124, 252, 0) if p["instance"] == instance else (95, 158, 95)
            draw.ellipse([cx - 4, cy - 4, cx + 4, cy + 4], outline=color, width=2)
            tx = max(label_x, cx + 14)
            if tx > cx + 14:
                draw.line([cx + 4, cy, tx - 3, cy], fill=color, width=1)
            draw.text((tx, cy), p["level"], fill=color, font=small, anchor="lm", stroke_width=2, stroke_fill="black")

    cx, cy = to_img(*lesion)                          # 병변 위치
    r = max(10, img.width // 40)
    draw.ellipse([cx - r, cy - r, cx + r, cy + r], outline=(255, 140, 0), width=3)
    draw.line([cx - r * 1.8, cy, cx - r * 1.2, cy], fill=(255, 140, 0), width=3)
    draw.line([cx + r * 1.2, cy, cx + r * 1.8, cy], fill=(255, 140, 0), width=3)

    draw.rectangle([0, 0, img.width, small.size * 2 if hasattr(small, "size") else 30], fill=(0, 0, 0))
    draw.text((8, 4), title, fill=(255, 200, 80), font=small)
    return img
