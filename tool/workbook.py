# -*- coding: utf-8 -*-
"""
LumbarDISC 엑셀 양식 전용 읽기/쓰기.
(일반적인 엑셀 도우미는 excel_io.py, 이 파일은 양식의 시트·열 이름을 아는 부분)
"""
from __future__ import annotations

from dataclasses import dataclass

import config
import dataset
import excel_io


@dataclass
class Patient:
    row: int            # sample_list 시트의 엑셀 행 번호
    sample_no: str
    study_id: str
    t2_series: str
    t1_series: str
    status: str


def map_sample_columns(cols: list[dict]) -> dict[str, dict]:
    """sample_list 머리글 → 역할(study_id, t2, t1, status, order)."""
    found = {}
    for c in cols:
        n = c["norm"]
        if n == "studyid":
            found.setdefault("study_id", c)
        elif "t2" in n and "axial" not in n and ("series" in n or "sag" in n):
            found.setdefault("t2", c)
        elif "t1" in n and ("series" in n or "sag" in n):
            found.setdefault("t1", c)
        elif n == "status":
            found.setdefault("status", c)
        elif n in {"no", "order", "seq", "index", "sampleno", "sampleorder", "순번", "번호"}:
            found.setdefault("order", c)
    return found


def sample_sheet(wb):
    return next((s for s in wb.worksheets if excel_io.norm(s.title) == "samplelist"), None)


def read_sample_list() -> list[Patient]:
    """sample_list 시트에서 study_id 가 있는 행만 읽어 순서대로 반환."""
    wb, _ = excel_io.load()
    ws = sample_sheet(wb)
    if ws is None:
        raise RuntimeError(f"sample_list 시트가 없습니다. 시트 목록: {wb.sheetnames}")
    hr = excel_io.find_header_row(ws)
    if hr is None:
        raise RuntimeError("sample_list 시트에서 study_id 머리글을 찾지 못했습니다.")
    cols = map_sample_columns(excel_io.header_columns(ws, hr, excel_io.theme_colors(wb)))

    def val(r, key):
        c = cols.get(key)
        v = ws.cell(r, c["col"]).value if c else None
        return "" if v is None else dataset._id(v)

    patients = []
    for r in range(hr + 1, ws.max_row + 1):
        if val(r, "study_id"):
            patients.append(Patient(r, val(r, "order"), val(r, "study_id"), val(r, "t2"),
                                    val(r, "t1"), val(r, "status") or config.STATUS_TODO))
    return patients
