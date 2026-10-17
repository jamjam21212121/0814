# -*- coding: utf-8 -*-
"""
라벨링 작업 정의 — 같은 프로그램으로 여러 연구를 기록하기 위한 설정.

- SCHMORL : 쉬모를 결절 (종판 11칸: L1~L5 상·하종판 + S1 상종판)   ← run_labeler.bat
- FM      : 압박골절 + 전이암 (척추체 L1~L5, S1 각각)                ← run_fracture.bat

영상 보기·확대·T1·레벨 표시·MicroDicom·이어하기 등은 두 작업이 똑같이 쓰고,
아래의 '무엇을 어떤 칸에 기록하는지'만 다릅니다.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

import config

VERTEBRAE = ["L1", "L2", "L3", "L4", "L5", "S1"]
ENDPLATES = [(v, e) for v in VERTEBRAE for e in ("superior", "inferior") if not (v == "S1" and e == "inferior")]
NORMAL, UNCERTAIN = "Normal", "Uncertain"


@dataclass
class Condition:
    key: str            # 내부 이름
    excel_name: str     # lesions 시트 condition 칸에 들어가는 글자
    positive: str       # 확실할 때 label 값 (labels 칸에도 이 값)
    ko: str             # 화면 이름
    short: str          # 영상 위 짧은 이름
    letter: str         # 입력 키
    color: str          # 표시 색


@dataclass
class Task:
    key: str
    title: str
    excel_path: Path
    capture_dir: Path
    state_path: Path
    conditions: list
    uses_endplate: bool
    sites: list = field(default_factory=list)      # labels 시트의 칸 (환자 1행에 들어가는 칸들)

    def condition(self, key_or_name: str) -> Condition:
        for c in self.conditions:
            if key_or_name in (c.key, c.excel_name, c.positive):
                return c
        return self.conditions[0]

    def condition_of(self, lesion) -> Condition:
        return self.condition(getattr(lesion, "condition", "") or self.conditions[0].excel_name)

    def site_of(self, lesion):
        """병변이 들어갈 labels 칸."""
        if self.uses_endplate:
            return (lesion.vertebra, lesion.endplate)
        return (self.condition_of(lesion).key, lesion.vertebra)

    def site_from_header(self, norm_text: str):
        """머리글 글자 → labels 칸 (해당 없으면 None)."""
        if self.uses_endplate:
            n = norm_text.replace("endplate", "")
            for v, e in ENDPLATES:
                if n.startswith(v.lower()) and e in n:
                    return (v, e)
            return None
        for c in self.conditions:
            for v in VERTEBRAE:
                if norm_text == f"{c.key}{v.lower()}":
                    return (c.key, v)
        return None

    def labels(self) -> set:
        return {c.positive for c in self.conditions} | {UNCERTAIN}


SCHMORL = Task(
    key="schmorl", title="쉬모를 결절",
    excel_path=config.EXCEL_PATH, capture_dir=config.CAPTURE_DIR, state_path=config.STATE_PATH,
    conditions=[Condition("schmorl", "Schmorl Node", "Schmorl", "쉬모를", "쉬모를", "s", "#ff4040")],
    uses_endplate=True, sites=list(ENDPLATES))

FM = Task(
    key="fm", title="압박골절 · 전이암",
    excel_path=config.FM_EXCEL_PATH, capture_dir=config.FM_CAPTURE_DIR, state_path=config.FM_STATE_PATH,
    conditions=[Condition("fracture", "Compression Fracture", "Fracture", "압박골절", "골절", "f", "#ff4040"),
                Condition("metastasis", "Metastasis", "Metastasis", "전이암", "전이", "c", "#e060ff")],
    uses_endplate=False,
    sites=[(c, v) for c in ("fracture", "metastasis") for v in VERTEBRAE])

TASKS = {"schmorl": SCHMORL, "fm": FM}
