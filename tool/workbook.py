# -*- coding: utf-8 -*-
"""
LumbarDISC 엑셀 양식 전용 읽기/쓰기.
(일반적인 엑셀 도우미는 excel_io.py, 이 파일은 양식의 시트·열 이름을 아는 부분)
쉬모를 / 압박골절·전이암 중 어느 엑셀에 어떤 칸을 쓰는지는 tasks.py 의 Task 가 정합니다.
"""
from __future__ import annotations

import copy
from dataclasses import dataclass, field
from datetime import datetime

from openpyxl.formula.translate import Translator

import config
import dataset
import excel_io
import tasks


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


def read_sample_list(path=None) -> list[Patient]:
    """sample_list 시트에서 study_id 가 있는 행만 읽어 순서대로 반환."""
    wb, _ = excel_io.load(path or config.EXCEL_PATH)
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


# ════════════════════════════════════════════════════════
# ④ labels / lesions / sample_list 기록
# ════════════════════════════════════════════════════════
VERTEBRAE = tasks.VERTEBRAE
ENDPLATES = tasks.ENDPLATES
NORMAL, UNCERTAIN = tasks.NORMAL, tasks.UNCERTAIN
SCHMORL = "Schmorl"
CONDITION = "Schmorl Node"
PENDING = "Pending"

LESION_FIELDS = {"lesion_id": "lesionid", "study_id": "studyid", "series_id": "seriesid",
                 "instance": "instancenumber", "condition": "condition", "vertebra": "vertebra",
                 "endplate": "endplate", "x": "x", "y": "y", "label": "label",
                 "capture_file": "capturefile", "labeled_at": "labeledat"}


@dataclass
class Lesion:
    lesion_id: str
    study_id: str
    series_id: str
    instance: int
    vertebra: str          # L1 ~ S1
    endplate: str          # superior / inferior (압박골절·전이암은 "")
    x: float
    y: float
    label: str             # Schmorl·Fracture·Metastasis / Uncertain
    capture_file: str = ""
    labeled_at: str = ""
    condition: str = CONDITION   # Schmorl Node / Compression Fracture / Metastasis


@dataclass
class Record:
    """환자 1명의 판독 결과 (labeled=False 이면 아직 미완료)."""
    labeled: bool = False
    lesions: list = field(default_factory=list)
    labeled_at: str = ""

    def site_values(self, task) -> dict:
        """labels 칸 값: 병변 있으면 확실한 값(우선) 또는 Uncertain, 없으면 Normal."""
        values = {s: NORMAL for s in task.sites}
        for les in self.lesions:
            key = task.site_of(les)
            if key in values and (les.label != UNCERTAIN or values[key] in (NORMAL, UNCERTAIN)):
                values[key] = les.label
        return values

    def endplate_values(self) -> dict:
        """쉬모를 종판 11칸 값."""
        return self.site_values(tasks.SCHMORL)

    def status(self) -> str:
        if not self.labeled:
            return config.STATUS_TODO
        if any(les.label == UNCERTAIN for les in self.lesions):
            return config.STATUS_UNCERTAIN
        return config.STATUS_DONE


def now_text() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M")


def _endplate_key(norm_text: str):
    """'endplatel3inferior' → ('L3', 'inferior')"""
    return tasks.SCHMORL.site_from_header(norm_text)


def site_name(site) -> str:
    return "_".join(site) if isinstance(site, tuple) else str(site)


class SheetMap:
    """시트의 머리글 행·열 위치·머리글 색을 기억하고, 회색 열에만 쓰도록 지켜 줌."""

    def __init__(self, wb, ws, fields: dict):
        self.ws = ws
        self.hr = excel_io.find_header_row(ws)
        if self.hr is None:
            raise RuntimeError(f"[{ws.title}] 시트에서 study_id 머리글을 찾지 못했습니다.")
        self.cols = excel_io.header_columns(ws, self.hr, excel_io.theme_colors(wb))
        self.col: dict = {}
        for c in self.cols:
            for key, wanted in fields.items():
                if (callable(wanted) and wanted(c["norm"]) == key) or c["norm"] == wanted:
                    self.col.setdefault(key, c)

    def missing(self, keys) -> list:
        return [k for k in keys if k not in self.col]

    def not_gray(self, keys) -> list:
        return [k for k in keys if k in self.col and self.col[k]["kind"] != "gray"]

    def get(self, r, key):
        c = self.col.get(key)
        return self.ws.cell(r, c["col"]).value if c else None

    def put(self, r, key, value) -> None:
        """회색 열이고 수식 칸이 아닐 때만 씀."""
        c = self.col.get(key)
        if not c or c["kind"] != "gray":
            return
        cell = self.ws.cell(r, c["col"])
        if not excel_io.is_formula(cell.value):
            cell.value = value

    def rows(self):
        return range(self.hr + 1, self.ws.max_row + 1)

    def find_row(self, key, value) -> int | None:
        for r in self.rows():
            v = self.get(r, key)
            if v not in (None, "") and dataset._id(v) == value:
                return r
        return None

    def free_row(self) -> int:
        """수식 칸을 뺀 모든 칸이 비어 있는 첫 행 (없으면 맨 아래 다음 행)."""
        for r in self.rows():
            if all(excel_io.is_formula(self.ws.cell(r, c["col"]).value)
                   or self.ws.cell(r, c["col"]).value in (None, "") for c in self.cols):
                return r
        return self.ws.max_row + 1

    def ensure_formulas(self, r: int) -> None:
        """양식에 미리 깔린 수식 범위를 넘어선 행이면, 바로 위 행 수식을 복사."""
        for c in self.cols:
            cell = self.ws.cell(r, c["col"])
            if c["kind"] == "green" and cell.value in (None, ""):
                src = self.ws.cell(self.hr + 1, c["col"])
                if excel_io.is_formula(src.value) and isinstance(src.value, str):
                    cell.value = Translator(src.value, origin=src.coordinate).translate_formula(cell.coordinate)
                    cell._style = copy.copy(src._style)


def _lesion_number(lesion_id: str) -> int:
    tail = str(lesion_id).rsplit("_", 1)[-1]
    return int(tail) if tail.isdigit() else 10 ** 6


def _num(s: str):
    return int(s) if str(s).isdigit() else s


class LabelStore:
    """엑셀 ↔ 프로그램 사이의 판독 결과 저장소."""

    def __init__(self, patients: list, task=None):
        self.task = task or tasks.SCHMORL
        self.patients = {p.study_id: p for p in patients}
        self.records: dict = {p.study_id: Record() for p in patients}
        self.dirty: set = set()          # 아직 엑셀에 저장되지 못한 환자
        self._backed_up = False
        self._used_numbers: dict = {}    # 환자별로 이미 쓴 병변 번호

    # ── 시트 찾기 ──
    def _maps(self, wb):
        def sheet(name):
            ws = next((s for s in wb.worksheets if excel_io.norm(s.title) == name), None)
            if ws is None:
                raise RuntimeError(f"'{name}' 시트가 없습니다. 시트 목록: {wb.sheetnames}")
            return ws
        label_fields = {"study_id": "studyid", "labeled_at": "labeledat"}
        task = self.task
        labels = SheetMap(wb, sheet("labels"), {**label_fields,
                                               **{s: (lambda n, s=s: s if task.site_from_header(n) == s else None)
                                                  for s in task.sites}})
        lesions = SheetMap(wb, sheet("lesions"), {**self.lesion_fields(), "review_status": "reviewstatus"})
        sample = SheetMap(wb, sample_sheet(wb), {"study_id": "studyid", "status": "status",
                                                 "labeled_at": "labeledat"})
        return labels, lesions, sample

    def lesion_fields(self) -> dict:
        if self.task.uses_endplate:
            return dict(LESION_FIELDS)
        return {k: v for k, v in LESION_FIELDS.items() if k != "endplate"}

    def _load(self):
        return excel_io.load(self.task.excel_path)

    def check_template(self) -> list:
        """필요한 열이 있는지, 회색인지 확인. 문제 목록 반환(빈 목록이면 정상)."""
        wb, _ = self._load()
        labels, lesions, sample = self._maps(wb)
        problems = []
        need_labels = ["study_id"] + self.task.sites
        need_lesions = list(self.lesion_fields())
        for name, m, keys in [("labels", labels, need_labels), ("lesions", lesions, need_lesions),
                              ("sample_list", sample, ["study_id", "status"])]:
            for k in m.missing(keys):
                problems.append(f"[{name}] '{site_name(k)}' 열을 찾지 못했습니다.")
            for k in m.not_gray(keys):
                problems.append(f"[{name}] '{m.col[k]['text']}' 열 머리글이 회색이 아니라 쓰지 않습니다.")
        return problems

    # ── 읽기 ──
    def load(self) -> None:
        wb, _ = self._load()
        labels, lesions, _ = self._maps(wb)
        task = self.task
        for r in labels.rows():
            sid = labels.get(r, "study_id")
            sid = dataset._id(sid) if sid not in (None, "") else ""
            if sid in self.records and any(labels.get(r, s) not in (None, "") for s in task.sites):
                rec = self.records[sid]
                rec.labeled = True
                rec.labeled_at = str(labels.get(r, "labeled_at") or "")
        for r in lesions.rows():
            sid = lesions.get(r, "study_id")
            sid = dataset._id(sid) if sid not in (None, "") else ""
            if sid not in self.records or lesions.get(r, "vertebra") in (None, ""):
                continue
            cond = task.condition(str(lesions.get(r, "condition") or "").strip())
            try:
                les = Lesion(
                    lesion_id=str(lesions.get(r, "lesion_id") or f"{sid}_row{r}"),
                    study_id=sid, series_id=dataset._id(lesions.get(r, "series_id") or ""),
                    instance=int(float(lesions.get(r, "instance") or 0)),
                    vertebra=str(lesions.get(r, "vertebra")).strip().upper(),
                    endplate=(str(lesions.get(r, "endplate") or "").strip().lower()
                              if task.uses_endplate else ""),
                    x=float(lesions.get(r, "x") or 0), y=float(lesions.get(r, "y") or 0),
                    label=str(lesions.get(r, "label") or cond.positive).strip().capitalize(),
                    capture_file=str(lesions.get(r, "capture_file") or ""),
                    labeled_at=str(lesions.get(r, "labeled_at") or ""),
                    condition=cond.excel_name)
            except (TypeError, ValueError):
                continue
            self.records[sid].lesions.append(les)
            self.records[sid].labeled = True
        for rec in self.records.values():        # 엑셀 행 순서와 상관없이 기록한 순서(번호 순)로
            rec.lesions.sort(key=lambda les: _lesion_number(les.lesion_id))

    # ── 병변 번호 ──
    def new_lesion_id(self, study_id: str) -> str:
        """'{study_id}_01' 형식. 이번 실행에서 지운 번호도 다시 쓰지 않음(기록 혼동 방지)."""
        used = self._used_numbers.setdefault(study_id, set())
        for les in self.records[study_id].lesions:
            tail = les.lesion_id.rsplit("_", 1)[-1]
            if tail.isdigit():
                used.add(int(tail))
        n = max(used, default=0) + 1
        used.add(n)
        return f"{study_id}_{n:02d}"

    def human_inputs(self, lesion_list) -> dict:
        """병변별로 노란 칸(사람 입력)에 적힌 내용 목록. {lesion_id: ["review_status=Confirmed", ...]}"""
        wb, _ = self._load()
        _, lesions, _ = self._maps(wb)
        result = {}
        for lesion in lesion_list:
            r = lesions.find_row("lesion_id", lesion.lesion_id)
            found = []
            if r is not None:
                for c in lesions.cols:
                    if c["kind"] != "yellow":
                        continue
                    v = lesions.ws.cell(r, c["col"]).value
                    if v in (None, "") or (c["norm"] == "reviewstatus" and v == PENDING):
                        continue
                    found.append(f"{c['text']}={v}")
            result[lesion.lesion_id] = found
        return result

    def has_human_input(self, lesion: Lesion) -> list:
        return self.human_inputs([lesion])[lesion.lesion_id]

    # ── 저장 ──
    def save(self) -> None:
        """dirty 환자들을 엑셀에 반영. 실패하면 ExcelLockedError (dirty 는 그대로 남음)."""
        if not self.dirty:
            return
        if not self._backed_up:
            excel_io.backup(self.task.excel_path, tag="session_start")
            self._backed_up = True
        wb, _ = self._load()
        labels, lesions, sample = self._maps(wb)
        for sid in [s for s in self.patients if s in self.dirty]:      # 환자 목록 순서대로
            self._sync(sid, labels, lesions, sample)
        excel_io.save(wb, self.task.excel_path)
        self.dirty.clear()

    def _sync(self, sid, labels, lesions, sample) -> None:
        rec = self.records[sid]
        patient = self.patients[sid]

        # labels: 환자 1행
        r = labels.find_row("study_id", sid)
        if rec.labeled:
            if r is None:
                r = labels.free_row()
                labels.ensure_formulas(r)
            labels.put(r, "study_id", _num(sid))
            for site, value in rec.site_values(self.task).items():
                labels.put(r, site, value)
            labels.put(r, "labeled_at", rec.labeled_at)
        elif r is not None:                       # 판독 취소 → 회색 칸 비우기
            for key in ["study_id", "labeled_at"] + self.task.sites:
                labels.put(r, key, None)

        # lesions: 병변 1개 = 1행
        keep = {les.lesion_id for les in rec.lesions}
        for r2 in lesions.rows():
            v = lesions.get(r2, "study_id")
            lid = lesions.get(r2, "lesion_id")
            # lesion_id 가 없는 행(사람이 직접 적은 행 등)은 건드리지 않음
            if v not in (None, "") and dataset._id(v) == sid and lid not in (None, "") and str(lid) not in keep:
                for key in LESION_FIELDS:
                    lesions.put(r2, key, None)
                c = lesions.col.get("review_status")
                if c and lesions.ws.cell(r2, c["col"]).value == PENDING:   # 프로그램이 넣은 값만 지움
                    lesions.ws.cell(r2, c["col"]).value = None
        for les in rec.lesions:
            r2 = lesions.find_row("lesion_id", les.lesion_id)
            if r2 is None:
                r2 = lesions.free_row()
                lesions.ensure_formulas(r2)
            values = {"lesion_id": les.lesion_id, "study_id": _num(sid), "series_id": _num(les.series_id),
                      "instance": les.instance, "condition": self.task.condition_of(les).excel_name,
                      "vertebra": les.vertebra,
                      "endplate": les.endplate, "x": round(les.x, 1), "y": round(les.y, 1),
                      "label": les.label, "capture_file": les.capture_file or None,
                      "labeled_at": les.labeled_at}
            for key, value in values.items():
                lesions.put(r2, key, value)
            c = lesions.col.get("review_status")
            if les.label == UNCERTAIN and c:          # 노란 칸이지만 비어 있을 때만 Pending
                cell = lesions.ws.cell(r2, c["col"])
                if cell.value in (None, ""):
                    cell.value = PENDING

        # sample_list: status, labeled_at
        r3 = patient.row if dataset._id(sample.get(patient.row, "study_id") or "") == sid \
            else sample.find_row("study_id", sid)
        if r3 is not None:
            sample.put(r3, "status", rec.status())
            sample.put(r3, "labeled_at", rec.labeled_at or None)
