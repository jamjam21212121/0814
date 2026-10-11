# -*- coding: utf-8 -*-
"""
LumbarDISC 쉬모를 결절 라벨링 프로그램
  ② 이미지 창과 슬라이스 넘기기
  ③ 레벨 위치 표시 (train_label_coordinates.csv)
  ④ 분류 입력과 엑셀 기록 (입력할 때마다 즉시 저장)
  ⑤ 이어하기: 마지막으로 보던 환자·슬라이스부터 다시 시작  ← 지금 단계
  (⑥ MicroDicom 열기는 다음 단계에서 추가)

실행: run_labeler.bat 더블클릭

조작
  환자      : 왼쪽 목록 클릭, ← → 키, [이전 환자]/[다음 환자] 버튼
  슬라이스  : 마우스 휠, ↑ ↓ 키, 이미지 아래 슬라이더
  밝기/대비 : 이미지 위에서 오른쪽 버튼 드래그 (위아래=밝기, 좌우=대비), 슬라이더, [초기화]
  판독      : N=정상, S=쉬모를 결절, A=애매
              S/A 뒤에 척추(1~6 = L1~S1)와 종판(U=상종판, D=하종판)을 고르고 영상 클릭
              병변이 여러 개면 척추·종판 선택 → 클릭을 반복, 끝나면 Esc
"""
from __future__ import annotations

import json
import os
import sys
import time
import tkinter as tk
import traceback
import tkinter.font as tkfont
from tkinter import messagebox, ttk

from PIL import Image, ImageTk

import config
import dataset
import excel_io
import imaging
import workbook
from workbook import SCHMORL, UNCERTAIN, VERTEBRAE


def log(text: str) -> None:
    """검은 창(콘솔)에 진행 상황 표시."""
    print(text, flush=True)


LEVEL_COLOR = "#7CFC00"        # 라벨 좌표를 찍은 슬라이스: 밝은 연두
LEVEL_COLOR_DIM = "#5f9e5f"    # 다른 슬라이스: 흐린 초록 (위치 참고용)
LESION_COLOR = {SCHMORL: "#ff4040", UNCERTAIN: "#ffa020"}     # 쉬모를=빨강, 애매=주황
LABEL_KO = {SCHMORL: "쉬모를", UNCERTAIN: "애매"}
ENDPLATE_KO = {"superior": "상종판", "inferior": "하종판"}
STATUS_COLORS = {config.STATUS_DONE: "#d6f0d6", config.STATUS_UNCERTAIN: "#ffe2b0"}
# 한글 입력 상태에서 눌린 키 → 영문 키 (두벌식 자판 기준)
HANGUL_TO_KEY = {"ㅜ": "n", "ㄴ": "s", "ㅁ": "a", "ㅕ": "u", "ㅇ": "d"}

HELP_TEXT = """조작법

■ 판독 입력 (오른쪽 칸)
  N = 정상 → 다음 환자로
  S = 쉬모를 결절, A = 애매
   ① 척추 1~6 (L1~L5, S1)
   ② 종판 U=상종판, D=하종판
   ③ 영상에서 병변 클릭
  병변이 여러 개면 ①~③ 반복
  Esc = 병변 입력 끝내기
  Del = 고른 병변 삭제
  Ctrl+S = 저장 다시 시도
  ※ 한/영 키로 영문 상태에서
    사용하세요 (버튼도 됩니다)

■ 환자 이동
  목록 클릭, ← → 키

■ 슬라이스 넘기기
  마우스 휠, ↑ ↓ 키, 슬라이더

■ 밝기 / 대비
  오른쪽 버튼 누른 채 드래그
   · 위아래 = 밝기 · 좌우 = 대비
  슬라이더, [초기화] 버튼

■ 레벨 표시 (L1/L2 ~ L5/S1)
  F2 키, [레벨 표시] 체크
  Home 키 = 좌표 슬라이스로

■ 이 칸 열기/닫기
  F1 키, [조작법 보기] 버튼"""


def key_letter(event) -> str | None:
    """눌린 키를 영문 소문자/숫자로 (한글 입력 상태·숫자패드도 처리)."""
    if event.state & 0x0004:                      # Ctrl 조합은 따로 처리
        return None
    ks = event.keysym
    if len(ks) == 1 and ks.isascii() and ks.isalnum():
        return ks.lower()
    if ks.startswith("KP_") and ks[3:].isdigit():
        return ks[3:]
    if event.char in HANGUL_TO_KEY:
        return HANGUL_TO_KEY[event.char]
    if sys.platform == "win32" and (65 <= event.keycode <= 90 or 48 <= event.keycode <= 57):
        return chr(event.keycode).lower()
    return None


class LabelerApp:
    def __init__(self, root: tk.Tk, patients: list[workbook.Patient], store: workbook.LabelStore):
        self.root = root
        self.patients = patients
        self.store = store
        self.current: int | None = None      # 현재 환자 번호(0부터)
        self.series: imaging.SeriesData | None = None
        self.slice_idx = 0
        self.level0, self.width0 = 0.0, 1.0  # 시리즈 기본 밝기/대비
        self.view = None                     # 화면 배치 (원본 픽셀 ↔ 화면 좌표 변환용)
        self.message: str | None = None      # 영상 대신 보여 줄 안내/오류 문구
        self.coords: list[dict] = []         # 현재 환자의 레벨 좌표 (train_label_coordinates.csv)
        self.mode: str | None = None         # 병변 입력 중이면 SCHMORL 또는 UNCERTAIN
        self._resume_instance: int | None = None
        self._photo = None
        self._drag = None
        self._syncing = False

        self._build_ui()
        self._bind_keys()
        root.protocol("WM_DELETE_WINDOW", self.on_close)
        start, self._resume_instance = self._resume_position()
        self.root.after(50, lambda: self.load_patient(start))

    # ── 화면 구성 ─────────────────────────────────────────
    def _build_ui(self) -> None:
        r = self.root
        r.title("LumbarDISC 쉬모를 결절 라벨링")
        w = min(1700, r.winfo_screenwidth() - 40)
        h = min(900, r.winfo_screenheight() - 80)
        r.geometry(f"{w}x{h}")
        r.minsize(1000, 600)

        status_bar = ttk.Frame(r, relief="sunken", padding=(6, 2))
        status_bar.pack(side="bottom", fill="x")
        self.info_var = tk.StringVar()
        self.cursor_var = tk.StringVar()
        ttk.Label(status_bar, textvariable=self.info_var).pack(side="left")
        ttk.Label(status_bar, textvariable=self.cursor_var).pack(side="right")

        main = ttk.Frame(r)
        main.pack(fill="both", expand=True)
        self._build_patient_list(main)
        self._build_input_panel(main)

        # 조작법 — 기본은 숨김, [조작법 보기] 버튼이나 F1 로 열고 닫음
        self.help_frame = ttk.Frame(main, padding=(4, 6, 8, 6))
        ttk.Label(self.help_frame, text=HELP_TEXT, justify="left").pack(anchor="nw")

        self._build_image_area(main)

    def _build_patient_list(self, main) -> None:
        left = ttk.Frame(main, padding=6)
        left.pack(side="left", fill="y")
        ttk.Label(left, text=f"환자 목록 ({len(self.patients)}명)").pack(anchor="w")
        # 이전/다음 버튼을 먼저 아래에 붙여야 창이 작아도 버튼이 가려지지 않습니다.
        nav = ttk.Frame(left)
        nav.pack(side="bottom", fill="x", pady=(6, 0))
        ttk.Button(nav, text="◀ 이전 환자", command=lambda: self.step_patient(-1)).pack(side="left", expand=True, fill="x")
        ttk.Button(nav, text="다음 환자 ▶", command=lambda: self.step_patient(1)).pack(side="left", expand=True, fill="x")
        box = ttk.Frame(left)
        box.pack(fill="y", expand=True)

        # 행 높이와 열 너비를 실제 글꼴 크기에 맞춤 (Windows 배율 125~150%에서 글자 잘림 방지)
        font = tkfont.nametofont("TkDefaultFont")
        line = font.metrics("linespace")
        style = ttk.Style(self.root)
        style.configure("Treeview", rowheight=line + max(6, line // 2))
        style.configure("Treeview.Heading", padding=(4, max(3, line // 5)))
        pad = font.measure("00")
        widths = {"no": max(font.measure("번호"), font.measure("100")) + pad,
                  "study": font.measure("0000000000") + pad,
                  "status": font.measure("애매포함") + pad}
        self.tree = ttk.Treeview(box, columns=("no", "study", "status"), show="headings",
                                 selectmode="browse", height=10)
        for col, text in [("no", "번호"), ("study", "study_id"), ("status", "상태")]:
            self.tree.heading(col, text=text)
            self.tree.column(col, width=widths[col], anchor="center", stretch=False)
        for status, color in STATUS_COLORS.items():
            self.tree.tag_configure(status, background=color)
        for i, p in enumerate(self.patients):
            status = self.record_of(i).status()
            self.tree.insert("", "end", iid=str(i), values=(p.sample_no or i + 1, p.study_id, status),
                             tags=(status,))
        sb = ttk.Scrollbar(box, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="y")
        sb.pack(side="left", fill="y")
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)

    def _build_input_panel(self, main) -> None:
        """오른쪽: 판독 입력 칸 (N/S/A, 척추·종판 선택, 기록된 병변 목록, 저장 상태)."""
        panel = self.input_panel = ttk.Frame(main, padding=(6, 6, 8, 6))
        panel.pack(side="right", fill="y")
        bold = self._bold_font()
        ttk.Label(panel, text="판독 입력", font=bold).pack(anchor="w")
        self.patient_status_var = tk.StringVar()
        ttk.Label(panel, textvariable=self.patient_status_var).pack(anchor="w", pady=(2, 6))

        row = ttk.Frame(panel)
        row.pack(fill="x")
        ttk.Button(row, text="N 정상", command=self.press_normal).pack(side="left", expand=True, fill="x")
        ttk.Button(row, text="S 쉬모를", command=lambda: self.start_lesion(SCHMORL)).pack(side="left", expand=True, fill="x")
        ttk.Button(row, text="A 애매", command=lambda: self.start_lesion(UNCERTAIN)).pack(side="left", expand=True, fill="x")

        box = self.lesion_box = ttk.LabelFrame(panel, text="병변 입력", padding=6)
        box.pack(fill="x", pady=(10, 0))
        self.mode_var = tk.StringVar(value="꺼짐 (S 또는 A를 누르세요)")
        ttk.Label(box, textvariable=self.mode_var, font=bold).pack(anchor="w")
        ttk.Label(box, text="① 척추 (키 1~6)").pack(anchor="w", pady=(6, 0))
        self.vertebra_var = tk.StringVar(value="")
        vrow = ttk.Frame(box)
        vrow.pack(fill="x")
        for v in VERTEBRAE:
            ttk.Radiobutton(vrow, text=v, value=v, variable=self.vertebra_var, style="Toolbutton",
                            command=lambda v=v: self.select_vertebra(v)).pack(side="left", expand=True, fill="x")
        ttk.Label(box, text="② 종판 (U=상종판, D=하종판)").pack(anchor="w", pady=(6, 0))
        self.endplate_var = tk.StringVar(value="")
        erow = ttk.Frame(box)
        erow.pack(fill="x")
        self.endplate_buttons = {}
        for e in ("superior", "inferior"):
            b = ttk.Radiobutton(erow, text=f"{ENDPLATE_KO[e]} ({'U' if e == 'superior' else 'D'})",
                                value=e, variable=self.endplate_var, style="Toolbutton",
                                command=lambda e=e: self.select_endplate(e))
            b.pack(side="left", expand=True, fill="x")
            self.endplate_buttons[e] = b
        self.hint_var = tk.StringVar()
        wrap = tkfont.nametofont("TkDefaultFont").measure("가") * 17   # 글꼴 크기에 맞춘 줄바꿈 폭
        ttk.Label(box, textvariable=self.hint_var, wraplength=wrap, justify="left",
                  foreground="#1f5fbf").pack(anchor="w", pady=(6, 0))
        ttk.Button(box, text="병변 입력 끝내기 (Esc)", command=self.end_lesion).pack(fill="x", pady=(6, 0))

        lf = ttk.LabelFrame(panel, text="기록된 병변", padding=6)
        lf.pack(fill="both", expand=True, pady=(10, 0))
        self.lesion_list = tk.Listbox(lf, height=7, activestyle="none", exportselection=False)
        self.lesion_list.pack(fill="both", expand=True)
        self.lesion_list.bind("<<ListboxSelect>>", lambda e: self.render())
        self.lesion_list.bind("<Double-Button-1>", lambda e: self.go_to_lesion())
        brow = ttk.Frame(lf)
        brow.pack(fill="x", pady=(4, 0))
        ttk.Button(brow, text="그 슬라이스로", command=self.go_to_lesion).pack(side="left", expand=True, fill="x")
        ttk.Button(brow, text="삭제 (Del)", command=self.delete_lesion).pack(side="left", expand=True, fill="x")

        self.save_var = tk.StringVar(value="")
        self.save_label = tk.Label(panel, textvariable=self.save_var, anchor="w", justify="left",
                                   wraplength=wrap)
        self.save_label.pack(fill="x", pady=(8, 0))

    def _build_image_area(self, main) -> None:
        center = self.center_frame = ttk.Frame(main, padding=6)
        center.pack(side="left", fill="both", expand=True)
        # 제목 줄: 오른쪽 [조작법] 버튼을 먼저 붙여서 창이 좁아도 버튼은 보이게 (제목이 대신 잘림)
        title_row = ttk.Frame(center)
        title_row.pack(side="top", fill="x", pady=(0, 4))
        self.help_button = ttk.Button(title_row, text="조작법 보기 (F1)", command=self.toggle_help)
        self.help_button.pack(side="right")
        self.title_var = tk.StringVar()
        ttk.Label(title_row, textvariable=self.title_var, font=self._bold_font()).pack(side="left")
        # 슬라이더·버튼 줄을 먼저 아래에 붙이고, 남는 공간을 영상이 차지합니다.
        row2 = ttk.Frame(center)
        row2.pack(side="bottom", fill="x", pady=(4, 0))
        row = ttk.Frame(center)
        row.pack(side="bottom", fill="x", pady=(6, 0))
        self.canvas = tk.Canvas(center, bg="black", width=480, height=480,
                                highlightthickness=0, cursor="crosshair")
        self.canvas.pack(fill="both", expand=True)

        # 슬라이스 줄: 오른쪽 글자·체크를 먼저 붙이고 슬라이더는 마지막에 (창이 좁으면 슬라이더가 줄어듦)
        ttk.Label(row, text="슬라이스").pack(side="left")
        self.show_levels = tk.BooleanVar(value=True)
        ttk.Checkbutton(row, text="레벨 표시 (F2)", variable=self.show_levels,
                        command=self.render).pack(side="right", padx=(10, 0))
        self.slice_text = tk.StringVar()
        ttk.Label(row, textvariable=self.slice_text, width=14).pack(side="right")
        self.slice_scale = ttk.Scale(row, from_=0, to=1, orient="horizontal", command=self._on_slice_scale)
        self.slice_scale.pack(side="left", fill="x", expand=True, padx=6)

        self.bright = tk.DoubleVar(value=0)
        self.contrast = tk.DoubleVar(value=0)
        # [초기화] 버튼을 먼저 붙이고, 밝기·대비 슬라이더는 남는 폭에 맞춰 늘고 줄어듦
        ttk.Button(row2, text="초기화", command=self.reset_window).pack(side="right")
        for text, var in [("밝기", self.bright), ("대비", self.contrast)]:
            ttk.Label(row2, text=text).pack(side="left")
            ttk.Scale(row2, from_=-100, to=100, orient="horizontal", variable=var, length=60,
                      command=lambda _v: self.render()).pack(side="left", padx=(4, 12), fill="x", expand=True)

        c = self.canvas
        c.bind("<Configure>", lambda e: self.render())
        c.bind("<Motion>", self._on_motion)
        c.bind("<Leave>", lambda e: self.cursor_var.set(""))
        c.bind("<Button-1>", self._on_left_click)
        c.bind("<ButtonPress-3>", self._on_drag_start)
        c.bind("<B3-Motion>", self._on_drag)
        c.bind("<ButtonRelease-3>", lambda e: setattr(self, "_drag", None))

    @property
    def _level_font(self):
        if not hasattr(self, "_level_font_obj"):
            f = tkfont.nametofont("TkDefaultFont").copy()
            f.configure(weight="bold")
            self._level_font_obj = f
        return self._level_font_obj

    def _bold_font(self):
        f = tkfont.nametofont("TkDefaultFont").copy()
        f.configure(weight="bold", size=f.cget("size") + 1)
        return f

    def _bind_keys(self) -> None:
        """
        모든 위젯 앞에 공통 태그를 붙여, 어느 칸이 선택돼 있어도 키가 같은 동작을 하게 합니다.
        (예: 목록이 선택된 상태에서 ↑↓ 를 눌러도 목록이 아니라 슬라이스가 넘어감)
        """
        tag = "LabelerKeys"
        actions = {"<Up>": lambda: self.step_slice(-1), "<Down>": lambda: self.step_slice(1),
                   "<Left>": lambda: self.step_patient(-1), "<Right>": lambda: self.step_patient(1),
                   "<F1>": self.toggle_help, "<F2>": self.toggle_levels,
                   "<Home>": self.go_to_label_slice, "<Escape>": self.end_lesion,
                   "<Delete>": self.delete_lesion, "<Control-s>": self.retry_save,
                   "<Control-S>": self.retry_save}
        for seq, fn in actions.items():
            self.root.bind_class(tag, seq, lambda e, f=fn: (f(), "break")[1])
        self.root.bind_class(tag, "<Key>", self._on_key)
        self.root.bind_class(tag, "<MouseWheel>", self._on_wheel)          # Windows
        self.root.bind_class(tag, "<Button-4>", lambda e: self._on_wheel(e, -1))  # Linux
        self.root.bind_class(tag, "<Button-5>", lambda e: self._on_wheel(e, 1))

        def add(w):
            w.bindtags((tag,) + tuple(t for t in w.bindtags() if t != tag))
            for child in w.winfo_children():
                add(child)
        add(self.root)

    def _on_key(self, event):
        k = key_letter(event)
        if k is None:
            return None
        if k == "n":
            self.press_normal()
        elif k == "s":
            self.start_lesion(SCHMORL)
        elif k == "a":
            self.start_lesion(UNCERTAIN)
        elif k in "123456" and len(k) == 1:
            self.select_vertebra(VERTEBRAE[int(k) - 1])
        elif k == "u":
            self.select_endplate("superior")
        elif k == "d":
            self.select_endplate("inferior")
        else:
            return None
        return "break"

    def toggle_help(self) -> None:
        if self.help_frame.winfo_ismapped():
            self.help_frame.pack_forget()
            self.help_button.configure(text="조작법 보기 (F1)")
        else:
            self.help_frame.pack(side="right", fill="y", before=self.center_frame)
            self.help_button.configure(text="조작법 닫기 (F1)")

    # ── ⑤ 이어하기 ───────────────────────────────────────
    def _resume_position(self) -> tuple[int, int | None]:
        """지난번 마지막 환자·슬라이스. 기록이 없으면 첫 '미완료' 환자."""
        first_todo = next((i for i in range(len(self.patients))
                           if self.record_of(i).status() == config.STATUS_TODO), 0)
        try:
            state = json.loads(config.STATE_PATH.read_text(encoding="utf-8"))
            sid = str(state.get("study_id", ""))
            index = next(i for i, p in enumerate(self.patients) if p.study_id == sid)
        except (OSError, ValueError, StopIteration, AttributeError):
            log(f"이어하기 기록 없음 → 첫 미완료 환자({first_todo + 1}번)부터 시작")
            return first_todo, None
        log(f"이어하기: 지난번 마지막 환자 {index + 1}번 (study {sid}, instance {state.get('instance')}) "
            f"— {state.get('saved_at', '')}")
        return index, state.get("instance")

    def save_state(self) -> None:
        """현재 환자·슬라이스를 기록 (임시 파일에 쓴 뒤 교체해서 파일이 깨지지 않게)."""
        if self.current is None:
            return
        state = {"study_id": self.patients[self.current].study_id,
                 "patient_no": self.current + 1,
                 "instance": self.current_instance(),
                 "saved_at": workbook.now_text()}
        try:
            tmp = config.STATE_PATH.with_suffix(".tmp")
            tmp.write_text(json.dumps(state, ensure_ascii=False, indent=1), encoding="utf-8")
            os.replace(tmp, config.STATE_PATH)
        except OSError:
            log("[경고] 이어하기 위치 저장 실패\n" + traceback.format_exc())

    # ── 환자 ──────────────────────────────────────────────
    def record_of(self, index: int) -> workbook.Record:
        return self.store.records[self.patients[index].study_id]

    @property
    def record(self) -> workbook.Record | None:
        return self.record_of(self.current) if self.current is not None else None

    def _on_tree_select(self, _event=None) -> None:
        sel = self.tree.selection()
        if sel:
            self.load_patient(int(sel[0]))

    def step_patient(self, delta: int) -> None:
        if self.current is None:
            return
        new = min(max(self.current + delta, 0), len(self.patients) - 1)
        self.load_patient(new)

    def load_patient(self, index: int) -> None:
        if index == self.current and self.series is not None:
            return
        self.current = index
        self.mode = None
        p = self.patients[index]
        self.tree.selection_set(str(index))
        self.tree.see(str(index))
        self.title_var.set(f"[{index + 1}/{len(self.patients)}] study {p.study_id}")
        self.hint_var.set("")
        self.lesion_list.selection_clear(0, "end")
        self.refresh_panel()

        self.series = None                       # 이전 환자 영상은 메모리에서 해제
        self.coords = []
        self.bright.set(0)
        self.contrast.set(0)
        folder = dataset.series_dir(p.study_id, p.t2_series)
        log(f"[환자 {index + 1}] study {p.study_id} 영상 읽는 중: {folder}")
        self.message = "영상 불러오는 중...\n(OneDrive 온라인 전용 파일이면 처음엔 오래 걸릴 수 있습니다)"
        self.render()
        self.root.update_idletasks()

        t0 = time.time()
        series = imaging.SeriesData(folder)
        self.series = series
        if not len(series):
            self.slice_idx = 0
            self.message = f"Sagittal T2 영상 파일이 없습니다.\n{folder}"
            log("  → " + self.message.replace("\n", " "))
            self.render()
            return
        try:
            self.coords = dataset.read_coordinates(p.study_id, p.t2_series)
        except Exception:
            self.coords = []
            log("  → 좌표 CSV 읽기 실패\n" + traceback.format_exc())
        label_idx = self.label_slice_index()
        # 레벨 좌표를 찍은 슬라이스부터, 없으면 가운데 슬라이스부터 (이어하기면 지난번 슬라이스)
        self.slice_idx = label_idx if label_idx is not None else len(series) // 2
        if self._resume_instance is not None:
            resume_idx = series.index_of_instance(self._resume_instance)
            if resume_idx is not None:
                self.slice_idx = resume_idx
            self._resume_instance = None
        try:
            self.level0, self.width0 = imaging.auto_window(series.get(self.slice_idx))
        except Exception as e:
            self.message = f"DICOM을 읽지 못했습니다.\n{series.files[self.slice_idx]}\n\n{e!r}"
            log("  → DICOM 읽기 실패\n" + traceback.format_exc())
            self.render()
            return
        self.message = None
        log(f"  → {len(series)}장, 레벨 좌표 {len(self.coords)}개, 첫 화면 instance "
            f"{series.instances[self.slice_idx]} ({time.time() - t0:.1f}초)")
        self._syncing = True
        self.slice_scale.configure(to=max(len(series) - 1, 1))
        self.slice_scale.set(self.slice_idx)
        self._syncing = False
        self.render()
        self.save_state()

    # ── ④ 판독 입력 ───────────────────────────────────────
    def press_normal(self) -> None:
        rec = self.record
        if rec is None:
            return
        if rec.lesions:
            blocked = self._human_input_of(rec.lesions)
            if blocked:
                messagebox.showwarning("정상으로 바꿀 수 없음",
                                       "교수님 입력이 있는 병변이 있어 지울 수 없습니다.\n\n" + "\n".join(blocked)
                                       + "\n\n엑셀에서 직접 확인해 주세요.")
                return
            if not messagebox.askyesno("정상으로 바꾸기",
                                       f"기록된 병변 {len(rec.lesions)}개를 지우고 '정상'으로 바꿀까요?"):
                return
            for les in rec.lesions:
                self._remove_capture(les)
        self.mode = None
        rec.lesions = []
        rec.labeled = True
        rec.labeled_at = workbook.now_text()
        self.commit()
        if config.AUTO_NEXT_AFTER_NORMAL and self.current < len(self.patients) - 1:
            self.load_patient(self.current + 1)

    def start_lesion(self, label: str) -> None:
        if self.current is None:
            return
        self.mode = label
        self.vertebra_var.set("")
        self.endplate_var.set("")
        self.refresh_panel()
        self.render()

    def end_lesion(self) -> None:
        if self.mode is not None:
            self.hint_var.set("")
        self.mode = None
        self.vertebra_var.set("")
        self.endplate_var.set("")
        self.refresh_panel()
        self.render()

    def select_vertebra(self, v: str) -> None:
        if self.mode is None:
            self.hint_var.set("먼저 S(쉬모를) 또는 A(애매)를 누르세요.")
            return
        self.vertebra_var.set(v)
        if v == "S1":                              # S1 은 상종판만
            self.endplate_var.set("superior")
        self.refresh_panel()
        self.render()

    def select_endplate(self, e: str) -> None:
        if self.mode is None:
            self.hint_var.set("먼저 S(쉬모를) 또는 A(애매)를 누르세요.")
            return
        if e == "inferior" and self.vertebra_var.get() == "S1":
            self.hint_var.set("S1은 상종판만 기록합니다.")
            self.root.bell()
            return
        self.endplate_var.set(e)
        self.refresh_panel()
        self.render()

    def add_lesion(self, x: float, y: float) -> None:
        p = self.patients[self.current]
        rec = self.record
        v, e = self.vertebra_var.get(), self.endplate_var.get()
        les = workbook.Lesion(lesion_id=self.store.new_lesion_id(p.study_id), study_id=p.study_id,
                              series_id=p.t2_series, instance=self.current_instance(), vertebra=v,
                              endplate=e, x=x, y=y, label=self.mode, labeled_at=workbook.now_text())
        if self.mode == UNCERTAIN:
            les.capture_file = self._save_capture(les, rec)
        rec.lesions.append(les)
        rec.labeled = True
        rec.labeled_at = les.labeled_at
        log(f"  병변 기록: {les.lesion_id} {v} {e} {les.label} instance {les.instance} "
            f"x={x:.1f} y={y:.1f}" + (f" 캡처 {les.capture_file}" if les.capture_file else ""))
        self.vertebra_var.set("")                  # 다음 병변은 척추·종판을 새로 고르도록
        self.endplate_var.set("")
        self.commit()
        self.lesion_list.selection_clear(0, "end")
        self.hint_var.set(f"기록됨: {v} {ENDPLATE_KO[e]} ({LABEL_KO[les.label]}).\n"
                          "병변이 더 있으면 척추·종판을 고르고 클릭, 끝났으면 Esc 또는 → 다음 환자.")

    def _capture_name(self, les: workbook.Lesion, rec: workbook.Record) -> str:
        base = f"{les.study_id}_{les.vertebra}{'sup' if les.endplate == 'superior' else 'inf'}"
        used = {x.capture_file for x in rec.lesions}
        name, n = f"{base}.png", 2
        while name in used or (config.CAPTURE_DIR / name).exists():
            name, n = f"{base}_{n}.png", n + 1
        return name

    def _save_capture(self, les: workbook.Lesion, rec: workbook.Record) -> str:
        try:
            config.CAPTURE_DIR.mkdir(parents=True, exist_ok=True)
            name = self._capture_name(les, rec)
            level, width = imaging.adjusted_window(self.level0, self.width0,
                                                   self.bright.get(), self.contrast.get())
            title = (f"study {les.study_id} | {les.vertebra} {les.endplate} | Uncertain | "
                     f"instance {les.instance} | x={les.x:.1f}, y={les.y:.1f}")
            img = imaging.make_capture(self.series.get(self.slice_idx), level, width, (les.x, les.y),
                                       title, self.coords, les.instance)
            img.save(config.CAPTURE_DIR / name)
            return name
        except Exception as e:
            log("[오류] 캡처 저장 실패\n" + traceback.format_exc())
            messagebox.showwarning("캡처 저장 실패", f"병변은 기록하지만 캡처 이미지를 저장하지 못했습니다.\n{e}")
            return ""

    def _remove_capture(self, les: workbook.Lesion) -> None:
        if les.capture_file:
            try:
                (config.CAPTURE_DIR / les.capture_file).unlink()
            except OSError:
                pass

    def _human_input_of(self, lesions) -> list[str]:
        """노란 칸(교수님 입력)에 내용이 있는 병변 목록."""
        try:
            inputs = self.store.human_inputs(lesions)
        except Exception:
            log("[경고] 교수님 입력 확인 실패\n" + traceback.format_exc())
            inputs = {}
        return [f"{lid}: " + ", ".join(items) for lid, items in inputs.items() if items]

    def selected_lesion(self) -> workbook.Lesion | None:
        rec = self.record
        sel = self.lesion_list.curselection()
        if rec is None or not sel or sel[0] >= len(rec.lesions):
            return None
        return rec.lesions[sel[0]]

    def delete_lesion(self) -> None:
        les = self.selected_lesion()
        if les is None:
            self.hint_var.set("지울 병변을 '기록된 병변' 목록에서 먼저 고르세요.")
            return
        blocked = self._human_input_of([les])
        if blocked:
            messagebox.showwarning("삭제할 수 없음", "교수님 입력이 있는 병변이라 지울 수 없습니다.\n\n"
                                   + "\n".join(blocked) + "\n\n엑셀에서 직접 확인해 주세요.")
            return
        if not messagebox.askyesno("병변 삭제", f"{self._lesion_text(les)}\n이 병변을 지울까요?"):
            return
        rec = self.record
        rec.lesions.remove(les)
        self._remove_capture(les)
        if not rec.lesions:                        # 병변이 하나도 없으면 다시 미완료
            rec.labeled = False
            rec.labeled_at = ""
            self.mode = None
            self.vertebra_var.set("")
            self.endplate_var.set("")
            self.hint_var.set("병변이 모두 지워져 '미완료'로 돌아갔습니다. 정상이면 N을 누르세요.")
        else:
            rec.labeled_at = workbook.now_text()
        self.commit()

    def go_to_lesion(self) -> None:
        les = self.selected_lesion()
        if les is None or not self.series:
            return
        idx = self.series.index_of_instance(les.instance)
        if idx is not None:
            self.step_slice(idx - self.slice_idx)
            self.render()

    def _lesion_text(self, les: workbook.Lesion) -> str:
        num = les.lesion_id.rsplit("_", 1)[-1]
        return f"#{num} {les.vertebra} {ENDPLATE_KO.get(les.endplate, les.endplate)} " \
               f"{LABEL_KO.get(les.label, les.label)}  inst {les.instance}"

    # ── 저장 ──────────────────────────────────────────────
    def commit(self) -> None:
        """현재 환자 결과를 바로 엑셀에 저장하고 화면을 갱신."""
        self.store.dirty.add(self.patients[self.current].study_id)
        self.save()
        self.save_state()
        self.refresh_panel()
        self.refresh_patient_row(self.current)
        self.render()

    def save(self) -> bool:
        try:
            self.store.save()
        except excel_io.ExcelLockedError as e:
            log("[저장 실패] " + str(e).replace("\n", " "))
            self._set_save_state(False)
            messagebox.showwarning("엑셀 저장 실패",
                                   "엑셀 파일이 열려 있어 저장하지 못했습니다.\n\n"
                                   "엑셀을 닫은 뒤 Ctrl+S 를 누르거나, 다음 입력 때 자동으로 다시 저장합니다.\n"
                                   "(입력한 내용은 프로그램이 기억하고 있습니다)")
            return False
        except Exception as e:
            log("[저장 실패]\n" + traceback.format_exc())
            self._set_save_state(False)
            messagebox.showerror("엑셀 저장 실패", f"{e!r}\n\n자세한 내용은 검은 창을 확인하세요.")
            return False
        self._set_save_state(True)
        return True

    def retry_save(self) -> None:
        if not self.store.dirty:
            self._set_save_state(True)
            return
        if self.save():
            for i in range(len(self.patients)):
                self.refresh_patient_row(i)

    def _set_save_state(self, ok: bool) -> None:
        if ok:
            self.save_var.set(f"✔ 엑셀 저장됨 ({time.strftime('%H:%M:%S')})")
            self.save_label.configure(fg="#1a7f37")
        else:
            self.save_var.set(f"✖ 저장 안 됨 ({len(self.store.dirty)}명)\n엑셀을 닫고 Ctrl+S")
            self.save_label.configure(fg="#c62828")

    def on_close(self) -> None:
        while self.store.dirty:
            again = messagebox.askretrycancel(
                "저장되지 않은 내용", f"아직 엑셀에 저장되지 않은 환자가 {len(self.store.dirty)}명 있습니다.\n"
                "엑셀을 닫고 [다시 시도]를 누르세요.")
            if not again:
                if messagebox.askyesno("그래도 종료", "저장하지 않고 종료하면 이 내용은 사라집니다.\n종료할까요?"):
                    break
                continue
            try:
                self.store.save()
            except Exception:
                log("[저장 실패]\n" + traceback.format_exc())
        self.save_state()
        self.root.destroy()

    # ── 화면 갱신 ─────────────────────────────────────────
    def refresh_patient_row(self, index: int) -> None:
        p = self.patients[index]
        status = self.record_of(index).status()
        self.tree.item(str(index), values=(p.sample_no or index + 1, p.study_id, status), tags=(status,))

    def refresh_panel(self) -> None:
        rec = self.record
        if rec is None:
            return
        status = rec.status()
        detail = {config.STATUS_TODO: "아직 판독 전",
                  config.STATUS_DONE: "정상" if not rec.lesions else f"병변 {len(rec.lesions)}개",
                  config.STATUS_UNCERTAIN: f"병변 {len(rec.lesions)}개 (애매 포함)"}[status]
        self.patient_status_var.set(f"현재 환자: {status}  ·  {detail}")
        if self.mode is None:
            self.mode_var.set("꺼짐 (S 또는 A를 누르세요)")
        else:
            self.mode_var.set(f"{'S 쉬모를 결절' if self.mode == SCHMORL else 'A 애매'} 입력 중")
            v, e = self.vertebra_var.get(), self.endplate_var.get()
            if not v:
                self.hint_var.set("① 척추를 고르세요 (키 1~6).")
            elif not e:
                self.hint_var.set(f"{v} 선택됨. ② 종판을 고르세요 (U=상종판, D=하종판).")
            else:
                self.hint_var.set(f"{v} {ENDPLATE_KO[e]} → ③ 영상에서 병변 위치를 클릭하세요.")
        state = ["!disabled"] if self.vertebra_var.get() != "S1" else ["disabled"]
        self.endplate_buttons["inferior"].state(state)
        sel = self.lesion_list.curselection()
        self.lesion_list.delete(0, "end")
        for les in rec.lesions:
            self.lesion_list.insert("end", self._lesion_text(les))
            self.lesion_list.itemconfigure("end", foreground=LESION_COLOR[les.label] if les.label in LESION_COLOR else "black")
        if sel and sel[0] < len(rec.lesions):
            self.lesion_list.selection_set(sel[0])

    # ── 레벨 좌표 ─────────────────────────────────────────
    def label_slice_index(self) -> int | None:
        """레벨 좌표가 가장 많이 찍힌 instance 의 슬라이스 번호 (없으면 None)."""
        if not self.coords or not self.series:
            return None
        counts: dict[int, int] = {}
        for c in self.coords:
            counts[c["instance"]] = counts.get(c["instance"], 0) + 1
        for inst in sorted(counts, key=lambda k: -counts[k]):
            idx = self.series.index_of_instance(inst)
            if idx is not None:
                return idx
        return None

    def go_to_label_slice(self) -> None:
        idx = self.label_slice_index()
        if idx is not None:
            self.step_slice(idx - self.slice_idx)

    def toggle_levels(self) -> None:
        self.show_levels.set(not self.show_levels.get())
        self.render()

    def _outlined_text(self, x, y, text, color, anchor, tags) -> None:
        for dx, dy in ((1, 1), (-1, -1), (1, -1), (-1, 1)):       # 글자 테두리(검정)
            self.canvas.create_text(x + dx, y + dy, text=text, anchor=anchor, fill="black",
                                    font=self._level_font, tags=tags)
        self.canvas.create_text(x, y, text=text, anchor=anchor, fill=color, font=self._level_font, tags=tags)

    def draw_levels(self) -> None:
        """
        L1/L2 ~ L5/S1 위치 표시. 점은 원래 좌표에 두고, 글자는 오른쪽으로 떨어진 곳에 한 줄로 모아
        점과 선으로 이어 줍니다 (글자가 종판 경계를 가리지 않도록).
        좌표를 찍은 슬라이스는 진하게, 다른 슬라이스는 흐리게.
        """
        c = self.canvas
        inst = self.current_instance()
        points = []
        for p in self.coords:
            if 0 <= p["x"] < self.view["w"] and 0 <= p["y"] < self.view["h"]:
                cx, cy = self.pixel_to_canvas(p["x"], p["y"])
                same = p["instance"] == inst
                text = p["level"] if same else f'{p["level"]} (inst {p["instance"]})'
                points.append((cx, cy, same, text))
        if not points:
            return
        img_w = self.view["w"] * self.view["sx"]
        label_x = max(pt[0] for pt in points) + max(60, img_w * 0.15)        # 글자를 쓸 세로줄
        widest = max(self._level_font.measure(pt[3]) for pt in points)
        label_x = min(label_x, c.winfo_width() - widest - 8)                  # 화면 밖으로 안 나가게
        for cx, cy, same, text in points:
            color = LEVEL_COLOR if same else LEVEL_COLOR_DIM
            r = 5 if same else 4
            c.create_oval(cx - r, cy - r, cx + r, cy + r, outline=color, width=2,
                          fill=color if same else "", tags="level")
            if label_x > cx + r + 10:
                c.create_line(cx + r, cy, label_x - 3, cy, fill=color, width=1,
                              dash=() if same else (3, 3), tags="level")
            self._outlined_text(max(label_x, cx + r + 10), cy, text, color, "w", "level")

    def draw_lesions(self) -> None:
        """기록된 병변: 같은 슬라이스면 진한 원 + 번호, 다른 슬라이스면 점선 원 (inst 표시)."""
        rec = self.record
        if rec is None:
            return
        inst = self.current_instance()
        selected = self.selected_lesion()
        for les in rec.lesions:
            cx, cy = self.pixel_to_canvas(les.x, les.y)
            color = LESION_COLOR.get(les.label, "#ff4040")
            num = les.lesion_id.rsplit("_", 1)[-1]
            short = f"#{num} {les.vertebra}{'상' if les.endplate == 'superior' else '하'}"
            if les.instance == inst:
                r = 9
                width = 4 if les is selected else 2
                self.canvas.create_oval(cx - r, cy - r, cx + r, cy + r, outline=color, width=width, tags="lesion")
                self.canvas.create_line(cx - 3, cy, cx + 3, cy, fill=color, tags="lesion")
                self.canvas.create_line(cx, cy - 3, cx, cy + 3, fill=color, tags="lesion")
                self._outlined_text(cx + r + 4, cy, short, color, "w", "lesion")
            else:
                r = 7
                self.canvas.create_oval(cx - r, cy - r, cx + r, cy + r, outline=color, width=1,
                                        dash=(3, 3), tags="lesion")
                self._outlined_text(cx + r + 4, cy, f"{short} (inst {les.instance})", color, "w", "lesion")

    def draw_mode_banner(self) -> None:
        if self.mode is None:
            return
        v, e = self.vertebra_var.get(), self.endplate_var.get()
        target = f"{v} {ENDPLATE_KO[e]} → 병변 클릭" if v and e else "척추·종판을 고르세요"
        text = f"{'S 쉬모를' if self.mode == SCHMORL else 'A 애매'} 입력 중: {target}   (Esc 끝내기)"
        self._outlined_text(10, 12, text, LESION_COLOR[self.mode], "w", "banner")

    # ── 슬라이스 ──────────────────────────────────────────
    def step_slice(self, delta: int) -> None:
        if not self.series or not len(self.series):
            return
        new = min(max(self.slice_idx + delta, 0), len(self.series) - 1)
        if new != self.slice_idx:
            self.slice_idx = new
            self._syncing = True
            self.slice_scale.set(new)
            self._syncing = False
            self.render()

    def _on_slice_scale(self, value) -> None:
        if self._syncing or not self.series or not len(self.series):
            return
        new = min(max(int(round(float(value))), 0), len(self.series) - 1)
        if new != self.slice_idx:
            self.slice_idx = new
            self.render()

    def _on_wheel(self, event, direction: int | None = None):
        # 이미지 위에서만 슬라이스를 넘기고, 목록 위에서는 목록이 스크롤되게 둡니다.
        if self.root.winfo_containing(event.x_root, event.y_root) is not self.canvas:
            return None
        if direction is None:
            direction = -1 if event.delta > 0 else 1   # 휠 아래로 = 다음 슬라이스
        self.step_slice(direction)
        return "break"

    # ── 밝기/대비 ─────────────────────────────────────────
    def reset_window(self) -> None:
        self.bright.set(0)
        self.contrast.set(0)
        self.render()

    def _on_drag_start(self, e) -> None:
        self._drag = (e.x, e.y, self.bright.get(), self.contrast.get())

    def _on_drag(self, e) -> None:
        if not self._drag:
            return
        x0, y0, b0, c0 = self._drag
        self.contrast.set(min(max(c0 + (e.x - x0) * 0.5, -100), 100))
        self.bright.set(min(max(b0 - (e.y - y0) * 0.5, -100), 100))
        self.render()

    # ── 그리기와 좌표 변환 ────────────────────────────────
    def render(self) -> None:
        c = self.canvas
        cw, ch = c.winfo_width(), c.winfo_height()
        if cw < 50 or ch < 50:                   # 창이 아직 배치되기 전이면 잠시 뒤 다시
            self.root.after(100, self.render)
            return
        c.delete("all")
        self.view = None
        arr, message = None, self.message
        if message is None and self.series is not None and len(self.series):
            try:
                arr = self.series.get(self.slice_idx)
            except Exception as e:
                message = f"DICOM을 읽지 못했습니다.\n{self.series.files[self.slice_idx]}\n\n{e!r}"
                log(traceback.format_exc())
        if arr is None:
            c.create_text(cw / 2, ch / 2, text=message or "", fill="white", justify="center",
                          width=cw - 40)
            self._update_info()
            return

        h, w = arr.shape
        scale = min(cw / w, ch / h)
        dw, dh = max(1, round(w * scale)), max(1, round(h * scale))
        ox, oy = (cw - dw) / 2, (ch - dh) / 2
        level, width = imaging.adjusted_window(self.level0, self.width0,
                                               self.bright.get(), self.contrast.get())
        img = imaging.to_display(arr, level, width).resize((dw, dh), Image.Resampling.BILINEAR)
        self._photo = ImageTk.PhotoImage(img)
        c.create_image(ox, oy, anchor="nw", image=self._photo, tags="image")
        self.view = {"ox": ox, "oy": oy, "sx": dw / w, "sy": dh / h, "w": w, "h": h}
        if self.show_levels.get():
            self.draw_levels()
        self.draw_lesions()
        self.draw_mode_banner()
        self._update_info()

    def pixel_to_canvas(self, x: float, y: float) -> tuple[float, float]:
        v = self.view
        return v["ox"] + (x + 0.5) * v["sx"], v["oy"] + (y + 0.5) * v["sy"]

    def canvas_to_pixel(self, cx: float, cy: float) -> tuple[float, float] | None:
        """화면 좌표 → 원본 픽셀 좌표 (이미지 밖이면 None)."""
        v = self.view
        if not v:
            return None
        x = (cx - v["ox"]) / v["sx"] - 0.5
        y = (cy - v["oy"]) / v["sy"] - 0.5
        if -0.5 <= x < v["w"] - 0.5 and -0.5 <= y < v["h"] - 0.5:
            return x, y
        return None

    def current_instance(self) -> int | None:
        if not self.series or not len(self.series):
            return None
        return self.series.instances[self.slice_idx]

    def _update_info(self) -> None:
        if not self.series or not len(self.series) or self.message:
            self.slice_text.set("")
            self.info_var.set("영상 없음" if self.series is not None else "")
            return
        n = len(self.series)
        self.slice_text.set(f"{self.slice_idx + 1}/{n}  inst {self.current_instance()}")
        gap = "   ※ instance 번호 일부 누락 (다운로드 미완료 의심)" if self.series.has_gaps() else ""
        series_id = self.patients[self.current].t2_series if self.current is not None else ""
        if self.coords:
            insts = sorted({c["instance"] for c in self.coords})
            levels = f"   레벨 좌표 {len(self.coords)}개 (instance {', '.join(map(str, insts))})"
        else:
            levels = "   레벨 좌표 없음"
        self.info_var.set(f"series {series_id}   슬라이스 {self.slice_idx + 1}/{n}   "
                          f"밝기 {self.bright.get():+.0f}   대비 {self.contrast.get():+.0f}{levels}{gap}")

    def _pixel_text(self, cx, cy) -> str | None:
        pos = self.canvas_to_pixel(cx, cy)
        if pos is None or self.message:
            return None
        x, y = pos
        value = self.series.get(self.slice_idx)[int(round(y)), int(round(x))]
        return f"x={x:.1f}, y={y:.1f}  (픽셀값 {value:.0f})"

    def _on_motion(self, e) -> None:
        text = self._pixel_text(e.x, e.y)
        self.cursor_var.set(f"커서  {text}" if text else "")

    def _on_left_click(self, e) -> None:
        pos = self.canvas_to_pixel(e.x, e.y)
        if pos is None or self.message:
            return
        if self.mode is not None:
            if not (self.vertebra_var.get() and self.endplate_var.get()):
                self.hint_var.set("척추와 종판을 먼저 고른 뒤 클릭하세요.")
                self.root.bell()
                return
            self.add_lesion(*pos)
            return
        # 입력 중이 아닐 때: 클릭 위치 확인용 표시
        self.cursor_var.set(f"클릭  instance {self.current_instance()},  {self._pixel_text(e.x, e.y)}")
        r = 5
        self.canvas.delete("probe")
        self.canvas.create_oval(e.x - r, e.y - r, e.x + r, e.y + r, outline="#00e5ff",
                                width=2, tags="probe")


def _setup_fonts() -> None:
    if sys.platform == "win32":
        for name in ("TkDefaultFont", "TkTextFont", "TkHeadingFont", "TkMenuFont"):
            try:
                tkfont.nametofont(name).configure(family="Malgun Gothic", size=10)
            except tk.TclError:
                pass


def main() -> None:
    if sys.platform == "win32":
        try:  # 고해상도 화면에서 글자가 흐릿하지 않게
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
    log(f"LumbarDISC 라벨링 프로그램 시작 (코드 버전 {config.TOOL_VERSION})")
    log("이 검은 창을 닫으면 프로그램도 꺼집니다. 오류가 나면 이 창의 글자를 복사해 알려 주세요.")
    root = tk.Tk()
    _setup_fonts()

    def on_error(exc, val, tb):
        text = "".join(traceback.format_exception(exc, val, tb))
        log("[오류]\n" + text)
        messagebox.showerror("오류", f"{val!r}\n\n자세한 내용은 검은 창을 확인하세요.")
    root.report_callback_exception = on_error

    log(f"엑셀 읽는 중: {config.EXCEL_PATH}")
    try:
        patients = workbook.read_sample_list()
        store = workbook.LabelStore(patients)
        problems = store.check_template()
        store.load()
    except Exception as e:
        root.withdraw()
        log("[오류] 엑셀 읽기 실패\n" + traceback.format_exc())
        messagebox.showerror("엑셀 읽기 실패", f"{config.EXCEL_PATH}\n\n{e}")
        return
    if not patients:
        root.withdraw()
        messagebox.showerror("환자 목록 없음",
                             "sample_list 시트가 비어 있습니다.\n먼저 step1_make_sample.bat 을 실행하세요.")
        return
    if problems:
        root.withdraw()
        log("[오류] 엑셀 양식 문제\n  " + "\n  ".join(problems))
        messagebox.showerror("엑셀 양식 문제", "기록에 필요한 열을 찾지 못했습니다.\n\n" + "\n".join(problems))
        return
    done = sum(r.labeled for r in store.records.values())
    log(f"환자 {len(patients)}명 읽음 (판독 완료 {done}명). 프로그램 창을 엽니다.")
    if excel_io.lock_file_exists():
        messagebox.showwarning("엑셀이 열려 있음",
                               "엑셀에서 결과 파일이 열려 있는 것 같습니다.\n"
                               "열려 있으면 저장이 실패하니 엑셀을 닫아 주세요.")
    LabelerApp(root, patients, store)
    root.mainloop()
    log("프로그램 종료")


if __name__ == "__main__":
    main()
