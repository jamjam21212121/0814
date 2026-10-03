# -*- coding: utf-8 -*-
"""
LumbarDISC 쉬모를 결절 라벨링 프로그램
  ② 이미지 창과 슬라이스 넘기기  ← 지금 단계
  (③ 레벨 표시, ④ 분류 입력·엑셀 기록, ⑤ 이어하기, ⑥ MicroDicom 열기는 단계별로 추가)

실행: run_labeler.bat 더블클릭

조작
  환자      : 왼쪽 목록 클릭, ← → 키, [이전 환자]/[다음 환자] 버튼
  슬라이스  : 마우스 휠, ↑ ↓ 키, 이미지 아래 슬라이더
  밝기/대비 : 이미지 위에서 오른쪽 버튼 드래그 (위아래=밝기, 좌우=대비), 슬라이더, [초기화]
"""
from __future__ import annotations

import sys
import time
import tkinter as tk
import traceback
import tkinter.font as tkfont
from tkinter import messagebox, ttk

from PIL import Image, ImageTk

import config
import dataset
import imaging
import workbook

def log(text: str) -> None:
    """검은 창(콘솔)에 진행 상황 표시."""
    print(text, flush=True)


STATUS_COLORS = {config.STATUS_DONE: "#d6f0d6", config.STATUS_UNCERTAIN: "#ffe2b0"}
HELP_TEXT = """조작법

■ 환자 이동
  목록 클릭, ← → 키

■ 슬라이스 넘기기
  마우스 휠, ↑ ↓ 키,
  이미지 아래 슬라이더

■ 밝기 / 대비
  이미지 위에서
  오른쪽 버튼 누른 채 드래그
   · 위아래 = 밝기
   · 좌우 = 대비
  슬라이더, [초기화] 버튼

■ 좌표 확인
  마우스를 올리면 아래
  상태줄에 원본 픽셀 x, y
  (클릭하면 그 위치 고정 표시)

■ 이 칸 열기/닫기
  F1 키, [조작법 보기] 버튼

※ 레벨 표시(③), 분류 입력
   N/S/A(④)는 다음 단계에서
   추가됩니다."""


class LabelerApp:
    def __init__(self, root: tk.Tk, patients: list[workbook.Patient]):
        self.root = root
        self.patients = patients
        self.current: int | None = None      # 현재 환자 번호(0부터)
        self.series: imaging.SeriesData | None = None
        self.slice_idx = 0
        self.level0, self.width0 = 0.0, 1.0  # 시리즈 기본 밝기/대비
        self.view = None                     # 화면 배치 (원본 픽셀 ↔ 화면 좌표 변환용)
        self.message: str | None = None      # 영상 대신 보여 줄 안내/오류 문구
        self._photo = None
        self._drag = None
        self._syncing = False

        self._build_ui()
        self._bind_keys()
        start = next((i for i, p in enumerate(patients) if p.status == config.STATUS_TODO), 0)
        self.root.after(50, lambda: self.load_patient(start))

    # ── 화면 구성 ─────────────────────────────────────────
    def _build_ui(self) -> None:
        r = self.root
        r.title("LumbarDISC 쉬모를 결절 라벨링")
        r.geometry("1150x780")
        r.minsize(900, 600)

        status_bar = ttk.Frame(r, relief="sunken", padding=(6, 2))
        status_bar.pack(side="bottom", fill="x")
        self.info_var = tk.StringVar()
        self.cursor_var = tk.StringVar()
        ttk.Label(status_bar, textvariable=self.info_var).pack(side="left")
        ttk.Label(status_bar, textvariable=self.cursor_var).pack(side="right")

        main = ttk.Frame(r)
        main.pack(fill="both", expand=True)

        # 왼쪽: 환자 목록
        left = ttk.Frame(main, padding=6)
        left.pack(side="left", fill="y")
        ttk.Label(left, text=f"환자 목록 ({len(self.patients)}명)").pack(anchor="w")
        box = ttk.Frame(left)
        box.pack(fill="y", expand=True)
        self.tree = ttk.Treeview(box, columns=("no", "study", "status"), show="headings",
                                 selectmode="browse", height=25)
        for col, text, width in [("no", "번호", 46), ("study", "study_id", 112), ("status", "상태", 70)]:
            self.tree.heading(col, text=text)
            self.tree.column(col, width=width, anchor="center", stretch=False)
        for status, color in STATUS_COLORS.items():
            self.tree.tag_configure(status, background=color)
        for i, p in enumerate(self.patients):
            self.tree.insert("", "end", iid=str(i), values=(p.sample_no or i + 1, p.study_id, p.status),
                             tags=(p.status,))
        sb = ttk.Scrollbar(box, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="y")
        sb.pack(side="left", fill="y")
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)
        nav = ttk.Frame(left)
        nav.pack(fill="x", pady=(6, 0))
        ttk.Button(nav, text="◀ 이전 환자", command=lambda: self.step_patient(-1)).pack(side="left", expand=True, fill="x")
        ttk.Button(nav, text="다음 환자 ▶", command=lambda: self.step_patient(1)).pack(side="left", expand=True, fill="x")

        # 오른쪽: 조작법 — 기본은 숨김, [조작법 보기] 버튼이나 F1 로 열고 닫음
        self.help_frame = ttk.Frame(main, padding=(4, 6, 8, 6))
        ttk.Label(self.help_frame, text=HELP_TEXT, justify="left").pack(anchor="nw")

        # 가운데: 이미지
        center = self.center_frame = ttk.Frame(main, padding=6)
        center.pack(side="left", fill="both", expand=True)
        self.title_var = tk.StringVar()
        ttk.Label(center, textvariable=self.title_var, font=self._bold_font()).pack(anchor="w", pady=(0, 4))
        self.canvas = tk.Canvas(center, bg="black", width=560, height=560,
                                highlightthickness=0, cursor="crosshair")
        self.canvas.pack(fill="both", expand=True)

        row = ttk.Frame(center)
        row.pack(fill="x", pady=(6, 0))
        ttk.Label(row, text="슬라이스").pack(side="left")
        self.slice_scale = ttk.Scale(row, from_=0, to=1, orient="horizontal", command=self._on_slice_scale)
        self.slice_scale.pack(side="left", fill="x", expand=True, padx=6)
        self.slice_text = tk.StringVar()
        ttk.Label(row, textvariable=self.slice_text, width=28).pack(side="left")

        row2 = ttk.Frame(center)
        row2.pack(fill="x", pady=(4, 0))
        self.bright = tk.DoubleVar(value=0)
        self.contrast = tk.DoubleVar(value=0)
        for text, var in [("밝기", self.bright), ("대비", self.contrast)]:
            ttk.Label(row2, text=text).pack(side="left")
            ttk.Scale(row2, from_=-100, to=100, orient="horizontal", variable=var, length=170,
                      command=lambda _v: self.render()).pack(side="left", padx=(4, 14))
        ttk.Button(row2, text="밝기/대비 초기화", command=self.reset_window).pack(side="left")
        self.help_button = ttk.Button(row2, text="조작법 보기 (F1)", command=self.toggle_help)
        self.help_button.pack(side="right")

        c = self.canvas
        c.bind("<Configure>", lambda e: self.render())
        c.bind("<Motion>", self._on_motion)
        c.bind("<Leave>", lambda e: self.cursor_var.set(""))
        c.bind("<Button-1>", self._on_left_click)
        c.bind("<ButtonPress-3>", self._on_drag_start)
        c.bind("<B3-Motion>", self._on_drag)
        c.bind("<ButtonRelease-3>", lambda e: setattr(self, "_drag", None))

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
                   "<F1>": self.toggle_help}
        for seq, fn in actions.items():
            self.root.bind_class(tag, seq, lambda e, f=fn: (f(), "break")[1])
        self.root.bind_class(tag, "<MouseWheel>", self._on_wheel)          # Windows
        self.root.bind_class(tag, "<Button-4>", lambda e: self._on_wheel(e, -1))  # Linux
        self.root.bind_class(tag, "<Button-5>", lambda e: self._on_wheel(e, 1))

        def add(w):
            w.bindtags((tag,) + tuple(t for t in w.bindtags() if t != tag))
            for child in w.winfo_children():
                add(child)
        add(self.root)

    def toggle_help(self) -> None:
        if self.help_frame.winfo_ismapped():
            self.help_frame.pack_forget()
            self.help_button.configure(text="조작법 보기 (F1)")
        else:
            self.help_frame.pack(side="right", fill="y", before=self.center_frame)
            self.help_button.configure(text="조작법 닫기 (F1)")

    # ── 환자 ──────────────────────────────────────────────
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
        p = self.patients[index]
        self.tree.selection_set(str(index))
        self.tree.see(str(index))
        self.title_var.set(f"[{index + 1}/{len(self.patients)}] study {p.study_id}  —  "
                           f"Sagittal T2 (series {p.t2_series})")

        self.series = None                       # 이전 환자 영상은 메모리에서 해제
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
        self.slice_idx = len(series) // 2        # 가운데 슬라이스부터
        try:
            self.level0, self.width0 = imaging.auto_window(series.get(self.slice_idx))
        except Exception as e:
            self.message = f"DICOM을 읽지 못했습니다.\n{series.files[self.slice_idx]}\n\n{e!r}"
            log("  → DICOM 읽기 실패\n" + traceback.format_exc())
            self.render()
            return
        self.message = None
        log(f"  → {len(series)}장, 가운데 슬라이스 읽기 완료 ({time.time() - t0:.1f}초)")
        self._syncing = True
        self.slice_scale.configure(to=max(len(series) - 1, 1))
        self.slice_scale.set(self.slice_idx)
        self._syncing = False
        self.render()

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
        self.slice_text.set(f"{self.slice_idx + 1} / {n}  (instance_number {self.current_instance()})")
        level, width = imaging.adjusted_window(self.level0, self.width0,
                                               self.bright.get(), self.contrast.get())
        gap = "   ※ instance 번호 일부 누락 (다운로드 미완료 의심)" if self.series.has_gaps() else ""
        self.info_var.set(f"슬라이스 {self.slice_idx + 1}/{n}   밝기 {self.bright.get():+.0f}  "
                          f"대비 {self.contrast.get():+.0f}   (윈도우 L {level:.0f} / W {width:.0f}){gap}")

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
        # ②단계: 클릭 위치 확인용 (④단계에서 병변 기록으로 바뀝니다)
        text = self._pixel_text(e.x, e.y)
        if text:
            self.cursor_var.set(f"클릭  instance {self.current_instance()},  {text}")
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
    log(f"환자 {len(patients)}명 읽음. 프로그램 창을 엽니다.")
    LabelerApp(root, patients)
    root.mainloop()
    log("프로그램 종료")


if __name__ == "__main__":
    main()
