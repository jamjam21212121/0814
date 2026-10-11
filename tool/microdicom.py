# -*- coding: utf-8 -*-
"""
⑥ MicroDicom 으로 현재 환자의 Sagittal T2 시리즈 열기.

- 이 프로그램이 연 MicroDicom 창만 기억했다가, 다음 환자를 열 때 그 창을 닫고 새로 엽니다.
  (사용자가 직접 연 MicroDicom 창은 건드리지 않음)
- 실행 파일 위치·인자는 config.py 의 MICRODICOM_PATH / MICRODICOM_ARGS 에서 바꿀 수 있습니다.
"""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path

import config

_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)


def find_exe() -> Path | None:
    """config 경로 → Program Files 의 MicroDicom 폴더 순서로 mDicom.exe 를 찾음."""
    if config.MICRODICOM_PATH.exists():
        return config.MICRODICOM_PATH
    for base in (Path(r"C:\Program Files"), Path(r"C:\Program Files (x86)")):
        if not base.is_dir():
            continue
        for d in base.iterdir():
            exe = d / "mDicom.exe"
            if "microdicom" in d.name.lower() and exe.exists():
                return exe
    return None


class MicroDicom:
    def __init__(self):
        self.proc: subprocess.Popen | None = None
        self.opened_folder: Path | None = None

    def close_previous(self) -> None:
        """이 프로그램이 열었던 MicroDicom 창을 닫음."""
        proc, self.proc = self.proc, None
        if proc is None or proc.poll() is not None:
            return
        try:
            if sys.platform == "win32":
                subprocess.run(["taskkill", "/PID", str(proc.pid), "/T", "/F"],
                               capture_output=True, creationflags=_NO_WINDOW, timeout=5)
            else:
                proc.terminate()
                proc.wait(timeout=5)
        except Exception:
            pass

    def open(self, folder: Path) -> str | None:
        """폴더를 MicroDicom 으로 엶. 성공하면 None, 실패하면 이유(문자열)."""
        exe = find_exe()
        if exe is None:
            return (f"MicroDicom 실행 파일을 찾지 못했습니다: {config.MICRODICOM_PATH}\n"
                    "config.py 의 MICRODICOM_PATH 를 실제 위치로 고쳐 주세요.")
        if not folder.is_dir():
            return f"영상 폴더가 없습니다: {folder}"
        self.close_previous()
        args = [str(exe)] + [a.replace("{folder}", str(folder)) for a in config.MICRODICOM_ARGS]
        try:
            self.proc = subprocess.Popen(args, cwd=str(exe.parent))
        except OSError as e:
            return f"MicroDicom 을 실행하지 못했습니다: {e}"
        self.opened_folder = folder
        return None
