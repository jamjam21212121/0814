# -*- coding: utf-8 -*-
r"""
엑셀 파일 진단 — "엑셀 읽기 실패"가 날 때 원인을 찾는 도구 (파일은 고치지 않고 읽기만 함)

- 결과 엑셀과 labeling\backup\ 의 백업 파일을 하나씩 열어 보고 어느 것이 정상인지 알려 줍니다.
- 결과 엑셀의 내부 구조(어느 프로그램으로 저장됐는지, 스타일 번호, 이름 정의 등)를 점검합니다.
- 보고서: labeling\excel_diagnosis.txt

실행: diagnose_excel.bat 더블클릭
"""
from __future__ import annotations

import re
import traceback
import zipfile
from pathlib import Path

import openpyxl

import config

lines: list[str] = []


def out(text: str = "") -> None:
    print(text)
    lines.append(text)


def try_load(path: Path) -> str | None:
    """열리면 None, 안 열리면 오류 설명."""
    try:
        openpyxl.load_workbook(path)
        return None
    except Exception as e:
        return f"{type(e).__name__}: {e}"


def inspect_zip(path: Path) -> None:
    out(f"■ 내부 구조: {path.name} ({path.stat().st_size:,} bytes)")
    try:
        z = zipfile.ZipFile(path)
    except zipfile.BadZipFile as e:
        out(f"  [문제] 엑셀(zip) 파일이 깨졌습니다: {e}")
        return
    names = z.namelist()
    out(f"  내부 파일 {len(names)}개")
    head = z.read("xl/workbook.xml")[:300].decode("utf-8", "replace")
    if re.search(r"<\w+:workbook", head):
        out("  [참고] XML에 접두사(x: 등)가 붙어 있습니다 → Excel이 아닌 다른 프로그램/웹에서 저장된 파일로 보입니다.")
    app = z.read("docProps/app.xml").decode("utf-8", "replace") if "docProps/app.xml" in names else ""
    m = re.search(r"<(?:\w+:)?Application>(.*?)</(?:\w+:)?Application>", app)
    out(f"  마지막으로 저장한 프로그램: {m.group(1) if m else '(기록 없음)'}")
    m = re.search(r"<(?:\w+:)?AppVersion>(.*?)</(?:\w+:)?AppVersion>", app)
    if m:
        out(f"  프로그램 버전: {m.group(1)}")

    styles = z.read("xl/styles.xml").decode("utf-8", "replace") if "xl/styles.xml" in names else ""
    xfs = re.search(r"<(?:\w+:)?cellXfs[^>]*>(.*?)</(?:\w+:)?cellXfs>", styles, re.S)
    n_xf = len(re.findall(r"<(?:\w+:)?xf[ >/]", xfs.group(1))) if xfs else 0
    m = re.search(r"<(?:\w+:)?cellXfs[^>]*count=\"(\d+)\"", styles)
    out(f"  셀 스타일(cellXfs): 선언 {m.group(1) if m else '-'}개, 실제 {n_xf}개")

    wb_xml = z.read("xl/workbook.xml").decode("utf-8", "replace")
    sheets = re.findall(r"<(?:\w+:)?sheet [^>]*name=\"([^\"]+)\"", wb_xml)
    out(f"  시트: {sheets}")
    for dn in re.findall(r"<(?:\w+:)?definedName[^>]*>.*?</(?:\w+:)?definedName>", wb_xml, re.S):
        out(f"  이름 정의: {dn[:160]}")
        lid = re.search(r"localSheetId=\"(\d+)\"", dn)
        if lid and int(lid.group(1)) >= len(sheets):
            out("    [문제] 없는 시트 번호를 가리킵니다")

    for name in sorted(n for n in names if n.startswith("xl/worksheets/sheet")):
        xml = z.read(name).decode("utf-8", "replace")
        used = [int(s) for s in re.findall(r"<(?:\w+:)?c [^>]*\bs=\"(\d+)\"", xml)]
        used += [int(s) for s in re.findall(r"<(?:\w+:)?(?:row|col) [^>]*\bs=\"(\d+)\"", xml)]
        bad = sorted({s for s in used if s >= n_xf})
        out(f"  {name}: 셀 {len(re.findall(r'<(?:\w+:)?c ', xml))}개, 스타일 번호 최대 {max(used) if used else '-'}"
            f" (스타일 {n_xf}개)" + (f"  [문제] 없는 스타일 번호 {bad[:10]}" if bad else ""))


def main() -> None:
    out("엑셀 파일 진단 (파일은 수정하지 않습니다)")
    out(f"코드 버전: {config.TOOL_VERSION}, openpyxl {openpyxl.__version__}")
    out()
    target = config.EXCEL_PATH
    if not target.exists():
        out(f"[없음] {target}")
    else:
        out(f"■ 결과 엑셀: {target}")
        err = try_load(target)
        if err is None:
            out("  [OK] 정상적으로 열립니다.")
        else:
            out(f"  [문제] 열리지 않습니다 → {err}")
            out("  ── 자세한 오류 ──")
            try:
                openpyxl.load_workbook(target)
            except Exception:
                for line in traceback.format_exc().splitlines():
                    out("  " + line)
        out()
        inspect_zip(target)

    out()
    out("■ 백업 파일 (최근 것부터)")
    backups = sorted(config.BACKUP_DIR.glob("*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True) \
        if config.BACKUP_DIR.is_dir() else []
    if not backups:
        out("  백업이 없습니다.")
    for p in backups[:15]:
        err = try_load(p)
        out(f"  {'[OK]  ' if err is None else '[문제]'} {p.name}" + ("" if err is None else f" → {err[:80]}"))

    report = config.LABELING_DIR / "excel_diagnosis.txt"
    try:
        report.write_text("\n".join(lines), encoding="utf-8")
        print(f"\n보고서 저장: {report}")
    except OSError as e:
        print(f"\n보고서 저장 실패: {e}")


if __name__ == "__main__":
    main()
