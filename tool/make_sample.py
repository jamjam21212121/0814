# -*- coding: utf-8 -*-
r"""
① 무작위 100명 추출 → 엑셀 sample_list 시트에 기록

대상(모집단): train_images 에 Sagittal T2/STIR 시리즈가 실제로 다운로드된 study
추출 방법   : study_id 를 숫자 순으로 정렬한 뒤 random.Random(시드).sample(...)
              → 같은 데이터·같은 시드면 언제 돌려도 같은 100명
기록 위치   : sample_list 시트의 '회색 머리글' 열 (study_id, Sagittal T2/T1 series_id, status 등)
추가 기록   : labeling\sample_selection_log.txt (시드, 모집단 수, 추출 목록 — 논문 방법 기술용)

실행: step1_make_sample.bat 더블클릭  (또는  py make_sample.py)
"""
from __future__ import annotations

import random
import sys
from datetime import datetime

import config
import dataset
import excel_io
import workbook


def fail(msg: str) -> None:
    print("\n[중단] " + msg)
    sys.exit(1)


# ── 1) 모집단과 표본 ──────────────────────────────────────
def build_sample():
    for p in (config.SERIES_CSV, config.COORDS_CSV, config.IMAGES_DIR, config.EXCEL_PATH):
        if not p.exists():
            hint = ""
            if p.suffix == ".csv":
                hint = ("\n        → CSV 파일을 data 폴더 바로 아래나 그 하위 폴더(예: data\\새 폴더\\)에 두세요."
                        "\n          step0_check.bat 보고서의 '2. 폴더와 파일'에서 실제 위치를 확인할 수 있습니다.")
            fail(f"파일/폴더가 없습니다: {p}{hint}")

    series_map = dataset.read_series_descriptions()
    labeled = dataset.studies_with_coordinates()

    eligible = []          # [(study_id, t2_series, t1_series, t1_downloaded)]
    skipped = []
    for study in dataset.downloaded_study_ids():
        t2, t2_ok = dataset.pick_series(study, config.DESC_SAG_T2, series_map, labeled)
        if not t2_ok:
            skipped.append(study)
            continue
        t1, t1_ok = dataset.pick_series(study, config.DESC_SAG_T1, series_map, labeled)
        eligible.append((study, t2, t1, t1_ok))

    print(f"다운로드된 study 폴더: {len(eligible) + len(skipped)}개")
    print(f"  └ Sagittal T2/STIR 가 있어 추출 대상: {len(eligible)}개 (제외 {len(skipped)}개)")
    if len(eligible) < config.SAMPLE_SIZE:
        fail(f"추출 대상이 {len(eligible)}명뿐이라 {config.SAMPLE_SIZE}명을 뽑을 수 없습니다.")

    eligible.sort(key=lambda e: dataset.id_sort_key(e[0]))
    rng = random.Random(config.RANDOM_SEED)
    sample = rng.sample(eligible, config.SAMPLE_SIZE)
    return sample, len(eligible), skipped


# ── 2) 엑셀 sample_list 열 찾기 ───────────────────────────
def as_number(s: str):
    return int(s) if s.isdigit() else s


def write_excel(sample, n_eligible: int) -> None:
    if excel_io.lock_file_exists():
        print("[주의] Excel에서 파일이 열려 있는 것 같습니다. 닫지 않으면 저장이 실패합니다.")

    wb, warns = excel_io.load()
    for w in warns:
        print(f"[openpyxl 경고] {w}")
    ws = workbook.sample_sheet(wb)
    if ws is None:
        fail(f"sample_list 시트를 찾지 못했습니다. 시트 목록: {wb.sheetnames}")

    hr = excel_io.find_header_row(ws)
    if hr is None:
        fail("sample_list 시트에서 'study_id' 머리글을 찾지 못했습니다.")
    theme = excel_io.theme_colors(wb)
    cols = excel_io.header_columns(ws, hr, theme)
    mapping = workbook.map_sample_columns(cols)

    print(f"\nsample_list 머리글(행 {hr}):")
    for c in cols:
        role = next((k for k, v in mapping.items() if v is c), "")
        print(f"  {c['letter']:>3} {c['text']:<28} {excel_io.KIND_KO[c['kind']]:<10} {('→ ' + role) if role else ''}")

    if "study_id" not in mapping:
        fail("study_id 열이 없습니다.")
    writable = {}
    for key, c in mapping.items():
        if c["kind"] == "gray":
            writable[key] = c
        else:
            print(f"[건너뜀] '{c['text']}' 열은 회색 머리글이 아니라서 쓰지 않습니다 ({excel_io.KIND_KO[c['kind']]}).")
    if "study_id" not in writable:
        fail("study_id 열 머리글이 회색이 아닙니다. 양식을 확인해 주세요.")
    for key, label in [("t2", "Sagittal T2 series_id"), ("t1", "Sagittal T1 series_id"), ("status", "status")]:
        if key not in mapping:
            print(f"[주의] {label} 열을 찾지 못해 기록하지 않습니다.")

    status_value = config.STATUS_TODO
    if "status" in writable:
        options = excel_io.list_validation_options(ws, writable["status"]["col"], hr + 1)
        if options and status_value not in options:
            print(f"[주의] status 드롭다운 값 {options} 에 '{status_value}' 가 없어 status 는 비워 둡니다."
                  " → config.py 의 STATUS_* 값을 드롭다운과 맞춰 주세요.")
            writable.pop("status")

    # 이미 기록된 표본이 있는지 확인
    sid_col = writable["study_id"]["col"]
    existing = [dataset._id(ws.cell(r, sid_col).value) for r in range(hr + 1, ws.max_row + 1)
                if ws.cell(r, sid_col).value not in (None, "")]
    new_ids = [s[0] for s in sample]
    if existing:
        if existing == new_ids:
            print(f"\n이미 같은 {len(existing)}명이 sample_list 에 기록되어 있습니다. 변경 없이 종료합니다.")
            return
        fail(f"sample_list 에 이미 다른 study {len(existing)}명이 있습니다. 덮어쓰지 않습니다.\n"
             "        (다시 뽑으려면 labeling\\backup 의 백업본으로 되돌리거나 study_id 칸을 비운 뒤 실행하세요.)")

    # 쓰기
    for i, (study, t2, t1, _) in enumerate(sample):
        r = hr + 1 + i
        values = {"study_id": as_number(study), "t2": as_number(t2) if t2 else None,
                  "t1": as_number(t1) if t1 else None, "status": status_value, "order": i + 1}
        for key, c in writable.items():
            cell = ws.cell(r, c["col"])
            if excel_io.is_formula(cell.value):
                print(f"[건너뜀] {cell.coordinate} 에 수식이 있어 쓰지 않습니다.")
                continue
            cell.value = values[key]
    excel_io.extend_tables(ws, hr + len(sample))

    bk = excel_io.backup(tag="before_sample")
    print(f"\n백업 저장: {bk}")
    try:
        excel_io.save(wb)
    except excel_io.ExcelLockedError as e:
        fail(str(e))
    print(f"엑셀 저장 완료: {config.EXCEL_PATH}")


def write_log(sample, n_eligible: int, skipped: list[str]) -> None:
    log = config.LABELING_DIR / "sample_selection_log.txt"
    rows = [
        f"작성: {datetime.now():%Y-%m-%d %H:%M:%S}",
        f"모집단: train_images 에 '{config.DESC_SAG_T2}' 시리즈가 다운로드된 study {n_eligible}명",
        f"제외(Sagittal T2 미다운로드): {len(skipped)}명 {skipped}",
        f"추출: Python random.Random({config.RANDOM_SEED}).sample(study_id 오름차순 목록, {config.SAMPLE_SIZE})",
        "",
        "순번\tstudy_id\tSagittal_T2_series_id\tSagittal_T1_series_id\tT1_다운로드",
    ]
    for i, (study, t2, t1, t1_ok) in enumerate(sample, 1):
        rows.append(f"{i}\t{study}\t{t2}\t{t1}\t{'O' if t1_ok else 'X'}")
    log.write_text("\n".join(rows), encoding="utf-8")
    print(f"추출 기록 저장: {log}")


def main() -> None:
    print(f"① 무작위 {config.SAMPLE_SIZE}명 추출 (시드 {config.RANDOM_SEED})   [코드 버전 {config.TOOL_VERSION}]\n")
    sample, n_eligible, skipped = build_sample()
    no_t1 = [s[0] for s in sample if not s[3]]
    print(f"\n추출 결과 {len(sample)}명 — 처음 10명:")
    for i, (study, t2, t1, t1_ok) in enumerate(sample[:10], 1):
        print(f"  {i:>3}. study {study}  T2 {t2}  T1 {t1 or '(없음)'}{'' if t1_ok else ' (T1 미다운로드)'}")
    if no_t1:
        print(f"[참고] T1 이 다운로드되지 않은 study {len(no_t1)}명: {no_t1}")
    write_excel(sample, n_eligible)
    write_log(sample, n_eligible, skipped)


if __name__ == "__main__":
    main()
