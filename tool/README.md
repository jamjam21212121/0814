# LumbarDISC 쉬모를 결절 라벨링 도구

RSNA 2024 Lumbar Spine 데이터의 Sagittal T2 영상을 보면서 쉬모를 결절(Schmorl node)을
엑셀 양식(`LumbarDISC_Schmorl_labeling_v0.1.xlsx`)에 기록하는 도구입니다.
판독은 MicroDicom으로 하고, 이 프로그램은 기록용입니다.

## 폴더 배치

모든 데이터와 결과물은 **`C:\Users\kay85\OneDrive\바탕 화면\연구멘토링`** 안에서만 읽고 씁니다.
기준 경로는 `config.py`의 `PROJECT_DIR`이고, `tool` 폴더를 어디서 실행하든 이 경로가 기준입니다.
관리하기 쉽도록 `tool` 폴더도 이 안에 두는 것을 권장합니다.

```
연구멘토링\
├─ data\                      ← RSNA 데이터 (train_images, CSV 2개)
├─ labeling\
│   ├─ LumbarDISC_Schmorl_labeling_v0.1.xlsx   ← 결과 엑셀
│   ├─ captures\              ← 애매(Uncertain) 병변 캡처 이미지
│   ├─ backup\                ← 엑셀 수정 전 자동 백업 (자동 생성)
│   ├─ setup_report.txt       ← ⓪ 점검 보고서
│   └─ sample_selection_log.txt ← ① 추출 기록 (시드, 모집단, 목록)
└─ tool\                      ← 이 폴더 (코드)
```

폴더 이름이나 위치가 바뀌면 `config.py`의 `PROJECT_DIR` 한 줄만 고치면 됩니다.
실행 파일(.bat)은 `__pycache__` 같은 부산물을 만들지 않도록 설정돼 있습니다.

## 처음 한 번: 준비

1. **Python 설치**: https://www.python.org/downloads/ 에서 3.10 이상을 설치합니다.
   설치 첫 화면에서 **"Add python.exe to PATH"를 꼭 체크**하세요.
2. **패키지 설치**: `tool\install.bat`을 더블클릭합니다. 끝에 `Successfully installed ...`가
   나오거나 이미 설치돼 있다는 메시지가 나오면 정상입니다.

## 단계별 실행

| 단계 | 실행 파일 | 하는 일 | 파일 수정 |
|---|---|---|---|
| ⓪ 점검 | `step0_check.bat` | 폴더·CSV·DICOM·엑셀 양식·MicroDicom 위치를 확인하고 보고서를 만듭니다 | 없음 (보고서만 저장) |
| ① 추출 | `step1_make_sample.bat` | 무작위 100명을 뽑아 `sample_list` 시트에 기록합니다 | 엑셀 (수정 전 자동 백업) |

엑셀을 수정하는 단계는 **실행 전에 엑셀 파일을 닫아 두세요.** 열려 있으면 저장이 실패하고,
원본은 그대로 남습니다.

### ⓪ 점검 결과 보는 법
- `[OK]`: 정상, `[없음]`: 경로나 파일을 확인해야 합니다.
- 보고서 파일은 `labeling\setup_report.txt`에 저장됩니다.
- **6. 엑셀 양식** 항목은 각 열의 머리글 색을 `회색(프로그램)`, `초록(수식)`, `노랑(사람)`으로
  구분해서 보여 줍니다. 프로그램은 **회색 열에만** 값을 씁니다.

### ① 추출 방식 (논문 방법 기술용)
- 모집단: `train_images`에 Sagittal T2/STIR 시리즈가 실제로 다운로드된 study
- study_id를 오름차순으로 정렬한 뒤 `random.Random(2024).sample(목록, 100)`으로 추출합니다.
  시드(2024)와 인원(100)은 `config.py`에서 바꿀 수 있습니다.
- 같은 데이터와 같은 시드로 실행하면 언제나 같은 100명이 뽑힙니다.
  **데이터를 더 내려받으면 모집단이 바뀌어** 결과도 달라지니 주의하세요.
- 시드, 모집단 수, 추출 목록은 `labeling\sample_selection_log.txt`에 남습니다.
- `sample_list`에 이미 표본이 기록돼 있으면 덮어쓰지 않습니다.
  다시 뽑으려면 `labeling\backup\`의 백업본으로 되돌리세요.

## 파일 설명

| 파일 | 역할 |
|---|---|
| `config.py` | 경로, 시드, 표본 수, status 표기 같은 설정 |
| `dataset.py` | CSV와 DICOM 폴더를 읽는 공통 함수 |
| `excel_io.py` | 엑셀 머리글 색 판별, 백업, 안전 저장(열려 있으면 경고) |
| `check_setup.py` | ⓪ 점검 |
| `make_sample.py` | ① 100명 추출 |
